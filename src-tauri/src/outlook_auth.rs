use std::io::{Read, Write};
use std::net::TcpListener;
use std::time::Duration;

use base64::engine::general_purpose::URL_SAFE_NO_PAD;
use base64::Engine;
use keyring::Entry;
use rand::distributions::Alphanumeric;
use rand::{thread_rng, Rng};
use reqwest::Client;
use serde::Deserialize;
use sha2::{Digest, Sha256};
use url::Url;

const AUTHORITY: &str = "https://login.microsoftonline.com/common/oauth2/v2.0";
const GRAPH_BASE: &str = "https://graph.microsoft.com/v1.0";
const KEYRING_SERVICE: &str = "calendar-notes-outlook";
pub const OUTLOOK_SCOPES: &str = "openid profile email offline_access User.Read Calendars.Read";

#[derive(Debug, Clone, Deserialize)]
pub struct TokenResponse {
    pub access_token: String,
    pub refresh_token: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct GraphUser {
    pub id: String,
    #[serde(rename = "displayName")]
    pub display_name: Option<String>,
    pub mail: Option<String>,
    #[serde(rename = "userPrincipalName")]
    pub user_principal_name: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct GraphCalendar {
    pub id: String,
    pub name: String,
}

pub fn microsoft_client_id() -> Result<String, String> {
    std::env::var("MICROSOFT_CLIENT_ID")
        .ok()
        .filter(|value| !value.trim().is_empty())
        .ok_or_else(|| "未配置 MICROSOFT_CLIENT_ID。请在 Microsoft Entra 应用注册中创建桌面/移动客户端，并设置本机环境变量。".to_string())
}

pub fn save_refresh_token(account_id: &str, refresh_token: &str) -> Result<(), String> {
    let entry = Entry::new(KEYRING_SERVICE, account_id).map_err(|error| error.to_string())?;
    entry
        .set_password(refresh_token)
        .map_err(|error| error.to_string())
}

pub fn load_refresh_token(account_id: &str) -> Result<String, String> {
    let entry = Entry::new(KEYRING_SERVICE, account_id).map_err(|error| error.to_string())?;
    entry
        .get_password()
        .map_err(|_| "未找到 Outlook 刷新令牌，请重新连接账号".to_string())
}

pub fn delete_refresh_token(account_id: &str) -> Result<(), String> {
    let entry = Entry::new(KEYRING_SERVICE, account_id).map_err(|error| error.to_string())?;
    match entry.delete_credential() {
        Ok(()) => Ok(()),
        Err(_) => Ok(()),
    }
}

pub async fn interactive_login(http: &Client) -> Result<TokenResponse, String> {
    let client_id = microsoft_client_id()?;
    let state = random_string(32);
    let verifier = random_string(96);
    let challenge = pkce_challenge(&verifier);
    let listener = TcpListener::bind("127.0.0.1:0").map_err(|error| error.to_string())?;
    listener
        .set_nonblocking(false)
        .map_err(|error| error.to_string())?;
    let port = listener
        .local_addr()
        .map_err(|error| error.to_string())?
        .port();
    let redirect_uri = format!("http://127.0.0.1:{port}/auth/callback");

    let mut auth_url =
        Url::parse(&format!("{AUTHORITY}/authorize")).map_err(|error| error.to_string())?;
    auth_url
        .query_pairs_mut()
        .append_pair("client_id", &client_id)
        .append_pair("response_type", "code")
        .append_pair("redirect_uri", &redirect_uri)
        .append_pair("response_mode", "query")
        .append_pair("scope", OUTLOOK_SCOPES)
        .append_pair("state", &state)
        .append_pair("code_challenge", &challenge)
        .append_pair("code_challenge_method", "S256")
        .append_pair("prompt", "select_account");

    open::that(auth_url.as_str()).map_err(|error| format!("无法打开系统浏览器：{error}"))?;

    let expected_state = state.clone();
    let auth_code =
        tauri::async_runtime::spawn_blocking(move || wait_for_auth_code(listener, &expected_state))
            .await
            .map_err(|error| error.to_string())??;

    exchange_code(http, &client_id, &redirect_uri, &auth_code, &verifier).await
}

pub async fn exchange_code(
    http: &Client,
    client_id: &str,
    redirect_uri: &str,
    code: &str,
    verifier: &str,
) -> Result<TokenResponse, String> {
    let response = http
        .post(format!("{AUTHORITY}/token"))
        .form(&[
            ("client_id", client_id),
            ("scope", OUTLOOK_SCOPES),
            ("code", code),
            ("redirect_uri", redirect_uri),
            ("grant_type", "authorization_code"),
            ("code_verifier", verifier),
        ])
        .send()
        .await
        .map_err(|error| error.to_string())?;
    parse_token_response(response).await
}

pub async fn refresh_access_token(
    http: &Client,
    account_id: &str,
) -> Result<TokenResponse, String> {
    let client_id = microsoft_client_id()?;
    let refresh_token = load_refresh_token(account_id)?;
    let response = http
        .post(format!("{AUTHORITY}/token"))
        .form(&[
            ("client_id", client_id.as_str()),
            ("scope", OUTLOOK_SCOPES),
            ("refresh_token", refresh_token.as_str()),
            ("grant_type", "refresh_token"),
        ])
        .send()
        .await
        .map_err(|error| error.to_string())?;
    let token = parse_token_response(response).await?;
    if let Some(new_refresh_token) = token.refresh_token.as_deref() {
        save_refresh_token(account_id, new_refresh_token)?;
    }
    Ok(token)
}

pub async fn fetch_user(http: &Client, access_token: &str) -> Result<GraphUser, String> {
    graph_get(http, access_token, &format!("{GRAPH_BASE}/me")).await
}

pub async fn fetch_primary_calendar(
    http: &Client,
    access_token: &str,
) -> Result<GraphCalendar, String> {
    graph_get(http, access_token, &format!("{GRAPH_BASE}/me/calendar")).await
}

pub async fn graph_get<T: for<'de> Deserialize<'de>>(
    http: &Client,
    access_token: &str,
    url: &str,
) -> Result<T, String> {
    let response = http
        .get(url)
        .bearer_auth(access_token)
        .header("Prefer", "outlook.timezone=\"UTC\"")
        .send()
        .await
        .map_err(|error| error.to_string())?;
    if !response.status().is_success() {
        let status = response.status();
        let body = response.text().await.unwrap_or_default();
        return Err(format!("Graph 请求失败：{status} {body}"));
    }
    response
        .json::<T>()
        .await
        .map_err(|error| error.to_string())
}

async fn parse_token_response(response: reqwest::Response) -> Result<TokenResponse, String> {
    if !response.status().is_success() {
        let status = response.status();
        let body = response.text().await.unwrap_or_default();
        return Err(format!("Microsoft 登录失败：{status} {body}"));
    }
    response
        .json::<TokenResponse>()
        .await
        .map_err(|error| error.to_string())
}

fn wait_for_auth_code(listener: TcpListener, expected_state: &str) -> Result<String, String> {
    let (mut stream, _) = listener.accept().map_err(|error| error.to_string())?;
    stream
        .set_read_timeout(Some(Duration::from_secs(300)))
        .map_err(|error| error.to_string())?;
    let mut buffer = [0_u8; 4096];
    let bytes_read = stream
        .read(&mut buffer)
        .map_err(|error| error.to_string())?;
    let request = String::from_utf8_lossy(&buffer[..bytes_read]);
    let first_line = request
        .lines()
        .next()
        .ok_or_else(|| "无法读取 Microsoft 登录回调".to_string())?;
    let path = first_line
        .split_whitespace()
        .nth(1)
        .ok_or_else(|| "Microsoft 登录回调格式无效".to_string())?;
    let url = Url::parse(&format!("http://localhost{path}"))
        .map_err(|error| format!("Microsoft 登录回调 URL 无效：{error}"))?;
    let code = url
        .query_pairs()
        .find_map(|(key, value)| (key == "code").then(|| value.to_string()))
        .ok_or_else(|| "Microsoft 登录未返回授权码".to_string())?;
    let returned_state = url
        .query_pairs()
        .find_map(|(key, value)| (key == "state").then(|| value.to_string()))
        .ok_or_else(|| "Microsoft 登录未返回 state".to_string())?;
    let body = if returned_state == expected_state {
        "<html><body><h1>Calendar Notes 已连接 Outlook</h1><p>可以关闭此页面并返回应用。</p></body></html>"
    } else {
        "<html><body><h1>Calendar Notes 登录校验失败</h1><p>请关闭此页面并重试。</p></body></html>"
    };
    let response = format!(
        "HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{}",
        body.as_bytes().len(),
        body
    );
    let _ = stream.write_all(response.as_bytes());
    if returned_state != expected_state {
        return Err("Microsoft 登录 state 校验失败".to_string());
    }
    Ok(code)
}

fn random_string(length: usize) -> String {
    thread_rng()
        .sample_iter(&Alphanumeric)
        .take(length)
        .map(char::from)
        .collect()
}

fn pkce_challenge(verifier: &str) -> String {
    let digest = Sha256::digest(verifier.as_bytes());
    URL_SAFE_NO_PAD.encode(digest)
}
