use chrono::{Datelike, Local, NaiveDate, Utc};
use reqwest::StatusCode;
use serde::Deserialize;
use serde_json::Value;
use tokio::time::{sleep, Duration};
use url::Url;

use crate::db;
use crate::models::{Account, SyncResult};
use crate::outlook_auth;
use crate::outlook_com_sync;
use crate::state::AppServices;

const GRAPH_BASE: &str = "https://graph.microsoft.com/v1.0";

#[derive(Debug, Deserialize)]
struct DeltaResponse {
    #[serde(default)]
    value: Vec<GraphEvent>,
    #[serde(rename = "@odata.nextLink")]
    next_link: Option<String>,
    #[serde(rename = "@odata.deltaLink")]
    delta_link: Option<String>,
}

#[derive(Debug, Deserialize)]
struct RemovedMarker {}

#[derive(Debug, Deserialize)]
struct GraphDateTime {
    #[serde(rename = "dateTime")]
    date_time: Option<String>,
    #[serde(rename = "timeZone")]
    time_zone: Option<String>,
}

#[derive(Debug, Deserialize)]
struct GraphBody {
    #[serde(rename = "contentType")]
    content_type: Option<String>,
    content: Option<String>,
}

#[derive(Debug, Deserialize)]
struct GraphLocation {
    #[serde(rename = "displayName")]
    display_name: Option<String>,
}

#[derive(Debug, Deserialize)]
struct GraphEvent {
    id: String,
    #[serde(rename = "@removed")]
    removed: Option<RemovedMarker>,
    subject: Option<String>,
    body: Option<GraphBody>,
    start: Option<GraphDateTime>,
    end: Option<GraphDateTime>,
    location: Option<GraphLocation>,
    attendees: Option<Vec<Value>>,
    organizer: Option<Value>,
    #[serde(rename = "isAllDay")]
    is_all_day: Option<bool>,
    #[serde(rename = "webLink")]
    web_link: Option<String>,
    #[serde(rename = "onlineMeetingUrl")]
    online_meeting_url: Option<String>,
    categories: Option<Vec<String>>,
    #[serde(rename = "reminderMinutesBeforeStart")]
    reminder_minutes_before_start: Option<i64>,
    #[serde(rename = "isReminderOn")]
    is_reminder_on: Option<bool>,
    sensitivity: Option<String>,
    #[serde(rename = "lastModifiedDateTime")]
    last_modified_date_time: Option<String>,
}

pub async fn sync_all_accounts(services: &AppServices) -> Result<SyncResult, String> {
    let _guard = services.sync_lock.lock().await;
    let accounts = {
        let connection = services.db.lock().map_err(|error| error.to_string())?;
        db::list_accounts(&connection)?
    };

    let mut synced_accounts = 0;
    let mut synced_events = 0;
    let mut errors = Vec::new();

    for account in accounts {
        let sync_result = if outlook_com_sync::is_outlook_com_account(&account) {
            outlook_com_sync::sync_account(services, &account).await
        } else {
            sync_account(services, &account).await
        };
        match sync_result {
            Ok(count) => {
                synced_accounts += 1;
                synced_events += count;
            }
            Err(error) => errors.push(format!("{}: {}", account.email, error)),
        }
    }

    if !errors.is_empty() && synced_accounts == 0 {
        return Err(errors.join("；"));
    }

    let message = if errors.is_empty() {
        "Outlook 同步完成".to_string()
    } else {
        format!("Outlook 部分同步完成：{}", errors.join("；"))
    };

    Ok(SyncResult {
        synced_accounts,
        synced_events,
        message,
    })
}

pub async fn sync_single_account(
    services: &AppServices,
    account_id: &str,
) -> Result<SyncResult, String> {
    let _guard = services.sync_lock.lock().await;
    let account = {
        let connection = services.db.lock().map_err(|error| error.to_string())?;
        db::get_account(&connection, account_id)?.ok_or_else(|| "账号不存在".to_string())?
    };
    let synced_events = if outlook_com_sync::is_outlook_com_account(&account) {
        outlook_com_sync::sync_account(services, &account).await?
    } else {
        sync_account(services, &account).await?
    };
    Ok(SyncResult {
        synced_accounts: 1,
        synced_events,
        message: "Outlook 同步完成".to_string(),
    })
}

async fn sync_account(services: &AppServices, account: &Account) -> Result<usize, String> {
    let token = outlook_auth::refresh_access_token(&services.http, &account.id).await?;
    let calendar_id = {
        let connection = services.db.lock().map_err(|error| error.to_string())?;
        db::primary_calendar_id(&connection, &account.id)?
            .ok_or_else(|| "未找到 Outlook 主日历".to_string())?
    };
    let (window_start, window_end) = current_sync_window()?;
    let (sync_state_id, delta_link) = {
        let connection = services.db.lock().map_err(|error| error.to_string())?;
        let (id, delta_link, _) = db::sync_state_for_window(
            &connection,
            &account.id,
            &calendar_id,
            &window_start,
            &window_end,
        )?;
        (id, delta_link)
    };

    let sync_result = sync_delta_loop(
        services,
        account,
        &calendar_id,
        &token.access_token,
        delta_link.as_deref(),
        &window_start,
        &window_end,
    )
    .await;

    match sync_result {
        Ok((delta_link, count)) => {
            let synced_at = Utc::now().to_rfc3339();
            let connection = services.db.lock().map_err(|error| error.to_string())?;
            db::mark_sync_success(&connection, &sync_state_id, &delta_link, &synced_at)?;
            Ok(count)
        }
        Err(error) => {
            let connection = services
                .db
                .lock()
                .map_err(|lock_error| lock_error.to_string())?;
            let _ = db::mark_sync_error(&connection, &sync_state_id, &error);
            Err(error)
        }
    }
}

async fn sync_delta_loop(
    services: &AppServices,
    account: &Account,
    calendar_id: &str,
    access_token: &str,
    existing_delta_link: Option<&str>,
    window_start: &str,
    window_end: &str,
) -> Result<(String, usize), String> {
    let mut next_url = existing_delta_link
        .map(ToOwned::to_owned)
        .unwrap_or_else(|| build_delta_url(window_start, window_end));
    let mut synced_count = 0;

    loop {
        let response = fetch_delta_page(&services.http, access_token, &next_url).await?;
        for event in response.value {
            persist_graph_event(services, account, calendar_id, event).await?;
            synced_count += 1;
        }
        if let Some(next_link) = response.next_link {
            next_url = next_link;
            continue;
        }
        return response
            .delta_link
            .map(|link| (link, synced_count))
            .ok_or_else(|| "Graph delta 响应缺少 deltaLink".to_string());
    }
}

async fn fetch_delta_page(
    http: &reqwest::Client,
    access_token: &str,
    url: &str,
) -> Result<DeltaResponse, String> {
    let mut last_error = None;
    for attempt in 0..3 {
        let response = http
            .get(url)
            .bearer_auth(access_token)
            .header("Prefer", "outlook.timezone=\"UTC\"")
            .send()
            .await
            .map_err(|error| error.to_string())?;
        let status = response.status();
        if status.is_success() {
            return response
                .json::<DeltaResponse>()
                .await
                .map_err(|error| error.to_string());
        }
        let body = response.text().await.unwrap_or_default();
        if should_retry(status) && attempt < 2 {
            last_error = Some(format!("Graph 暂时不可用：{status} {body}"));
            sleep(Duration::from_secs(2_u64.pow(attempt + 1))).await;
            continue;
        }
        if status == StatusCode::GONE {
            return Err("Graph delta token 已失效，请断开并重新连接 Outlook 账号".to_string());
        }
        return Err(format!("Graph 同步失败：{status} {body}"));
    }
    Err(last_error.unwrap_or_else(|| "Graph 同步失败".to_string()))
}

fn should_retry(status: StatusCode) -> bool {
    status == StatusCode::TOO_MANY_REQUESTS || status.is_server_error()
}

async fn persist_graph_event(
    services: &AppServices,
    account: &Account,
    calendar_id: &str,
    event: GraphEvent,
) -> Result<(), String> {
    let connection = services.db.lock().map_err(|error| error.to_string())?;
    if event.removed.is_some() {
        return db::mark_external_event_deleted(&connection, &account.id, &event.id);
    }

    let start = event
        .start
        .as_ref()
        .ok_or_else(|| "Outlook 日程缺少开始时间".to_string())?;
    let end = event
        .end
        .as_ref()
        .ok_or_else(|| "Outlook 日程缺少结束时间".to_string())?;
    let start_utc = normalize_graph_datetime(start.date_time.as_deref())?;
    let end_utc = normalize_graph_datetime(end.date_time.as_deref())?;
    let attendees_json = event
        .attendees
        .as_ref()
        .map(serde_json::to_string)
        .transpose()
        .map_err(|error| error.to_string())?;
    let organizer_json = event
        .organizer
        .as_ref()
        .map(serde_json::to_string)
        .transpose()
        .map_err(|error| error.to_string())?;
    let categories_json = event
        .categories
        .as_ref()
        .map(serde_json::to_string)
        .transpose()
        .map_err(|error| error.to_string())?;

    db::upsert_external_event(
        &connection,
        &account.id,
        calendar_id,
        &event.id,
        event.subject.as_deref().unwrap_or("无标题日程"),
        event
            .body
            .as_ref()
            .and_then(|body| body.content_type.as_deref()),
        event.body.as_ref().and_then(|body| body.content.as_deref()),
        &start_utc,
        &end_utc,
        start.time_zone.as_deref(),
        end.time_zone.as_deref(),
        event.is_all_day.unwrap_or(false),
        event
            .location
            .as_ref()
            .and_then(|location| location.display_name.as_deref()),
        attendees_json.as_deref(),
        organizer_json.as_deref(),
        event.web_link.as_deref(),
        event.online_meeting_url.as_deref(),
        categories_json.as_deref(),
        event.reminder_minutes_before_start,
        event.is_reminder_on.unwrap_or(false),
        event.sensitivity.as_deref(),
        event.last_modified_date_time.as_deref(),
    )
}

fn build_delta_url(window_start: &str, window_end: &str) -> String {
    let mut url =
        Url::parse(&format!("{GRAPH_BASE}/me/calendarView/delta")).expect("valid graph url");
    url.query_pairs_mut()
        .append_pair("startDateTime", window_start)
        .append_pair("endDateTime", window_end)
        .append_pair("$select", "id,subject,body,start,end,location,attendees,organizer,isAllDay,webLink,onlineMeetingUrl,categories,reminderMinutesBeforeStart,isReminderOn,sensitivity,lastModifiedDateTime");
    url.to_string()
}

pub fn current_sync_window() -> Result<(String, String), String> {
    let today = Local::now().date_naive();
    let current_month = first_of_month(today.year(), today.month())?;
    let previous_month = add_months(current_month, -1)?;
    let after_next_month = add_months(current_month, 2)?;
    let start = previous_month
        .and_hms_opt(0, 0, 0)
        .ok_or_else(|| "无法计算同步开始时间".to_string())?
        .and_utc();
    let end = after_next_month
        .and_hms_opt(0, 0, 0)
        .ok_or_else(|| "无法计算同步结束时间".to_string())?
        .and_utc();
    Ok((start.to_rfc3339(), end.to_rfc3339()))
}

fn first_of_month(year: i32, month: u32) -> Result<NaiveDate, String> {
    NaiveDate::from_ymd_opt(year, month, 1).ok_or_else(|| "年月参数无效".to_string())
}

fn add_months(date: NaiveDate, offset: i32) -> Result<NaiveDate, String> {
    let month_index = date.year() * 12 + date.month0() as i32 + offset;
    let year = month_index.div_euclid(12);
    let month = month_index.rem_euclid(12) as u32 + 1;
    first_of_month(year, month)
}

fn normalize_graph_datetime(value: Option<&str>) -> Result<String, String> {
    let value = value.ok_or_else(|| "Outlook 日程时间为空".to_string())?;
    if let Ok(parsed) = chrono::DateTime::parse_from_rfc3339(value) {
        return Ok(parsed.with_timezone(&Utc).to_rfc3339());
    }
    let trimmed = value.trim_end_matches('Z');
    let normalized = format!("{trimmed}Z");
    chrono::DateTime::parse_from_rfc3339(&normalized)
        .map(|parsed| parsed.with_timezone(&Utc).to_rfc3339())
        .map_err(|error| format!("Outlook 日程时间格式无效：{value} ({error})"))
}
