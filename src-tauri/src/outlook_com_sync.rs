use std::fs;
use std::process::Command;

use chrono::Utc;
use serde::Deserialize;
use serde_json::Value;
use uuid::Uuid;

use crate::db;
use crate::models::{Account, SyncResult};
use crate::outlook_sync;
use crate::state::AppServices;

const OUTLOOK_COM_SCRIPT: &str = include_str!("../scripts/outlook-com-sync.ps1");
const OUTLOOK_COM_PROVIDER: &str = "outlook-com";

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ComSnapshot {
    account: ComAccount,
    events: Vec<ComEvent>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ComAccount {
    provider_user_id: String,
    email: String,
    display_name: String,
    calendar_id: String,
    calendar_name: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ComEvent {
    provider_event_id: String,
    title: String,
    body_content_type: Option<String>,
    body_content: Option<String>,
    start_utc: String,
    end_utc: String,
    start_timezone: Option<String>,
    end_timezone: Option<String>,
    is_all_day: bool,
    location: Option<String>,
    attendees: Option<Value>,
    organizer: Option<Value>,
    web_link: Option<String>,
    online_meeting_url: Option<String>,
    categories: Option<Vec<String>>,
    reminder_minutes_before_start: Option<i64>,
    is_reminder_on: bool,
    sensitivity: Option<String>,
    last_modified_utc: Option<String>,
}

pub fn is_outlook_com_account(account: &Account) -> bool {
    account.provider == OUTLOOK_COM_PROVIDER
}

pub async fn connect_local_outlook(services: &AppServices) -> Result<Account, String> {
    let _guard = services.sync_lock.lock().await;
    connect_local_outlook_without_lock(services).await
}

pub(crate) async fn sync_account(
    services: &AppServices,
    account: &Account,
) -> Result<usize, String> {
    let (window_start, window_end) = outlook_sync::current_sync_window()?;
    let snapshot = read_snapshot(window_start.clone(), window_end.clone()).await?;
    let (_, _, count) = persist_snapshot(
        services,
        snapshot,
        &window_start,
        &window_end,
        Some(account),
    )?;
    Ok(count)
}

pub async fn sync_local_outlook_now(services: &AppServices) -> Result<SyncResult, String> {
    let _guard = services.sync_lock.lock().await;
    let account = {
        let connection = services.db.lock().map_err(|error| error.to_string())?;
        db::list_accounts(&connection)?
            .into_iter()
            .find(is_outlook_com_account)
    };

    let (synced_events, message) = if let Some(account) = account {
        (
            sync_account(services, &account).await?,
            "经典 Outlook 日程已刷新".to_string(),
        )
    } else {
        let (window_start, window_end) = outlook_sync::current_sync_window()?;
        let snapshot = read_snapshot(window_start.clone(), window_end.clone()).await?;
        let (_, _, synced_events) =
            persist_snapshot(services, snapshot, &window_start, &window_end, None)?;
        (synced_events, "经典 Outlook 日程已连接并刷新".to_string())
    };

    Ok(SyncResult {
        synced_accounts: 1,
        synced_events,
        message,
    })
}

async fn connect_local_outlook_without_lock(services: &AppServices) -> Result<Account, String> {
    let (window_start, window_end) = outlook_sync::current_sync_window()?;
    let snapshot = read_snapshot(window_start.clone(), window_end.clone()).await?;
    let (account, _, _) = persist_snapshot(services, snapshot, &window_start, &window_end, None)?;
    Ok(account)
}

async fn read_snapshot(window_start: String, window_end: String) -> Result<ComSnapshot, String> {
    tauri::async_runtime::spawn_blocking(move || run_script(&window_start, &window_end))
        .await
        .map_err(|error| error.to_string())?
}

fn run_script(window_start: &str, window_end: &str) -> Result<ComSnapshot, String> {
    if !cfg!(target_os = "windows") {
        return Err("经典 Outlook COM 同步只支持 Windows".to_string());
    }

    let script_path = std::env::temp_dir().join(format!(
        "calendar-notes-outlook-com-sync-{}.ps1",
        Uuid::new_v4()
    ));
    let mut script_bytes = vec![0xEF, 0xBB, 0xBF];
    script_bytes.extend_from_slice(OUTLOOK_COM_SCRIPT.as_bytes());
    fs::write(&script_path, script_bytes).map_err(|error| error.to_string())?;

    let include_private_details =
        std::env::var("CALENDAR_NOTES_OUTLOOK_COM_INCLUDE_PRIVATE_DETAILS")
            .map(|value| matches!(value.as_str(), "1" | "true" | "TRUE" | "yes" | "YES"))
            .unwrap_or(false);

    let mut command = Command::new("powershell.exe");
    command
        .arg("-NoProfile")
        .arg("-NonInteractive")
        .arg("-ExecutionPolicy")
        .arg("Bypass")
        .arg("-File")
        .arg(&script_path)
        .arg("-WindowStartUtc")
        .arg(window_start)
        .arg("-WindowEndUtc")
        .arg(window_end);
    if include_private_details {
        command.arg("-IncludePrivateDetails");
    }

    let output = command.output().map_err(|error| error.to_string());
    let _ = fs::remove_file(&script_path);
    let output = output?;

    let stdout = String::from_utf8(output.stdout).map_err(|error| error.to_string())?;
    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
    if !output.status.success() {
        let detail = if stderr.is_empty() {
            stdout.trim()
        } else {
            &stderr
        };
        return Err(format!("经典 Outlook COM 读取失败：{detail}"));
    }
    serde_json::from_str::<ComSnapshot>(stdout.trim()).map_err(|error| {
        format!(
            "经典 Outlook COM 输出解析失败：{error}；原始输出：{}",
            stdout.trim()
        )
    })
}

fn persist_snapshot(
    services: &AppServices,
    snapshot: ComSnapshot,
    window_start: &str,
    window_end: &str,
    expected_account: Option<&Account>,
) -> Result<(Account, String, usize), String> {
    if let Some(account) = expected_account {
        if account.provider_user_id != snapshot.account.provider_user_id {
            return Err(
                "当前经典 Outlook 配置文件与已连接的本机账号不一致，请断开后重新连接".to_string(),
            );
        }
    }

    let connection = services.db.lock().map_err(|error| error.to_string())?;
    let account = db::upsert_account_for_provider(
        &connection,
        OUTLOOK_COM_PROVIDER,
        &snapshot.account.provider_user_id,
        &snapshot.account.email,
        &snapshot.account.display_name,
    )?;
    let calendar_id = db::upsert_primary_calendar(
        &connection,
        &account.id,
        &snapshot.account.calendar_id,
        &snapshot.account.calendar_name,
    )?;
    let (sync_state_id, _, _) = db::sync_state_for_window(
        &connection,
        &account.id,
        &calendar_id,
        window_start,
        window_end,
    )?;

    let event_count = snapshot.events.len();
    let mut seen_provider_event_ids = Vec::with_capacity(event_count);
    let persist_result = (|| -> Result<(), String> {
        for event in snapshot.events {
            let provider_event_id = event.provider_event_id.trim().to_string();
            if provider_event_id.is_empty() {
                continue;
            }
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
                &calendar_id,
                &provider_event_id,
                event
                    .title
                    .trim()
                    .is_empty()
                    .then_some("无标题日程")
                    .unwrap_or(&event.title),
                event.body_content_type.as_deref(),
                event.body_content.as_deref(),
                &event.start_utc,
                &event.end_utc,
                event.start_timezone.as_deref(),
                event.end_timezone.as_deref(),
                event.is_all_day,
                event
                    .location
                    .as_deref()
                    .filter(|value| !value.trim().is_empty()),
                attendees_json.as_deref(),
                organizer_json.as_deref(),
                event.web_link.as_deref(),
                event.online_meeting_url.as_deref(),
                categories_json.as_deref(),
                event.reminder_minutes_before_start,
                event.is_reminder_on,
                event.sensitivity.as_deref(),
                event.last_modified_utc.as_deref(),
            )?;
            seen_provider_event_ids.push(provider_event_id);
        }
        db::mark_missing_external_events_deleted(
            &connection,
            &account.id,
            &calendar_id,
            window_start,
            window_end,
            &seen_provider_event_ids,
        )
    })();

    match persist_result {
        Ok(()) => {
            db::mark_sync_success(
                &connection,
                &sync_state_id,
                "outlook-com-snapshot",
                &Utc::now().to_rfc3339(),
            )?;
            let account = db::get_account(&connection, &account.id)?
                .ok_or_else(|| "账号同步状态保存失败".to_string())?;
            Ok((account, calendar_id, event_count))
        }
        Err(error) => {
            let _ = db::mark_sync_error(&connection, &sync_state_id, &error);
            Err(error)
        }
    }
}
