use tauri::State;

use crate::db;
use crate::models::{
    Account, CalendarItem, CreateNoteRequest, ExternalEventDetail, MonthView, SyncResult,
    UpdateNoteCompletionRequest,
};
use crate::outlook_auth;
use crate::outlook_com_sync;
use crate::outlook_sync;
use crate::state::AppServices;
use chrono::{Local, TimeZone, Utc};

#[tauri::command]
pub async fn connect_outlook(state: State<'_, AppServices>) -> Result<Account, String> {
    let token = outlook_auth::interactive_login(&state.http).await?;
    let refresh_token = token.refresh_token.as_deref().ok_or_else(|| {
        "Microsoft 未返回 refresh token，请确认已请求 offline_access 权限".to_string()
    })?;
    let user = outlook_auth::fetch_user(&state.http, &token.access_token).await?;
    let email = user
        .mail
        .clone()
        .or(user.user_principal_name.clone())
        .unwrap_or_else(|| "unknown@outlook".to_string());
    let display_name = user.display_name.clone().unwrap_or_else(|| email.clone());
    let calendar = outlook_auth::fetch_primary_calendar(&state.http, &token.access_token).await?;

    let account = {
        let connection = state.db.lock().map_err(|error| error.to_string())?;
        let account = db::upsert_account(&connection, &user.id, &email, &display_name)?;
        db::upsert_primary_calendar(&connection, &account.id, &calendar.id, &calendar.name)?;
        account
    };

    outlook_auth::save_refresh_token(&account.id, refresh_token)?;
    let _ = outlook_sync::sync_single_account(&state, &account.id).await;
    Ok(account)
}

#[tauri::command]
pub async fn connect_local_outlook(state: State<'_, AppServices>) -> Result<Account, String> {
    outlook_com_sync::connect_local_outlook(&state).await
}

#[tauri::command]
pub async fn refresh_local_outlook(state: State<'_, AppServices>) -> Result<SyncResult, String> {
    outlook_com_sync::sync_local_outlook_now(&state).await
}

#[tauri::command]
pub async fn disconnect_outlook(
    account_id: String,
    state: State<'_, AppServices>,
) -> Result<(), String> {
    let provider = {
        let connection = state.db.lock().map_err(|error| error.to_string())?;
        db::get_account(&connection, &account_id)?
            .map(|account| account.provider)
            .unwrap_or_default()
    };
    if provider == "outlook" {
        outlook_auth::delete_refresh_token(&account_id)?;
    }
    let connection = state.db.lock().map_err(|error| error.to_string())?;
    db::delete_account(&connection, &account_id)
}

#[tauri::command]
pub async fn sync_outlook_now(
    account_id: Option<String>,
    state: State<'_, AppServices>,
) -> Result<SyncResult, String> {
    if let Some(account_id) = account_id {
        outlook_sync::sync_single_account(&state, &account_id).await
    } else {
        outlook_sync::sync_all_accounts(&state).await
    }
}

#[tauri::command]
pub fn get_month_view(
    year: i32,
    month: u32,
    state: State<'_, AppServices>,
) -> Result<MonthView, String> {
    let connection = state.db.lock().map_err(|error| error.to_string())?;
    db::get_month_view(&connection, year, month)
}

#[tauri::command]
pub fn create_note(
    request: CreateNoteRequest,
    state: State<'_, AppServices>,
) -> Result<CalendarItem, String> {
    let connection = state.db.lock().map_err(|error| error.to_string())?;
    db::create_local_note(&connection, request)
}

#[tauri::command]
pub fn set_note_completed(
    request: UpdateNoteCompletionRequest,
    state: State<'_, AppServices>,
) -> Result<CalendarItem, String> {
    let connection = state.db.lock().map_err(|error| error.to_string())?;
    db::update_local_note_completion(&connection, request)
}

#[tauri::command]
pub fn get_external_event_detail(
    event_id: String,
    state: State<'_, AppServices>,
) -> Result<Option<ExternalEventDetail>, String> {
    let connection = state.db.lock().map_err(|error| error.to_string())?;
    db::get_external_event_detail(&connection, &event_id)
}

#[tauri::command]
pub fn get_local_note(
    note_id: String,
    state: State<'_, AppServices>,
) -> Result<CalendarItem, String> {
    let connection = state.db.lock().map_err(|error| error.to_string())?;
    db::get_local_note_item(&connection, &note_id)
}

#[tauri::command]
pub fn get_upcoming_local_notes(
    state: State<'_, AppServices>,
) -> Result<Vec<CalendarItem>, String> {
    let today = Local::now().date_naive();
    let start_naive = today
        .and_hms_opt(0, 0, 0)
        .ok_or_else(|| "当天开始时间无效".to_string())?;
    let start = Local
        .from_local_datetime(&start_naive)
        .single()
        .or_else(|| Local.from_local_datetime(&start_naive).earliest())
        .ok_or_else(|| "当天开始时间无效".to_string())?
        .with_timezone(&Utc);
    let connection = state.db.lock().map_err(|error| error.to_string())?;
    db::list_upcoming_local_note_items(&connection, &start.to_rfc3339())
}
