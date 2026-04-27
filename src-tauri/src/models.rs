use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Account {
    pub id: String,
    pub provider: String,
    pub provider_user_id: String,
    pub email: String,
    pub display_name: String,
    pub status: String,
    pub last_synced_at: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CalendarItem {
    pub id: String,
    pub source: String,
    pub source_label: String,
    pub read_only: bool,
    pub title: String,
    pub start_utc: String,
    pub end_utc: String,
    pub start_timezone: Option<String>,
    pub end_timezone: Option<String>,
    pub is_all_day: bool,
    pub location: Option<String>,
    pub sensitivity: Option<String>,
    pub category: Option<String>,
    pub note_color: Option<String>,
    pub note_body: Option<String>,
    pub reminder_at_utc: Option<String>,
    pub completed_at: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CreateNoteRequest {
    pub date_key: String,
    pub title: String,
    pub body: Option<String>,
    pub color: String,
    pub reminder_at_utc: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UpdateNoteCompletionRequest {
    pub note_id: String,
    pub completed: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ReminderNotice {
    pub id: String,
    pub title: String,
    pub body: Option<String>,
    pub color: String,
    pub reminder_at_utc: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MonthView {
    pub year: i32,
    pub month: u32,
    pub items: Vec<CalendarItem>,
    pub accounts: Vec<Account>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalEventDetail {
    pub id: String,
    pub source: String,
    pub provider_event_id: String,
    pub title: String,
    pub body_content_type: Option<String>,
    pub body_content: Option<String>,
    pub start_utc: String,
    pub end_utc: String,
    pub start_timezone: Option<String>,
    pub end_timezone: Option<String>,
    pub is_all_day: bool,
    pub location: Option<String>,
    pub attendees_json: Option<String>,
    pub organizer_json: Option<String>,
    pub web_link: Option<String>,
    pub online_meeting_url: Option<String>,
    pub categories_json: Option<String>,
    pub reminder_minutes_before_start: Option<i64>,
    pub is_reminder_on: bool,
    pub sensitivity: Option<String>,
    pub last_modified_utc: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SyncResult {
    pub synced_accounts: usize,
    pub synced_events: usize,
    pub message: String,
}
