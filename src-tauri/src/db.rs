use std::fs;
use std::path::Path;

use chrono::{Local, TimeZone, Utc};
use rusqlite::{params, params_from_iter, Connection, OptionalExtension};
use uuid::Uuid;

use crate::models::{
    Account, CalendarItem, CreateNoteRequest, ExternalEventDetail, MonthView, ReminderNotice,
    UpdateNoteCompletionRequest,
};

pub fn open_database(app_data_dir: &Path) -> Result<Connection, String> {
    fs::create_dir_all(app_data_dir).map_err(|error| error.to_string())?;
    let db_path = app_data_dir.join("calendar-notes.sqlite3");
    let connection = Connection::open(db_path).map_err(|error| error.to_string())?;
    connection
        .pragma_update(None, "foreign_keys", "ON")
        .map_err(|error| error.to_string())?;
    init_schema(&connection)?;
    Ok(connection)
}

fn init_schema(connection: &Connection) -> Result<(), String> {
    connection
        .execute_batch(
            r#"
            CREATE TABLE IF NOT EXISTS accounts (
                id TEXT PRIMARY KEY,
                provider TEXT NOT NULL,
                provider_user_id TEXT NOT NULL,
                email TEXT NOT NULL,
                display_name TEXT NOT NULL,
                status TEXT NOT NULL,
                last_synced_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(provider, provider_user_id)
            );

            CREATE TABLE IF NOT EXISTS external_calendars (
                id TEXT PRIMARY KEY,
                account_id TEXT NOT NULL,
                provider_calendar_id TEXT NOT NULL,
                name TEXT NOT NULL,
                is_primary INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(account_id, provider_calendar_id),
                FOREIGN KEY(account_id) REFERENCES accounts(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS external_events (
                id TEXT PRIMARY KEY,
                account_id TEXT NOT NULL,
                calendar_id TEXT NOT NULL,
                provider TEXT NOT NULL,
                provider_event_id TEXT NOT NULL,
                title TEXT NOT NULL,
                body_content_type TEXT,
                body_content TEXT,
                start_utc TEXT NOT NULL,
                end_utc TEXT NOT NULL,
                start_timezone TEXT,
                end_timezone TEXT,
                is_all_day INTEGER NOT NULL DEFAULT 0,
                location TEXT,
                attendees_json TEXT,
                organizer_json TEXT,
                web_link TEXT,
                online_meeting_url TEXT,
                categories_json TEXT,
                reminder_minutes_before_start INTEGER,
                is_reminder_on INTEGER NOT NULL DEFAULT 0,
                sensitivity TEXT,
                last_modified_utc TEXT,
                deleted_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(account_id, provider_event_id),
                FOREIGN KEY(account_id) REFERENCES accounts(id) ON DELETE CASCADE,
                FOREIGN KEY(calendar_id) REFERENCES external_calendars(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_external_events_time
                ON external_events(start_utc, end_utc, deleted_at);

            CREATE TABLE IF NOT EXISTS local_notes (
                id TEXT PRIMARY KEY,
                title TEXT NOT NULL,
                body TEXT,
                color TEXT NOT NULL,
                note_date TEXT NOT NULL,
                start_utc TEXT NOT NULL,
                end_utc TEXT NOT NULL,
                reminder_at_utc TEXT,
                completed_at TEXT,
                reminded_at TEXT,
                deleted_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_local_notes_date
                ON local_notes(note_date, deleted_at);


            CREATE TABLE IF NOT EXISTS sync_state (
                id TEXT PRIMARY KEY,
                account_id TEXT NOT NULL,
                calendar_id TEXT NOT NULL,
                window_start TEXT NOT NULL,
                window_end TEXT NOT NULL,
                delta_link TEXT,
                last_success_at TEXT,
                last_error TEXT,
                failure_count INTEGER NOT NULL DEFAULT 0,
                updated_at TEXT NOT NULL,
                UNIQUE(account_id, calendar_id, window_start, window_end),
                FOREIGN KEY(account_id) REFERENCES accounts(id) ON DELETE CASCADE,
                FOREIGN KEY(calendar_id) REFERENCES external_calendars(id) ON DELETE CASCADE
            );
            "#,
        )
        .map_err(|error| error.to_string())?;
    ensure_local_note_column(connection, "reminder_at_utc", "TEXT")?;
    ensure_local_note_column(connection, "completed_at", "TEXT")?;
    ensure_local_note_column(connection, "reminded_at", "TEXT")?;
    connection
        .execute(
            "CREATE INDEX IF NOT EXISTS idx_local_notes_reminder ON local_notes(reminder_at_utc, completed_at, reminded_at, deleted_at)",
            [],
        )
        .map_err(|error| error.to_string())?;
    Ok(())
}

fn ensure_local_note_column(
    connection: &Connection,
    column_name: &str,
    column_type: &str,
) -> Result<(), String> {
    let columns = connection
        .prepare("PRAGMA table_info(local_notes)")
        .map_err(|error| error.to_string())?
        .query_map([], |row| row.get::<_, String>(1))
        .map_err(|error| error.to_string())?
        .collect::<Result<Vec<_>, _>>()
        .map_err(|error| error.to_string())?;

    if !columns.iter().any(|name| name == column_name) {
        connection
            .execute(
                &format!("ALTER TABLE local_notes ADD COLUMN {column_name} {column_type}"),
                [],
            )
            .map_err(|error| error.to_string())?;
    }
    Ok(())
}

pub fn now_string() -> String {
    Utc::now().to_rfc3339()
}

pub fn upsert_account(
    connection: &Connection,
    provider_user_id: &str,
    email: &str,
    display_name: &str,
) -> Result<Account, String> {
    upsert_account_for_provider(connection, "outlook", provider_user_id, email, display_name)
}

pub fn upsert_account_for_provider(
    connection: &Connection,
    provider: &str,
    provider_user_id: &str,
    email: &str,
    display_name: &str,
) -> Result<Account, String> {
    let now = now_string();
    let existing_id: Option<String> = connection
        .query_row(
            "SELECT id FROM accounts WHERE provider = ?1 AND provider_user_id = ?2",
            params![provider, provider_user_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|error| error.to_string())?;

    let id = existing_id.unwrap_or_else(|| Uuid::new_v4().to_string());
    connection
        .execute(
            r#"
            INSERT INTO accounts (id, provider, provider_user_id, email, display_name, status, created_at, updated_at)
            VALUES (?1, ?2, ?3, ?4, ?5, 'connected', ?6, ?6)
            ON CONFLICT(provider, provider_user_id) DO UPDATE SET
                email = excluded.email,
                display_name = excluded.display_name,
                status = 'connected',
                updated_at = excluded.updated_at
            "#,
            params![id, provider, provider_user_id, email, display_name, now],
        )
        .map_err(|error| error.to_string())?;

    get_account(connection, &id)?.ok_or_else(|| "账号保存失败".to_string())
}

pub fn get_account(connection: &Connection, account_id: &str) -> Result<Option<Account>, String> {
    connection
        .query_row(
            r#"
            SELECT id, provider, provider_user_id, email, display_name, status, last_synced_at
            FROM accounts WHERE id = ?1
            "#,
            params![account_id],
            |row| {
                Ok(Account {
                    id: row.get(0)?,
                    provider: row.get(1)?,
                    provider_user_id: row.get(2)?,
                    email: row.get(3)?,
                    display_name: row.get(4)?,
                    status: row.get(5)?,
                    last_synced_at: row.get(6)?,
                })
            },
        )
        .optional()
        .map_err(|error| error.to_string())
}

pub fn list_accounts(connection: &Connection) -> Result<Vec<Account>, String> {
    let mut statement = connection
        .prepare(
            r#"
            SELECT id, provider, provider_user_id, email, display_name, status, last_synced_at
            FROM accounts ORDER BY created_at ASC
            "#,
        )
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map([], |row| {
            Ok(Account {
                id: row.get(0)?,
                provider: row.get(1)?,
                provider_user_id: row.get(2)?,
                email: row.get(3)?,
                display_name: row.get(4)?,
                status: row.get(5)?,
                last_synced_at: row.get(6)?,
            })
        })
        .map_err(|error| error.to_string())?;
    rows.collect::<Result<Vec<_>, _>>()
        .map_err(|error| error.to_string())
}

pub fn delete_account(connection: &Connection, account_id: &str) -> Result<(), String> {
    connection
        .execute("DELETE FROM accounts WHERE id = ?1", params![account_id])
        .map_err(|error| error.to_string())?;
    Ok(())
}

pub fn upsert_primary_calendar(
    connection: &Connection,
    account_id: &str,
    provider_calendar_id: &str,
    name: &str,
) -> Result<String, String> {
    let now = now_string();
    let existing_id: Option<String> = connection
        .query_row(
            "SELECT id FROM external_calendars WHERE account_id = ?1 AND provider_calendar_id = ?2",
            params![account_id, provider_calendar_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|error| error.to_string())?;
    let id = existing_id.unwrap_or_else(|| Uuid::new_v4().to_string());
    connection
        .execute(
            r#"
            INSERT INTO external_calendars (id, account_id, provider_calendar_id, name, is_primary, created_at, updated_at)
            VALUES (?1, ?2, ?3, ?4, 1, ?5, ?5)
            ON CONFLICT(account_id, provider_calendar_id) DO UPDATE SET
                name = excluded.name,
                is_primary = 1,
                updated_at = excluded.updated_at
            "#,
            params![id, account_id, provider_calendar_id, name, now],
        )
        .map_err(|error| error.to_string())?;
    Ok(id)
}

pub fn primary_calendar_id(
    connection: &Connection,
    account_id: &str,
) -> Result<Option<String>, String> {
    connection
        .query_row(
            "SELECT id FROM external_calendars WHERE account_id = ?1 AND is_primary = 1 LIMIT 1",
            params![account_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|error| error.to_string())
}

pub fn get_month_view(connection: &Connection, year: i32, month: u32) -> Result<MonthView, String> {
    let start = chrono::NaiveDate::from_ymd_opt(year, month, 1)
        .ok_or_else(|| "年月参数无效".to_string())?
        .and_hms_opt(0, 0, 0)
        .ok_or_else(|| "日期参数无效".to_string())?
        .and_utc();
    let next_month = if month == 12 {
        chrono::NaiveDate::from_ymd_opt(year + 1, 1, 1)
    } else {
        chrono::NaiveDate::from_ymd_opt(year, month + 1, 1)
    }
    .ok_or_else(|| "年月参数无效".to_string())?
    .and_hms_opt(0, 0, 0)
    .ok_or_else(|| "日期参数无效".to_string())?
    .and_utc();

    let mut statement = connection
        .prepare(
            r#"
            SELECT external_events.id, external_events.title, external_events.start_utc,
                   external_events.end_utc, external_events.start_timezone,
                   external_events.end_timezone, external_events.is_all_day,
                   external_events.location, external_events.sensitivity,
                   external_events.categories_json, accounts.provider
            FROM external_events
            JOIN accounts ON accounts.id = external_events.account_id
            WHERE deleted_at IS NULL
              AND start_utc < ?2
              AND end_utc >= ?1
            ORDER BY start_utc ASC, title ASC
            "#,
        )
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map(
            params![start.to_rfc3339(), next_month.to_rfc3339()],
            |row| {
                let categories_json: Option<String> = row.get(9)?;
                let category = categories_json
                    .and_then(|value| serde_json::from_str::<Vec<String>>(&value).ok())
                    .and_then(|values| values.first().cloned());
                Ok(CalendarItem {
                    id: row.get(0)?,
                    source: "outlook".to_string(),
                    source_label: if row.get::<_, String>(10)? == "outlook-com" {
                        "经典 Outlook".to_string()
                    } else {
                        "Outlook".to_string()
                    },
                    read_only: true,
                    title: row.get(1)?,
                    start_utc: row.get(2)?,
                    end_utc: row.get(3)?,
                    start_timezone: row.get(4)?,
                    end_timezone: row.get(5)?,
                    is_all_day: row.get::<_, i64>(6)? != 0,
                    location: row.get(7)?,
                    sensitivity: row.get(8)?,
                    category,
                    note_color: None,
                    note_body: None,
                    reminder_at_utc: None,
                    completed_at: None,
                })
            },
        )
        .map_err(|error| error.to_string())?;
    let mut items = rows
        .collect::<Result<Vec<_>, _>>()
        .map_err(|error| error.to_string())?;
    items.extend(list_local_note_items(
        connection,
        &start.to_rfc3339(),
        &next_month.to_rfc3339(),
    )?);
    items.sort_by(|left, right| {
        left.start_utc
            .cmp(&right.start_utc)
            .then(left.title.cmp(&right.title))
    });
    Ok(MonthView {
        year,
        month,
        items,
        accounts: list_accounts(connection)?,
    })
}

pub fn get_external_event_detail(
    connection: &Connection,
    event_id: &str,
) -> Result<Option<ExternalEventDetail>, String> {
    connection
        .query_row(
            r#"
            SELECT id, provider, provider_event_id, title, body_content_type, body_content,
                   start_utc, end_utc, start_timezone, end_timezone, is_all_day, location,
                   attendees_json, organizer_json, web_link, online_meeting_url, categories_json,
                   reminder_minutes_before_start, is_reminder_on, sensitivity, last_modified_utc
            FROM external_events WHERE id = ?1 AND deleted_at IS NULL
            "#,
            params![event_id],
            |row| {
                Ok(ExternalEventDetail {
                    id: row.get(0)?,
                    source: row.get(1)?,
                    provider_event_id: row.get(2)?,
                    title: row.get(3)?,
                    body_content_type: row.get(4)?,
                    body_content: row.get(5)?,
                    start_utc: row.get(6)?,
                    end_utc: row.get(7)?,
                    start_timezone: row.get(8)?,
                    end_timezone: row.get(9)?,
                    is_all_day: row.get::<_, i64>(10)? != 0,
                    location: row.get(11)?,
                    attendees_json: row.get(12)?,
                    organizer_json: row.get(13)?,
                    web_link: row.get(14)?,
                    online_meeting_url: row.get(15)?,
                    categories_json: row.get(16)?,
                    reminder_minutes_before_start: row.get(17)?,
                    is_reminder_on: row.get::<_, i64>(18)? != 0,
                    sensitivity: row.get(19)?,
                    last_modified_utc: row.get(20)?,
                })
            },
        )
        .optional()
        .map_err(|error| error.to_string())
}

pub fn sync_state_for_window(
    connection: &Connection,
    account_id: &str,
    calendar_id: &str,
    window_start: &str,
    window_end: &str,
) -> Result<(String, Option<String>, i64), String> {
    let existing: Option<(String, Option<String>, i64)> = connection
        .query_row(
            r#"
            SELECT id, delta_link, failure_count
            FROM sync_state
            WHERE account_id = ?1 AND calendar_id = ?2 AND window_start = ?3 AND window_end = ?4
            "#,
            params![account_id, calendar_id, window_start, window_end],
            |row| Ok((row.get(0)?, row.get(1)?, row.get(2)?)),
        )
        .optional()
        .map_err(|error| error.to_string())?;

    if let Some(existing) = existing {
        return Ok(existing);
    }

    let id = Uuid::new_v4().to_string();
    connection
        .execute(
            r#"
            INSERT INTO sync_state (id, account_id, calendar_id, window_start, window_end, updated_at)
            VALUES (?1, ?2, ?3, ?4, ?5, ?6)
            "#,
            params![id, account_id, calendar_id, window_start, window_end, now_string()],
        )
        .map_err(|error| error.to_string())?;
    Ok((id, None, 0))
}

pub fn mark_sync_success(
    connection: &Connection,
    sync_state_id: &str,
    delta_link: &str,
    synced_at: &str,
) -> Result<(), String> {
    connection
        .execute(
            r#"
            UPDATE sync_state
            SET delta_link = ?2,
                last_success_at = ?3,
                last_error = NULL,
                failure_count = 0,
                updated_at = ?3
            WHERE id = ?1
            "#,
            params![sync_state_id, delta_link, synced_at],
        )
        .map_err(|error| error.to_string())?;
    connection
        .execute(
            r#"
            UPDATE accounts
            SET last_synced_at = ?2,
                status = 'connected',
                updated_at = ?2
            WHERE id = ?1
            "#,
            params![
                account_id_from_sync_state(connection, sync_state_id)?,
                synced_at
            ],
        )
        .map_err(|error| error.to_string())?;
    Ok(())
}

pub fn mark_sync_error(
    connection: &Connection,
    sync_state_id: &str,
    error_message: &str,
) -> Result<i64, String> {
    let failure_count = connection
        .query_row(
            "SELECT failure_count FROM sync_state WHERE id = ?1",
            params![sync_state_id],
            |row| row.get::<_, i64>(0),
        )
        .optional()
        .map_err(|error| error.to_string())?
        .unwrap_or(0)
        + 1;
    let now = now_string();
    connection
        .execute(
            r#"
            UPDATE sync_state
            SET last_error = ?2,
                failure_count = ?3,
                updated_at = ?4
            WHERE id = ?1
            "#,
            params![sync_state_id, error_message, failure_count, now],
        )
        .map_err(|error| error.to_string())?;
    Ok(failure_count)
}

fn account_id_from_sync_state(
    connection: &Connection,
    sync_state_id: &str,
) -> Result<String, String> {
    connection
        .query_row(
            "SELECT account_id FROM sync_state WHERE id = ?1",
            params![sync_state_id],
            |row| row.get(0),
        )
        .map_err(|error| error.to_string())
}

#[allow(clippy::too_many_arguments)]
pub fn upsert_external_event(
    connection: &Connection,
    account_id: &str,
    calendar_id: &str,
    provider_event_id: &str,
    title: &str,
    body_content_type: Option<&str>,
    body_content: Option<&str>,
    start_utc: &str,
    end_utc: &str,
    start_timezone: Option<&str>,
    end_timezone: Option<&str>,
    is_all_day: bool,
    location: Option<&str>,
    attendees_json: Option<&str>,
    organizer_json: Option<&str>,
    web_link: Option<&str>,
    online_meeting_url: Option<&str>,
    categories_json: Option<&str>,
    reminder_minutes_before_start: Option<i64>,
    is_reminder_on: bool,
    sensitivity: Option<&str>,
    last_modified_utc: Option<&str>,
) -> Result<(), String> {
    let now = now_string();
    let existing_id: Option<String> = connection
        .query_row(
            "SELECT id FROM external_events WHERE account_id = ?1 AND provider_event_id = ?2",
            params![account_id, provider_event_id],
            |row| row.get(0),
        )
        .optional()
        .map_err(|error| error.to_string())?;
    let id = existing_id.unwrap_or_else(|| Uuid::new_v4().to_string());
    connection
        .execute(
            r#"
            INSERT INTO external_events (
                id, account_id, calendar_id, provider, provider_event_id, title,
                body_content_type, body_content, start_utc, end_utc, start_timezone, end_timezone,
                is_all_day, location, attendees_json, organizer_json, web_link, online_meeting_url,
                categories_json, reminder_minutes_before_start, is_reminder_on, sensitivity,
                last_modified_utc, deleted_at, created_at, updated_at
            )
            VALUES (
                ?1, ?2, ?3, 'outlook', ?4, ?5,
                ?6, ?7, ?8, ?9, ?10, ?11,
                ?12, ?13, ?14, ?15, ?16, ?17,
                ?18, ?19, ?20, ?21,
                ?22, NULL, ?23, ?23
            )
            ON CONFLICT(account_id, provider_event_id) DO UPDATE SET
                calendar_id = excluded.calendar_id,
                title = excluded.title,
                body_content_type = excluded.body_content_type,
                body_content = excluded.body_content,
                start_utc = excluded.start_utc,
                end_utc = excluded.end_utc,
                start_timezone = excluded.start_timezone,
                end_timezone = excluded.end_timezone,
                is_all_day = excluded.is_all_day,
                location = excluded.location,
                attendees_json = excluded.attendees_json,
                organizer_json = excluded.organizer_json,
                web_link = excluded.web_link,
                online_meeting_url = excluded.online_meeting_url,
                categories_json = excluded.categories_json,
                reminder_minutes_before_start = excluded.reminder_minutes_before_start,
                is_reminder_on = excluded.is_reminder_on,
                sensitivity = excluded.sensitivity,
                last_modified_utc = excluded.last_modified_utc,
                deleted_at = NULL,
                updated_at = excluded.updated_at
            "#,
            params![
                id,
                account_id,
                calendar_id,
                provider_event_id,
                title,
                body_content_type,
                body_content,
                start_utc,
                end_utc,
                start_timezone,
                end_timezone,
                i64::from(is_all_day),
                location,
                attendees_json,
                organizer_json,
                web_link,
                online_meeting_url,
                categories_json,
                reminder_minutes_before_start,
                i64::from(is_reminder_on),
                sensitivity,
                last_modified_utc,
                now,
            ],
        )
        .map_err(|error| error.to_string())?;
    Ok(())
}

pub fn mark_external_event_deleted(
    connection: &Connection,
    account_id: &str,
    provider_event_id: &str,
) -> Result<(), String> {
    let now = now_string();
    connection
        .execute(
            r#"
            UPDATE external_events
            SET deleted_at = ?3,
                updated_at = ?3
            WHERE account_id = ?1 AND provider_event_id = ?2
            "#,
            params![account_id, provider_event_id, now],
        )
        .map_err(|error| error.to_string())?;
    Ok(())
}

pub fn mark_missing_external_events_deleted(
    connection: &Connection,
    account_id: &str,
    calendar_id: &str,
    window_start: &str,
    window_end: &str,
    seen_provider_event_ids: &[String],
) -> Result<(), String> {
    let now = now_string();
    let mut sql = String::from(
        r#"
        UPDATE external_events
        SET deleted_at = ?1,
            updated_at = ?1
        WHERE account_id = ?2
          AND calendar_id = ?3
          AND deleted_at IS NULL
          AND start_utc < ?5
          AND end_utc > ?4
        "#,
    );
    let mut values = vec![
        now,
        account_id.to_string(),
        calendar_id.to_string(),
        window_start.to_string(),
        window_end.to_string(),
    ];

    if !seen_provider_event_ids.is_empty() {
        let placeholders = (0..seen_provider_event_ids.len())
            .map(|index| format!("?{}", index + 6))
            .collect::<Vec<_>>()
            .join(", ");
        sql.push_str(" AND provider_event_id NOT IN (");
        sql.push_str(&placeholders);
        sql.push(')');
        values.extend(seen_provider_event_ids.iter().cloned());
    }

    connection
        .execute(&sql, params_from_iter(values))
        .map_err(|error| error.to_string())?;
    Ok(())
}

pub fn create_local_note(
    connection: &Connection,
    request: CreateNoteRequest,
) -> Result<CalendarItem, String> {
    let title = request.title.trim().to_string();
    if title.is_empty() {
        return Err("待办标题不能为空".to_string());
    }

    let date_key = request.date_key.trim().to_string();
    let date = chrono::NaiveDate::parse_from_str(&date_key, "%Y-%m-%d")
        .map_err(|_| "待办日期格式无效".to_string())?;
    let start_naive = date
        .and_hms_opt(0, 0, 0)
        .ok_or_else(|| "待办开始时间无效".to_string())?;
    let end_naive = date
        .succ_opt()
        .ok_or_else(|| "待办结束日期无效".to_string())?
        .and_hms_opt(0, 0, 0)
        .ok_or_else(|| "待办结束时间无效".to_string())?;
    let start = Local
        .from_local_datetime(&start_naive)
        .single()
        .or_else(|| Local.from_local_datetime(&start_naive).earliest())
        .ok_or_else(|| "待办开始时间无效".to_string())?
        .with_timezone(&Utc);
    let end = Local
        .from_local_datetime(&end_naive)
        .single()
        .or_else(|| Local.from_local_datetime(&end_naive).earliest())
        .ok_or_else(|| "待办结束时间无效".to_string())?
        .with_timezone(&Utc);
    let color = normalize_note_color(&request.color).to_string();
    let reminder_at_utc = normalize_reminder_at(request.reminder_at_utc)?;
    let body = request
        .body
        .as_deref()
        .map(str::trim)
        .filter(|body| !body.is_empty())
        .map(ToOwned::to_owned);
    let id = Uuid::new_v4().to_string();
    let now = now_string();

    connection
        .execute(
            r#"
            INSERT INTO local_notes (
                id, title, body, color, note_date, start_utc, end_utc,
                reminder_at_utc, completed_at, reminded_at, created_at, updated_at
            )
            VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, NULL, NULL, ?9, ?9)
            "#,
            params![
                &id,
                &title,
                body.as_deref(),
                &color,
                &date_key,
                start.to_rfc3339(),
                end.to_rfc3339(),
                reminder_at_utc.as_deref(),
                &now,
            ],
        )
        .map_err(|error| error.to_string())?;

    Ok(CalendarItem {
        id,
        source: "local_note".to_string(),
        source_label: "待办".to_string(),
        read_only: false,
        title,
        start_utc: start.to_rfc3339(),
        end_utc: end.to_rfc3339(),
        start_timezone: None,
        end_timezone: None,
        is_all_day: true,
        location: None,
        sensitivity: None,
        category: Some(color.clone()),
        note_color: Some(color),
        note_body: body,
        reminder_at_utc,
        completed_at: None,
    })
}

pub fn update_local_note_completion(
    connection: &Connection,
    request: UpdateNoteCompletionRequest,
) -> Result<CalendarItem, String> {
    let now = now_string();
    let completed_at = if request.completed {
        Some(now.as_str())
    } else {
        None
    };
    let changed = connection
        .execute(
            r#"
            UPDATE local_notes
            SET completed_at = ?2,
                updated_at = ?3
            WHERE id = ?1 AND deleted_at IS NULL
            "#,
            params![&request.note_id, completed_at, &now],
        )
        .map_err(|error| error.to_string())?;

    if changed == 0 {
        return Err("待办不存在或已删除".to_string());
    }

    get_local_note_item(connection, &request.note_id)
}

pub fn take_due_local_note_reminders(
    connection: &Connection,
) -> Result<Vec<ReminderNotice>, String> {
    let now = now_string();
    let mut statement = connection
        .prepare(
            r#"
            SELECT id, title, body, color, reminder_at_utc
            FROM local_notes
            WHERE deleted_at IS NULL
              AND completed_at IS NULL
              AND reminded_at IS NULL
              AND reminder_at_utc IS NOT NULL
              AND reminder_at_utc <= ?1
            ORDER BY reminder_at_utc ASC, created_at ASC
            "#,
        )
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map(params![&now], |row| {
            Ok(ReminderNotice {
                id: row.get(0)?,
                title: row.get(1)?,
                body: row.get(2)?,
                color: row.get(3)?,
                reminder_at_utc: row.get(4)?,
            })
        })
        .map_err(|error| error.to_string())?;
    let candidates = rows
        .collect::<Result<Vec<_>, _>>()
        .map_err(|error| error.to_string())?;
    drop(statement);

    let mut reminders = Vec::new();
    for reminder in candidates {
        let changed = connection
            .execute(
                r#"
                UPDATE local_notes
                SET reminded_at = ?2,
                    updated_at = ?2
                WHERE id = ?1
                  AND deleted_at IS NULL
                  AND completed_at IS NULL
                  AND reminded_at IS NULL
                "#,
                params![&reminder.id, &now],
            )
            .map_err(|error| error.to_string())?;
        if changed > 0 {
            reminders.push(reminder);
        }
    }

    Ok(reminders)
}

fn list_local_note_items(
    connection: &Connection,
    window_start: &str,
    window_end: &str,
) -> Result<Vec<CalendarItem>, String> {
    let mut statement = connection
        .prepare(
            r#"
            SELECT id, title, body, color, start_utc, end_utc, reminder_at_utc, completed_at
            FROM local_notes
            WHERE deleted_at IS NULL
              AND start_utc < ?2
              AND end_utc >= ?1
            ORDER BY start_utc ASC, created_at ASC
            "#,
        )
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map(params![window_start, window_end], |row| {
            local_note_item_from_row(row)
        })
        .map_err(|error| error.to_string())?;
    rows.collect::<Result<Vec<_>, _>>()
        .map_err(|error| error.to_string())
}

pub fn list_upcoming_local_note_items(
    connection: &Connection,
    from_utc: &str,
) -> Result<Vec<CalendarItem>, String> {
    let mut statement = connection
        .prepare(
            r#"
            SELECT id, title, body, color, start_utc, end_utc, reminder_at_utc, completed_at
            FROM local_notes
            WHERE deleted_at IS NULL
              AND completed_at IS NULL
              AND end_utc >= ?1
            ORDER BY start_utc ASC, created_at ASC
            "#,
        )
        .map_err(|error| error.to_string())?;
    let rows = statement
        .query_map(params![from_utc], |row| local_note_item_from_row(row))
        .map_err(|error| error.to_string())?;
    rows.collect::<Result<Vec<_>, _>>()
        .map_err(|error| error.to_string())
}

pub fn get_local_note_item(connection: &Connection, id: &str) -> Result<CalendarItem, String> {
    connection
        .query_row(
            r#"
            SELECT id, title, body, color, start_utc, end_utc, reminder_at_utc, completed_at
            FROM local_notes
            WHERE id = ?1 AND deleted_at IS NULL
            "#,
            params![id],
            |row| local_note_item_from_row(row),
        )
        .optional()
        .map_err(|error| error.to_string())?
        .ok_or_else(|| "待办不存在或已删除".to_string())
}

fn local_note_item_from_row(row: &rusqlite::Row<'_>) -> rusqlite::Result<CalendarItem> {
    let color: String = row.get(3)?;
    Ok(CalendarItem {
        id: row.get(0)?,
        source: "local_note".to_string(),
        source_label: "待办".to_string(),
        read_only: false,
        title: row.get(1)?,
        start_utc: row.get(4)?,
        end_utc: row.get(5)?,
        start_timezone: None,
        end_timezone: None,
        is_all_day: true,
        location: None,
        sensitivity: None,
        category: Some(color.clone()),
        note_color: Some(color),
        note_body: row.get(2)?,
        reminder_at_utc: row.get(6)?,
        completed_at: row.get(7)?,
    })
}

fn normalize_reminder_at(value: Option<String>) -> Result<Option<String>, String> {
    let Some(value) = value
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty())
    else {
        return Ok(None);
    };

    chrono::DateTime::parse_from_rfc3339(&value)
        .map(|datetime| Some(datetime.with_timezone(&Utc).to_rfc3339()))
        .map_err(|_| "提醒时间格式无效".to_string())
}

fn normalize_note_color(color: &str) -> &'static str {
    match color.trim() {
        "sage" => "sage",
        "rose" => "rose",
        "blue" => "blue",
        "amber" => "amber",
        _ => "paper",
    }
}
