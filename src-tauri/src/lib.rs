#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod commands;
mod db;
mod models;
mod outlook_auth;
mod outlook_com_sync;
mod outlook_sync;
mod state;

use std::time::Duration;

use state::AppServices;
use tauri::menu::{Menu, MenuItem};
use tauri::tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent};
use tauri::{Emitter, Manager, WindowEvent};
use tauri_plugin_notification::NotificationExt;

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_notification::init())
        .setup(|app| {
            let app_data_dir = app
                .path()
                .app_data_dir()
                .unwrap_or_else(|_| fallback_app_data_dir());
            let connection = db::open_database(&app_data_dir)?;
            let services = AppServices::new(connection);
            app.manage(services.clone());
            setup_tray(app)?;
            start_background_reminders(app.handle().clone(), services);
            Ok(())
        })
        .on_window_event(|window, event| {
            if let WindowEvent::CloseRequested { api, .. } = event {
                api.prevent_close();
                let _ = window.hide();
            }
        })
        .invoke_handler(tauri::generate_handler![
            commands::connect_outlook,
            commands::connect_local_outlook,
            commands::refresh_local_outlook,
            commands::disconnect_outlook,
            commands::sync_outlook_now,
            commands::get_month_view,
            commands::create_note,
            commands::set_note_completed,
            commands::get_external_event_detail,
            commands::get_local_note,
            commands::get_upcoming_local_notes,
            open_floating_note,
            get_active_floating_note_id,
            set_floating_note_always_on_top,
        ])
        .run(tauri::generate_context!())
        .expect("error while running Calendar Notes");
}

fn setup_tray(app: &tauri::App) -> Result<(), Box<dyn std::error::Error>> {
    let show_item = MenuItem::with_id(app, "show", "显示 Calendar Notes", true, None::<&str>)?;
    let quit_item = MenuItem::with_id(app, "quit", "退出", true, None::<&str>)?;
    let menu = Menu::with_items(app, &[&show_item, &quit_item])?;

    TrayIconBuilder::new()
        .tooltip("Calendar Notes")
        .menu(&menu)
        .show_menu_on_left_click(false)
        .on_menu_event(|app, event| match event.id.as_ref() {
            "show" => show_main_window(app),
            "quit" => app.exit(0),
            _ => {}
        })
        .on_tray_icon_event(|tray, event| {
            if let TrayIconEvent::Click {
                button: MouseButton::Left,
                button_state: MouseButtonState::Up,
                ..
            } = event
            {
                show_main_window(tray.app_handle());
            }
        })
        .build(app)?;
    Ok(())
}

fn show_main_window(app: &tauri::AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.unminimize();
        let _ = window.show();
        let _ = window.set_focus();
    }
}

#[tauri::command]
fn open_floating_note(
    app: tauri::AppHandle,
    services: tauri::State<'_, AppServices>,
    note_id: String,
) -> Result<(), String> {
    let safe_note_id: String = note_id
        .chars()
        .filter(|character| character.is_ascii_alphanumeric() || *character == '-')
        .collect();
    if safe_note_id.is_empty() {
        return Err("待办 ID 无效".to_string());
    }

    let mut current_note_id = services
        .floating_note_id
        .lock()
        .map_err(|error| error.to_string())?;
    *current_note_id = Some(safe_note_id.clone());
    drop(current_note_id);

    let window = app
        .get_webview_window("floating-note")
        .ok_or_else(|| "悬浮便签窗口未初始化".to_string())?;
    window.set_shadow(true).map_err(|error| error.to_string())?;
    window
        .set_always_on_top(true)
        .map_err(|error| error.to_string())?;
    window.show().map_err(|error| error.to_string())?;
    window.unminimize().map_err(|error| error.to_string())?;
    window.set_focus().map_err(|error| error.to_string())?;
    app.emit_to("floating-note", "floating-note-open", safe_note_id)
        .map_err(|error| error.to_string())?;
    Ok(())
}

#[tauri::command]
fn get_active_floating_note_id(
    services: tauri::State<'_, AppServices>,
) -> Result<Option<String>, String> {
    services
        .floating_note_id
        .lock()
        .map(|note_id| note_id.clone())
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn set_floating_note_always_on_top(
    window: tauri::WebviewWindow,
    always_on_top: bool,
) -> Result<(), String> {
    window
        .set_always_on_top(always_on_top)
        .map_err(|error| error.to_string())
}

fn start_background_reminders(app: tauri::AppHandle, services: AppServices) {
    tauri::async_runtime::spawn(async move {
        loop {
            let reminders = {
                let connection = services.db.lock().map_err(|error| error.to_string());
                match connection {
                    Ok(connection) => {
                        db::take_due_local_note_reminders(&connection).unwrap_or_default()
                    }
                    Err(_) => Vec::new(),
                }
            };

            for reminder in reminders {
                show_main_window(&app);
                let _ = app.emit("todo-reminder", reminder.clone());
                let body = reminder
                    .body
                    .as_deref()
                    .filter(|body| !body.trim().is_empty())
                    .unwrap_or("提醒时间已到，当前还没有标记完成。");
                let _ = app
                    .notification()
                    .builder()
                    .title(format!("待办提醒：{}", reminder.title))
                    .body(body)
                    .show();
            }

            tokio::time::sleep(Duration::from_secs(30)).await;
        }
    });
}
fn fallback_app_data_dir() -> std::path::PathBuf {
    dirs::data_local_dir()
        .unwrap_or_else(std::env::temp_dir)
        .join("Calendar Notes")
}
