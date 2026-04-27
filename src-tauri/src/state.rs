use std::sync::{Arc, Mutex};

use reqwest::Client;
use rusqlite::Connection;
use tokio::sync::Mutex as AsyncMutex;

#[derive(Clone)]
pub struct AppServices {
    pub db: Arc<Mutex<Connection>>,
    pub http: Client,
    pub sync_lock: Arc<AsyncMutex<()>>,
    pub floating_note_id: Arc<Mutex<Option<String>>>,
}

impl AppServices {
    pub fn new(connection: Connection) -> Self {
        Self {
            db: Arc::new(Mutex::new(connection)),
            http: Client::new(),
            sync_lock: Arc::new(AsyncMutex::new(())),
            floating_note_id: Arc::new(Mutex::new(None)),
        }
    }
}
