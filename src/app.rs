use windows::core::Result;

use crate::{
    document::export::AutoSaveManager,
    document::model::DocumentModel,
    settings::{SettingsStore, schema::Settings},
    theme::ThemeManager,
    window::AppWindow,
};

pub struct App {
    window: AppWindow,
    #[allow(dead_code)]
    state: AppState,
    #[allow(dead_code)]
    startup: StartupState,
}

pub struct AppState {
    pub show_tabs: bool,
    pub show_sidebar: bool,
    pub sidebar_width: f32,
    pub show_toolbar: bool,
    pub show_statusbar: bool,
    pub show_settings: bool,
    pub show_debug_panel: bool,
    pub status_text: String,
    pub document: DocumentModel,
    pub autosave: AutoSaveManager,
    pub settings: Settings,
}

impl Default for AppState {
    fn default() -> Self {
        Self {
            show_tabs: true,
            show_sidebar: true,
            sidebar_width: 260.0,
            show_toolbar: true,
            show_statusbar: true,
            show_settings: false,
            show_debug_panel: false,
            status_text: "Ready".to_string(),
            document: DocumentModel::default(),
            autosave: AutoSaveManager::new(60),
            settings: Settings::default(),
        }
    }
}

pub struct StartupState {
    pub window_shown_immediately: bool,
    pub theme_loaded: bool,
    pub settings_loaded: bool,
    pub recent_files_loaded: bool,
    pub startup_budget_ms: u32,
}

impl Default for StartupState {
    fn default() -> Self {
        Self {
            window_shown_immediately: true,
            theme_loaded: false,
            settings_loaded: false,
            recent_files_loaded: false,
            startup_budget_ms: 500,
        }
    }
}

impl StartupState {
    fn finish_theme_load(&mut self) {
        self.theme_loaded = true;
    }

    fn finish_settings_load(&mut self) {
        self.settings_loaded = true;
    }

    fn finish_recent_files_load(&mut self) {
        self.recent_files_loaded = true;
    }
}

impl App {
    pub fn new() -> Result<Self> {
        crate::profile_scope!("app.new");

        let mut startup = StartupState::default();
        let themes = ThemeManager::load()?;
        startup.finish_theme_load();

        let settings = SettingsStore::load()
            .map(|store| store.settings().clone())
            .unwrap_or_default();
        startup.finish_settings_load();
        startup.finish_recent_files_load();

        let window = AppWindow::new(themes.active().clone())?;

        let mut state = AppState::default();
        state.settings = settings;
        state.show_toolbar = state.settings.appearance.show_toolbar;
        state.show_sidebar = state.settings.appearance.show_sidebar;
        state.show_statusbar = state.settings.appearance.show_status_bar;
        state.show_tabs = state.settings.appearance.show_tab_bar;
        state.autosave = AutoSaveManager::new(
            state
                .settings
                .files
                .auto_save_interval
                .as_seconds()
                .unwrap_or(60),
        );

        Ok(Self {
            window,
            state,
            startup,
        })
    }

    pub fn run(self) -> Result<()> {
        self.window.run()
    }
}
