use std::time::Instant;

use windows::core::Result;

use crate::{
    document::export::AutoSaveManager,
    render::perf::emit_startup_marker,
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
    pub theme_load_ms: u32,
    pub settings_load_ms: u32,
    pub window_init_ms: u32,
    pub total_startup_ms: u32,
}

impl Default for StartupState {
    fn default() -> Self {
        Self {
            window_shown_immediately: true,
            theme_loaded: false,
            settings_loaded: false,
            recent_files_loaded: false,
            startup_budget_ms: 500,
            theme_load_ms: 0,
            settings_load_ms: 0,
            window_init_ms: 0,
            total_startup_ms: 0,
        }
    }
}

impl StartupState {
    fn finish_theme_load(&mut self, elapsed_ms: u32) {
        self.theme_loaded = true;
        self.theme_load_ms = elapsed_ms;
    }

    fn finish_settings_load(&mut self, elapsed_ms: u32) {
        self.settings_loaded = true;
        self.settings_load_ms = elapsed_ms;
    }

    fn finish_recent_files_load(&mut self) {
        self.recent_files_loaded = true;
    }

    fn finish_window_init(&mut self, elapsed_ms: u32) {
        self.window_init_ms = elapsed_ms;
        self.window_shown_immediately = elapsed_ms <= 200;
    }

    fn finish_startup(&mut self, elapsed_ms: u32) {
        self.total_startup_ms = elapsed_ms;
    }
}

impl App {
    pub fn new() -> Result<Self> {
        crate::profile_scope!("app.new");

        let startup_begin = Instant::now();
        let mut startup = StartupState::default();

        let theme_begin = Instant::now();
        let themes = ThemeManager::load()?;
        let theme_ms = theme_begin.elapsed().as_millis() as u32;
        startup.finish_theme_load(theme_ms);
        emit_startup_marker("theme_load", theme_ms as f64);

        let settings_begin = Instant::now();
        let settings = SettingsStore::load()
            .map(|store| store.settings().clone())
            .unwrap_or_default();
        let settings_ms = settings_begin.elapsed().as_millis() as u32;
        startup.finish_settings_load(settings_ms);
        emit_startup_marker("settings_load", settings_ms as f64);
        startup.finish_recent_files_load();

        let _ = themes.apply_preference(&settings.appearance.theme);

        let window_begin = Instant::now();
        let window = AppWindow::new(themes, settings.clone())?;
        let window_ms = window_begin.elapsed().as_millis() as u32;
        startup.finish_window_init(window_ms);
        emit_startup_marker("window_init", window_ms as f64);

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

        let total_ms = startup_begin.elapsed().as_millis() as u32;
        startup.finish_startup(total_ms);
        emit_startup_marker("total", total_ms as f64);
        if total_ms > startup.startup_budget_ms {
            eprintln!(
                "Startup budget exceeded: {} ms > {} ms",
                total_ms, startup.startup_budget_ms
            );
        }

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
