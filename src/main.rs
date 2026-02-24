#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod app;
mod document;
mod editor;
mod render;
mod settings;
mod theme;
mod ui;
mod window;

fn main() -> windows::core::Result<()> {
    app::App::new()?.run()
}
