//! Native shell for Sheaf: file access, the application menu and
//! OS integration (command-line arguments, Finder "Open With", drag and drop,
//! unsaved-changes protection). Parsing, editing and rendering happen in the web frontend.

use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use std::time::Duration;

use serde::Serialize;
use tauri::menu::{Menu, MenuBuilder, MenuItem, MenuItemBuilder, PredefinedMenuItem, Submenu, SubmenuBuilder};
use tauri::{AppHandle, Emitter, Manager, Runtime, WindowEvent, Wry};

const SUPPORTED_EXTENSIONS: [&str; 5] = ["xlsx", "xlsm", "xls", "csv", "tsv"];
/// Formats the app writes (never .xls).
const WRITABLE_EXTENSIONS: [&str; 4] = ["xlsx", "xlsm", "csv", "tsv"];
/// Limits match the frontend (src/services/workbook/formatDetection.ts).
const MAX_WORKBOOK_BYTES: u64 = 200 * 1024 * 1024;
const MAX_TEXT_BYTES: u64 = 500 * 1024 * 1024;

/// Files to open that arrived before the frontend was listening.
#[derive(Default)]
struct PendingFiles {
    frontend_ready: bool,
    paths: Vec<String>,
}

#[derive(Default)]
struct AppState {
    pending: Mutex<PendingFiles>,
    recent: Mutex<Vec<String>>,
    recent_menu: Mutex<Option<Submenu<Wry>>>,
    /// The open document has unsaved changes: closing asks the frontend first.
    document_edited: AtomicBool,
    /// Set once the user confirmed quitting, so the close isn't intercepted again.
    quitting: AtomicBool,
}

/// Error returned to the frontend; `code` maps to a user-facing message there.
#[derive(Debug, Serialize)]
struct FileError {
    code: &'static str,
    message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    size: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    limit: Option<u64>,
}

impl FileError {
    fn new(code: &'static str, message: impl Into<String>) -> Self {
        Self { code, message: message.into(), size: None, limit: None }
    }
}

impl From<std::io::Error> for FileError {
    fn from(error: std::io::Error) -> Self {
        let code = match error.kind() {
            std::io::ErrorKind::NotFound => "NOT_FOUND",
            std::io::ErrorKind::PermissionDenied => "PERMISSION_DENIED",
            _ => "READ_FAILED",
        };
        FileError::new(code, error.to_string())
    }
}

fn extension(path: &Path) -> String {
    path.extension().and_then(|e| e.to_str()).unwrap_or_default().to_ascii_lowercase()
}

/// Reads a spreadsheet file and returns its raw bytes (sent as an ArrayBuffer, not JSON).
/// Only spreadsheet extensions are readable, and oversized files are refused before reading.
#[tauri::command]
async fn read_workbook_file(path: String) -> Result<tauri::ipc::Response, FileError> {
    let path = PathBuf::from(path);
    let ext = extension(&path);
    if !ext.is_empty() && !SUPPORTED_EXTENSIONS.contains(&ext.as_str()) {
        return Err(FileError::new("UNSUPPORTED_FORMAT", format!("Unsupported extension .{ext}")));
    }

    let metadata = std::fs::metadata(&path)?;
    if !metadata.is_file() {
        return Err(FileError::new("READ_FAILED", "Not a regular file"));
    }
    let limit = if ext == "csv" || ext == "tsv" { MAX_TEXT_BYTES } else { MAX_WORKBOOK_BYTES };
    if metadata.len() > limit {
        return Err(FileError {
            code: "TOO_LARGE",
            message: format!("{} bytes exceeds the {limit} byte limit", metadata.len()),
            size: Some(metadata.len()),
            limit: Some(limit),
        });
    }

    let bytes = tauri::async_runtime::spawn_blocking(move || std::fs::read(path))
        .await
        .map_err(|e| FileError::new("READ_FAILED", e.to_string()))??;
    Ok(tauri::ipc::Response::new(bytes))
}

fn decode_percent(text: &str) -> Option<String> {
    let bytes = text.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            let hex = std::str::from_utf8(&bytes[i + 1..i + 3]).ok()?;
            out.push(u8::from_str_radix(hex, 16).ok()?);
            i += 3;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    String::from_utf8(out).ok()
}

/// Writes a workbook the frontend built. The body is the raw file; the path is in
/// the `x-path` header (percent-encoded). Written to a temporary file first and
/// renamed into place, so a failed save never leaves a half-written file.
#[tauri::command]
async fn write_workbook_file(request: tauri::ipc::Request<'_>) -> Result<(), FileError> {
    let tauri::ipc::InvokeBody::Raw(data) = request.body() else {
        return Err(FileError::new("READ_FAILED", "Expected raw file contents"));
    };
    let path = request
        .headers()
        .get("x-path")
        .and_then(|value| value.to_str().ok())
        .and_then(decode_percent)
        .map(PathBuf::from)
        .ok_or_else(|| FileError::new("READ_FAILED", "Missing destination path"))?;
    if !WRITABLE_EXTENSIONS.contains(&extension(&path).as_str()) {
        return Err(FileError::new("UNSUPPORTED_FORMAT", "Only .xlsx, .xlsm, .csv and .tsv can be saved"));
    }
    let data = data.clone();
    tauri::async_runtime::spawn_blocking(move || write_atomically(&path, &data))
        .await
        .map_err(|e| FileError::new("READ_FAILED", e.to_string()))??;
    Ok(())
}

/// Writes next to the target, then renames over it (atomic on the same volume).
fn write_atomically(path: &Path, data: &[u8]) -> std::io::Result<()> {
    let name = path.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default();
    let temp = path.with_file_name(format!(".{name}.saving-{}", std::process::id()));
    let result = (|| {
        let mut file = std::fs::File::create(&temp)?;
        file.write_all(data)?;
        file.sync_all()?;
        std::fs::rename(&temp, path)
    })();
    if result.is_err() {
        let _ = std::fs::remove_file(&temp);
    }
    result
}

/// The frontend reports whether there are unsaved changes.
#[tauri::command]
fn set_document_edited(state: tauri::State<'_, AppState>, edited: bool) {
    state.document_edited.store(edited, Ordering::SeqCst);
}

/// Quits after the frontend has dealt with unsaved changes.
#[tauri::command]
fn quit_app(app: AppHandle, state: tauri::State<'_, AppState>) {
    state.quitting.store(true, Ordering::SeqCst);
    app.exit(0);
}

/// True when a close/quit should be handed to the frontend (unsaved changes).
fn should_confirm_close<R: Runtime>(app: &AppHandle<R>) -> bool {
    let state = app.state::<AppState>();
    state.document_edited.load(Ordering::SeqCst) && !state.quitting.load(Ordering::SeqCst)
}

/// Returns files passed at launch (once) and marks the frontend as ready for `open-files` events.
#[tauri::command]
fn take_startup_files(state: tauri::State<'_, AppState>) -> Vec<String> {
    let mut pending = state.pending.lock().unwrap();
    pending.frontend_ready = true;
    std::mem::take(&mut pending.paths)
}

/// Rebuilds File ▸ Open Recent from the frontend's list.
#[tauri::command]
fn set_recent_files(app: AppHandle, state: tauri::State<'_, AppState>, paths: Vec<String>) -> Result<(), String> {
    *state.recent.lock().unwrap() = paths.clone();
    if let Some(menu) = state.recent_menu.lock().unwrap().as_ref() {
        fill_recent_menu(&app, menu, &paths).map_err(|e| e.to_string())?;
    }
    Ok(())
}

/// Opens files now if the frontend is listening, otherwise queues them.
#[cfg_attr(not(any(target_os = "macos", target_os = "ios")), allow(dead_code))]
fn open_paths<R: Runtime>(app: &AppHandle<R>, paths: Vec<String>) {
    if paths.is_empty() {
        return;
    }
    let state = app.state::<AppState>();
    let mut pending = state.pending.lock().unwrap();
    if pending.frontend_ready {
        drop(pending);
        let _ = app.emit("open-files", paths);
        if let Some(window) = app.get_webview_window("main") {
            let _ = window.unminimize();
            let _ = window.set_focus();
        }
    } else {
        pending.paths.extend(paths);
    }
}

/// Spreadsheet paths given on the command line (Windows/Linux "Open with", terminals).
fn startup_paths() -> Vec<String> {
    std::env::args_os()
        .skip(1)
        .map(PathBuf::from)
        .filter(|p| !p.to_string_lossy().starts_with('-') && p.is_file())
        .map(|p| std::path::absolute(&p).unwrap_or(p).to_string_lossy().into_owned())
        .collect()
}

fn item<R: Runtime>(app: &AppHandle<R>, id: &str, label: &str, accelerator: Option<&str>) -> tauri::Result<MenuItem<R>> {
    let builder = MenuItemBuilder::with_id(id, label);
    match accelerator {
        Some(keys) => builder.accelerator(keys).build(app),
        None => builder.build(app),
    }
}

fn fill_recent_menu<R: Runtime>(app: &AppHandle<R>, menu: &Submenu<R>, paths: &[String]) -> tauri::Result<()> {
    for existing in menu.items()? {
        menu.remove(&existing)?;
    }
    if paths.is_empty() {
        let empty = MenuItemBuilder::with_id("recent.none", "No Recent Files").enabled(false).build(app)?;
        menu.append(&empty)?;
        return Ok(());
    }
    for (index, path) in paths.iter().enumerate() {
        let label = Path::new(path)
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| path.clone());
        menu.append(&MenuItemBuilder::with_id(format!("recent:{index}"), label).build(app)?)?;
    }
    menu.append(&PredefinedMenuItem::separator(app)?)?;
    menu.append(&item(app, "recent.clear", "Clear Menu", None)?)?;
    Ok(())
}

/// Application menu. Item ids are frontend command ids (src/state/commands.ts).
fn build_menu<R: Runtime>(app: &AppHandle<R>) -> tauri::Result<(Menu<R>, Submenu<R>)> {
    let recent = SubmenuBuilder::new(app, "Open Recent").build()?;

    let file = SubmenuBuilder::new(app, "File")
        .item(&item(app, "file.new", "New", Some("CmdOrCtrl+N"))?)
        .item(&item(app, "file.open", "Open…", Some("CmdOrCtrl+O"))?)
        .item(&recent)
        .separator()
        .item(&item(app, "file.close", "Close Workbook", Some("CmdOrCtrl+W"))?)
        .separator()
        .item(&item(app, "file.save", "Save", Some("CmdOrCtrl+S"))?)
        .item(&item(app, "file.saveAs", "Save As…", Some("CmdOrCtrl+Shift+S"))?);
    // Quit is a custom item so unsaved changes can be confirmed first.
    #[cfg(not(target_os = "macos"))]
    let file = file.separator().item(&item(app, "app.quit", "Exit", None)?);
    let file = file.build()?;

    #[cfg(target_os = "macos")]
    let (redo_keys, replace_keys, strike_keys) = ("Cmd+Shift+Z", "Cmd+Shift+H", "Cmd+Shift+X");
    #[cfg(not(target_os = "macos"))]
    let (redo_keys, replace_keys, strike_keys) = ("Ctrl+Y", "Ctrl+H", "Ctrl+5");
    let edit = SubmenuBuilder::new(app, "Edit")
        .item(&item(app, "edit.undo", "Undo", Some("CmdOrCtrl+Z"))?)
        .item(&item(app, "edit.redo", "Redo", Some(redo_keys))?)
        .separator()
        .cut()
        .copy()
        .paste()
        .item(&item(app, "edit.pasteValues", "Paste Values", Some("CmdOrCtrl+Shift+V"))?)
        .item(&item(app, "edit.clear", "Clear Contents", None)?)
        .select_all()
        .separator()
        .item(&item(app, "edit.fillDown", "Fill Down", Some("CmdOrCtrl+D"))?)
        .item(&item(app, "edit.fillRight", "Fill Right", Some("CmdOrCtrl+R"))?)
        .separator()
        .item(&item(app, "delete.cells", "Delete Rows or Columns", Some("CmdOrCtrl+Alt+-"))?)
        .item(&item(app, "delete.rows", "Delete Rows", None)?)
        .item(&item(app, "delete.columns", "Delete Columns", None)?)
        .separator()
        .item(&item(app, "edit.find", "Find…", Some("CmdOrCtrl+F"))?)
        .item(&item(app, "edit.replace", "Find and Replace…", Some(replace_keys))?)
        .item(&item(app, "edit.findNext", "Find Next", Some("F3"))?)
        .item(&item(app, "edit.findPrevious", "Find Previous", Some("Shift+F3"))?)
        .separator()
        .item(&item(app, "edit.goTo", "Go to Cell…", Some("CmdOrCtrl+G"))?)
        .build()?;

    let insert = SubmenuBuilder::new(app, "Insert")
        .item(&item(app, "insert.cells", "Rows or Columns", Some("CmdOrCtrl+Alt+="))?)
        .separator()
        .item(&item(app, "insert.rowsAbove", "Rows Above", None)?)
        .item(&item(app, "insert.rowsBelow", "Rows Below", None)?)
        .item(&item(app, "insert.columnsLeft", "Columns Left", None)?)
        .item(&item(app, "insert.columnsRight", "Columns Right", None)?)
        .separator()
        .item(&item(app, "insert.sheet", "Sheet", Some("Shift+F11"))?)
        .separator()
        .item(&item(app, "insert.date", "Today’s Date", Some("CmdOrCtrl+;"))?)
        .item(&item(app, "insert.time", "Current Time", None)?)
        .build()?;

    let number = SubmenuBuilder::new(app, "Number")
        .item(&item(app, "format.numberGeneral", "General", None)?)
        .item(&item(app, "format.numberNumber", "Number", None)?)
        .item(&item(app, "format.numberCurrency", "Currency", None)?)
        .item(&item(app, "format.numberPercent", "Percentage", None)?)
        .item(&item(app, "format.numberDate", "Date", None)?)
        .separator()
        .item(&item(app, "format.increaseDecimals", "Increase Decimals", None)?)
        .item(&item(app, "format.decreaseDecimals", "Decrease Decimals", None)?)
        .build()?;
    let rows = SubmenuBuilder::new(app, "Rows")
        .item(&item(app, "format.hideRows", "Hide", None)?)
        .item(&item(app, "format.unhideRows", "Unhide", None)?)
        .build()?;
    let columns = SubmenuBuilder::new(app, "Columns")
        .item(&item(app, "format.hideColumns", "Hide", None)?)
        .item(&item(app, "format.unhideColumns", "Unhide", None)?)
        .build()?;
    let sheet = SubmenuBuilder::new(app, "Sheet")
        .item(&item(app, "format.renameSheet", "Rename…", None)?)
        .item(&item(app, "delete.sheet", "Delete…", None)?)
        .build()?;
    let format = SubmenuBuilder::new(app, "Format")
        .item(&item(app, "format.bold", "Bold", Some("CmdOrCtrl+B"))?)
        .item(&item(app, "format.italic", "Italic", Some("CmdOrCtrl+I"))?)
        .item(&item(app, "format.underline", "Underline", Some("CmdOrCtrl+U"))?)
        .item(&item(app, "format.strikethrough", "Strikethrough", Some(strike_keys))?)
        .separator()
        .item(&item(app, "format.increaseFont", "Bigger Font", None)?)
        .item(&item(app, "format.decreaseFont", "Smaller Font", None)?)
        .separator()
        .item(&item(app, "format.alignLeft", "Align Left", None)?)
        .item(&item(app, "format.alignCenter", "Center", None)?)
        .item(&item(app, "format.alignRight", "Align Right", None)?)
        .item(&item(app, "format.wrap", "Wrap Text", None)?)
        .separator()
        .item(&item(app, "format.mergeCenter", "Merge & Center", None)?)
        .item(&item(app, "format.unmerge", "Unmerge Cells", None)?)
        .separator()
        .item(&number)
        .item(&rows)
        .item(&columns)
        .item(&sheet)
        .separator()
        .item(&item(app, "format.clear", "Clear Formatting", Some("CmdOrCtrl+Backslash"))?)
        .build()?;

    let data = SubmenuBuilder::new(app, "Data")
        .item(&item(app, "data.sortAscending", "Sort A to Z", None)?)
        .item(&item(app, "data.sortDescending", "Sort Z to A", None)?)
        .build()?;

    let freeze = SubmenuBuilder::new(app, "Freeze Panes")
        .item(&item(app, "view.freezeTopRow", "Freeze Top Row", None)?)
        .item(&item(app, "view.freezeFirstColumn", "Freeze First Column", None)?)
        .item(&item(app, "view.freezeAtSelection", "Freeze at Selected Cell", None)?)
        .separator()
        .item(&item(app, "view.unfreeze", "Unfreeze Panes", None)?)
        .build()?;

    let view = SubmenuBuilder::new(app, "View")
        .item(&item(app, "view.zoomIn", "Zoom In", Some("CmdOrCtrl+="))?)
        .item(&item(app, "view.zoomOut", "Zoom Out", Some("CmdOrCtrl+-"))?)
        .item(&item(app, "view.zoomReset", "Actual Size", Some("CmdOrCtrl+0"))?)
        .separator()
        .item(&freeze)
        .item(&item(app, "view.toggleGridlines", "Show/Hide Gridlines", None)?)
        .separator()
        .item(&item(app, "view.nextSheet", "Next Sheet", Some("Ctrl+PageDown"))?)
        .item(&item(app, "view.previousSheet", "Previous Sheet", Some("Ctrl+PageUp"))?)
        .separator()
        .item(&item(app, "view.toggleInfo", "Workbook Information", Some("CmdOrCtrl+Alt+I"))?);
    #[cfg(target_os = "macos")]
    let view = view.separator().fullscreen();
    let view = view.build()?;

    let window = SubmenuBuilder::new(app, "Window").minimize().maximize().build()?;

    let menu = MenuBuilder::new(app);
    #[cfg(target_os = "macos")]
    let menu = {
        let app_menu = SubmenuBuilder::new(app, app.package_info().name.clone())
            .about(None)
            .separator()
            .services()
            .separator()
            .hide()
            .hide_others()
            .show_all()
            .separator()
            .item(&item(app, "app.quit", &format!("Quit {}", app.package_info().name), Some("Cmd+Q"))?)
            .build()?;
        menu.item(&app_menu)
    };
    let menu = menu
        .item(&file)
        .item(&edit)
        .item(&insert)
        .item(&format)
        .item(&data)
        .item(&view)
        .item(&window)
        .build()?;
    Ok((menu, recent))
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let state = AppState::default();
    state.pending.lock().unwrap().paths = startup_paths();

    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_clipboard_manager::init())
        .manage(state)
        .setup(|app| {
            let handle = app.handle();
            let (menu, recent) = build_menu(handle)?;
            fill_recent_menu(handle, &recent, &[])?;
            handle.set_menu(menu)?;
            *app.state::<AppState>().recent_menu.lock().unwrap() = Some(recent);

            // The window starts hidden to avoid a blank flash; the frontend shows it
            // after its first render. Show it anyway if that never happens.
            if let Some(window) = app.get_webview_window("main") {
                std::thread::spawn(move || {
                    std::thread::sleep(Duration::from_secs(3));
                    let _ = window.show();
                });
            }
            Ok(())
        })
        .on_menu_event(|app, event| {
            let id = event.id().as_ref();
            if let Some(index) = id.strip_prefix("recent:") {
                let path = index
                    .parse::<usize>()
                    .ok()
                    .and_then(|i| app.state::<AppState>().recent.lock().unwrap().get(i).cloned());
                if let Some(path) = path {
                    let _ = app.emit("open-files", vec![path]);
                }
            } else {
                let _ = app.emit("menu-command", id);
            }
        })
        .on_window_event(|window, event| {
            // Closing the window with unsaved changes: let the frontend ask Save / Don't Save / Cancel.
            if let WindowEvent::CloseRequested { api, .. } = event {
                if should_confirm_close(window.app_handle()) {
                    api.prevent_close();
                    let _ = window.emit("close-requested", ());
                }
            }
        })
        .invoke_handler(tauri::generate_handler![
            read_workbook_file,
            write_workbook_file,
            take_startup_files,
            set_recent_files,
            set_document_edited,
            quit_app
        ])
        .build(tauri::generate_context!())
        .expect("error while starting Sheaf")
        .run(|app, event| {
            // Files opened from Finder ("Open With", double-click, drag onto the Dock icon).
            #[cfg(any(target_os = "macos", target_os = "ios"))]
            if let tauri::RunEvent::Opened { urls } = &event {
                let paths = urls
                    .iter()
                    .filter_map(|url| url.to_file_path().ok())
                    .map(|path| path.to_string_lossy().into_owned())
                    .collect();
                open_paths(app, paths);
            }
            // Quitting some other way (e.g. from the Dock) with unsaved changes.
            if let tauri::RunEvent::ExitRequested { api, code, .. } = &event {
                if code.is_none() && should_confirm_close(app) {
                    api.prevent_exit();
                    let _ = app.emit("close-requested", ());
                }
            }
            let _ = (app, event);
        });
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decodes_percent_encoded_paths() {
        assert_eq!(
            decode_percent("%2FUsers%2Fme%2FQ1%20GST%20%E2%82%B9.xlsx").as_deref(),
            Some("/Users/me/Q1 GST ₹.xlsx")
        );
        assert_eq!(decode_percent("C%3A%5CData%5Cbook.csv").as_deref(), Some("C:\\Data\\book.csv"));
        assert_eq!(decode_percent("%zz"), None);
        assert_eq!(decode_percent("%2"), Some("%2".to_string()));
    }

    #[test]
    fn writes_atomically_replacing_the_old_file() {
        let dir = std::env::temp_dir().join(format!("sheaf-test-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("book.xlsx");
        std::fs::write(&path, b"old contents").unwrap();

        write_atomically(&path, b"new contents").unwrap();

        assert_eq!(std::fs::read(&path).unwrap(), b"new contents");
        let leftovers: Vec<_> = std::fs::read_dir(&dir).unwrap().filter_map(Result::ok).map(|e| e.file_name()).collect();
        assert_eq!(leftovers.len(), 1, "temporary file was not cleaned up: {leftovers:?}");
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn refuses_to_write_into_a_missing_folder() {
        let path = std::env::temp_dir().join("sheaf-missing-folder").join("book.xlsx");
        assert!(write_atomically(&path, b"x").is_err());
    }
}
