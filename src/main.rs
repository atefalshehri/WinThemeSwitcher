#![windows_subsystem = "windows"]

use std::error::Error;
use std::ffi::c_void;
use std::fs;
use std::path::{Path, PathBuf};
use std::ptr;
use std::sync::OnceLock;
use std::time::{Duration, Instant};

use chrono::{DateTime, Local, NaiveDate, TimeZone};
use serde::{Deserialize, Serialize};
use sun_times::sun_times;
use tray_icon::{
    menu::{Menu, MenuEvent, MenuId, MenuItem, PredefinedMenuItem},
    TrayIconBuilder,
};
use windows::Devices::Geolocation::{GeolocationAccessStatus, Geolocator};
use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
use windows_sys::core::{GUID, HRESULT};
use windows_sys::Win32::Foundation::{SysFreeString, HWND, LPARAM, LRESULT, WPARAM};
use windows_sys::Win32::Graphics::Dwm::DwmFlush;
use windows_sys::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows_sys::Win32::System::LibraryLoader::GetModuleHandleW;
use windows_sys::Win32::System::Power::PowerRegisterSuspendResumeNotification;
use windows_sys::Win32::System::RemoteDesktop::{
    WTSRegisterSessionNotification, NOTIFY_FOR_THIS_SESSION,
};
use windows_sys::Win32::System::Registry::{
    RegCloseKey, RegDeleteValueW, RegOpenKeyExW, RegQueryValueExW, RegSetValueExW, HKEY,
    HKEY_CURRENT_USER, KEY_QUERY_VALUE, KEY_SET_VALUE, REG_DWORD, REG_SZ,
};
use windows_sys::Win32::UI::Shell::{SHLoadIndirectString, ShellExecuteW};
use windows_sys::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DispatchMessageW, FindWindowW, GetMessageW, MessageBoxW,
    PostMessageW, RegisterClassW, SendMessageTimeoutW, TranslateMessage, HWND_BROADCAST,
    HWND_MESSAGE, IDYES, MB_ICONINFORMATION, MB_ICONQUESTION, MB_OK, MB_YESNO, MSG,
    SMTO_ABORTIFHUNG, SW_HIDE, SW_SHOWNORMAL, WM_CLOSE, WM_POWERBROADCAST, WM_SETTINGCHANGE,
    WM_THEMECHANGED, WM_WTSSESSION_CHANGE, WNDCLASSW,
};
use winit::event::{Event, StartCause};
use winit::event_loop::{ActiveEventLoop, ControlFlow, EventLoop, EventLoopProxy};

const PBT_APMRESUMEAUTOMATIC: WPARAM = 0x12;
const WTS_SESSION_UNLOCK: WPARAM = 0x8;
const DEVICE_NOTIFY_WINDOW_HANDLE: u32 = 0x0;

const THEME_KEY: &str = "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";
const RUN_KEY: &str = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
const APP_NAME: &str = "WinThemeSwitcher";

// === IThemeManager2 (themeui.dll, undocumented but stable since Win10 1809) ===
//
// The Settings UWP wraps this same interface. AutoDarkMode and similar tools use
// it as the canonical theme-apply path. Going through this instead of
// ShellExecuteW(.theme) avoids the silent-fail problem we hit with the UWP
// activation pipeline (post-unlock / scheduled-while-away contexts), AND removes
// most of the heuristic AV signals (no HWND_BROADCAST WM_SETTINGCHANGE, no
// direct WM_THEMECHANGED to Shell_TrayWnd — SetCurrentTheme does the broadcast
// itself from inside themeui.dll, where it's expected by AV behavior models).
//
// References:
//   - https://gist.github.com/namazso/0fde102c2fc56049c7c37f7fdf9ac3cd (C#)
//   - https://github.com/HenriquedoVal/wtheme/blob/main/ThemeManager2.h (C)
//   - https://github.com/AutoDarkMode/Windows-Auto-Night-Mode/blob/master/AutoDarkModeSvc/Handlers/IThemeManager2/Tm2Handler.cs

const CLSID_THEME_MANAGER2: GUID = GUID {
    data1: 0x9324da94,
    data2: 0x50ec,
    data3: 0x4a14,
    data4: [0xa7, 0x70, 0xe9, 0x0c, 0xa0, 0x3e, 0x7c, 0x8f],
};

const IID_THEME_MANAGER2: GUID = GUID {
    data1: 0xc1e8c83e,
    data2: 0x845d,
    data3: 0x4d95,
    data4: [0x81, 0xdb, 0xe2, 0x83, 0xfd, 0xff, 0xc0, 0x00],
};

const THEME_INIT_NO_FLAGS: i32 = 0;
// THEME_APPLY_FLAGS bitmask. 0 = apply everything (matches Settings UWP). NO_HOURGLASS
// suppresses the wait cursor for unattended apply.
const THEME_APPLY_FLAG_NO_HOURGLASS: i32 = 1 << 8;

#[repr(C)]
struct IThemeManager2Vtbl {
    // IUnknown
    query_interface: unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> HRESULT,
    add_ref: unsafe extern "system" fn(*mut c_void) -> u32,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
    // IThemeManager2 — only the slots we actually call. ORDER MATTERS — must
    // match vtable layout exactly. Reference: namazso C# gist + wtheme C header.
    init: unsafe extern "system" fn(*mut c_void, i32) -> HRESULT,
    _init_async: unsafe extern "system" fn(*mut c_void, HWND, i32) -> HRESULT,
    _refresh: unsafe extern "system" fn(*mut c_void) -> HRESULT,
    _refresh_async: unsafe extern "system" fn(*mut c_void, HWND, i32) -> HRESULT,
    _refresh_complete: unsafe extern "system" fn(*mut c_void) -> HRESULT,
    get_theme_count: unsafe extern "system" fn(*mut c_void, *mut i32) -> HRESULT,
    get_theme: unsafe extern "system" fn(*mut c_void, i32, *mut *mut c_void) -> HRESULT,
    _is_theme_disabled: unsafe extern "system" fn(*mut c_void, i32, *mut i32) -> HRESULT,
    _get_current_theme: unsafe extern "system" fn(*mut c_void, *mut i32) -> HRESULT,
    set_current_theme: unsafe extern "system" fn(*mut c_void, HWND, i32, i32, i32, i32) -> HRESULT,
    // Remaining slots (GetCustomTheme, GetDefaultTheme, CreateThemePack, ...) omitted —
    // not called from this app. The struct only needs to expose what we call;
    // unused trailing slots don't affect ABI.
}

#[repr(C)]
struct IThemeVtbl {
    query_interface: unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> HRESULT,
    add_ref: unsafe extern "system" fn(*mut c_void) -> u32,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
    // GetDisplayName returns a BSTR (allocated with SysAllocString — release with SysFreeString).
    get_display_name: unsafe extern "system" fn(*mut c_void, *mut *mut u16) -> HRESULT,
    // PutDisplayName + later methods omitted — wtheme header notes they vary across
    // Windows versions and aren't safe to call.
}

/// RAII wrapper around an IThemeManager2 COM pointer. Calls Release on drop.
struct ThemeMgr {
    ptr: *mut c_void,
}

impl ThemeMgr {
    /// CoCreateInstance + Init. Caller must already be on an STA thread
    /// (we are — main thread does CoInitializeEx(APARTMENTTHREADED) at startup).
    unsafe fn create() -> Result<Self, HRESULT> {
        let mut ptr: *mut c_void = ptr::null_mut();
        let hr = CoCreateInstance(
            &CLSID_THEME_MANAGER2,
            ptr::null_mut(),
            CLSCTX_INPROC_SERVER,
            &IID_THEME_MANAGER2,
            &mut ptr,
        );
        if hr < 0 || ptr.is_null() {
            return Err(hr);
        }
        let vtbl = Self::vtbl_of(ptr);
        let hr = (vtbl.init)(ptr, THEME_INIT_NO_FLAGS);
        if hr < 0 {
            (vtbl.release)(ptr);
            return Err(hr);
        }
        Ok(Self { ptr })
    }

    unsafe fn vtbl_of(ptr: *mut c_void) -> &'static IThemeManager2Vtbl {
        &**(ptr as *const *const IThemeManager2Vtbl)
    }

    unsafe fn vtbl(&self) -> &IThemeManager2Vtbl {
        Self::vtbl_of(self.ptr)
    }

    unsafe fn count(&self) -> Result<i32, HRESULT> {
        let mut n = 0i32;
        let hr = (self.vtbl().get_theme_count)(self.ptr, &mut n);
        if hr < 0 {
            return Err(hr);
        }
        Ok(n)
    }

    /// Returns the display name of the theme at `index`, or an HRESULT error.
    /// Note: enumeration order is not stable across launches — re-enumerate every
    /// apply rather than caching indices.
    unsafe fn theme_display_name(&self, index: i32) -> Result<String, HRESULT> {
        let mut theme_ptr: *mut c_void = ptr::null_mut();
        let hr = (self.vtbl().get_theme)(self.ptr, index, &mut theme_ptr);
        if hr < 0 || theme_ptr.is_null() {
            return Err(hr);
        }
        let theme_vtbl = &**(theme_ptr as *const *const IThemeVtbl);
        let mut bstr: *mut u16 = ptr::null_mut();
        let hr = (theme_vtbl.get_display_name)(theme_ptr, &mut bstr);
        if hr < 0 || bstr.is_null() {
            (theme_vtbl.release)(theme_ptr);
            return Err(hr);
        }
        let name = read_wide_string(bstr);
        SysFreeString(bstr);
        (theme_vtbl.release)(theme_ptr);
        Ok(name)
    }

    /// Apply the theme at `index`. `apply_now=1` makes it take effect immediately
    /// (registry write + WM_THEMECHANGED + WM_SETTINGCHANGE broadcast all happen
    /// inside SetCurrentTheme). `pack_flags=0` matches Settings UWP defaults.
    unsafe fn set_current(&self, index: i32, apply_flags: i32) -> Result<(), HRESULT> {
        let hr = (self.vtbl().set_current_theme)(
            self.ptr,
            ptr::null_mut(),
            index,
            1,
            apply_flags,
            0,
        );
        if hr < 0 {
            return Err(hr);
        }
        Ok(())
    }
}

impl Drop for ThemeMgr {
    fn drop(&mut self) {
        unsafe { (self.vtbl().release)(self.ptr) };
    }
}

unsafe fn read_wide_string(p: *const u16) -> String {
    if p.is_null() {
        return String::new();
    }
    let mut len = 0usize;
    while *p.add(len) != 0 {
        len += 1;
    }
    String::from_utf16_lossy(std::slice::from_raw_parts(p, len))
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(default)]
struct Config {
    latitude: f64,
    longitude: f64,
    auto_start: bool,
    theme_day: Option<String>,
    theme_night: Option<String>,
}

impl Default for Config {
    fn default() -> Self {
        Self {
            latitude: 0.0,
            longitude: 0.0,
            auto_start: true,
            theme_day: None,
            theme_night: None,
        }
    }
}

impl Config {
    fn has_location(&self) -> bool {
        !(self.latitude == 0.0 && self.longitude == 0.0)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Theme {
    Light,
    Dark,
}

#[derive(Debug, Clone)]
enum AppEvent {
    Menu(MenuId),
    Wake(WakeKind),
}

#[derive(Debug, Clone, Copy)]
enum WakeKind {
    Unlock,
    Power,
}

static EVENT_PROXY: OnceLock<EventLoopProxy<AppEvent>> = OnceLock::new();

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

fn config_path() -> PathBuf {
    std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|d| d.to_path_buf()))
        .unwrap_or_else(|| PathBuf::from("."))
        .join("config.json")
}

fn log_path() -> PathBuf {
    std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|d| d.to_path_buf()))
        .unwrap_or_else(|| PathBuf::from("."))
        .join("events.log")
}

fn log_event(line: &str) {
    use std::io::Write;
    let path = log_path();
    if let Ok(meta) = fs::metadata(&path) {
        if meta.len() > 256 * 1024 {
            let _ = fs::rename(&path, path.with_extension("log.old"));
        }
    }
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
    {
        let _ = writeln!(f, "{}", line);
    }
}

fn theme_str(t: Option<Theme>) -> &'static str {
    match t {
        Some(Theme::Light) => "light",
        Some(Theme::Dark) => "dark",
        None => "unknown",
    }
}

fn load_or_create_config() -> Config {
    let path = config_path();
    if let Ok(content) = fs::read_to_string(&path) {
        if let Ok(cfg) = serde_json::from_str::<Config>(&content) {
            return cfg;
        }
    }
    let cfg = Config::default();
    if let Ok(json) = serde_json::to_string_pretty(&cfg) {
        let _ = fs::write(&path, json);
    }
    cfg
}

fn save_config(cfg: &Config) -> Result<(), Box<dyn Error>> {
    let json = serde_json::to_string_pretty(cfg)?;
    fs::write(config_path(), json)?;
    Ok(())
}

fn ensure_com_initialized() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }
}

fn try_get_windows_location() -> Option<(f64, f64)> {
    let access = Geolocator::RequestAccessAsync().ok()?.get().ok()?;
    if access != GeolocationAccessStatus::Allowed {
        return None;
    }
    let geo = Geolocator::new().ok()?;
    let pos = geo.GetGeopositionAsync().ok()?.get().ok()?;
    let coord = pos.Coordinate().ok()?;
    let point = coord.Point().ok()?;
    let p = point.Position().ok()?;
    Some((p.Latitude, p.Longitude))
}

fn open_config_in_editor() {
    let path = config_path();
    let path_w = wide(&path.to_string_lossy());
    let verb = wide("open");
    unsafe {
        ShellExecuteW(
            ptr::null_mut(),
            verb.as_ptr(),
            path_w.as_ptr(),
            ptr::null(),
            ptr::null(),
            SW_SHOWNORMAL,
        );
    }
}

fn open_location_settings() {
    let verb = wide("open");
    let uri = wide("ms-settings:privacy-location");
    unsafe {
        ShellExecuteW(
            ptr::null_mut(),
            verb.as_ptr(),
            uri.as_ptr(),
            ptr::null(),
            ptr::null(),
            SW_SHOWNORMAL,
        );
    }
}

fn ask_enable_location() -> bool {
    let title = wide("WinThemeSwitcher — Location");
    let body = wide(
        "Windows Location is off or not allowed for desktop apps.\n\n\
         Enable it so sunrise and sunset can be computed automatically?\n\n\
         Yes opens Windows Settings. No lets you enter coordinates manually in config.json.",
    );
    unsafe {
        let r = MessageBoxW(
            ptr::null_mut(),
            body.as_ptr(),
            title.as_ptr(),
            MB_YESNO | MB_ICONQUESTION,
        );
        r == IDYES as i32
    }
}

fn show_enable_pending_message() {
    let title = wide("WinThemeSwitcher");
    let body = wide(
        "Turn on \"Location services\" in the Settings window that just opened. \
         Then right-click the WinThemeSwitcher tray icon and choose Refresh.",
    );
    unsafe {
        MessageBoxW(
            ptr::null_mut(),
            body.as_ptr(),
            title.as_ptr(),
            MB_OK | MB_ICONINFORMATION,
        );
    }
}

fn show_manual_setup_prompt() {
    let title = wide("WinThemeSwitcher — Setup");
    let body = wide(
        "Please set latitude and longitude in config.json (opening now), \
         then right-click the tray icon and choose Refresh.",
    );
    unsafe {
        MessageBoxW(
            ptr::null_mut(),
            body.as_ptr(),
            title.as_ptr(),
            MB_OK | MB_ICONINFORMATION,
        );
    }
    open_config_in_editor();
}

fn acquire_location(cfg: &mut Config) {
    if let Some((lat, lon)) = try_get_windows_location() {
        cfg.latitude = lat;
        cfg.longitude = lon;
        let _ = save_config(cfg);
        return;
    }
    if ask_enable_location() {
        open_location_settings();
        show_enable_pending_message();
    } else {
        show_manual_setup_prompt();
    }
}

fn current_theme() -> Option<Theme> {
    let subkey = wide(THEME_KEY);
    let value = wide("SystemUsesLightTheme");
    unsafe {
        let mut hkey: HKEY = ptr::null_mut();
        if RegOpenKeyExW(HKEY_CURRENT_USER, subkey.as_ptr(), 0, KEY_QUERY_VALUE, &mut hkey) != 0 {
            return None;
        }
        let mut data: u32 = 0;
        let mut size: u32 = 4;
        let mut kind: u32 = 0;
        let r = RegQueryValueExW(
            hkey,
            value.as_ptr(),
            ptr::null_mut(),
            &mut kind,
            &mut data as *mut u32 as *mut u8,
            &mut size,
        );
        RegCloseKey(hkey);
        if r != 0 {
            return None;
        }
        Some(if data == 0 { Theme::Dark } else { Theme::Light })
    }
}

fn write_theme_registry(theme: Theme) -> Result<(), Box<dyn Error>> {
    let value: u32 = if theme == Theme::Light { 1 } else { 0 };
    let subkey = wide(THEME_KEY);
    let apps = wide("AppsUseLightTheme");
    let sys = wide("SystemUsesLightTheme");
    unsafe {
        let mut hkey: HKEY = ptr::null_mut();
        if RegOpenKeyExW(HKEY_CURRENT_USER, subkey.as_ptr(), 0, KEY_SET_VALUE, &mut hkey) != 0 {
            return Err("RegOpenKeyExW failed for Personalize".into());
        }
        RegSetValueExW(
            hkey,
            apps.as_ptr(),
            0,
            REG_DWORD,
            &value as *const u32 as *const u8,
            4,
        );
        RegSetValueExW(
            hkey,
            sys.as_ptr(),
            0,
            REG_DWORD,
            &value as *const u32 as *const u8,
            4,
        );
        RegCloseKey(hkey);
    }
    Ok(())
}

fn broadcast_setting_change() {
    let param = wide("ImmersiveColorSet");
    let mut result: usize = 0;
    unsafe {
        SendMessageTimeoutW(
            HWND_BROADCAST,
            WM_SETTINGCHANGE,
            0,
            param.as_ptr() as isize,
            SMTO_ABORTIFHUNG,
            500,
            &mut result,
        );
    }
}

fn poke_shell() {
    let param = wide("ImmersiveColorSet");
    for class in ["Shell_TrayWnd", "Shell_SecondaryTrayWnd"] {
        let cls = wide(class);
        unsafe {
            let hwnd = FindWindowW(cls.as_ptr(), ptr::null());
            if (hwnd as usize) == 0 {
                continue;
            }
            let mut result: usize = 0;
            SendMessageTimeoutW(
                hwnd,
                WM_THEMECHANGED,
                0,
                0,
                SMTO_ABORTIFHUNG,
                500,
                &mut result,
            );
            SendMessageTimeoutW(
                hwnd,
                WM_SETTINGCHANGE,
                0,
                param.as_ptr() as isize,
                SMTO_ABORTIFHUNG,
                500,
                &mut result,
            );
        }
    }
    unsafe {
        DwmFlush();
    }
}

fn resolve_theme_file(theme: Theme, cfg: &Config) -> PathBuf {
    let custom = match theme {
        Theme::Light => cfg.theme_day.as_deref(),
        Theme::Dark => cfg.theme_night.as_deref(),
    };
    if let Some(p) = custom {
        let path = PathBuf::from(p);
        if path.exists() {
            return path;
        }
    }
    let win_dir = std::env::var("SystemRoot").unwrap_or_else(|_| "C:\\Windows".to_string());
    let leaf = match theme {
        Theme::Light => "aero.theme",
        Theme::Dark => "dark.theme",
    };
    PathBuf::from(win_dir).join("Resources").join("Themes").join(leaf)
}

fn apply_theme_file(path: &std::path::Path) -> bool {
    let s = path.to_string_lossy();
    let path_w = wide(&s);
    let verb = wide("open");
    unsafe {
        let h = ShellExecuteW(
            ptr::null_mut(),
            verb.as_ptr(),
            path_w.as_ptr(),
            ptr::null(),
            ptr::null(),
            SW_HIDE,
        );
        (h as isize) > 32
    }
}

fn start_settings_closer() {
    let started = Instant::now();
    std::thread::spawn(move || {
        let class = wide("ApplicationFrameWindow");
        let titles = [
            ("Settings", wide("Settings")),
            ("Themes", wide("Themes")),
            ("Personalization", wide("Personalization")),
        ];
        let mut first_close_logged = false;
        for _ in 0..14 {
            std::thread::sleep(Duration::from_millis(150));
            for (name, title_w) in &titles {
                unsafe {
                    let hwnd = FindWindowW(class.as_ptr(), title_w.as_ptr());
                    if (hwnd as usize) != 0 {
                        PostMessageW(hwnd, WM_CLOSE, 0, 0);
                        if !first_close_logged {
                            log_event(&format!(
                                "{} settings_closed title={} after_ms={}",
                                Local::now().to_rfc3339(),
                                name,
                                started.elapsed().as_millis(),
                            ));
                            first_close_logged = true;
                        }
                    }
                }
            }
        }
    });
}

fn start_commit_watcher(target: Theme) {
    let started = Instant::now();
    std::thread::spawn(move || {
        for _ in 0..25 {
            std::thread::sleep(Duration::from_millis(200));
            if current_theme() == Some(target) {
                log_event(&format!(
                    "{} commit_observed target={} after_ms={}",
                    Local::now().to_rfc3339(),
                    theme_str(Some(target)),
                    started.elapsed().as_millis(),
                ));
                return;
            }
        }
        log_event(&format!(
            "{} commit_timeout target={} actual={} after_ms={}",
            Local::now().to_rfc3339(),
            theme_str(Some(target)),
            theme_str(current_theme()),
            started.elapsed().as_millis(),
        ));
        // ShellExecute(.theme) lied about success — Settings UWP didn't actually apply.
        // Observed when the schedule fires while the user isn't interactive (sunset while
        // away, immediately after WTS_SESSION_UNLOCK, immediately after PBT_APMRESUMEAUTOMATIC).
        // Force the mode flip via direct registry write so at minimum light/dark is correct;
        // wallpaper won't change on this path (would require IThemeManager2 — see CLAUDE.md).
        let fb_started = Instant::now();
        match write_theme_registry(target) {
            Ok(()) => {
                broadcast_setting_change();
                poke_shell();
                let mut confirmed = false;
                for _ in 0..10 {
                    std::thread::sleep(Duration::from_millis(100));
                    if current_theme() == Some(target) {
                        confirmed = true;
                        break;
                    }
                }
                log_event(&format!(
                    "{} fallback_registry target={} confirmed={} after_ms={}",
                    Local::now().to_rfc3339(),
                    theme_str(Some(target)),
                    confirmed,
                    fb_started.elapsed().as_millis(),
                ));
            }
            Err(e) => {
                log_event(&format!(
                    "{} fallback_registry_err target={} err=\"{}\"",
                    Local::now().to_rfc3339(),
                    theme_str(Some(target)),
                    e,
                ));
            }
        }
    });
}

/// Reads `[Theme]\nDisplayName=...` from a `.theme` (INI) file.
/// `DisplayName` may be a literal string, OR an SHLoadIndirectString resource
/// reference of the form `@%SystemRoot%\System32\themeui.dll,-2060` (system themes
/// use this — the actual user-visible name is in a localized string table).
/// Returns the resolved literal string, or None if the file is unreadable / has no
/// DisplayName / the indirect-string resolution fails.
fn resolve_theme_display_name(theme_file: &Path) -> Option<String> {
    let content = fs::read_to_string(theme_file).ok()?;
    let mut in_theme_section = false;
    for raw_line in content.lines() {
        let line = raw_line.trim();
        if line.starts_with(';') || line.is_empty() {
            continue;
        }
        if line.starts_with('[') && line.ends_with(']') {
            in_theme_section = line.eq_ignore_ascii_case("[Theme]");
            continue;
        }
        if !in_theme_section {
            continue;
        }
        if let Some(rest) = line.strip_prefix("DisplayName=") {
            let raw = rest.trim();
            if raw.starts_with('@') {
                return resolve_indirect_string(raw);
            }
            return Some(raw.to_string());
        }
    }
    None
}

/// Resolves `@dll,-id` resource string references using SHLoadIndirectString.
fn resolve_indirect_string(source: &str) -> Option<String> {
    let src_w = wide(source);
    let mut buf = [0u16; 512];
    let hr = unsafe {
        SHLoadIndirectString(
            src_w.as_ptr(),
            buf.as_mut_ptr(),
            buf.len() as u32,
            ptr::null_mut(),
        )
    };
    if hr < 0 {
        return None;
    }
    let len = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
    if len == 0 {
        return None;
    }
    Some(String::from_utf16_lossy(&buf[..len]))
}

/// Apply a `.theme` file via the IThemeManager2 COM interface — the same API
/// the Settings UWP itself wraps. Reliable from any context (post-unlock,
/// scheduled-while-away, background tick) — unlike the ShellExecuteW(.theme)
/// path which silently fails when the user isn't actively interactive.
///
/// The interface enumerates installed themes by index (no by-path lookup), so
/// we resolve the target's DisplayName from the .theme file and match it
/// against `ITheme::GetDisplayName` in the enumeration. System themes
/// (aero.theme, dark.theme) are always present after `Init`. Custom user
/// themes need to have been installed first (e.g. via Settings → Themes, or
/// AddAndSelectTheme — not implemented here; custom themes fall through to
/// the legacy path).
///
/// Logs `theme_manager2_apply` on success; the caller logs the err string.
fn apply_via_theme_manager2(theme: Theme, theme_file: &Path) -> Result<(), Box<dyn Error>> {
    let target_name = resolve_theme_display_name(theme_file)
        .ok_or("could not resolve DisplayName from .theme file")?;
    let started = Instant::now();
    unsafe {
        let mgr =
            ThemeMgr::create().map_err(|hr| format!("CoCreateInstance/Init hr=0x{:08x}", hr))?;
        let n = mgr
            .count()
            .map_err(|hr| format!("GetThemeCount hr=0x{:08x}", hr))?;
        for i in 0..n {
            let name = match mgr.theme_display_name(i) {
                Ok(n) => n,
                Err(hr) => {
                    log_event(&format!(
                        "{} theme_manager2_enum_skip i={} hr=0x{:08x}",
                        Local::now().to_rfc3339(),
                        i,
                        hr
                    ));
                    continue;
                }
            };
            if name == target_name {
                mgr.set_current(i, THEME_APPLY_FLAG_NO_HOURGLASS)
                    .map_err(|hr| format!("SetCurrentTheme i={} hr=0x{:08x}", i, hr))?;
                log_event(&format!(
                    "{} theme_manager2_apply target={} display=\"{}\" idx={} after_ms={}",
                    Local::now().to_rfc3339(),
                    theme_str(Some(theme)),
                    name,
                    i,
                    started.elapsed().as_millis(),
                ));
                return Ok(());
            }
        }
        Err(format!("no installed theme matches DisplayName \"{}\"", target_name).into())
    }
}

/// Three-tier apply, best-to-worst:
///   1. IThemeManager2  — atomic, reliable, no Settings UWP, no AV-tripping broadcast.
///   2. ShellExecuteW(.theme) + commit_watcher — legacy. Watcher promotes to (3) on silent fail.
///   3. Direct registry write — flips light/dark mode but not wallpaper. Last resort.
fn apply_theme(theme: Theme, cfg: &Config) -> Result<&'static str, Box<dyn Error>> {
    let theme_file = resolve_theme_file(theme, cfg);

    if theme_file.exists() {
        match apply_via_theme_manager2(theme, &theme_file) {
            Ok(()) => return Ok("theme-manager2"),
            Err(e) => log_event(&format!(
                "{} theme_manager2_err target={} msg=\"{}\"",
                Local::now().to_rfc3339(),
                theme_str(Some(theme)),
                e,
            )),
        }
    }

    if theme_file.exists() && apply_theme_file(&theme_file) {
        start_commit_watcher(theme);
        start_settings_closer();
        std::thread::sleep(Duration::from_millis(300));
        poke_shell();
        return Ok("theme-file");
    }

    write_theme_registry(theme)?;
    broadcast_setting_change();
    poke_shell();
    Ok("registry")
}

fn set_auto_start(enable: bool) -> Result<(), Box<dyn Error>> {
    let subkey = wide(RUN_KEY);
    let name = wide(APP_NAME);
    unsafe {
        let mut hkey: HKEY = ptr::null_mut();
        if RegOpenKeyExW(HKEY_CURRENT_USER, subkey.as_ptr(), 0, KEY_SET_VALUE, &mut hkey) != 0 {
            return Err("RegOpenKeyExW failed for Run".into());
        }
        if enable {
            let exe = std::env::current_exe()?;
            let exe_w = wide(&format!("\"{}\"", exe.to_string_lossy()));
            RegSetValueExW(
                hkey,
                name.as_ptr(),
                0,
                REG_SZ,
                exe_w.as_ptr() as *const u8,
                (exe_w.len() * 2) as u32,
            );
        } else {
            RegDeleteValueW(hkey, name.as_ptr());
        }
        RegCloseKey(hkey);
    }
    Ok(())
}

fn sun_times_for(date: NaiveDate, lat: f64, lon: f64) -> (DateTime<Local>, DateTime<Local>) {
    match sun_times(date, lat, lon, 0.0) {
        Some((sr, ss)) => (sr.with_timezone(&Local), ss.with_timezone(&Local)),
        None => {
            let sr = Local
                .from_local_datetime(&date.and_hms_opt(6, 0, 0).unwrap())
                .single()
                .unwrap();
            let ss = Local
                .from_local_datetime(&date.and_hms_opt(18, 0, 0).unwrap())
                .single()
                .unwrap();
            (sr, ss)
        }
    }
}

fn target_theme(now: DateTime<Local>, lat: f64, lon: f64) -> Theme {
    let (sr, ss) = sun_times_for(now.date_naive(), lat, lon);
    if now >= sr && now < ss {
        Theme::Light
    } else {
        Theme::Dark
    }
}

fn next_transition(now: DateTime<Local>, lat: f64, lon: f64) -> DateTime<Local> {
    let today = now.date_naive();
    let (sr, ss) = sun_times_for(today, lat, lon);
    if now < sr {
        sr
    } else if now < ss {
        ss
    } else {
        let tomorrow = today.succ_opt().expect("date overflow");
        let (tom_sr, _) = sun_times_for(tomorrow, lat, lon);
        tom_sr
    }
}

fn deadline_instant(target: DateTime<Local>) -> Instant {
    let now = Local::now();
    let delta = (target - now).to_std().unwrap_or(Duration::from_secs(1));
    Instant::now() + delta
}

fn make_tray_icon() -> Option<tray_icon::Icon> {
    const SIZE: u32 = 32;
    let mut rgba = vec![0u8; (SIZE * SIZE * 4) as usize];
    let cx = SIZE as f32 / 2.0;
    let cy = SIZE as f32 / 2.0;
    let r = SIZE as f32 / 2.0 - 1.5;
    for y in 0..SIZE {
        for x in 0..SIZE {
            let dx = x as f32 + 0.5 - cx;
            let dy = y as f32 + 0.5 - cy;
            let d = (dx * dx + dy * dy).sqrt();
            let idx = ((y * SIZE + x) * 4) as usize;
            if d <= r {
                if dx < 0.0 {
                    rgba[idx] = 255;
                    rgba[idx + 1] = 140;
                    rgba[idx + 2] = 0;
                } else {
                    rgba[idx] = 44;
                    rgba[idx + 1] = 62;
                    rgba[idx + 2] = 100;
                }
                rgba[idx + 3] = 255;
            }
        }
    }
    tray_icon::Icon::from_rgba(rgba, SIZE, SIZE).ok()
}

fn tick(cfg: &Config, elwt: &ActiveEventLoop, cause: &str, force: bool) {
    let now = Local::now();
    let now_str = now.to_rfc3339();

    if !cfg.has_location() {
        log_event(&format!("{} cause={} skipped=no-location", now_str, cause));
        elwt.set_control_flow(ControlFlow::Wait);
        return;
    }

    let want = target_theme(now, cfg.latitude, cfg.longitude);
    let current = current_theme();

    let outcome = if force || current != Some(want) {
        match apply_theme(want, cfg) {
            Ok(method) => format!("applied={}", method),
            Err(e) => format!("err=\"{}\"", e),
        }
    } else {
        "applied=skip".to_string()
    };

    let next = next_transition(now, cfg.latitude, cfg.longitude);

    log_event(&format!(
        "{} cause={} current={} target={} {} next={}",
        now_str,
        cause,
        theme_str(current),
        theme_str(Some(want)),
        outcome,
        next.to_rfc3339(),
    ));

    elwt.set_control_flow(ControlFlow::WaitUntil(deadline_instant(next)));
}

unsafe extern "system" fn wake_window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let kind = match msg {
        WM_WTSSESSION_CHANGE if wparam == WTS_SESSION_UNLOCK => Some(WakeKind::Unlock),
        WM_POWERBROADCAST if wparam == PBT_APMRESUMEAUTOMATIC => Some(WakeKind::Power),
        _ => None,
    };
    if let Some(k) = kind {
        if let Some(proxy) = EVENT_PROXY.get() {
            let _ = proxy.send_event(AppEvent::Wake(k));
        }
    }
    DefWindowProcW(hwnd, msg, wparam, lparam)
}

fn start_wake_listener() {
    std::thread::spawn(|| {
        let class_name = wide("WinThemeSwitcherWakeListener");
        unsafe {
            let hinstance = GetModuleHandleW(ptr::null());
            let mut wc: WNDCLASSW = std::mem::zeroed();
            wc.lpfnWndProc = Some(wake_window_proc);
            wc.hInstance = hinstance;
            wc.lpszClassName = class_name.as_ptr();
            RegisterClassW(&wc);

            let hwnd = CreateWindowExW(
                0,
                class_name.as_ptr(),
                ptr::null(),
                0,
                0,
                0,
                0,
                0,
                HWND_MESSAGE,
                ptr::null_mut(),
                hinstance,
                ptr::null(),
            );
            if (hwnd as usize) == 0 {
                return;
            }
            WTSRegisterSessionNotification(hwnd, NOTIFY_FOR_THIS_SESSION);
            let mut handle = ptr::null_mut();
            PowerRegisterSuspendResumeNotification(
                DEVICE_NOTIFY_WINDOW_HANDLE,
                hwnd as _,
                &mut handle,
            );

            let mut msg: MSG = std::mem::zeroed();
            while GetMessageW(&mut msg, ptr::null_mut(), 0, 0) > 0 {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
    });
}

fn main() -> Result<(), Box<dyn Error>> {
    ensure_com_initialized();

    let mut cfg = load_or_create_config();

    if !cfg.has_location() {
        acquire_location(&mut cfg);
    }

    if cfg.auto_start {
        let _ = set_auto_start(true);
    }

    let event_loop = EventLoop::<AppEvent>::with_user_event().build()?;
    let proxy = event_loop.create_proxy();
    let _ = EVENT_PROXY.set(event_loop.create_proxy());
    start_wake_listener();

    let tray_menu = Menu::new();
    let open_cfg_i = MenuItem::new("Open Config", true, None);
    let refresh_i = MenuItem::new("Refresh", true, None);
    let quit_i = MenuItem::new("Quit", true, None);
    tray_menu.append_items(&[
        &open_cfg_i,
        &refresh_i,
        &PredefinedMenuItem::separator(),
        &quit_i,
    ])?;

    let mut tray_builder = TrayIconBuilder::new()
        .with_menu(Box::new(tray_menu))
        .with_tooltip("WinThemeSwitcher");
    if let Some(icon) = make_tray_icon() {
        tray_builder = tray_builder.with_icon(icon);
    }
    let _tray = tray_builder.build()?;

    MenuEvent::set_event_handler(Some(move |event: MenuEvent| {
        let _ = proxy.send_event(AppEvent::Menu(event.id));
    }));

    let open_cfg_id = open_cfg_i.id().clone();
    let refresh_id = refresh_i.id().clone();
    let quit_id = quit_i.id().clone();

    event_loop.run(move |event, elwt| match event {
        Event::NewEvents(StartCause::Init) => tick(&cfg, elwt, "init", false),
        Event::NewEvents(StartCause::ResumeTimeReached { .. }) => {
            tick(&cfg, elwt, "resume-time", false);
        }
        Event::UserEvent(AppEvent::Wake(kind)) => {
            let cause = match kind {
                WakeKind::Unlock => "wake-unlock",
                WakeKind::Power => "wake-power",
            };
            tick(&cfg, elwt, cause, false);
        }
        Event::UserEvent(AppEvent::Menu(id)) => {
            if id == quit_id {
                elwt.exit();
            } else if id == open_cfg_id {
                open_config_in_editor();
            } else if id == refresh_id {
                cfg = load_or_create_config();
                if !cfg.has_location() {
                    if let Some((lat, lon)) = try_get_windows_location() {
                        cfg.latitude = lat;
                        cfg.longitude = lon;
                        let _ = save_config(&cfg);
                    }
                }
                tick(&cfg, elwt, "refresh", true);
            }
        }
        _ => {}
    })?;

    Ok(())
}
