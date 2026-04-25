#![windows_subsystem = "windows"]

use std::error::Error;
use std::fs;
use std::path::PathBuf;
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
use windows_sys::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows_sys::Win32::Graphics::Dwm::DwmFlush;
use windows_sys::Win32::System::LibraryLoader::GetModuleHandleW;
use windows_sys::Win32::System::Power::PowerRegisterSuspendResumeNotification;
use windows_sys::Win32::System::RemoteDesktop::{
    WTSRegisterSessionNotification, NOTIFY_FOR_THIS_SESSION,
};
use windows_sys::Win32::System::Registry::{
    RegCloseKey, RegDeleteValueW, RegOpenKeyExW, RegQueryValueExW, RegSetValueExW, HKEY,
    HKEY_CURRENT_USER, KEY_QUERY_VALUE, KEY_SET_VALUE, REG_DWORD, REG_SZ,
};
use windows_sys::Win32::UI::Shell::ShellExecuteW;
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
    Wake,
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
    std::thread::spawn(|| {
        let class = wide("ApplicationFrameWindow");
        let titles = [wide("Settings"), wide("Themes"), wide("Personalization")];
        for _ in 0..14 {
            std::thread::sleep(Duration::from_millis(150));
            for title in &titles {
                unsafe {
                    let hwnd = FindWindowW(class.as_ptr(), title.as_ptr());
                    if (hwnd as usize) != 0 {
                        PostMessageW(hwnd, WM_CLOSE, 0, 0);
                    }
                }
            }
        }
    });
}

fn apply_theme(theme: Theme, cfg: &Config) -> Result<(), Box<dyn Error>> {
    let theme_file = resolve_theme_file(theme, cfg);
    if theme_file.exists() && apply_theme_file(&theme_file) {
        start_settings_closer();
        std::thread::sleep(Duration::from_millis(300));
        poke_shell();
        return Ok(());
    }
    write_theme_registry(theme)?;
    broadcast_setting_change();
    poke_shell();
    Ok(())
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

fn tick(cfg: &Config, elwt: &ActiveEventLoop) {
    if !cfg.has_location() {
        elwt.set_control_flow(ControlFlow::Wait);
        return;
    }
    let now = Local::now();
    let want = target_theme(now, cfg.latitude, cfg.longitude);
    if current_theme() != Some(want) {
        let _ = apply_theme(want, cfg);
    }
    let next = next_transition(now, cfg.latitude, cfg.longitude);
    elwt.set_control_flow(ControlFlow::WaitUntil(deadline_instant(next)));
}

unsafe extern "system" fn wake_window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let trigger = match msg {
        WM_WTSSESSION_CHANGE => wparam == WTS_SESSION_UNLOCK,
        WM_POWERBROADCAST => wparam == PBT_APMRESUMEAUTOMATIC,
        _ => false,
    };
    if trigger {
        if let Some(proxy) = EVENT_PROXY.get() {
            let _ = proxy.send_event(AppEvent::Wake);
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
        Event::NewEvents(StartCause::Init)
        | Event::NewEvents(StartCause::ResumeTimeReached { .. }) => {
            tick(&cfg, elwt);
        }
        Event::UserEvent(AppEvent::Wake) => {
            tick(&cfg, elwt);
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
                if cfg.has_location() {
                    let now = Local::now();
                    let want = target_theme(now, cfg.latitude, cfg.longitude);
                    let _ = apply_theme(want, &cfg);
                }
                tick(&cfg, elwt);
            }
        }
        _ => {}
    })?;

    Ok(())
}
