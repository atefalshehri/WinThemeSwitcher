#![windows_subsystem = "windows"]

use std::error::Error;
use std::ffi::c_void;
use std::fs;
use std::path::{Path, PathBuf};
use std::ptr;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::OnceLock;
use std::time::{Duration, Instant};

use chrono::{DateTime, Local, Utc};
use serde::{Deserialize, Serialize};
use sun_times::sun_times;
use tray_icon::{
    menu::{Menu, MenuEvent, MenuId, MenuItem, PredefinedMenuItem},
    TrayIconBuilder, TrayIconEvent,
};
use windows::Devices::Geolocation::{GeolocationAccessStatus, Geolocator};
use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
use windows_sys::core::{GUID, HRESULT};
use windows_sys::Win32::Foundation::{
    GetLastError, SysFreeString, ERROR_ALREADY_EXISTS, HWND, LPARAM, LRESULT, WPARAM,
};
use windows_sys::Win32::Graphics::Dwm::DwmFlush;
use windows_sys::Win32::Storage::FileSystem::{
    GetFileInformationByHandle, BY_HANDLE_FILE_INFORMATION,
};
use windows_sys::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows_sys::Win32::System::LibraryLoader::{
    GetModuleHandleW, SetDefaultDllDirectories, LOAD_LIBRARY_SEARCH_SYSTEM32,
};
use windows_sys::Win32::System::Power::{
    PowerRegisterSuspendResumeNotification, DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS,
};
use windows_sys::Win32::System::Registry::{
    RegCloseKey, RegDeleteValueW, RegOpenKeyExW, RegQueryValueExW, RegSetValueExW, HKEY,
    HKEY_CURRENT_USER, KEY_QUERY_VALUE, KEY_SET_VALUE, REG_DWORD, REG_SZ,
};
use windows_sys::Win32::System::RemoteDesktop::{
    WTSRegisterSessionNotification, NOTIFY_FOR_THIS_SESSION,
};
use windows_sys::Win32::System::SystemInformation::GetSystemDirectoryW;
use windows_sys::Win32::System::Threading::CreateMutexW;
use windows_sys::Win32::System::WindowsProgramming::GetPrivateProfileStringW;
use windows_sys::Win32::UI::Shell::{
    AssocQueryStringW, SHLoadIndirectString, ShellExecuteW, ASSOCF_INIT_IGNOREUNKNOWN,
    ASSOCSTR_FRIENDLYAPPNAME,
};
use windows_sys::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DispatchMessageW, FindWindowW, GetMessageW, MessageBoxW,
    PostMessageW, RegisterClassW, SendMessageTimeoutW, TranslateMessage, DEVICE_NOTIFY_CALLBACK,
    HWND_BROADCAST, HWND_MESSAGE, IDYES, MB_ICONINFORMATION, MB_ICONQUESTION, MB_ICONWARNING,
    MB_OK, MB_YESNO, MSG, PBT_APMRESUMEAUTOMATIC, SMTO_ABORTIFHUNG, SW_HIDE, SW_SHOWNORMAL,
    WM_CLOSE, WM_SETTINGCHANGE, WM_THEMECHANGED, WM_WTSSESSION_CHANGE, WNDCLASSW,
};
use winit::event::{Event, StartCause};
use winit::event_loop::{ActiveEventLoop, ControlFlow, EventLoop, EventLoopProxy};

const WTS_SESSION_UNLOCK: WPARAM = 0x8;

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
    query_interface:
        unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> HRESULT,
    add_ref: unsafe extern "system" fn(*mut c_void) -> u32,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
    // IThemeManager2 — every slot up through the last one we call.
    // ORDER MATTERS — must match the vtable layout exactly; the `_`-prefixed
    // entries are never called but are mandatory padding that keeps the
    // called slots at the right offsets. Reference: namazso C# gist + wtheme
    // C header.
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
    // only TRAILING slots after the last called method may be left out; they
    // don't affect the offsets above.
}

#[repr(C)]
struct IThemeVtbl {
    query_interface:
        unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> HRESULT,
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
        let hr =
            (self.vtbl().set_current_theme)(self.ptr, ptr::null_mut(), index, 1, apply_flags, 0);
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
    /// Keys this version doesn't know (a user's note, a newer version's
    /// setting). Kept so save_config writes them back instead of silently
    /// deleting them.
    #[serde(flatten)]
    extra: serde_json::Map<String, serde_json::Value>,
}

impl Default for Config {
    fn default() -> Self {
        Self {
            latitude: 0.0,
            longitude: 0.0,
            auto_start: true,
            theme_day: None,
            theme_night: None,
            extra: serde_json::Map::new(),
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

impl Theme {
    fn opposite(self) -> Theme {
        match self {
            Theme::Light => Theme::Dark,
            Theme::Dark => Theme::Light,
        }
    }
}

/// Why a tick is running — decides whether an apply is forced, state-aware, or
/// override-preserving (see `decide_tick`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TickKind {
    /// First tick after launch.
    Init,
    /// ResumeTimeReached — a scheduled sunrise/sunset (or an apply retry).
    Scheduled,
    /// Session unlock / power resume.
    Wake,
    /// User clicked Refresh.
    Refresh,
}

/// Per-session tick state, owned by the event-loop closure.
struct TickState {
    /// Next scheduled transition (UTC) recorded by the last tick that ended
    /// reconciled — applied successfully, found the screen already matching,
    /// or deliberately preserved an override. `now >= reconciled_next` on a
    /// later tick means at least one transition has passed since we were
    /// last in sync, regardless of how many were missed (a same-THEME parity
    /// comparison would wrongly preserve an override across an ordinary
    /// overnight lock that spans sunset AND sunrise). A FAILED apply never
    /// advances it — a retry leaves it stale, an exhausted budget clears it —
    /// so every subsequent wake sees the transition as still-unreconciled
    /// and re-applies, instead of misreading the failure as a user override.
    reconciled_next: Option<DateTime<Utc>>,
    /// Consecutive failed applies in the current failure episode.
    retry_count: u32,
    /// current_theme() observed when the last apply failed. If the screen no
    /// longer matches this, the user intervened during the retry window and
    /// the retry must stand down rather than clobber their choice.
    retry_baseline: Option<Theme>,
    /// The next-transition instant computed at the tick whose apply failed —
    /// the failure episode's own window. An intervention only cancels the
    /// retry while `now < episode_next`; past it, a transition has passed
    /// and reconciling to the schedule outranks the stand-down (otherwise an
    /// intervention right before a suspend that spans a transition would be
    /// promoted to a day-long override).
    episode_next: Option<DateTime<Utc>>,
    /// Wall-clock deadline of the pending wait (None = plain Wait, e.g. no
    /// location). See `plan_wake`.
    armed: Option<DateTime<Utc>>,
    /// (wall, monotonic) clock readings the last tick decided from — the
    /// reference for detecting a clock step. Only `tick` writes it (via `arm`,
    /// or clearing it on the no-location path); wake handling never re-marks.
    mark: Option<(DateTime<Utc>, Instant)>,
}

impl TickState {
    fn new() -> Self {
        Self {
            reconciled_next: None,
            retry_count: 0,
            retry_baseline: None,
            episode_next: None,
            armed: None,
            mark: None,
        }
    }
}

/// What a tick decided to do — see `decide_tick`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TickAction {
    /// Call apply_theme.
    Apply,
    /// Screen already matches the schedule.
    SkipInSync,
    /// Wake (or early Scheduled) tick, screen diverges, but no transition
    /// passed since the last reconciled tick: a manual override is being
    /// preserved.
    SkipOverride,
    /// The user changed the theme during a pending retry window — cancel the
    /// retry episode and let their choice stand until the next transition.
    CancelRetry,
}

/// The tick decision, minus all I/O. Pure — unit-tested, including across
/// sequences of ticks mutating one TickState via note_reconciled /
/// note_apply_failed. IMPORTANT: called with the PRE-tick state; the state
/// notes are recorded after the apply outcome is known.
fn decide_tick(
    kind: TickKind,
    current: Option<Theme>,
    target: Theme,
    now: DateTime<Utc>,
    next: DateTime<Utc>,
    clock_stepped_back: bool,
    state: &TickState,
) -> TickAction {
    // Refresh is fresh user intent: always force-apply.
    if kind == TickKind::Refresh {
        return TickAction::Apply;
    }
    // Pending-retry gate: the screen moved away from the failure snapshot,
    // so the user intervened mid-episode. Only an OBSERVED move counts — an
    // unreadable reading (None) on EITHER side proves nothing and must not
    // cancel. And the stand-down only applies within the failure episode's
    // own window (now < episode_next): once a transition has passed,
    // reconciling to the schedule outranks it, exactly like any override
    // ending at the next natural transition.
    if state.retry_count > 0
        && state.retry_baseline.is_some()
        && current.is_some()
        && current != state.retry_baseline
        && state.episode_next.is_some_and(|n| now < n)
    {
        return TickAction::CancelRetry;
    }
    if current == Some(target) {
        return TickAction::SkipInSync;
    }
    // A transition has passed since we were last in sync if now is at/after
    // the recorded next transition — OR if the wall clock was observed to
    // step BACKWARDS and the schedule's upcoming transition is now earlier
    // than the recorded one: a clock that had been running ahead got
    // corrected across a transition, so the frame we reconciled in no longer
    // exists and a diverged screen is the app's own stale apply, not a user
    // override. Gated on an observed step because `next` can also move
    // earlier without one — e.g. after Refresh adopts new coordinates.
    let transition_passed = state.reconciled_next.is_none_or(|n| {
        now >= n || (clock_stepped_back && next + chrono::Duration::seconds(60) < n)
    });
    // No transition has passed since we were last in sync, so a diverged
    // screen is a manual override: preserve it. This covers Wake ticks and
    // ALSO Scheduled ticks that run before reconciled_next — which, since
    // plan_wake never ticks before the armed deadline, are clock-jump ticks
    // (a forward step or suspend artifact that didn't reach the transition,
    // or a backward step within the same frame). Re-applying the schedule
    // there would revert the user's override for no transition at all.
    // Override preservation yields to a pending retry:
    // during an active failure episode the divergence is the FAILURE, not
    // an override (including episodes started mid-window by a failed
    // Refresh, where reconciled_next is still in the future), so retry
    // ticks and wakes re-apply.
    if matches!(kind, TickKind::Wake | TickKind::Scheduled)
        && !transition_passed
        && state.retry_count == 0
    {
        return TickAction::SkipOverride;
    }
    TickAction::Apply
}

/// Record a tick that ended in sync with the schedule (any non-Err outcome).
fn note_reconciled(state: &mut TickState, next: DateTime<Utc>) {
    state.reconciled_next = Some(next);
    state.retry_count = 0;
    state.retry_baseline = None;
    state.episode_next = None;
}

/// Record a failed apply. Returns true when a quick retry should be
/// scheduled (reconciled_next is left stale); false when the budget is
/// exhausted (the episode resets so the NEXT transition window gets a fresh
/// budget, and reconciled_next is cleared so later wakes remain free retry
/// opportunities).
fn note_apply_failed(
    state: &mut TickState,
    observed: Option<Theme>,
    next_utc: DateTime<Utc>,
) -> bool {
    state.retry_count += 1;
    state.retry_baseline = observed;
    if state.retry_count > MAX_APPLY_RETRIES {
        state.retry_count = 0;
        state.retry_baseline = None;
        state.episode_next = None;
        // Budget spent: the screen is still wrong, so later wakes must keep
        // treating it as a failure to retry, never as an override to
        // preserve — including an episode a failed Refresh started mid-
        // window, where reconciled_next is still in the future.
        state.reconciled_next = None;
        false
    } else {
        state.episode_next = Some(next_utc);
        true
    }
}

/// What the Toggle menu item should apply. Pure — unit-tested. An unreadable
/// current theme (registry read failure) defaults the base to Light, so the
/// first toggle lands on Dark.
fn toggle_target(current: Option<Theme>) -> Theme {
    current.unwrap_or(Theme::Light).opposite()
}

/// Keep a message single-line and parseable inside a `msg="..."` log field
/// (inner quotes swapped to apostrophes, newlines flattened). Shared by the
/// panic hook, fatal-error reporting, config-error reporting, and tick
/// apply-error lines.
fn sanitize_log_msg(s: &str) -> String {
    s.replace('"', "'").replace(['\n', '\r'], " ")
}

/// How long after a failed apply the bounded retry fires.
const APPLY_RETRY_DELAY_SECS: i64 = 60;
/// Consecutive failures after which we stop retrying until the next
/// scheduled transition.
const MAX_APPLY_RETRIES: u32 = 3;

/// Deadline for the next tick after a failed apply: retry soon, but never
/// past the scheduled transition itself. Pure — unit-tested.
fn retry_deadline(now: DateTime<Local>, next: DateTime<Local>) -> DateTime<Local> {
    std::cmp::min(
        next,
        now + chrono::Duration::seconds(APPLY_RETRY_DELAY_SECS),
    )
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

/// Serializes rotation + append across threads (main loop, commit watcher,
/// settings closer, wake listener, detached MessageBox threads).
static LOG_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn log_event(line: &str) {
    // A poisoned lock just means another thread panicked mid-write; logging
    // must keep working.
    let _guard = LOG_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    write_log_line(line);
}

/// The panic hook's variant: never blocks. std's Mutex isn't reentrant, so a
/// panic raised while this thread holds LOG_LOCK would otherwise deadlock
/// instead of aborting. Best effort is right for a dying process.
fn log_event_from_panic(line: &str) {
    let _guard = LOG_LOCK.try_lock();
    write_log_line(line);
}

fn write_log_line(line: &str) {
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
        // One write per line, so concurrent lines can't interleave.
        let _ = f.write_all(format!("{line}\n").as_bytes());
    }
}

fn theme_str(t: Option<Theme>) -> &'static str {
    match t {
        Some(Theme::Light) => "light",
        Some(Theme::Dark) => "dark",
        None => "unknown",
    }
}

/// A self-heal performed by `load_config_ex`.
#[derive(Debug)]
struct Healed {
    reason: &'static str,
    /// Set when writing the healed defaults back failed (the reset then only
    /// exists in memory, and repeats next launch).
    save_err: Option<String>,
}

/// `load_config_ex` without the heal flag (tests use this form).
#[cfg(test)]
fn load_config_at(path: &Path) -> Result<Config, String> {
    load_config_ex(path).map(|(cfg, _)| cfg)
}

/// Load the config from `path`. A missing file is first-run: defaults are
/// written and returned. `Err` means the file EXISTS but could not be read or
/// parsed — it is left untouched on disk so a hand-edit typo can be fixed
/// instead of silently wiping the user's coordinates and theme paths.
///
/// Also reports whether the file was self-healed (so callers can log it — a
/// heal resets settings, and must not be silent). Logging stays out of here
/// because tests call this directly and log_path() is exe-relative.
fn load_config_ex(path: &Path) -> Result<(Config, Option<Healed>), String> {
    match read_config_text(path) {
        Ok(content) if config_content_is_empty(&content) => {
            // A crash mid-write can leave a 0-byte or NUL-filled file (the
            // latter when the size was extended but the data never flushed).
            // Nothing in it to preserve — self-heal like first run instead
            // of erroring on every launch.
            let cfg = Config::default();
            let save_err = save_config_at(path, &cfg).err().map(|e| e.to_string());
            Ok((
                cfg,
                Some(Healed {
                    reason: "empty",
                    save_err,
                }),
            ))
        }
        Ok(content) => serde_json::from_str::<Config>(&content)
            .map(|cfg| (cfg, None))
            .map_err(|e| format!("config.json is not valid JSON: {e}")),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
            let cfg = Config::default();
            if let Ok(json) = serde_json::to_string_pretty(&cfg) {
                // create_new, not a replace: if the file appears between our
                // read and this write (an editor saving via delete-then-
                // rename), the user's file wins and the defaults are dropped.
                use std::io::Write;
                if let Ok(mut f) = fs::OpenOptions::new()
                    .write(true)
                    .create_new(true)
                    .open(path)
                {
                    let _ = f.write_all(json.as_bytes());
                }
            }
            Ok((cfg, None))
        }
        Err(e) => Err(format!("config.json could not be read: {e}")),
    }
}

/// config.json as text, whatever encoding a Windows editor saved it in:
/// UTF-8 with or without a BOM (Notepad), or UTF-16 LE/BE with a BOM
/// (PowerShell 5.1's `>` / Out-File default is LE; Notepad offers both).
fn read_config_text(path: &Path) -> std::io::Result<String> {
    use std::io::{Error, ErrorKind};
    let bytes = fs::read(path)?;
    let utf16 = |rest: &[u8], le: bool| -> std::io::Result<String> {
        if !rest.len().is_multiple_of(2) {
            return Err(Error::new(ErrorKind::InvalidData, "truncated UTF-16 text"));
        }
        let units: Vec<u16> = rest
            .chunks_exact(2)
            .map(|c| {
                if le {
                    u16::from_le_bytes([c[0], c[1]])
                } else {
                    u16::from_be_bytes([c[0], c[1]])
                }
            })
            .collect();
        String::from_utf16(&units)
            .map_err(|_| Error::new(ErrorKind::InvalidData, "not valid UTF-16 text"))
    };
    if let Some(rest) = bytes.strip_prefix(&[0xFF, 0xFE]) {
        return utf16(rest, true);
    }
    if let Some(rest) = bytes.strip_prefix(&[0xFE, 0xFF]) {
        return utf16(rest, false);
    }
    let rest = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(&bytes);
    String::from_utf8(rest.to_vec()).map_err(|_| {
        Error::new(
            ErrorKind::InvalidData,
            "not UTF-8 or UTF-16 text (save it as UTF-8)",
        )
    })
}

/// Nothing worth preserving: only whitespace, NULs, and/or a BOM.
fn config_content_is_empty(content: &str) -> bool {
    content
        .trim_matches(|c: char| c.is_whitespace() || c == '\0' || c == '\u{feff}')
        .is_empty()
}

/// Serialize `cfg` over config.json unconditionally. Only call with a Config
/// that was successfully loaded from disk this session — persisting a
/// default/fallback Config here is exactly the settings-wipe bug fixed in
/// v0.3.2 (broken files must stay on disk for the user to repair).
fn save_config(cfg: &Config) -> Result<(), Box<dyn Error>> {
    save_config_at(&config_path(), cfg)
}

/// Atomic replace: write a sibling temp file, then rename it over `path`
/// (std's rename replaces an existing file on Windows). A crash leaves either
/// the old file or the new one — never the truncated/NUL-filled middle state
/// that a direct fs::write can.
fn save_config_at(path: &Path, cfg: &Config) -> Result<(), Box<dyn Error>> {
    use std::io::Write;
    let json = serde_json::to_string_pretty(cfg)?;
    if is_linked_file(path) {
        // A rename would detach this name from the shared file (e.g. a
        // package manager's hard-linked "persisted" config) — write in place
        // instead, accepting the non-atomic window for this rare case.
        let mut f = fs::OpenOptions::new()
            .write(true)
            .truncate(true)
            .open(path)?;
        f.write_all(json.as_bytes())?;
        f.sync_all()?;
        return Ok(());
    }
    let tmp = path.with_extension("json.tmp");
    let written = (|| -> std::io::Result<()> {
        let mut f = fs::File::create(&tmp)?;
        f.write_all(json.as_bytes())?;
        // Data must be on disk before the rename publishes it.
        f.sync_all()
    })();
    if let Err(e) = written {
        let _ = fs::remove_file(&tmp);
        return Err(e.into());
    }
    // An on-access AV scan of the fresh temp file can briefly hold it
    // (ERROR_ACCESS_DENIED / ERROR_SHARING_VIOLATION) — retry a moment
    // before giving up. A read-only config.json fails the same way, for good.
    let mut last = None;
    for attempt in 0..5 {
        match fs::rename(&tmp, path) {
            Ok(()) => return Ok(()),
            Err(e) if matches!(e.raw_os_error(), Some(5) | Some(32)) && attempt < 4 => {
                std::thread::sleep(Duration::from_millis(100));
                last = Some(e);
            }
            Err(e) => {
                last = Some(e);
                break;
            }
        }
    }
    let _ = fs::remove_file(&tmp);
    Err(last
        .map(|e| e.into())
        .unwrap_or_else(|| "rename failed".into()))
}

/// Load config.json from its real location, logging a self-heal.
fn load_config_logged() -> Result<Config, String> {
    load_config_ex(&config_path()).map(|(cfg, healed)| {
        if let Some(h) = healed {
            let action = if h.save_err.is_some() {
                "reset-in-memory"
            } else {
                "reset-to-defaults"
            };
            log_event(&format!(
                "{} config_healed reason={} action={}",
                Local::now().to_rfc3339(),
                h.reason,
                action,
            ));
            if let Some(e) = h.save_err {
                log_event(&format!(
                    "{} config_save_err msg=\"{}\"",
                    Local::now().to_rfc3339(),
                    sanitize_log_msg(&e),
                ));
            }
        }
        cfg
    })
}

/// Whether `path` is a symlink/junction or a hard link with other names —
/// replacing it by rename would silently detach it from the shared file.
/// Only name-surrogate reparse points count (`is_symlink()`: symlinks,
/// junctions, LX symlinks): OneDrive/cloud and dedup files are reparse
/// points too, but renaming over them is fine.
fn is_linked_file(path: &Path) -> bool {
    use std::os::windows::io::AsRawHandle;
    match fs::symlink_metadata(path) {
        Ok(meta) if meta.file_type().is_symlink() => return true,
        Ok(_) => {}
        Err(_) => return false,
    }
    let Ok(f) = fs::File::open(path) else {
        return false;
    };
    let mut info: BY_HANDLE_FILE_INFORMATION = unsafe { std::mem::zeroed() };
    let ok = unsafe { GetFileInformationByHandle(f.as_raw_handle() as _, &mut info) };
    ok != 0 && info.nNumberOfLinks > 1
}

/// save_config with the failure logged instead of dropped.
fn persist_config(cfg: &Config) {
    if let Err(e) = save_config(cfg) {
        log_event(&format!(
            "{} config_save_err msg=\"{}\"",
            Local::now().to_rfc3339(),
            sanitize_log_msg(&e.to_string()),
        ));
    }
}

/// Refresh's config step, minus I/O. Adopts a freshly (re)loaded config when
/// it parsed. Returns whether the autostart registration may be re-asserted —
/// only from a config that came from disk at some point this session, never
/// from the in-memory fallback of a session that started with a broken file
/// (whose auto_start=true default must not override the file's false) — and
/// the load error to report, if any. Pure — unit-tested.
fn adopt_reloaded_config(
    load: Result<Config, String>,
    cfg: &mut Config,
    have_disk_cfg: &mut bool,
) -> (bool, Option<String>) {
    match load {
        Ok(new_cfg) => {
            *cfg = new_cfg;
            *have_disk_cfg = true;
            (true, None)
        }
        // Keep the last-known-good (or fallback) config; the broken file
        // stays on disk for the user to fix.
        Err(e) => (*have_disk_cfg, Some(e)),
    }
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

/// Whether `.json` has a real "open" handler. Without one (common on a fresh
/// Windows — and true on the maintainer's machine), ShellExecute("open") on
/// config.json doesn't fail: it resolves to the "Unknown" class and shows the
/// "How do you want to open this file?" picker. ASSOCF_INIT_IGNOREUNKNOWN
/// makes that case fail here instead.
fn json_has_open_handler() -> bool {
    let ext = wide(".json");
    let verb = wide("open");
    let mut len: u32 = 0;
    // FRIENDLYAPPNAME, not EXECUTABLE: packaged (Store) handlers have no
    // executable path but do have a name. A null buffer just asks for the
    // length — S_FALSE/S_OK both mean "there is a handler".
    let hr = unsafe {
        AssocQueryStringW(
            ASSOCF_INIT_IGNOREUNKNOWN,
            ASSOCSTR_FRIENDLYAPPNAME,
            ext.as_ptr(),
            verb.as_ptr(),
            ptr::null_mut(),
            &mut len,
        )
    };
    hr >= 0
}

/// Full path of System32\notepad.exe — never a bare "notepad.exe", which
/// ShellExecute would resolve by search starting in the working directory
/// (the exe's own folder on a double-click launch).
fn system_notepad() -> PathBuf {
    let mut buf = [0u16; 260];
    let n = unsafe { GetSystemDirectoryW(buf.as_mut_ptr(), buf.len() as u32) } as usize;
    let dir = if n > 0 && n < buf.len() {
        PathBuf::from(String::from_utf16_lossy(&buf[..n]))
    } else {
        PathBuf::from("C:\\Windows\\System32")
    };
    dir.join("notepad.exe")
}

fn open_config_in_editor() {
    let path = config_path();
    let path_w = wide(&path.to_string_lossy());
    let verb = wide("open");
    unsafe {
        let mut code: isize = -1;
        if json_has_open_handler() {
            let h = ShellExecuteW(
                ptr::null_mut(),
                verb.as_ptr(),
                path_w.as_ptr(),
                ptr::null(),
                ptr::null(),
                SW_SHOWNORMAL,
            );
            code = h as isize;
            if code > 32 {
                return;
            }
        }
        // No usable .json handler: Notepad (by full path).
        let notepad = wide(&system_notepad().to_string_lossy());
        let arg = wide(&format!("\"{}\"", path.to_string_lossy()));
        let h2 = ShellExecuteW(
            ptr::null_mut(),
            verb.as_ptr(),
            notepad.as_ptr(),
            arg.as_ptr(),
            ptr::null(),
            SW_SHOWNORMAL,
        );
        if (h2 as isize) <= 32 {
            log_event(&format!(
                "{} open_config_err code={} fallback_code={}",
                Local::now().to_rfc3339(),
                code,
                h2 as isize,
            ));
        }
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

fn show_message_box(title: &str, body: &str, flags: u32) -> i32 {
    let title_w = wide(title);
    let body_w = wide(body);
    unsafe { MessageBoxW(ptr::null_mut(), body_w.as_ptr(), title_w.as_ptr(), flags) }
}

fn ask_enable_location() -> bool {
    show_message_box(
        "WinThemeSwitcher — Location",
        "Windows Location is off or not allowed for desktop apps.\n\n\
         Enable it so sunrise and sunset can be computed automatically?\n\n\
         Yes opens Windows Settings. No lets you enter coordinates manually in config.json.",
        MB_YESNO | MB_ICONQUESTION,
    ) == IDYES
}

fn show_enable_pending_message() {
    show_message_box(
        "WinThemeSwitcher",
        "Turn on \"Location services\" in the Settings window that just opened. \
         Then right-click the WinThemeSwitcher tray icon and choose Refresh.",
        MB_OK | MB_ICONINFORMATION,
    );
}

fn show_manual_setup_prompt() {
    show_message_box(
        "WinThemeSwitcher — Setup",
        "Please set latitude and longitude in config.json (opening now), \
         then right-click the tray icon and choose Refresh.",
        MB_OK | MB_ICONINFORMATION,
    );
    open_config_in_editor();
}

/// A config-error box is already on screen (repeated Refresh clicks with a
/// still-broken file must not stack duplicates).
static CONFIG_ERROR_BOX_OPEN: AtomicBool = AtomicBool::new(false);

/// Log + tell the user their hand-edited config is broken and was left
/// untouched. Called from startup and from a user-initiated Refresh. The
/// MessageBox runs on a detached thread — a modal here would otherwise park
/// startup before the tray exists, or stall the event loop (scheduled
/// transitions, wake events) until dismissed. Plain MessageBoxW has no STA
/// requirement, so a worker thread is fine.
fn report_config_error(err: &str) {
    log_event(&format!(
        "{} config_error msg=\"{}\"",
        Local::now().to_rfc3339(),
        // serde_json errors quote the offending token; keep the log's
        // quoted-field convention parseable.
        sanitize_log_msg(err),
    ));
    if CONFIG_ERROR_BOX_OPEN.swap(true, Ordering::SeqCst) {
        return;
    }
    let body = format!(
        "{err}\n\nThe file was left unchanged — your settings are still in it. \
         Fix the error (tray menu → Open Config), then choose Refresh.",
    );
    std::thread::spawn(move || {
        show_message_box(
            "WinThemeSwitcher — Config error",
            &body,
            MB_OK | MB_ICONWARNING,
        );
        CONFIG_ERROR_BOX_OPEN.store(false, Ordering::SeqCst);
    });
}

fn acquire_location(cfg: &mut Config) {
    if let Some((lat, lon)) = try_get_windows_location() {
        cfg.latitude = lat;
        cfg.longitude = lon;
        persist_config(cfg);
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
        if RegOpenKeyExW(
            HKEY_CURRENT_USER,
            subkey.as_ptr(),
            0,
            KEY_QUERY_VALUE,
            &mut hkey,
        ) != 0
        {
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
        if RegOpenKeyExW(
            HKEY_CURRENT_USER,
            subkey.as_ptr(),
            0,
            KEY_SET_VALUE,
            &mut hkey,
        ) != 0
        {
            return Err("RegOpenKeyExW failed for Personalize".into());
        }
        // Both results matter: a blocked write (AV/HIPS, ACL, hive error)
        // must surface as Err so tick's bounded retry runs, instead of being
        // logged as applied=registry while nothing changed.
        // Order matters: SystemUsesLightTheme — the value current_theme()
        // reads — is written LAST and only if Apps succeeded, so the only
        // possible half state ("Apps new, System old") reads as not-applied
        // and gets retried. The reverse would read as the screen having
        // moved, which the retry gate would take for a user intervention.
        let mut failed: Option<String> = None;
        let mut any_ok = false;
        for (name, name_w) in [("AppsUseLightTheme", &apps), ("SystemUsesLightTheme", &sys)] {
            let rc = RegSetValueExW(
                hkey,
                name_w.as_ptr(),
                0,
                REG_DWORD,
                &value as *const u32 as *const u8,
                4,
            );
            if rc != 0 {
                failed = Some(format!("RegSetValueExW {name} rc={rc}"));
                break;
            }
            any_ok = true;
        }
        RegCloseKey(hkey);
        if let Some(e) = failed {
            if any_ok {
                // Half-written (one mode flipped): callers skip their
                // broadcast on Err, so repaint here rather than leave a
                // changed value un-announced.
                broadcast_setting_change();
                poke_shell();
            }
            return Err(e.into());
        }
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
    PathBuf::from(win_dir)
        .join("Resources")
        .join("Themes")
        .join(leaf)
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
                    sanitize_log_msg(&e.to_string()),
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
///
/// Parsed with Windows' own INI reader (GetPrivateProfileStringW), not by hand,
/// so the name we match in tier 1 is exactly the one Windows reads: section
/// and key names case-insensitive, whitespace around `=` trimmed, enclosing
/// quotes stripped, and the file decoded the way Windows decodes it (UTF-16
/// with a BOM, otherwise the system ANSI code page — system .theme files are
/// Windows-1252, e.g. `aero.theme`'s raw `0xa9` copyright byte).
fn resolve_theme_display_name(theme_file: &Path) -> Option<String> {
    if !theme_file.is_file() {
        return None;
    }
    let section = wide("Theme");
    let key = wide("DisplayName");
    let empty = wide("");
    // Absolute: given a bare file name, GetPrivateProfileStringW looks in
    // %WINDIR%, not the directory is_file() just checked.
    let abs = std::path::absolute(theme_file).unwrap_or_else(|_| theme_file.to_path_buf());
    let file = wide(&abs.to_string_lossy());
    let mut buf = vec![0u16; 1024];
    let n = unsafe {
        GetPrivateProfileStringW(
            section.as_ptr(),
            key.as_ptr(),
            empty.as_ptr(),
            buf.as_mut_ptr(),
            buf.len() as u32,
            file.as_ptr(),
        )
    } as usize;
    if n == 0 {
        return None;
    }
    let raw = String::from_utf16_lossy(&buf[..n.min(buf.len())]);
    let raw = raw.trim();
    if raw.starts_with('@') {
        return resolve_indirect_string(raw);
    }
    Some(raw.to_string())
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
                    sanitize_log_msg(&name),
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
    let configured = match theme {
        Theme::Light => cfg.theme_day.as_deref(),
        Theme::Dark => cfg.theme_night.as_deref(),
    };
    if let Some(p) = configured {
        if Path::new(p) != theme_file {
            // resolve_theme_file fell back to the stock theme — say so, or a
            // typo in theme_day/theme_night is invisible.
            log_event(&format!(
                "{} theme_path_missing target={} path=\"{}\" using=\"{}\"",
                Local::now().to_rfc3339(),
                theme_str(Some(theme)),
                sanitize_log_msg(p),
                sanitize_log_msg(&theme_file.to_string_lossy()),
            ));
        }
    }

    if theme_file.exists() {
        match apply_via_theme_manager2(theme, &theme_file) {
            Ok(()) => return Ok("theme-manager2"),
            Err(e) => log_event(&format!(
                "{} theme_manager2_err target={} msg=\"{}\"",
                Local::now().to_rfc3339(),
                theme_str(Some(theme)),
                sanitize_log_msg(&e.to_string()),
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

/// Make the HKCU Run value match `enable`. Writes only when the value differs
/// (a missing value — e.g. deleted by an AV quarantine — always differs, so
/// the documented re-assert-on-launch/Refresh recovery still works), and
/// surfaces every registry failure as Err instead of dropping it.
fn set_auto_start(enable: bool) -> Result<(), Box<dyn Error>> {
    const ERROR_FILE_NOT_FOUND: u32 = 2;
    let subkey = wide(RUN_KEY);
    let name = wide(APP_NAME);
    unsafe {
        let mut hkey: HKEY = ptr::null_mut();
        let rc = RegOpenKeyExW(
            HKEY_CURRENT_USER,
            subkey.as_ptr(),
            0,
            KEY_QUERY_VALUE | KEY_SET_VALUE,
            &mut hkey,
        );
        if rc == ERROR_FILE_NOT_FOUND && !enable {
            // No Run key at all → certainly no Run value to remove.
            return Ok(());
        }
        if rc != 0 {
            return Err(format!("RegOpenKeyExW Run rc={rc}").into());
        }
        let result = if enable {
            match std::env::current_exe() {
                Ok(exe) => {
                    let want = format!("\"{}\"", exe.to_string_lossy());
                    if read_reg_sz(hkey, &name).as_deref() == Some(want.as_str()) {
                        Ok(())
                    } else {
                        let want_w = wide(&want);
                        let rc = RegSetValueExW(
                            hkey,
                            name.as_ptr(),
                            0,
                            REG_SZ,
                            want_w.as_ptr() as *const u8,
                            (want_w.len() * 2) as u32,
                        );
                        if rc == 0 {
                            Ok(())
                        } else {
                            Err(format!("RegSetValueExW Run rc={rc}"))
                        }
                    }
                }
                Err(e) => Err(format!("current_exe: {e}")),
            }
        } else {
            let rc = RegDeleteValueW(hkey, name.as_ptr());
            if rc == 0 || rc == ERROR_FILE_NOT_FOUND {
                Ok(())
            } else {
                Err(format!("RegDeleteValueW Run rc={rc}"))
            }
        };
        RegCloseKey(hkey);
        result.map_err(|e| e.into())
    }
}

/// Read a REG_SZ value from an open key; None if absent, not REG_SZ, or
/// unreadable. `name` must be NUL-terminated (from `wide`).
unsafe fn read_reg_sz(hkey: HKEY, name: &[u16]) -> Option<String> {
    let mut kind: u32 = 0;
    let mut size: u32 = 0;
    if RegQueryValueExW(
        hkey,
        name.as_ptr(),
        ptr::null_mut(),
        &mut kind,
        ptr::null_mut(),
        &mut size,
    ) != 0
        || kind != REG_SZ
    {
        return None;
    }
    // +1 u16 so the buffer is NUL-terminated even if the stored data isn't.
    let mut buf = vec![0u16; (size as usize).div_ceil(2) + 1];
    let mut bytes = (buf.len() * 2) as u32;
    if RegQueryValueExW(
        hkey,
        name.as_ptr(),
        ptr::null_mut(),
        &mut kind,
        buf.as_mut_ptr() as *mut u8,
        &mut bytes,
    ) != 0
    {
        return None;
    }
    let len = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
    Some(String::from_utf16_lossy(&buf[..len]))
}

/// set_auto_start with the failure logged — a missing autostart entry must
/// leave a diagnosable trace in events.log.
fn apply_auto_start(enable: bool) {
    if let Err(e) = set_auto_start(enable) {
        log_event(&format!(
            "{} autostart_err op={} msg=\"{}\"",
            Local::now().to_rfc3339(),
            if enable { "set" } else { "delete" },
            sanitize_log_msg(&e.to_string()),
        ));
    }
}

/// Civil sunrise/sunset threshold: sun center 0.833° below the horizon
/// (accounts for refraction + solar radius; same convention as `sun_times`).
const SUNRISE_ALTITUDE_DEG: f64 = -0.833;

/// Solar altitude in degrees at `t` for the given location. Standard
/// low-precision solar position (declination + hour angle against Greenwich
/// sidereal time) — well under a degree of error, plenty for deciding polar
/// day vs. polar night. Implemented locally because the `sun_times` crate's
/// own `altitude` has math bugs (seconds term, spurious to_degrees on an
/// already-degrees longitude).
fn solar_altitude_deg(t: DateTime<Utc>, lat: f64, lon: f64) -> f64 {
    // Fractional days since J2000.0 (JD 2451545.0).
    let n = t.timestamp_millis() as f64 / 86_400_000.0 + 2_440_587.5 - 2_451_545.0;
    let mean_long = (280.460 + 0.985_647_4 * n).rem_euclid(360.0);
    let mean_anom = (357.528 + 0.985_600_3 * n).rem_euclid(360.0).to_radians();
    let ecl_long =
        (mean_long + 1.915 * mean_anom.sin() + 0.020 * (2.0 * mean_anom).sin()).to_radians();
    let obliquity = (23.439 - 0.000_000_4 * n).to_radians();
    let declination = (obliquity.sin() * ecl_long.sin()).asin();
    let right_ascension = f64::atan2(obliquity.cos() * ecl_long.sin(), ecl_long.cos());
    let gmst_deg = (280.460_618_37 + 360.985_647_366_29 * n).rem_euclid(360.0);
    let hour_angle = (gmst_deg + lon - right_ascension.to_degrees())
        .rem_euclid(360.0)
        .to_radians();
    let lat_r = lat.to_radians();
    (lat_r.sin() * declination.sin() + lat_r.cos() * declination.cos() * hour_angle.cos())
        .asin()
        .to_degrees()
}

/// Sunrise/sunset instants for the UTC dates `d-1 ..= d+1` around `now`,
/// sorted, each tagged with the theme in effect AFTER it. All comparisons are
/// on UTC instants — an event must never be assumed to fall on any particular
/// LOCAL calendar date (`sun_times` takes a UTC date and keys events to the
/// solar day: in UTC+13/+14 the events for UTC date d land on local d+1, and
/// near the arctic circle a sunset crosses local midnight).
fn transitions_window(now: DateTime<Utc>, lat: f64, lon: f64) -> Vec<(DateTime<Utc>, Theme)> {
    let base = now.date_naive();
    let mut events = Vec::with_capacity(6);
    for off in -1..=1 {
        let date = base + chrono::Duration::days(off);
        if let Some((sunrise, sunset)) = sun_times(date, lat, lon, 0.0) {
            events.push((sunrise, Theme::Light));
            events.push((sunset, Theme::Dark));
        }
    }
    events.sort_by_key(|&(t, _)| t);
    events
}

/// Theme implied by the sun's position alone.
fn altitude_theme(t: DateTime<Utc>, lat: f64, lon: f64) -> Theme {
    if solar_altitude_deg(t, lat, lon) > SUNRISE_ALTITUDE_DEG {
        Theme::Light
    } else {
        Theme::Dark
    }
}

/// First instant after `now` (within 48 h) where `altitude_theme` flips,
/// refined to 1 ms and returned on the far side of the crossing (so a tick
/// scheduled there already sees the new theme; the ms precision leaves a
/// negligible window in which the theme has flipped but `next` hasn't come).
///
/// Probes are at least 15 s apart — fine enough that whether a short polar
/// "day" is found no longer depends on where the search starts (with a
/// 5-minute floor, a 2-minute segment was found from some start times and
/// not others, breaking `schedule`'s contract) — and further apart when the
/// sun is far from the threshold: the solar altitude changes by at most
/// 15.04°/h × |cos(latitude)| (Earth's rotation; exactly that times
/// sin(azimuth)) plus a little declination drift, so the threshold can't be
/// reached sooner than |altitude − threshold| / that rate. Near the poles the
/// bound shrinks toward the declination drift, so months-long seasons cost a
/// few hundred to ~2k probes.
fn next_altitude_crossing(now: DateTime<Utc>, lat: f64, lon: f64) -> Option<DateTime<Utc>> {
    let max_deg_per_hour = 15.05 * lat.to_radians().cos().abs() + 0.02;
    let start = altitude_theme(now, lat, lon);
    let min_step = chrono::Duration::seconds(15);
    let end = now + chrono::Duration::hours(48);
    let mut lo = now;
    while lo < end {
        let margin = (solar_altitude_deg(lo, lat, lon) - SUNRISE_ALTITUDE_DEG).abs();
        let safe = chrono::Duration::seconds((margin / max_deg_per_hour * 3600.0) as i64);
        let hi = std::cmp::min(lo + std::cmp::max(min_step, safe), end);
        if altitude_theme(hi, lat, lon) != start {
            let (mut a, mut b) = (lo, hi);
            while b - a > chrono::Duration::milliseconds(1) {
                let mid = a + (b - a) / 2;
                if altitude_theme(mid, lat, lon) == start {
                    a = mid;
                } else {
                    b = mid;
                }
            }
            return Some(b);
        }
        lo = hi;
    }
    None
}

/// Latitude from which the schedule uses the solar-altitude model instead of
/// `sun_times`. With the −0.833° threshold and the 23.44° axial tilt, whole
/// days without a sunrise or sunset begin at ~65.73°; below that
/// `sun_times` always has events and stays in charge (Reykjavik, 64.1°,
/// included).
const POLAR_MODEL_LAT: f64 = 65.5;

/// Current theme and next transition instant — the single source of truth
/// for tick(). Contract (unit-tested by an edge sweep): the theme stays
/// `current` for every instant in [now, next), and `next` is a real flip.
///
/// Below POLAR_MODEL_LAT: from the sorted `sun_times` sunrise/sunset events
/// around `now`.
///
/// At polar latitudes: purely from the solar altitude — current state from
/// `altitude_theme`, next from the next altitude crossing, however far.
/// Mixing in `sun_times` there is what went wrong before v0.4.1: after the
/// last sunset before the midnight sun it reports no following sunrise
/// (~68 days stuck on Dark in Tromsø), and switching between the two models
/// at UTC midnights made them disagree — phantom transitions that reverted
/// overrides, and minute-long flickers. One model per location avoids all
/// of that; it differs from `sun_times` by at most a few minutes on ordinary
/// days.
fn schedule(now: DateTime<Utc>, lat: f64, lon: f64) -> (Theme, DateTime<Utc>) {
    if lat.abs() >= POLAR_MODEL_LAT {
        return (
            altitude_theme(now, lat, lon),
            next_altitude_flip(now, lat, lon),
        );
    }
    let window = transitions_window(now, lat, lon);
    let current = window
        .iter()
        .rev()
        .find(|&&(t, _)| t <= now)
        .map(|&(_, theme)| theme)
        .unwrap_or_else(|| altitude_theme(now, lat, lon));
    let next = window
        .iter()
        .find(|&&(t, _)| t > now)
        .map(|&(t, _)| t)
        .unwrap_or_else(|| next_transition_beyond_window(now, lat, lon));
    (current, next)
}

/// The next `altitude_theme` flip after `now`, searched in 48 h chunks up to
/// 200 days (the poles' seasons are ~6 months). Adaptive probe steps keep
/// even a months-long polar season to a few thousand altitude evaluations.
fn next_altitude_flip(now: DateTime<Utc>, lat: f64, lon: f64) -> DateTime<Utc> {
    let limit = now + chrono::Duration::days(200);
    let mut from = now;
    while from < limit {
        if let Some(t) = next_altitude_crossing(from, lat, lon) {
            return t;
        }
        from += chrono::Duration::hours(48);
    }
    // Unreachable on Earth; re-check tomorrow rather than never.
    now + chrono::Duration::days(1)
}

/// Safety net for the `sun_times` path: the first event beyond the ±1-day
/// window. (Below POLAR_MODEL_LAT the window always has events, so this
/// shouldn't run.)
fn next_transition_beyond_window(now: DateTime<Utc>, lat: f64, lon: f64) -> DateTime<Utc> {
    let base = now.date_naive();
    for off in 2..=200 {
        if let Some((sunrise, sunset)) =
            sun_times(base + chrono::Duration::days(off), lat, lon, 0.0)
        {
            if sunrise > now {
                return sunrise;
            }
            if sunset > now {
                return sunset;
            }
        }
    }
    now + chrono::Duration::days(1)
}

// === Wall clock vs. monotonic clock ===
//
// The schedule is wall-clock (sunrise at 05:43 local), but winit's WaitUntil
// is a monotonic Instant. The mapping between them breaks when the system
// clock is STEPPED — w32time correcting a PC that booted 3 h behind after an
// Ubuntu session (Ubuntu keeps the RTC in UTC), a manual change, a VM
// restore. Before v0.4.1 a deadline armed from a wrong clock fired up to
// hours late, and the Init tick's decision (made from the wrong time) stood
// until then.
//
// So every arm records a (wall, mono) mark, and every wake of the loop —
// the deadline itself, a heartbeat at most HEARTBEAT later, or any other
// message (including the WM_TIMECHANGE broadcast winit's hidden top-level
// window receives) — compares how far each clock advanced since the mark.
// A divergence beyond CLOCK_STEP_MS is a step: re-evaluate now. The
// heartbeat makes this independent of whether WM_TIMECHANGE is delivered.

/// Longest the loop ever sleeps without re-checking the clocks.
const HEARTBEAT: Duration = Duration::from_secs(10 * 60);
/// Wall-vs-monotonic divergence treated as a clock step (drift and w32time's
/// gradual slewing are far smaller).
const CLOCK_STEP_MS: i64 = 60_000;

/// Map a wall-clock deadline onto the monotonic clock. A deadline already in
/// the past maps to `now_mono` (fire immediately). Pure — unit-tested.
fn wall_to_instant(target: DateTime<Utc>, now_wall: DateTime<Utc>, now_mono: Instant) -> Instant {
    now_mono + (target - now_wall).to_std().unwrap_or(Duration::ZERO)
}

/// How much further the wall clock moved than the monotonic clock since the
/// mark, in ms. ~0 normally; +10_800_000 after a 3 h forward step; negative
/// after a backward step. Pure — unit-tested.
fn clock_step_ms(
    mark: (DateTime<Utc>, Instant),
    now_wall: DateTime<Utc>,
    now_mono: Instant,
) -> i64 {
    let wall_ms = (now_wall - mark.0).num_milliseconds();
    let mono_ms = now_mono.saturating_duration_since(mark.1).as_millis() as i64;
    wall_ms - mono_ms
}

/// The clock step since `mark`, if it exceeds CLOCK_STEP_MS (positive =
/// the wall clock jumped forward). Pure — unit-tested.
fn detected_clock_step(
    mark: Option<(DateTime<Utc>, Instant)>,
    now_wall: DateTime<Utc>,
    now_mono: Instant,
) -> Option<i64> {
    let step = clock_step_ms(mark?, now_wall, now_mono);
    (step.abs() > CLOCK_STEP_MS).then_some(step)
}

/// What to do when the loop wakes without a tick-worthy event.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum WakePlan {
    /// Nothing armed (no location → plain Wait).
    Idle,
    /// Due, or the clock was stepped: run a Scheduled tick. `step_ms` is set
    /// when a clock step was detected.
    Tick { step_ms: Option<i64> },
    /// Not due: sleep until this Instant (the deadline, or the heartbeat).
    Arm(Instant),
}

/// Decide a wake from the armed wall deadline and the arming mark. Pure —
/// unit-tested.
fn plan_wake(
    armed: Option<DateTime<Utc>>,
    mark: Option<(DateTime<Utc>, Instant)>,
    now_wall: DateTime<Utc>,
    now_mono: Instant,
) -> WakePlan {
    let Some(deadline) = armed else {
        return WakePlan::Idle;
    };
    if let Some(step) = detected_clock_step(mark, now_wall, now_mono) {
        return WakePlan::Tick {
            step_ms: Some(step),
        };
    }
    if now_wall >= deadline {
        return WakePlan::Tick { step_ms: None };
    }
    WakePlan::Arm(std::cmp::min(
        wall_to_instant(deadline, now_wall, now_mono),
        now_mono + HEARTBEAT,
    ))
}

/// Arm the loop for the wall-clock `deadline` (capped at the heartbeat) and
/// record `mark` — the (wall, mono) readings the tick DECIDED from — as the
/// reference clock steps are measured against. Taking the mark at decision
/// time, not here after the apply, means a step during a slow apply is still
/// detected at the next wake.
fn arm(
    elwt: &ActiveEventLoop,
    state: &mut TickState,
    deadline: DateTime<Utc>,
    mark: (DateTime<Utc>, Instant),
) {
    let (now_wall, now_mono) = (Utc::now(), Instant::now());
    state.armed = Some(deadline);
    state.mark = Some(mark);
    elwt.set_control_flow(ControlFlow::WaitUntil(std::cmp::min(
        wall_to_instant(deadline, now_wall, now_mono),
        now_mono + HEARTBEAT,
    )));
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

fn tick(cfg: &Config, elwt: &ActiveEventLoop, kind: TickKind, cause: &str, state: &mut TickState) {
    let now = Local::now();
    let now_mono = Instant::now();
    let now_str = now.to_rfc3339();

    if !cfg.has_location() {
        log_event(&format!("{} cause={} skipped=no-location", now_str, cause));
        state.armed = None;
        state.mark = None;
        elwt.set_control_flow(ControlFlow::Wait);
        return;
    }

    let now_utc = now.with_timezone(&Utc);
    let (want, next_utc) = schedule(now_utc, cfg.latitude, cfg.longitude);
    let current = current_theme();
    let next = next_utc.with_timezone(&Local);

    // A clock step since the PRE-tick mark is consumed by THIS tick (the
    // clock-jump tick, or a wake/Refresh that beat it) — log it here so every
    // consumer logs it once, and let a backward one enable the frame rule.
    let step = detected_clock_step(state.mark, now_utc, now_mono);
    if let Some(ms) = step {
        log_event(&format!(
            "{} clock_jump step_s={:+} consumed_by={}",
            Local::now().to_rfc3339(),
            ms / 1000,
            cause,
        ));
    }
    let clock_stepped_back = step.is_some_and(|ms| ms < 0);

    // Decide with the PRE-tick state; record the outcome after the apply
    // result is known (a failed apply must never advance reconciled_next —
    // see TickState).
    let action = decide_tick(
        kind,
        current,
        want,
        now_utc,
        next_utc,
        clock_stepped_back,
        state,
    );

    let mut retry_note = String::new();
    let mut deadline = next;
    let outcome = match action {
        TickAction::Apply => {
            if kind == TickKind::Refresh {
                // Fresh user intent (likely a just-fixed config): a Refresh
                // never inherits a burned retry budget.
                state.retry_count = 0;
                state.retry_baseline = None;
                state.episode_next = None;
            }
            match apply_theme(want, cfg) {
                Ok(method) => {
                    note_reconciled(state, next_utc);
                    format!("applied={}", method)
                }
                Err(e) => {
                    // Bounded retry: reschedule soon instead of silently
                    // waiting up to ~12 h for the next transition. The retry
                    // arrives as a normal ResumeTimeReached tick; the
                    // pending-retry gate in decide_tick stands it down if
                    // the user changes the theme in the meantime.
                    if note_apply_failed(state, current, next_utc) {
                        retry_note = format!(" retry={}", state.retry_count);
                        deadline = retry_deadline(Local::now(), next);
                    } else {
                        retry_note = " retry=exhausted".to_string();
                    }
                    format!("err=\"{}\"", sanitize_log_msg(&e.to_string()))
                }
            }
        }
        TickAction::SkipInSync => {
            note_reconciled(state, next_utc);
            "applied=skip".to_string()
        }
        TickAction::SkipOverride => {
            note_reconciled(state, next_utc);
            "applied=skip-override".to_string()
        }
        TickAction::CancelRetry => {
            note_reconciled(state, next_utc);
            "applied=skip-user-intervened".to_string()
        }
    };

    // Stamp at write time, not tick start — apply_theme logs detail lines
    // (theme_manager2_apply, theme_manager2_err) mid-tick, and reusing the
    // tick-start timestamp here made this summary line sort before them.
    log_event(&format!(
        "{} cause={} current={} target={} {}{} next={}",
        Local::now().to_rfc3339(),
        cause,
        theme_str(current),
        theme_str(Some(want)),
        outcome,
        retry_note,
        deadline.to_rfc3339(),
    ));

    arm(
        elwt,
        state,
        deadline.with_timezone(&Utc),
        (now_utc, now_mono),
    );
}

unsafe extern "system" fn wake_window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == WM_WTSSESSION_CHANGE && wparam == WTS_SESSION_UNLOCK {
        if let Some(proxy) = EVENT_PROXY.get() {
            let _ = proxy.send_event(AppEvent::Wake(WakeKind::Unlock));
        }
    }
    DefWindowProcW(hwnd, msg, wparam, lparam)
}

/// Suspend/resume callback (DEVICE_NOTIFY_CALLBACK registration). Runs on a
/// system thread; EventLoopProxy::send_event is thread-safe. Only
/// PBT_APMRESUMEAUTOMATIC matters: the system sends it on EVERY resume
/// (PBT_APMRESUMESUSPEND follows only after user input, and would just be a
/// duplicate, idempotent wake). Note this wake can run while the session is
/// still at the lock screen; the unlock that follows is then a SkipInSync.
unsafe extern "system" fn power_callback(
    _context: *const c_void,
    kind: u32,
    _setting: *const c_void,
) -> u32 {
    if kind == PBT_APMRESUMEAUTOMATIC {
        if let Some(proxy) = EVENT_PROXY.get() {
            let _ = proxy.send_event(AppEvent::Wake(WakeKind::Power));
        }
    }
    0 // ERROR_SUCCESS
}

/// Register for resume-from-sleep via the documented callback form.
/// PowerRegisterSuspendResumeNotification accepts ONLY DEVICE_NOTIFY_CALLBACK
/// — until v0.4.1 this passed a window handle instead, failed with 87
/// (ERROR_INVALID_PARAMETER) on every launch, and the resume hook never fired.
/// The callback form also avoids depending on a message-only window receiving
/// WM_POWERBROADCAST, which Microsoft doesn't document. The subscribe
/// parameters must outlive the registration, which lasts as long as the
/// process — both are intentionally leaked.
fn register_power_resume() {
    let params: &'static mut DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS =
        Box::leak(Box::new(DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS {
            Callback: Some(power_callback),
            Context: ptr::null_mut(),
        }));
    let mut handle: *mut c_void = ptr::null_mut();
    let rc = unsafe {
        PowerRegisterSuspendResumeNotification(
            DEVICE_NOTIFY_CALLBACK,
            params as *mut DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS as _,
            &mut handle,
        )
    };
    if rc != 0 {
        // The error code IS the return value (WIN32_ERROR), not GetLastError.
        log_event(&format!(
            "{} wake_listener_err stage=power_register code={}",
            Local::now().to_rfc3339(),
            rc,
        ));
    }
}

fn start_wake_listener() {
    // Independent of the unlock window below, and first — so a slow WTS
    // registration at logon can't delay it.
    register_power_resume();
    std::thread::spawn(|| {
        let class_name = wide("WinThemeSwitcherWakeListener");
        unsafe {
            let hinstance = GetModuleHandleW(ptr::null());
            let mut wc: WNDCLASSW = std::mem::zeroed();
            wc.lpfnWndProc = Some(wake_window_proc);
            wc.hInstance = hinstance;
            wc.lpszClassName = class_name.as_ptr();
            if RegisterClassW(&wc) == 0 {
                // Capture immediately: Local::now()/log_event make Win32
                // calls that clobber the thread's last error.
                let err = GetLastError();
                log_event(&format!(
                    "{} wake_listener_err stage=register_class code={}",
                    Local::now().to_rfc3339(),
                    err,
                ));
                // CreateWindowExW will fail below and log; fall through.
            }

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
                let err = GetLastError();
                log_event(&format!(
                    "{} wake_listener_err stage=create_window code={}",
                    Local::now().to_rfc3339(),
                    err,
                ));
                return;
            }
            // Failures below degrade wake coverage (unlock or resume events
            // won't arrive) but don't kill the listener thread — log each so
            // a missing wake-tick has a diagnosable trace instead of silence.
            //
            // WTSRegisterSessionNotification depends on the terminal-services
            // machinery, which may not be up yet when we auto-start at logon
            // via HKCU\Run — retry briefly before settling for the log line.
            for attempt in 1..=3 {
                if WTSRegisterSessionNotification(hwnd, NOTIFY_FOR_THIS_SESSION) != 0 {
                    if attempt > 1 {
                        log_event(&format!(
                            "{} wake_listener_wts_ok attempt={}",
                            Local::now().to_rfc3339(),
                            attempt,
                        ));
                    }
                    break;
                }
                let err = GetLastError();
                log_event(&format!(
                    "{} wake_listener_err stage=wts_register attempt={} code={}",
                    Local::now().to_rfc3339(),
                    attempt,
                    err,
                ));
                if attempt < 3 {
                    std::thread::sleep(Duration::from_secs(2));
                }
            }

            let mut msg: MSG = std::mem::zeroed();
            while GetMessageW(&mut msg, ptr::null_mut(), 0, 0) > 0 {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
    });
}

/// Install a panic hook that leaves a trace in events.log. With
/// `panic = "abort"` and the windowed subsystem, an unhooked panic is a
/// zero-trace process death.
fn install_panic_hook() {
    std::panic::set_hook(Box::new(|info| {
        // location() is compile-time data (survives strip = true) and often
        // the only actionable part; payload_as_str covers &str and String
        // panics, which is all this codebase produces.
        let at = info
            .location()
            .map(|l| format!("{}:{}:{}", l.file(), l.line(), l.column()))
            .unwrap_or_else(|| "unknown".to_string());
        let msg = info.payload_as_str().unwrap_or("<non-string panic>");
        log_event_from_panic(&format!(
            "{} panic at={} msg=\"{}\"",
            Local::now().to_rfc3339(),
            at,
            sanitize_log_msg(msg),
        ));
    }));
}

/// Claim the per-session single-instance mutex. Returns false when another
/// instance already holds it. The handle is intentionally leaked — it must
/// live exactly as long as the process.
fn claim_single_instance() -> bool {
    let name = wide("Local\\WinThemeSwitcher.single-instance");
    unsafe {
        let handle = CreateMutexW(ptr::null(), 0, name.as_ptr());
        let last = GetLastError();
        if (handle as usize) != 0 && last == ERROR_ALREADY_EXISTS {
            // Second instance; the OS closes the extra handle at process exit.
            return false;
        }
        if (handle as usize) == 0 {
            // Failing to create the mutex is no reason to refuse to run.
            log_event(&format!(
                "{} single_instance_err code={}",
                Local::now().to_rfc3339(),
                last,
            ));
        }
        true
    }
}

fn main() {
    // Runtime LoadLibrary calls by bare name (dependencies load e.g.
    // uxtheme.dll that way) search System32 only — never the exe's own
    // folder. .cargo\config.toml's /DEPENDENTLOADFLAG does the same for the
    // static imports. Everything this app loads is a system DLL; COM/WinRT
    // activation uses full paths.
    unsafe {
        SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32);
    }
    install_panic_hook();
    if !claim_single_instance() {
        log_event(&format!(
            "{} duplicate_instance action=exit",
            Local::now().to_rfc3339(),
        ));
        show_message_box(
            "WinThemeSwitcher",
            "WinThemeSwitcher is already running — look for its icon in the \
             notification area.",
            MB_OK | MB_ICONINFORMATION,
        );
        std::process::exit(0);
    }
    if let Err(e) = run() {
        // Fail loudly: tray creation racing the taskbar at login, event-loop
        // build errors, and event-loop death all used to be silent exits.
        let msg = sanitize_log_msg(&e.to_string());
        log_event(&format!(
            "{} fatal_error msg=\"{}\"",
            Local::now().to_rfc3339(),
            msg,
        ));
        show_message_box(
            "WinThemeSwitcher — Error",
            &format!(
                "WinThemeSwitcher stopped because of an error:\n\n{e}\n\n\
                 See events.log next to the exe for details."
            ),
            MB_OK | MB_ICONWARNING,
        );
        std::process::exit(1);
    }
}

fn run() -> Result<(), Box<dyn Error>> {
    ensure_com_initialized();

    // On a broken config file the session runs read-only against it: theme
    // switching continues with in-memory defaults + best-effort coordinates,
    // but nothing is persisted (acquire_location saves on success, which
    // would overwrite the very file the user needs to fix) and the autostart
    // registration is left exactly as the user last set it (the fallback
    // default auto_start=true must not override a broken file's false).
    let loaded = load_config_logged();
    // Whether `cfg` has ever come from a successfully parsed file this
    // session. Until it has, the autostart registration must be left as the
    // user last set it — the in-memory fallback's auto_start=true must never
    // override a broken file's false, on startup OR on a Refresh.
    let mut have_disk_cfg = loaded.is_ok();
    let mut cfg = match loaded {
        Ok(mut cfg) => {
            if !cfg.has_location() {
                acquire_location(&mut cfg);
            }
            apply_auto_start(cfg.auto_start);
            cfg
        }
        Err(e) => {
            report_config_error(&e);
            let mut cfg = Config::default();
            if let Some((lat, lon)) = try_get_windows_location() {
                cfg.latitude = lat;
                cfg.longitude = lon;
            }
            cfg
        }
    };

    let event_loop = EventLoop::<AppEvent>::with_user_event().build()?;
    let proxy = event_loop.create_proxy();
    let _ = EVENT_PROXY.set(event_loop.create_proxy());
    start_wake_listener();

    let tray_menu = Menu::new();
    let toggle_i = MenuItem::new("Toggle Theme", true, None);
    let open_cfg_i = MenuItem::new("Open Config", true, None);
    let refresh_i = MenuItem::new("Refresh", true, None);
    let quit_i = MenuItem::new("Quit", true, None);
    tray_menu.append_items(&[
        &toggle_i,
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
    // Without a handler, tray-icon queues every mouse event over the icon
    // (dozens of Move events per hover) in an unbounded channel that nothing
    // drains — a slow leak over long uptimes. Nothing here needs them.
    TrayIconEvent::set_event_handler(Some(|_: TrayIconEvent| {}));

    let toggle_id = toggle_i.id().clone();
    let open_cfg_id = open_cfg_i.id().clone();
    let refresh_id = refresh_i.id().clone();
    let quit_id = quit_i.id().clone();

    let mut state = TickState::new();

    // winit 0.30 deprecates EventLoop::run in favor of ApplicationHandler;
    // the migration is deferred until a winit 0.31 bump forces it (README →
    // Roadmap → Maintenance notes).
    #[allow(deprecated)]
    event_loop.run(move |event, elwt| match event {
        Event::NewEvents(StartCause::Init) => tick(&cfg, elwt, TickKind::Init, "init", &mut state),
        // The armed deadline — or the heartbeat, or a clock step. Only a due
        // deadline or a detected step ticks; a heartbeat just re-arms.
        Event::NewEvents(StartCause::ResumeTimeReached { .. }) => {
            match plan_wake(state.armed, state.mark, Utc::now(), Instant::now()) {
                WakePlan::Tick { step_ms } => {
                    // tick() logs the clock_jump line itself.
                    let cause = if step_ms.is_some() {
                        "clock-jump"
                    } else {
                        "resume-time"
                    };
                    tick(&cfg, elwt, TickKind::Scheduled, cause, &mut state);
                }
                WakePlan::Arm(i) => elwt.set_control_flow(ControlFlow::WaitUntil(i)),
                WakePlan::Idle => {}
            }
        }
        // Any other wake (the WM_TIMECHANGE broadcast, tray/menu input):
        // never a tick (tick-scope invariant). If plan_wake says Tick (a
        // clock step, or the wall-clock deadline is already due), resume
        // immediately so the ResumeTimeReached arm ticks; otherwise re-arm
        // (idempotent).
        Event::NewEvents(StartCause::WaitCancelled { .. }) => {
            match plan_wake(state.armed, state.mark, Utc::now(), Instant::now()) {
                WakePlan::Tick { .. } => {
                    elwt.set_control_flow(ControlFlow::WaitUntil(Instant::now()));
                }
                WakePlan::Arm(i) => elwt.set_control_flow(ControlFlow::WaitUntil(i)),
                WakePlan::Idle => {}
            }
        }
        Event::UserEvent(AppEvent::Wake(kind)) => {
            let cause = match kind {
                WakeKind::Unlock => "wake-unlock",
                WakeKind::Power => "wake-power",
            };
            tick(&cfg, elwt, TickKind::Wake, cause, &mut state);
        }
        Event::UserEvent(AppEvent::Menu(id)) => {
            if id == quit_id {
                elwt.exit();
            } else if id == open_cfg_id {
                open_config_in_editor();
            } else if id == toggle_id {
                // A deliberate manual override: applies the opposite theme
                // and intentionally does NOT tick, touch TickState (it
                // tracks reconciliation with the schedule, not the screen),
                // or disturb the pending WaitUntil — so the override
                // survives lock/unlock (see decide_tick) and resets at the
                // next natural transition, exactly like an override made in
                // Settings. If a failed-apply retry is pending, one toggle
                // moves the screen off the retry baseline and the
                // pending-retry gate stands the retry down; a SECOND toggle
                // lands back on the baseline, which the gate can't tell from
                // the original failure, so the retry proceeds (known gap).
                let before = current_theme();
                let target = toggle_target(before);
                let outcome = match apply_theme(target, &cfg) {
                    Ok(method) => format!("applied={}", method),
                    Err(e) => format!("err=\"{}\"", sanitize_log_msg(&e.to_string())),
                };
                log_event(&format!(
                    "{} cause=toggle current={} target={} {}",
                    Local::now().to_rfc3339(),
                    theme_str(before),
                    theme_str(Some(target)),
                    outcome,
                ));
            } else if id == refresh_id {
                let (assert_autostart, error) =
                    adopt_reloaded_config(load_config_logged(), &mut cfg, &mut have_disk_cfg);
                match error {
                    Some(e) => report_config_error(&e),
                    None => {
                        if !cfg.has_location() {
                            if let Some((lat, lon)) = try_get_windows_location() {
                                cfg.latitude = lat;
                                cfg.longitude = lon;
                                persist_config(&cfg);
                            }
                        }
                    }
                }
                // Refresh re-asserting the Run value from the (possibly
                // last-known-good) config is the documented recovery path
                // when e.g. an AV quarantine deletes it — but never from the
                // in-memory fallback of a broken-config session.
                if assert_autostart {
                    apply_auto_start(cfg.auto_start);
                }
                tick(&cfg, elwt, TickKind::Refresh, "refresh", &mut state);
            }
        }
        _ => {}
    })?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::TimeZone;

    fn utc(y: i32, mo: u32, d: u32, h: u32, mi: u32, s: u32) -> DateTime<Utc> {
        Utc.with_ymd_and_hms(y, mo, d, h, mi, s).unwrap()
    }

    fn mins(m: i64) -> chrono::Duration {
        chrono::Duration::minutes(m)
    }

    // Apia, Samoa — UTC+13, west of the antimeridian. Regression fixture for
    // the wrong-solar-day bug: the old code passed the LOCAL date to
    // sun_times (which wants a UTC date), so every event landed on the wrong
    // local day and the app was permanently dark here.
    const APIA: (f64, f64) = (-13.83, -171.77);
    // Riyadh — baseline the deployed build has verified for months.
    const RIYADH: (f64, f64) = (24.753, 46.765);
    // Reykjavik — UTC+0 year-round; in June sunset falls just past midnight.
    const REYKJAVIK: (f64, f64) = (64.147, -21.94);
    // Tromsø — above the arctic circle: midnight sun in June, polar night in December.
    const TROMSO: (f64, f64) = (69.65, 18.96);

    #[test]
    fn apia_noon_is_light() {
        // 2026-01-15 12:00 local (UTC+13) = 2026-01-14 23:00 UTC
        let now = utc(2026, 1, 14, 23, 0, 0);
        let (theme, next) = schedule(now, APIA.0, APIA.1);
        assert_eq!(theme, Theme::Light);
        // next transition is that evening's sunset (~19:10 local)
        assert!(next > now && next - now < mins(9 * 60), "next = {next}");
    }

    #[test]
    fn apia_night_is_dark() {
        // 2026-01-15 22:00 local = 2026-01-15 09:00 UTC
        let now = utc(2026, 1, 15, 9, 0, 0);
        let (theme, next) = schedule(now, APIA.0, APIA.1);
        assert_eq!(theme, Theme::Dark);
        // next transition is the ~06:20 local sunrise
        assert!(next > now && next - now < mins(9 * 60), "next = {next}");
    }

    #[test]
    fn riyadh_matches_deployed_log() {
        // The deployed build logged next=2026-07-04T18:46:05+03:00 (15:46:05Z)
        // for a mid-day tick. Riyadh is a same-day timezone, where the old
        // math was correct — the new math must agree with it.
        let now = utc(2026, 7, 4, 9, 0, 0); // 12:00 local
        let (theme, next) = schedule(now, RIYADH.0, RIYADH.1);
        assert_eq!(theme, Theme::Light);
        let expected = utc(2026, 7, 4, 15, 46, 5);
        assert!((next - expected).abs() < mins(5), "next = {next}");
    }

    #[test]
    fn riyadh_evening_is_dark_until_sunrise() {
        let now = utc(2026, 7, 4, 19, 0, 0); // 22:00 local
        let (theme, next) = schedule(now, RIYADH.0, RIYADH.1);
        assert_eq!(theme, Theme::Dark);
        // sunrise is ~05:35 local = 02:35Z, ~7.6 h away
        assert!(next > now && next - now < mins(11 * 60), "next = {next}");
    }

    #[test]
    fn reykjavik_june_sunset_crosses_midnight() {
        // Sun sets a few minutes past local midnight on June 21; at 23:30 on
        // June 20 it is still up. The old single-local-date math missed the
        // post-midnight sunset entirely.
        let now = utc(2026, 6, 20, 23, 30, 0);
        let (theme, next) = schedule(now, REYKJAVIK.0, REYKJAVIK.1);
        assert_eq!(theme, Theme::Light);
        assert!(
            next - now < mins(120),
            "sunset should be < 2h away, next = {next}"
        );
        // Just after that sunset: dark until the ~03:00 sunrise.
        let later = next + mins(1);
        let (theme2, next2) = schedule(later, REYKJAVIK.0, REYKJAVIK.1);
        assert_eq!(theme2, Theme::Dark);
        assert!(
            next2 > later && next2 - later < mins(4 * 60),
            "next2 = {next2}"
        );
    }

    #[test]
    fn tromso_midnight_sun_is_light_with_far_next() {
        let now = utc(2026, 6, 21, 12, 0, 0);
        let (theme, next) = schedule(now, TROMSO.0, TROMSO.1);
        assert_eq!(theme, Theme::Light);
        // Polar day runs to ~late July — the next transition is weeks away
        // and must come from the forward scan, not a 24 h fallback.
        assert!(next - now > mins(5 * 24 * 60), "next = {next}");
        assert!(next - now < mins(60 * 24 * 60), "next = {next}");
    }

    #[test]
    fn tromso_polar_night_is_dark() {
        let now = utc(2026, 12, 21, 12, 0, 0);
        let (theme, next) = schedule(now, TROMSO.0, TROMSO.1);
        assert_eq!(theme, Theme::Dark);
        // Polar night ends mid-January (first sun_times event 2026-01-15 in
        // the previous season) — the forward scan must find it, not fall
        // back to a 24 h guess or run past the season.
        assert!(next - now > mins(10 * 24 * 60), "next = {next}");
        assert!(next - now < mins(40 * 24 * 60), "next = {next}");
    }

    // --- polar boundaries (v0.4.1) ---
    // Tromsø 2026: sun_times' last sunset before the midnight sun is
    // 2026-05-18T22:31:36Z, then None until the first events again on
    // 2026-07-26; polar night: last events 2026-11-27, None from 11-28 until
    // 2027-01-15. Expected instants below were cross-checked against an
    // independent port of sun_times + the altitude model (±5 min).

    /// Assert `schedule(at)` = (theme, next ≈ expect ± 5 min).
    fn assert_sched(at: DateTime<Utc>, loc: (f64, f64), theme: Theme, expect: DateTime<Utc>) {
        let (t, n) = schedule(at, loc.0, loc.1);
        assert_eq!(t, theme, "theme at {at}");
        assert!(
            (n - expect).abs() < mins(5),
            "next at {at} = {n}, expected ≈ {expect}"
        );
    }

    #[test]
    fn tromso_after_last_sunset_before_midnight_sun_is_not_dark_for_months() {
        // Regression: "the last event was a sunset" + no following sunrise
        // from sun_times made this Dark with next ≈ 68 days away. By the
        // altitude model the sun never gets below the threshold that night,
        // and the next real darkening is at the end of the polar day.
        let now = utc(2026, 5, 18, 22, 32, 36);
        let (theme, next) = schedule(now, TROMSO.0, TROMSO.1);
        assert_eq!(theme, Theme::Light);
        assert!(next - now > mins(30 * 24 * 60), "next = {next}");
        assert_sched(now, TROMSO, Theme::Light, utc(2026, 7, 25, 22, 37, 19));
    }

    #[test]
    fn tromso_polar_day_onset_details() {
        // By the altitude model the last darkness before the midnight sun is
        // 05-17 22:29–22:52Z; from then on the next flip is the end of the
        // polar day — no phantom transition at sun_times' 05-18 "sunset".
        assert_sched(
            utc(2026, 5, 17, 23, 59, 59),
            TROMSO,
            Theme::Light,
            utc(2026, 7, 25, 22, 37, 19),
        );
        assert_eq!(
            schedule(utc(2026, 5, 18, 22, 31, 36), TROMSO.0, TROMSO.1).0,
            Theme::Light
        );
        assert_eq!(
            schedule(utc(2026, 5, 19, 11, 31, 0), TROMSO.0, TROMSO.1).0,
            Theme::Light
        );
    }

    #[test]
    fn polar_model_threshold_and_southern_hemisphere() {
        // At exactly 65.5° the altitude model applies (>=); values from an
        // independent port, ±5 min.
        assert_sched(
            utc(2026, 6, 21, 0, 0, 0),
            (65.5, 0.0),
            Theme::Dark,
            utc(2026, 6, 21, 0, 35, 12),
        );
        assert_sched(
            utc(2026, 12, 21, 12, 0, 0),
            (-65.5, 0.0),
            Theme::Light,
            utc(2026, 12, 21, 23, 24, 50),
        );
        // Above the polar-day onset latitude the midnight sun is found
        // (guards against moving the threshold past ~65.73°).
        assert_sched(
            utc(2026, 6, 21, 0, 0, 0),
            (65.8, 0.0),
            Theme::Light,
            utc(2026, 6, 25, 23, 57, 18),
        );
    }

    #[test]
    fn short_polar_segments_are_found_from_any_start() {
        // Near the polar-night edge the altitude model can produce a day of
        // only minutes. Whether it's found must not depend on where the
        // search starts, or schedule's contract breaks.
        let starts = [
            utc(2026, 12, 2, 20, 0, 0),
            utc(2026, 12, 3, 6, 0, 0),
            utc(2026, 12, 3, 11, 2, 0),
        ];
        let nexts: Vec<_> = starts
            .iter()
            .map(|&t| schedule(t, TROMSO.0, TROMSO.1).1)
            .collect();
        for w in nexts.windows(2) {
            assert!(
                (w[0] - w[1]).abs() < chrono::Duration::seconds(1),
                "{nexts:?}"
            );
        }
    }

    #[test]
    fn below_the_polar_model_latitude_sun_times_still_decides() {
        // Reykjavik (64.1°) keeps sun_times' exact events.
        let now = utc(2026, 7, 4, 12, 0, 0);
        let (_, next) = schedule(now, REYKJAVIK.0, REYKJAVIK.1);
        let window = transitions_window(now, REYKJAVIK.0, REYKJAVIK.1);
        assert!(
            window.iter().any(|&(t, _)| t == next),
            "next {next} must be a sun_times event"
        );
    }

    #[test]
    fn tromso_polar_day_end_uses_the_first_real_darkening() {
        // The old forward scan pointed at sun_times' 23:05 "sunrise" while
        // it was still light; the altitude model darkens at ~22:37.
        assert_sched(
            utc(2026, 7, 25, 12, 0, 0),
            TROMSO,
            Theme::Light,
            utc(2026, 7, 25, 22, 37, 19),
        );
        assert_sched(
            utc(2026, 7, 25, 22, 45, 0),
            TROMSO,
            Theme::Dark,
            utc(2026, 7, 25, 23, 4, 59),
        );
        assert_sched(
            utc(2026, 6, 21, 12, 0, 0),
            TROMSO,
            Theme::Light,
            utc(2026, 7, 25, 22, 37, 19),
        );
    }

    #[test]
    fn tromso_polar_night_boundaries() {
        assert_sched(
            utc(2026, 11, 27, 10, 31, 0),
            TROMSO,
            Theme::Light,
            utc(2026, 11, 27, 10, 42, 38),
        );
        let (t, n) = schedule(utc(2026, 11, 27, 12, 0, 0), TROMSO.0, TROMSO.1);
        assert_eq!(t, Theme::Dark);
        let days = (n - utc(2026, 11, 27, 12, 0, 0)).num_days();
        assert!((40..=60).contains(&days), "next = {n}");
        // End of the polar night: the first real (short) day on 01-15, not
        // sun_times' first event on 01-16.
        assert_sched(
            utc(2027, 1, 14, 12, 0, 0),
            TROMSO,
            Theme::Dark,
            utc(2027, 1, 15, 10, 34, 44),
        );
        assert_sched(
            utc(2027, 1, 15, 10, 54, 0),
            TROMSO,
            Theme::Light,
            utc(2027, 1, 15, 11, 12, 55),
        );
        assert_sched(
            utc(2026, 12, 21, 12, 0, 0),
            TROMSO,
            Theme::Dark,
            utc(2027, 1, 15, 10, 34, 44),
        );
    }

    #[test]
    fn polar_day_straddling_utc_midnight_is_consistent_pevek() {
        // Pevek (UTC+12, east Siberia): the last short day before the polar
        // night straddles 00:00 UTC. Mixing sun_times and the altitude model
        // at that midnight once made a tick arm a `next` the other model
        // disagreed with (a wake in the gap "preserved" Light for 50 days).
        const PEVEK: (f64, f64) = (69.70, 170.31);
        let at = utc(2026, 11, 26, 23, 59, 52);
        let (theme, next) = schedule(at, PEVEK.0, PEVEK.1);
        assert_eq!(theme, altitude_theme(at, PEVEK.0, PEVEK.1));
        let horizon = next.min(at + mins(120));
        for m in (1..)
            .map(|k| at + chrono::Duration::seconds(30 * k))
            .take_while(|w| *w < horizon)
        {
            assert_eq!(
                schedule(m, PEVEK.0, PEVEK.1).0,
                theme,
                "theme changed at {m} before next = {next}"
            );
        }
        assert_ne!(
            schedule(next, PEVEK.0, PEVEK.1).0,
            theme,
            "next must be a real flip"
        );
    }

    /// The `schedule` contract, swept around polar edges: for sampled t,
    /// `next` is a real flip, and the theme stays schedule(t).0 at every
    /// sampled w in (t, next) — no phantom transitions, no flicker.
    fn sweep_contract(loc: (f64, f64), center: DateTime<Utc>) {
        let mut t = center - mins(2 * 24 * 60);
        while t < center + mins(2 * 24 * 60) {
            let (theme, next) = schedule(t, loc.0, loc.1);
            assert!(next > t, "next {next} not after {t}");
            assert_ne!(
                schedule(next, loc.0, loc.1).0,
                theme,
                "schedule({t}) = ({theme:?}, {next}) but {next} is not a flip"
            );
            let horizon = next.min(t + mins(36 * 60));
            let mut w = t + mins(20);
            while w < horizon {
                assert_eq!(
                    schedule(w, loc.0, loc.1).0,
                    theme,
                    "schedule({t}) = ({theme:?}, {next}) but theme differs at {w}"
                );
                w += mins(20);
            }
            t += mins(120);
        }
    }

    #[test]
    fn schedule_contract_holds_around_polar_edges() {
        sweep_contract(TROMSO, utc(2026, 5, 19, 0, 0, 0));
        sweep_contract(TROMSO, utc(2026, 7, 26, 0, 0, 0));
        sweep_contract(TROMSO, utc(2026, 11, 28, 0, 0, 0));
        sweep_contract(TROMSO, utc(2027, 1, 16, 0, 0, 0));
        sweep_contract((69.70, 170.31), utc(2026, 11, 28, 0, 0, 0));
        sweep_contract((-77.85, 166.67), utc(2026, 4, 25, 0, 0, 0));
    }

    #[test]
    fn tromso_first_polar_day_noon_is_light() {
        let now = utc(2026, 5, 19, 10, 45, 0);
        let (theme, next) = schedule(now, TROMSO.0, TROMSO.1);
        assert_eq!(theme, Theme::Light);
        assert!(next > now);
    }

    #[test]
    fn altitude_crossing_lands_on_the_far_side() {
        // Riyadh: an ordinary day has a crossing within 48 h, and the
        // returned instant already shows the flipped theme (so a tick
        // scheduled there can't hot-loop on the old one).
        let now = utc(2026, 7, 4, 9, 0, 0);
        let start = altitude_theme(now, RIYADH.0, RIYADH.1);
        let t = next_altitude_crossing(now, RIYADH.0, RIYADH.1).expect("crossing");
        assert!(t > now && t - now < mins(12 * 60), "t = {t}");
        assert_ne!(altitude_theme(t, RIYADH.0, RIYADH.1), start);
        assert_eq!(
            altitude_theme(t - chrono::Duration::seconds(2), RIYADH.0, RIYADH.1),
            start
        );
        // …and it agrees with sun_times' sunset to within a few minutes.
        let (_, sched_next) = schedule(now, RIYADH.0, RIYADH.1);
        assert!(
            (t - sched_next).abs() < mins(5),
            "t = {t}, schedule = {sched_next}"
        );
    }

    #[test]
    fn deep_polar_day_has_no_crossing_within_48h() {
        let now = utc(2026, 6, 21, 12, 0, 0);
        assert_eq!(next_altitude_crossing(now, TROMSO.0, TROMSO.1), None);
    }

    #[test]
    fn theme_flips_exactly_at_transition() {
        let now = utc(2026, 7, 4, 9, 0, 0);
        let (_, next) = schedule(now, RIYADH.0, RIYADH.1);
        let (before, _) = schedule(next - mins(1), RIYADH.0, RIYADH.1);
        let (at, next_after) = schedule(next, RIYADH.0, RIYADH.1);
        assert_eq!(before, Theme::Light);
        assert_eq!(at, Theme::Dark);
        assert!(next_after > next);
    }

    /// Temp file that cleans up after itself. Uniqueness comes from the
    /// caller-supplied `name` — every test must pass a distinct one or the
    /// parallel runner will clobber files across tests. The pid only guards
    /// against two simultaneous `cargo test` processes.
    struct TempConfig(PathBuf);

    impl TempConfig {
        fn new(name: &str, content: Option<&str>) -> Self {
            let path = std::env::temp_dir().join(format!(
                "wts-test-{}-{}-config.json",
                std::process::id(),
                name
            ));
            let _ = fs::remove_file(&path);
            if let Some(c) = content {
                fs::write(&path, c).unwrap();
            }
            Self(path)
        }
    }

    impl Drop for TempConfig {
        fn drop(&mut self) {
            let _ = fs::remove_file(&self.0);
        }
    }

    #[test]
    fn config_missing_file_creates_defaults() {
        let t = TempConfig::new("missing", None);
        let cfg = load_config_at(&t.0).expect("first run must succeed");
        assert!(!cfg.has_location());
        assert!(cfg.auto_start);
        // First run writes the defaults file, and it must round-trip.
        let written = fs::read_to_string(&t.0).expect("defaults file must be created");
        let reparsed: Config = serde_json::from_str(&written).unwrap();
        assert!(!reparsed.has_location());
    }

    #[test]
    fn config_valid_file_loads_values_and_stays_untouched() {
        let json = r#"{
  "latitude": 24.753,
  "longitude": 46.765,
  "auto_start": false,
  "theme_day": "C:\\Themes\\day.theme",
  "theme_night": null
}"#;
        let t = TempConfig::new("valid", Some(json));
        let cfg = load_config_at(&t.0).expect("valid config must parse");
        assert_eq!(cfg.latitude, 24.753);
        assert_eq!(cfg.longitude, 46.765);
        assert!(!cfg.auto_start);
        assert_eq!(cfg.theme_day.as_deref(), Some("C:\\Themes\\day.theme"));
        // Loading must never rewrite a readable file — byte-identical.
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json);
    }

    #[test]
    fn config_parse_error_is_reported_and_file_kept() {
        // The roadmap's exact failure: a hand-edited theme path with single
        // backslashes is invalid JSON. This must NOT be reset to defaults.
        let json =
            r#"{"latitude": 24.7, "longitude": 46.7, "theme_night": "C:\Tools\night.theme"}"#;
        let t = TempConfig::new("broken", Some(json));
        let err = load_config_at(&t.0).expect_err("parse error must be reported, not defaulted");
        assert!(!err.is_empty());
        assert_eq!(
            fs::read_to_string(&t.0).unwrap(),
            json,
            "a broken config file must be left byte-identical on disk"
        );
    }

    #[test]
    fn config_empty_file_is_healed_to_defaults() {
        // A crash mid-write (fs::write truncates first) leaves a 0-byte
        // config.json. There is nothing in it to preserve, so it must
        // self-heal like first run instead of erroring on every launch.
        let t = TempConfig::new("empty", Some("  \n"));
        let cfg = load_config_at(&t.0).expect("empty file must heal to defaults");
        assert!(!cfg.has_location());
        assert!(cfg.auto_start);
        let written = fs::read_to_string(&t.0).unwrap();
        let reparsed: Config = serde_json::from_str(&written).unwrap();
        assert!(!reparsed.has_location());
    }

    #[test]
    fn config_unknown_and_missing_fields_are_defaulted() {
        let json = r#"{"latitude": 1.0, "some_future_field": true}"#;
        let t = TempConfig::new("partial", Some(json));
        let cfg = load_config_at(&t.0).expect("unknown/missing fields are not errors");
        assert_eq!(cfg.latitude, 1.0);
        assert_eq!(cfg.longitude, 0.0);
        assert!(cfg.auto_start, "missing fields fall back to defaults");
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json);
    }

    #[test]
    fn config_with_utf8_bom_parses_and_stays_untouched() {
        // Notepad / PowerShell 5.1 Out-File can save UTF-8 with a BOM.
        let json = "\u{feff}{\"latitude\": 24.753, \"longitude\": 46.765}";
        let t = TempConfig::new("bom", Some(json));
        let cfg = load_config_at(&t.0).expect("a BOM must not make valid JSON invalid");
        assert_eq!(cfg.latitude, 24.753);
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json);
    }

    #[test]
    fn config_nul_filled_file_is_healed_to_defaults() {
        // A crash after the file was extended but before data was flushed
        // can leave it full of NULs — as unpreservable as an empty file.
        let t = TempConfig::new("nul", Some("\0\0\0\0\0\0\0\0"));
        let cfg = load_config_at(&t.0).expect("NUL-filled file must heal to defaults");
        assert!(!cfg.has_location());
        let reparsed: Config = serde_json::from_str(&fs::read_to_string(&t.0).unwrap()).unwrap();
        assert!(!reparsed.has_location());
    }

    #[test]
    fn config_utf16_file_loads() {
        // PowerShell 5.1's `>` / Out-File default is UTF-16LE with a BOM.
        let text = "{\"latitude\": 24.753, \"longitude\": 46.765}";
        let mut bytes = vec![0xFF, 0xFE];
        for u in text.encode_utf16() {
            bytes.extend_from_slice(&u.to_le_bytes());
        }
        let t = TempConfig::new("utf16", None);
        fs::write(&t.0, &bytes).unwrap();
        let cfg = load_config_at(&t.0).expect("UTF-16 config must load");
        assert_eq!(cfg.longitude, 46.765);
        assert_eq!(
            fs::read(&t.0).unwrap(),
            bytes,
            "loading must not rewrite it"
        );
    }

    #[test]
    fn config_utf16_be_file_loads_and_truncated_utf16_is_an_error() {
        let text = "{\"latitude\": 24.753, \"longitude\": 46.765}";
        let mut be = vec![0xFE, 0xFF];
        for u in text.encode_utf16() {
            be.extend_from_slice(&u.to_be_bytes());
        }
        let t = TempConfig::new("utf16be", None);
        fs::write(&t.0, &be).unwrap();
        assert_eq!(
            load_config_at(&t.0).expect("UTF-16 BE must load").latitude,
            24.753
        );
        let mut odd = be.clone();
        odd.push(0x00);
        let t2 = TempConfig::new("utf16-odd", None);
        fs::write(&t2.0, &odd).unwrap();
        assert!(
            load_config_at(&t2.0).is_err(),
            "truncated UTF-16 is broken, not healed"
        );
        assert_eq!(fs::read(&t2.0).unwrap(), odd);
    }

    #[test]
    fn save_config_at_writes_through_hard_links() {
        // A package manager's "persisted" config may be a hard link; a
        // rename-over save would silently detach it.
        let t = TempConfig::new("hardlink", Some("{\"latitude\": 1.0, \"longitude\": 2.0}"));
        let other = TempConfig::new("hardlink-other", None);
        fs::hard_link(&t.0, &other.0).unwrap();
        let cfg = Config {
            latitude: 24.753,
            longitude: 46.765,
            ..Config::default()
        };
        save_config_at(&t.0, &cfg).unwrap();
        assert_eq!(
            load_config_at(&other.0).unwrap().latitude,
            24.753,
            "the other link must see the saved content"
        );
    }

    #[test]
    fn config_partial_json_with_nul_padding_is_an_error_not_a_heal() {
        // The heal predicate must stay narrow: real content followed by NULs
        // is a broken file to preserve, not an empty one to reset.
        let mut bytes = b"{\"latitude\": 24.7".to_vec();
        bytes.extend(std::iter::repeat_n(0u8, 64));
        let t = TempConfig::new("nul-partial", None);
        fs::write(&t.0, &bytes).unwrap();
        assert!(load_config_at(&t.0).is_err());
        assert_eq!(fs::read(&t.0).unwrap(), bytes);
    }

    #[test]
    fn config_unknown_keys_survive_a_save() {
        let json =
            r#"{"latitude": 1.0, "longitude": 2.0, "my_note": "keep me", "future": {"a": 1}}"#;
        let t = TempConfig::new("extra", Some(json));
        let mut cfg = load_config_at(&t.0).unwrap();
        cfg.latitude = 3.0;
        save_config_at(&t.0, &cfg).unwrap();
        let back: serde_json::Value =
            serde_json::from_str(&fs::read_to_string(&t.0).unwrap()).unwrap();
        assert_eq!(back["latitude"], 3.0);
        assert_eq!(back["my_note"], "keep me");
        assert_eq!(back["future"]["a"], 1);
    }

    #[test]
    fn load_config_ex_reports_heals_and_their_save_failures() {
        let valid = TempConfig::new("heal-valid", Some("{\"latitude\": 1.0}"));
        assert!(load_config_ex(&valid.0).unwrap().1.is_none());
        let missing = TempConfig::new("heal-missing", None);
        assert!(
            load_config_ex(&missing.0).unwrap().1.is_none(),
            "first run is not a heal"
        );
        let empty = TempConfig::new("heal-empty", Some("\0\0 \n"));
        let h = load_config_ex(&empty.0).unwrap().1.expect("heal reported");
        assert_eq!(h.reason, "empty");
        assert!(h.save_err.is_none());
        // Healing a file that can't be written: reported, file untouched.
        let ro = TempConfig::new("heal-readonly", Some("  "));
        let mut perms = fs::metadata(&ro.0).unwrap().permissions();
        perms.set_readonly(true);
        fs::set_permissions(&ro.0, perms.clone()).unwrap();
        let result = load_config_ex(&ro.0);
        #[allow(clippy::permissions_set_readonly_false)]
        perms.set_readonly(false);
        fs::set_permissions(&ro.0, perms).unwrap();
        let h = result.unwrap().1.expect("heal reported");
        assert!(
            h.save_err.is_some(),
            "the failed write-back must be reported"
        );
        assert_eq!(fs::read_to_string(&ro.0).unwrap(), "  ");
    }

    #[test]
    fn save_config_at_fails_cleanly_on_a_read_only_target() {
        let json = "{\"latitude\": 1.0, \"longitude\": 2.0}";
        let t = TempConfig::new("readonly", Some(json));
        let mut perms = fs::metadata(&t.0).unwrap().permissions();
        perms.set_readonly(true);
        fs::set_permissions(&t.0, perms.clone()).unwrap();
        let result = save_config_at(&t.0, &Config::default());
        #[allow(clippy::permissions_set_readonly_false)]
        perms.set_readonly(false);
        fs::set_permissions(&t.0, perms).unwrap();
        assert!(result.is_err(), "replacing a read-only file must fail");
        assert!(
            !t.0.with_extension("json.tmp").exists(),
            "no stray temp file"
        );
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json, "target untouched");
    }

    #[test]
    fn refresh_adopts_good_configs_and_gates_autostart() {
        let good = |auto: bool| Config {
            latitude: 1.0,
            longitude: 2.0,
            auto_start: auto,
            ..Config::default()
        };
        // (a) Broken at startup, still broken on Refresh: keep the fallback,
        // and never re-assert autostart from its auto_start=true default.
        let mut cfg = Config::default();
        let mut have = false;
        let (assert_auto, err) = adopt_reloaded_config(Err("bad".into()), &mut cfg, &mut have);
        assert!(!assert_auto && err.is_some() && !have);
        // (b) Broken at startup, fixed before Refresh: adopt it, re-assert.
        let (assert_auto, err) = adopt_reloaded_config(Ok(good(false)), &mut cfg, &mut have);
        assert!(assert_auto && err.is_none() && have);
        assert!(!cfg.auto_start);
        // (c) Good earlier, broken now: keep last-known-good and re-assert
        // from it (the documented AV-quarantine recovery).
        let (assert_auto, err) = adopt_reloaded_config(Err("bad".into()), &mut cfg, &mut have);
        assert!(assert_auto && err.is_some());
        assert!(!cfg.auto_start, "last-known-good retained");
    }

    #[test]
    fn save_config_at_replaces_atomically_and_leaves_no_temp() {
        let t = TempConfig::new("atomic", Some("{\"latitude\": 1.0, \"longitude\": 2.0}"));
        let cfg = Config {
            latitude: 24.753,
            longitude: 46.765,
            auto_start: false,
            ..Config::default()
        };
        save_config_at(&t.0, &cfg).expect("save must succeed");
        let back = load_config_at(&t.0).expect("saved file must load");
        assert_eq!(back.latitude, 24.753);
        assert!(!back.auto_start);
        assert!(
            !t.0.with_extension("json.tmp").exists(),
            "the temp file must be renamed away, not left behind"
        );
    }

    #[test]
    fn solar_altitude_sanity() {
        // Riyadh at local solar noon in July: sun nearly overhead (~88°).
        assert!(solar_altitude_deg(utc(2026, 7, 4, 8, 53, 0), RIYADH.0, RIYADH.1) > 80.0);
        // Tromsø, December noon: polar night — below the sunrise threshold.
        assert!(
            solar_altitude_deg(utc(2026, 12, 21, 11, 0, 0), TROMSO.0, TROMSO.1)
                < SUNRISE_ALTITUDE_DEG
        );
        // Tromsø, June, near local solar midnight: midnight sun stays up (~3°).
        assert!(solar_altitude_deg(utc(2026, 6, 20, 22, 45, 0), TROMSO.0, TROMSO.1) > 0.0);
    }

    // --- tick decision: manual-override preservation (v0.4.0) ---

    /// decide_tick with a schedule `next` that doesn't invoke the
    /// backward-step frame rule (it equals the recorded transition, as on any
    /// tick with an unstepped clock). Frame-rule tests call decide_tick
    /// directly.
    fn dt(
        kind: TickKind,
        current: Option<Theme>,
        target: Theme,
        now: DateTime<Utc>,
        s: &TickState,
    ) -> TickAction {
        let next = s
            .reconciled_next
            .filter(|&n| n > now)
            .unwrap_or(now + chrono::Duration::hours(12));
        decide_tick(kind, current, target, now, next, false, s)
    }

    /// A reconciled state whose recorded next transition is at `next`.
    fn reconciled_at(next: DateTime<Utc>) -> TickState {
        let mut s = TickState::new();
        note_reconciled(&mut s, next);
        s
    }

    #[test]
    fn refresh_forces_even_when_matching() {
        // Refresh force-applies so config edits take effect immediately —
        // CLAUDE.md invariant. Must beat the in-sync skip.
        let s = reconciled_at(utc(2026, 7, 4, 15, 46, 5));
        let now = utc(2026, 7, 4, 9, 0, 0);
        assert_eq!(
            dt(TickKind::Refresh, Some(Theme::Light), Theme::Light, now, &s),
            TickAction::Apply
        );
    }

    #[test]
    fn matching_current_skips_apply() {
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let now = utc(2026, 7, 4, 9, 0, 0);
        let s = reconciled_at(sunset);
        for kind in [TickKind::Init, TickKind::Scheduled, TickKind::Wake] {
            assert_eq!(
                dt(kind, Some(Theme::Light), Theme::Light, now, &s),
                TickAction::SkipInSync,
                "kind {kind:?}"
            );
        }
    }

    #[test]
    fn scheduled_transition_applies_over_override() {
        // A natural transition resets any manual override — documented
        // behavior. The sunset tick fires at/after the recorded next.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let s = reconciled_at(sunset);
        assert_eq!(
            dt(
                TickKind::Scheduled,
                Some(Theme::Light),
                Theme::Dark,
                sunset,
                &s
            ),
            TickAction::Apply
        );
    }

    #[test]
    fn wake_before_next_transition_preserves_override() {
        // Mid-day the user picked Dark (schedule says Light until sunset).
        // Win+L → unlock before sunset: no transition passed → override
        // survives.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let s = reconciled_at(sunset);
        let now = utc(2026, 7, 4, 12, 0, 0);
        assert_eq!(
            dt(TickKind::Wake, Some(Theme::Dark), Theme::Light, now, &s),
            TickAction::SkipOverride
        );
    }

    #[test]
    fn wake_after_missed_transition_reconciles() {
        // Slept through sunset: now >= recorded next → reconcile.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let s = reconciled_at(sunset);
        let now = utc(2026, 7, 4, 20, 0, 0);
        assert_eq!(
            dt(TickKind::Wake, Some(Theme::Light), Theme::Dark, now, &s),
            TickAction::Apply
        );
    }

    #[test]
    fn wake_after_even_number_of_missed_transitions_reconciles() {
        // Overnight lock spanning sunset AND sunrise: the schedule's target
        // is back to Light — same THEME as when we reconciled, but two
        // transitions passed. A parity comparison of themes would wrongly
        // preserve yesterday's override; the time rule must reconcile.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let s = reconciled_at(sunset);
        let next_morning = utc(2026, 7, 5, 4, 0, 0);
        assert_eq!(
            dt(
                TickKind::Wake,
                Some(Theme::Dark), // yesterday's override, still on screen
                Theme::Light,
                next_morning,
                &s
            ),
            TickAction::Apply
        );
    }

    #[test]
    fn wake_without_baseline_reconciles() {
        // A wake before any reconciled tick (e.g. Init hit the no-location
        // path): the safe default is reconcile, not preserve.
        let s = TickState::new();
        let now = utc(2026, 7, 4, 12, 0, 0);
        assert_eq!(
            dt(TickKind::Wake, Some(Theme::Dark), Theme::Light, now, &s),
            TickAction::Apply
        );
    }

    #[test]
    fn wake_with_unreadable_current_skips_inside_window() {
        // current_theme() = None (registry read failed): inside a preserved
        // window the app declines to apply — pinned deliberately.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let s = reconciled_at(sunset);
        let now = utc(2026, 7, 4, 12, 0, 0);
        assert_eq!(
            dt(TickKind::Wake, None, Theme::Light, now, &s),
            TickAction::SkipOverride
        );
    }

    #[test]
    fn theme_opposite_flips() {
        assert_eq!(Theme::Light.opposite(), Theme::Dark);
        assert_eq!(Theme::Dark.opposite(), Theme::Light);
    }

    #[test]
    fn toggle_target_flips_current_and_defaults_dark() {
        assert_eq!(toggle_target(Some(Theme::Light)), Theme::Dark);
        assert_eq!(toggle_target(Some(Theme::Dark)), Theme::Light);
        // Unreadable current: base defaults to Light → toggle lands on Dark.
        assert_eq!(toggle_target(None), Theme::Dark);
    }

    #[test]
    fn sanitize_log_msg_keeps_field_parseable() {
        assert_eq!(
            sanitize_log_msg("bad \"path\" at\nline\r\ntwo"),
            "bad 'path' at line  two"
        );
    }

    // --- bounded apply retry (v0.4.0) ---

    #[test]
    fn retry_deadline_is_soon_but_never_past_next_transition() {
        let now = utc(2026, 7, 4, 9, 0, 0).with_timezone(&Local);
        let far_next = utc(2026, 7, 4, 15, 0, 0).with_timezone(&Local);
        let near_next = utc(2026, 7, 4, 9, 0, 30).with_timezone(&Local);
        assert_eq!(
            retry_deadline(now, far_next),
            now + chrono::Duration::seconds(APPLY_RETRY_DELAY_SECS)
        );
        assert_eq!(retry_deadline(now, near_next), near_next);
    }

    #[test]
    fn failed_apply_then_wake_must_reapply_not_preserve() {
        // THE design-review blocker: a failed apply must never be mistaken
        // for a user override. Sequence: noon tick reconciles; sunset passes
        // while asleep; resume tick's apply FAILS; seconds later the unlock
        // wake fires — it must Apply (acting as a free retry), not skip.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let mut s = reconciled_at(sunset);
        let resume_at = utc(2026, 7, 4, 17, 0, 0);
        assert_eq!(
            dt(
                TickKind::Wake,
                Some(Theme::Light),
                Theme::Dark,
                resume_at,
                &s
            ),
            TickAction::Apply
        );
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        let unlock_at = utc(2026, 7, 4, 17, 0, 10);
        assert_eq!(
            dt(
                TickKind::Wake,
                Some(Theme::Light),
                Theme::Dark,
                unlock_at,
                &s
            ),
            TickAction::Apply,
            "wake after failed apply must re-apply, not preserve the failure"
        );
    }

    #[test]
    fn user_intervention_during_retry_window_stands_down() {
        // Sunset apply fails (screen stuck Light); the user then explicitly
        // picks a theme. The pending retry must cancel instead of clobbering
        // their choice — and afterwards the override survives normally.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        // The user picks Dark (via Toggle or Settings), moving the screen off
        // the failure baseline (Light). Note the gate can only see a MOVE:
        // re-picking the stuck theme (Light) is indistinguishable from the
        // failure itself, so that case retries — an inherent limit of
        // observing current_theme().
        let retry_at = utc(2026, 7, 4, 16, 47, 5);
        assert_eq!(
            dt(
                TickKind::Scheduled,
                Some(Theme::Dark),
                Theme::Dark,
                retry_at,
                &s
            ),
            TickAction::CancelRetry
        );
        note_reconciled(&mut s, utc(2026, 7, 5, 2, 35, 0));
        assert_eq!(s.retry_count, 0);
    }

    #[test]
    fn unreadable_baseline_does_not_cancel_retry() {
        // The apply failed while the registry was unreadable (baseline None);
        // a later readable theme is not evidence the user did anything.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(&mut s, None, utc(2026, 7, 5, 2, 35, 0)));
        let retry_at = utc(2026, 7, 4, 16, 47, 5);
        assert_eq!(
            dt(
                TickKind::Scheduled,
                Some(Theme::Light),
                Theme::Dark,
                retry_at,
                &s
            ),
            TickAction::Apply
        );
    }

    #[test]
    fn intervention_after_the_episode_window_reconciles_instead_of_cancelling() {
        // Failed sunset apply → user fixes the theme by hand → the machine
        // sleeps through sunrise. On wake the episode's window is over: the
        // schedule outranks the stand-down.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let sunrise = utc(2026, 7, 5, 2, 35, 0);
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(&mut s, Some(Theme::Light), sunrise));
        let next_morning = utc(2026, 7, 5, 6, 0, 0);
        assert_eq!(
            dt(
                TickKind::Wake,
                Some(Theme::Dark),
                Theme::Light,
                next_morning,
                &s
            ),
            TickAction::Apply
        );
    }

    #[test]
    fn wake_during_retry_with_future_reconciled_next_still_retries() {
        // A failed REFRESH mid-window leaves reconciled_next in the future;
        // a wake inside the retry window must retry, not "preserve an
        // override" that is really the failure.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(&mut s, Some(Theme::Dark), sunset));
        let unlock_at = utc(2026, 7, 4, 12, 0, 30);
        assert_eq!(
            dt(
                TickKind::Wake,
                Some(Theme::Dark),
                Theme::Light,
                unlock_at,
                &s
            ),
            TickAction::Apply
        );
    }

    #[test]
    fn refresh_beats_the_pending_retry_gate() {
        // Refresh must force-apply even mid-episode with the screen moved
        // off the baseline (which would otherwise read as CancelRetry).
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        let now = utc(2026, 7, 4, 16, 0, 0);
        assert_eq!(
            dt(TickKind::Refresh, Some(Theme::Dark), Theme::Light, now, &s),
            TickAction::Apply
        );
    }

    #[test]
    fn early_scheduled_fire_preserves_override() {
        // The timer fires 1 ms before the recorded transition (drift, or the
        // wall clock stepped back after arming). The screen shows a manual
        // Dark override; the schedule still says Light for 1 more ms. This
        // must not re-apply the OUTGOING theme over the override.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let s = reconciled_at(sunset);
        let early = sunset - chrono::Duration::milliseconds(1);
        assert_eq!(
            dt(
                TickKind::Scheduled,
                Some(Theme::Dark),
                Theme::Light,
                early,
                &s
            ),
            TickAction::SkipOverride
        );
        // On time, the same tick applies the new theme as always.
        assert_eq!(
            dt(
                TickKind::Scheduled,
                Some(Theme::Light),
                Theme::Dark,
                sunset,
                &s
            ),
            TickAction::Apply
        );
    }

    #[test]
    fn early_scheduled_fire_does_not_block_a_pending_retry() {
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(&mut s, Some(Theme::Dark), sunset));
        let retry_at = utc(2026, 7, 4, 12, 1, 0);
        assert_eq!(
            dt(
                TickKind::Scheduled,
                Some(Theme::Dark),
                Theme::Light,
                retry_at,
                &s
            ),
            TickAction::Apply
        );
    }

    // --- wall clock → monotonic deadline (v0.4.1) ---

    #[test]
    fn wall_to_instant_maps_future_and_clamps_past() {
        let mono = Instant::now();
        let now = utc(2026, 9, 27, 14, 21, 59);
        assert_eq!(
            wall_to_instant(now + mins(23), now, mono),
            mono + Duration::from_secs(23 * 60)
        );
        // Already past: fire immediately (and repeated re-arms can't keep
        // pushing it into the future).
        assert_eq!(wall_to_instant(now - mins(180), now, mono), mono);
    }

    #[test]
    fn plan_wake_before_deadline_sleeps_at_most_a_heartbeat() {
        let mono = Instant::now();
        let now = utc(2026, 9, 27, 17, 22, 0);
        let mark = Some((now, mono));
        // Far deadline: capped at the heartbeat.
        assert_eq!(
            plan_wake(Some(now + mins(12 * 60)), mark, now, mono),
            WakePlan::Arm(mono + HEARTBEAT)
        );
        // Near deadline: the deadline itself.
        assert_eq!(
            plan_wake(Some(now + mins(3)), mark, now, mono),
            WakePlan::Arm(mono + Duration::from_secs(180))
        );
        // A timer that fires 1 ms early re-arms for the remaining 1 ms
        // instead of ticking early.
        let early_wall = now + mins(3) - chrono::Duration::milliseconds(1);
        let early_mono = mono + Duration::from_millis(180_000 - 1);
        assert_eq!(
            plan_wake(Some(now + mins(3)), mark, early_wall, early_mono),
            WakePlan::Arm(early_mono + Duration::from_millis(1))
        );
        // Nothing armed (no location): idle.
        assert_eq!(plan_wake(None, None, now, mono), WakePlan::Idle);
    }

    #[test]
    fn plan_wake_ticks_when_due() {
        let mono = Instant::now();
        let t0 = utc(2026, 9, 27, 17, 22, 0);
        let deadline = t0 + mins(23);
        assert_eq!(
            plan_wake(
                Some(deadline),
                Some((t0, mono)),
                deadline,
                mono + Duration::from_secs(23 * 60)
            ),
            WakePlan::Tick { step_ms: None }
        );
    }

    #[test]
    fn clock_step_forward_after_skewed_boot_ticks_immediately() {
        // The 2026-09-27 incident: booted 3 h behind (wall 17:21 = real
        // 20:21 local), Init armed next=17:44 wall. 19 s later w32time
        // stepped the clock +3 h. The next wake of the loop — the
        // WM_TIMECHANGE broadcast, or at worst the heartbeat — must detect
        // the step and tick, instead of sleeping ~23 min (or, before
        // v0.4.1, up to 3 h).
        let mono = Instant::now();
        let skewed = utc(2026, 9, 27, 14, 21, 59);
        let deadline = utc(2026, 9, 27, 14, 44, 54);
        let mark = Some((skewed, mono));
        let later_mono = mono + Duration::from_secs(19);
        let corrected = skewed + chrono::Duration::seconds(19) + mins(180);
        assert_eq!(
            plan_wake(Some(deadline), mark, corrected, later_mono),
            WakePlan::Tick {
                step_ms: Some(10_800_000)
            }
        );
        // …and that Scheduled tick applies, because a transition passed.
        let s = reconciled_at(deadline);
        assert_eq!(
            dt(
                TickKind::Scheduled,
                Some(Theme::Light),
                Theme::Dark,
                corrected,
                &s
            ),
            TickAction::Apply
        );
    }

    #[test]
    fn clock_step_backward_is_detected_too() {
        let mono = Instant::now();
        let t0 = utc(2026, 9, 27, 17, 22, 0);
        let plan = plan_wake(
            Some(t0 + mins(60)),
            Some((t0, mono)),
            t0 - mins(120) + mins(5),
            mono + Duration::from_secs(300),
        );
        assert_eq!(
            plan,
            WakePlan::Tick {
                step_ms: Some(-7_200_000)
            }
        );
    }

    #[test]
    fn small_drift_is_not_a_clock_step() {
        let mono = Instant::now();
        let t0 = utc(2026, 9, 27, 17, 22, 0);
        let plan = plan_wake(
            Some(t0 + mins(600)),
            Some((t0, mono)),
            t0 + mins(300) + chrono::Duration::seconds(5),
            mono + Duration::from_secs(300 * 60),
        );
        assert!(matches!(plan, WakePlan::Arm(_)), "plan = {plan:?}");
    }

    #[test]
    fn backward_step_across_a_transition_reapplies_not_preserves() {
        // The clock ran 3 h ahead: at real 14:44 (wall 17:44) the app
        // applied the sunset early and recorded next = tomorrow's sunrise.
        // w32time steps the clock back; at real 14:50 the schedule's upcoming
        // transition is TODAY's sunset (17:44) — earlier than the recorded
        // one. The Dark screen is the app's own stale apply, not a user
        // override: re-apply Light (and let the real sunset flip it).
        let sunset = utc(2026, 9, 27, 14, 44, 0);
        let tomorrow_sunrise = utc(2026, 9, 28, 2, 43, 0);
        let s = reconciled_at(tomorrow_sunrise);
        let real_now = utc(2026, 9, 27, 11, 50, 0);
        assert_eq!(
            decide_tick(
                TickKind::Scheduled,
                Some(Theme::Dark),
                Theme::Light,
                real_now,
                sunset,
                true,
                &s
            ),
            TickAction::Apply
        );
        // With an unstepped clock, the same diverged screen before the
        // recorded transition is an override and is preserved.
        assert_eq!(
            decide_tick(
                TickKind::Wake,
                Some(Theme::Dark),
                Theme::Light,
                real_now,
                tomorrow_sunrise,
                false,
                &s
            ),
            TickAction::SkipOverride
        );
    }

    #[test]
    fn detected_clock_step_sign_and_threshold() {
        let mono = Instant::now();
        let t0 = utc(2026, 9, 27, 17, 22, 0);
        let mark = Some((t0, mono));
        let later = mono + Duration::from_secs(600);
        // Forward +3 h, backward −2 h, and 30 s of drift (below threshold).
        assert_eq!(
            detected_clock_step(mark, t0 + mins(10 + 180), later),
            Some(10_800_000)
        );
        assert_eq!(
            detected_clock_step(mark, t0 + mins(10 - 120), later),
            Some(-7_200_000)
        );
        assert_eq!(
            detected_clock_step(mark, t0 + mins(10) + chrono::Duration::seconds(30), later),
            None
        );
        assert_eq!(detected_clock_step(None, t0, later), None);
    }

    #[test]
    fn backward_step_within_the_same_frame_preserves_override() {
        // The clock steps back but doesn't cross the recorded transition:
        // the schedule's next is still the same sunset (or only seconds
        // off), so a diverged screen is still the user's override.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let s = reconciled_at(sunset);
        let now = sunset - mins(120);
        for next in [sunset, sunset - chrono::Duration::seconds(30)] {
            for kind in [TickKind::Wake, TickKind::Scheduled] {
                assert_eq!(
                    decide_tick(kind, Some(Theme::Dark), Theme::Light, now, next, true, &s),
                    TickAction::SkipOverride,
                    "kind {kind:?}, next {next}"
                );
            }
        }
    }

    #[test]
    fn earlier_next_without_a_clock_step_still_preserves_override() {
        // `next` can legitimately move earlier with no clock step (e.g. a
        // polar-season altitude crossing coming within the 48 h search).
        // Without an observed backward step that must not read as a new
        // frame — the user's override stays.
        let recorded = utc(2027, 1, 16, 10, 23, 39);
        let s = reconciled_at(recorded);
        let now = utc(2027, 1, 14, 12, 0, 0);
        let earlier_next = utc(2027, 1, 15, 10, 34, 44);
        assert_eq!(
            decide_tick(
                TickKind::Wake,
                Some(Theme::Light),
                Theme::Dark,
                now,
                earlier_next,
                false,
                &s
            ),
            TickAction::SkipOverride
        );
    }

    #[test]
    fn exhausted_refresh_episode_keeps_wakes_as_retries() {
        // A failed Refresh at noon (reconciled_next still = the future
        // sunset) that burns its whole retry budget must not leave the wrong
        // screen looking like a manual override to later wakes.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let mut s = reconciled_at(sunset);
        for _ in 0..MAX_APPLY_RETRIES {
            assert!(note_apply_failed(&mut s, Some(Theme::Dark), sunset));
        }
        assert!(!note_apply_failed(&mut s, Some(Theme::Dark), sunset));
        let later = utc(2026, 7, 4, 10, 0, 0);
        assert_eq!(
            dt(TickKind::Wake, Some(Theme::Dark), Theme::Light, later, &s),
            TickAction::Apply
        );
    }

    #[test]
    fn unreadable_current_does_not_cancel_retry() {
        // None proves nothing about user intent — the retry must proceed.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        let retry_at = utc(2026, 7, 4, 16, 47, 5);
        assert_eq!(
            dt(TickKind::Scheduled, None, Theme::Dark, retry_at, &s),
            TickAction::Apply
        );
    }

    #[test]
    fn retry_budget_is_bounded_and_resets_per_episode() {
        let mut s = TickState::new();
        // Three failures schedule retries; the fourth gives up AND resets,
        // so the next transition window gets a fresh budget instead of
        // inheriting a permanently burned one.
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        assert!(!note_apply_failed(
            &mut s,
            Some(Theme::Light),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        assert_eq!(s.retry_count, 0);
        assert_eq!(s.retry_baseline, None);
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light),
            utc(2026, 7, 5, 2, 35, 0)
        ));
    }

    #[test]
    fn reconcile_clears_retry_episode() {
        let next = utc(2026, 7, 5, 2, 35, 0);
        let mut s = TickState::new();
        assert!(note_apply_failed(&mut s, Some(Theme::Light), next));
        note_reconciled(&mut s, next);
        assert_eq!(s.retry_count, 0);
        assert_eq!(s.retry_baseline, None);
        assert_eq!(s.reconciled_next, Some(next));
    }

    // --- .theme DisplayName resolution (v0.4.0) ---

    /// Temp .theme file that cleans up after itself; same uniqueness contract
    /// as TempConfig (distinct `name` per test).
    struct TempTheme(PathBuf);

    impl TempTheme {
        fn new(name: &str, bytes: &[u8]) -> Self {
            let path = std::env::temp_dir().join(format!(
                "wts-test-{}-{}.theme",
                std::process::id(),
                name
            ));
            fs::write(&path, bytes).unwrap();
            Self(path)
        }
    }

    impl Drop for TempTheme {
        fn drop(&mut self) {
            let _ = fs::remove_file(&self.0);
        }
    }

    #[test]
    fn theme_display_name_literal() {
        let t = TempTheme::new(
            "literal",
            b"; comment\r\n[Theme]\r\nDisplayName=My Custom Theme\r\nColor=1\r\n",
        );
        assert_eq!(
            resolve_theme_display_name(&t.0).as_deref(),
            Some("My Custom Theme")
        );
    }

    #[test]
    fn theme_display_name_only_read_from_theme_section() {
        let t = TempTheme::new(
            "wrong-section",
            b"[Control Panel\\Desktop]\r\nDisplayName=Nope\r\n[Slideshow]\r\nInterval=1\r\n",
        );
        assert_eq!(resolve_theme_display_name(&t.0), None);
    }

    #[test]
    fn theme_display_name_section_header_case_insensitive() {
        let t = TempTheme::new("case", b"[THEME]\r\nDisplayName=Loud\r\n");
        assert_eq!(resolve_theme_display_name(&t.0).as_deref(), Some("Loud"));
    }

    #[test]
    fn theme_display_name_survives_windows_1252_comment_bytes() {
        // aero.theme's copyright comment carries a raw 0xa9 (©) — invalid
        // UTF-8. The lossy decode + comment skipping must not derail parsing.
        let t = TempTheme::new(
            "cp1252",
            b"; Copyright \xa9 Microsoft\r\n[Theme]\r\nDisplayName=Real\r\n",
        );
        assert_eq!(resolve_theme_display_name(&t.0).as_deref(), Some("Real"));
    }

    #[test]
    fn theme_display_name_key_is_case_insensitive() {
        let t = TempTheme::new("key-case", b"[Theme]\r\ndisplayname = Lower Case Key\r\n");
        assert_eq!(
            resolve_theme_display_name(&t.0).as_deref(),
            Some("Lower Case Key")
        );
    }

    #[test]
    fn theme_display_name_quotes_are_stripped() {
        let t = TempTheme::new("quoted", b"[Theme]\r\nDisplayName=\"My Theme\"\r\n");
        assert_eq!(
            resolve_theme_display_name(&t.0).as_deref(),
            Some("My Theme")
        );
    }

    #[test]
    fn theme_display_name_reads_utf16_files_with_non_ascii_names() {
        // Windows reads a BOM'd UTF-16LE .theme as Unicode — so must we, or a
        // non-English DisplayName never matches in tier 1.
        let text = "[Theme]\r\nDisplayName=الوضع الليلي\r\n";
        let mut bytes = vec![0xFF, 0xFE];
        for u in text.encode_utf16() {
            bytes.extend_from_slice(&u.to_le_bytes());
        }
        let t = TempTheme::new("utf16", &bytes);
        assert_eq!(
            resolve_theme_display_name(&t.0).as_deref(),
            Some("الوضع الليلي")
        );
    }

    #[test]
    fn theme_display_name_missing_file_or_key_is_none() {
        let missing =
            std::env::temp_dir().join(format!("wts-test-{}-nonexistent.theme", std::process::id()));
        assert_eq!(resolve_theme_display_name(&missing), None);
        let t = TempTheme::new("no-name", b"[Theme]\r\nColor=1\r\n");
        assert_eq!(resolve_theme_display_name(&t.0), None);
    }

    #[test]
    fn system_theme_display_names_resolve_via_indirect_strings() {
        // Guarded: meaningful anywhere the stock themes exist (any normal
        // Windows, incl. GitHub windows-latest runners). Exercises the
        // SHLoadIndirectString path used for tier-1 apply of system themes.
        let root = std::env::var("SystemRoot").unwrap_or_else(|_| "C:\\Windows".into());
        let aero = PathBuf::from(&root).join("Resources\\Themes\\aero.theme");
        let dark = PathBuf::from(&root).join("Resources\\Themes\\dark.theme");
        if !aero.exists() || !dark.exists() {
            // Rust has no test-skip. libtest captures this line unless run
            // with --nocapture or --show-output, so it documents the skip
            // rather than announcing it; this test is the only coverage of
            // resolve_indirect_string.
            eprintln!("SKIP: system .theme files absent - indirect-string path not exercised");
            return;
        }
        let a = resolve_theme_display_name(&aero).expect("aero.theme display name");
        let d = resolve_theme_display_name(&dark).expect("dark.theme display name");
        assert!(!a.is_empty() && !d.is_empty());
        assert_ne!(a, d);
    }

    // --- resolve_theme_file fallback chain (v0.4.0) ---

    #[test]
    fn theme_file_custom_path_wins_when_present() {
        let t = TempTheme::new("custom-day", b"[Theme]\r\nDisplayName=Custom\r\n");
        let cfg = Config {
            theme_day: Some(t.0.to_string_lossy().into_owned()),
            ..Config::default()
        };
        assert_eq!(resolve_theme_file(Theme::Light, &cfg), t.0);
    }

    #[test]
    fn theme_file_falls_back_when_custom_path_missing() {
        let cfg = Config {
            theme_night: Some("C:\\definitely\\not\\here.theme".into()),
            ..Config::default()
        };
        let p = resolve_theme_file(Theme::Dark, &cfg);
        assert!(p.ends_with("dark.theme"), "got {p:?}");
    }

    #[test]
    fn theme_file_defaults_by_theme() {
        let cfg = Config::default();
        assert!(resolve_theme_file(Theme::Light, &cfg).ends_with("aero.theme"));
        assert!(resolve_theme_file(Theme::Dark, &cfg).ends_with("dark.theme"));
    }
}
