#![windows_subsystem = "windows"]

use std::error::Error;
use std::ffi::c_void;
use std::fs;
use std::path::{Path, PathBuf};
use std::ptr;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;
use std::sync::OnceLock;
use std::time::{Duration, Instant};

use chrono::{DateTime, Local, Utc};
use serde::{Deserialize, Serialize};
use sun_times::sun_times;
use tray_icon::{
    menu::{Menu, MenuEvent, MenuId, MenuItem, PredefinedMenuItem},
    TrayIcon, TrayIconBuilder, TrayIconEvent,
};
use windows::Devices::Geolocation::{GeolocationAccessStatus, Geolocator};
use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
use windows_sys::core::{GUID, HRESULT};
use windows_sys::Win32::Foundation::{
    CloseHandle, GetLastError, SysFreeString, ERROR_ACCESS_DENIED, ERROR_ALREADY_EXISTS,
    ERROR_FILE_NOT_FOUND, HWND, LPARAM, LRESULT, WPARAM,
};
use windows_sys::Win32::Graphics::Dwm::DwmFlush;
use windows_sys::Win32::Storage::FileSystem::{
    GetDriveTypeW, GetFileInformationByHandle, BY_HANDLE_FILE_INFORMATION, FILE_SHARE_READ,
};
use windows_sys::Win32::System::Com::{
    CoCreateInstance, CoIncrementMTAUsage, CLSCTX_INPROC_SERVER,
};
use windows_sys::Win32::System::Environment::ExpandEnvironmentStringsW;
use windows_sys::Win32::System::LibraryLoader::{
    GetModuleHandleW, SetDefaultDllDirectories, LOAD_LIBRARY_SEARCH_SYSTEM32,
};
use windows_sys::Win32::System::Power::{
    PowerRegisterSuspendResumeNotification, DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS,
};
use windows_sys::Win32::System::Registry::{
    RegCloseKey, RegCreateKeyExW, RegDeleteValueW, RegOpenKeyExW, RegQueryValueExW, RegSetValueExW,
    HKEY, HKEY_CURRENT_USER, KEY_QUERY_VALUE, KEY_SET_VALUE, REG_BINARY, REG_DWORD,
    REG_OPTION_NON_VOLATILE, REG_SZ,
};
use windows_sys::Win32::System::RemoteDesktop::{
    WTSRegisterSessionNotification, NOTIFY_FOR_THIS_SESSION,
};
use windows_sys::Win32::System::SystemInformation::GetSystemDirectoryW;
use windows_sys::Win32::System::Threading::{
    CreateMutexW, OpenMutexW, SYNCHRONIZATION_SYNCHRONIZE,
};
use windows_sys::Win32::System::WindowsProgramming::{GetPrivateProfileStringW, DRIVE_REMOTE};
use windows_sys::Win32::UI::Shell::{
    AssocQueryStringW, SHLoadIndirectString, ShellExecuteW, ASSOCF_INIT_IGNOREUNKNOWN,
    ASSOCSTR_FRIENDLYAPPNAME,
};
use windows_sys::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DispatchMessageW, FindWindowW, GetMessageW, MessageBoxW,
    PostMessageW, RegisterClassW, RegisterWindowMessageW, SendMessageTimeoutW, SendMessageW,
    TranslateMessage, DEVICE_NOTIFY_CALLBACK, HWND_BROADCAST, HWND_MESSAGE, IDYES,
    MB_ICONINFORMATION, MB_ICONQUESTION, MB_ICONWARNING, MB_OK, MB_YESNO, MSG,
    PBT_APMRESUMEAUTOMATIC, SMTO_ABORTIFHUNG, SW_HIDE, SW_SHOWNORMAL, WM_CLOSE, WM_SETTINGCHANGE,
    WM_THEMECHANGED, WM_WTSSESSION_CHANGE, WNDCLASSW,
};
use winit::event::{Event, StartCause};
use winit::event_loop::{ActiveEventLoop, ControlFlow, EventLoop, EventLoopProxy};

const WTS_SESSION_UNLOCK: WPARAM = 0x8;

const THEME_KEY: &str = "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";
const RUN_KEY: &str = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
/// Where Windows keeps the user's on/off switch for each Run entry (Settings
/// → Apps → Startup, Task Manager → Startup apps). Read, never written: the
/// user's switch wins.
const STARTUP_APPROVED_RUN_KEY: &str =
    "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StartupApproved\\Run";
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
    /// setting). Kept so a save writes them back instead of silently
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

/// Whether the pair is a place on Earth, in decimal degrees. False for NaN
/// and the infinities. Checked wherever coordinates enter (config.json,
/// Windows Location): `schedule` computes something for any numbers — 225
/// typed for 22.5 gives a schedule that's hours off, and an out-of-range
/// latitude can swap day and night.
fn valid_coordinates(lat: f64, lon: f64) -> bool {
    (-90.0..=90.0).contains(&lat) && (-180.0..=180.0).contains(&lon)
}

/// Latitude from which a Windows Location reading keeps a third decimal.
const FINE_COORDINATES_LAT: f64 = 60.0;

/// A Windows Location reading at the precision worth keeping: 2 decimals
/// (about 1 km; it moves a sunrise or sunset by a few seconds), 3 from
/// FINE_COORDINATES_LAT on — a margin for the latitudes from about 65.7°,
/// where polar day and night begin and a small change in position can move
/// their start by tens of minutes. Full sensor precision would only put the
/// user's address into config.json. Pure — unit-tested.
fn round_coordinates(lat: f64, lon: f64) -> (f64, f64) {
    let scale = if lat.abs() >= FINE_COORDINATES_LAT {
        1000.0
    } else {
        100.0
    };
    // `+ 0.0` turns a rounded -0.0 into 0.0 (it would be saved as "-0.0").
    let round = |v: f64| (v * scale).round() / scale + 0.0;
    let (rlat, rlon) = (round(lat), round(lon));
    if rlat == 0.0 && rlon == 0.0 {
        // 0/0 means "no location yet" (has_location): a reading next to it
        // stays as it is rather than rounding into that.
        return (lat + 0.0, lon + 0.0);
    }
    (rlat, rlon)
}

/// A Windows Location reading as the app keeps it — rounded — or why it
/// can't be used: "out-of-range" when it isn't a place on Earth, "zero" for
/// exactly 0/0, which a Config can't tell from "no location yet". Checked
/// before rounding, so a reading just outside a limit isn't rounded into
/// range. Pure — unit-tested.
fn usable_location(lat: f64, lon: f64) -> Result<(f64, f64), &'static str> {
    if !valid_coordinates(lat, lon) {
        return Err("out-of-range");
    }
    if lat == 0.0 && lon == 0.0 {
        return Err("zero");
    }
    Ok(round_coordinates(lat, lon))
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

    /// The schedule slot's name in log lines (`slot=`).
    fn slot_str(self) -> &'static str {
        match self {
            Theme::Light => "day",
            Theme::Dark => "night",
        }
    }
}

/// The two Personalize values a theme sets — all the app can observe of
/// which theme is on screen. A theme may set them differently (the
/// `[VisualStyles]` keys `AppMode` / `SystemMode`), e.g. light apps with a
/// dark taskbar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Modes {
    /// AppsUseLightTheme.
    apps_light: bool,
    /// SystemUsesLightTheme (taskbar, Start).
    system_light: bool,
}

impl From<Theme> for Modes {
    fn from(theme: Theme) -> Self {
        let light = theme == Theme::Light;
        Modes {
            apps_light: light,
            system_light: light,
        }
    }
}

/// Log form of a Modes reading, used for every current=/target=/actual=
/// field. No `=` or `,` inside, so key=value lines stay parseable.
fn modes_str(m: Option<Modes>) -> &'static str {
    match m {
        None => "unknown",
        Some(Modes {
            apps_light: true,
            system_light: true,
        }) => "light",
        Some(Modes {
            apps_light: false,
            system_light: false,
        }) => "dark",
        Some(Modes {
            apps_light: true,
            system_light: false,
        }) => "apps-light/system-dark",
        Some(Modes {
            apps_light: false,
            system_light: true,
        }) => "apps-dark/system-light",
    }
}

/// The modes a theme file declares, from its `[VisualStyles]` `SystemMode`
/// and `AppMode` values ("Light"/"Dark", case-insensitive). A value that is
/// absent, empty, or unrecognized falls back to the slot's default (day =
/// light, night = dark) — and the second result is then false: what Windows
/// does with a missing key is undocumented (high-contrast themes have none),
/// so such a slot's modes are only a guess. Pure — unit-tested.
fn modes_from(slot: Theme, system: Option<&str>, app: Option<&str>) -> (Modes, bool) {
    let parse = |v: Option<&str>| match v.map(|s| s.trim().to_ascii_lowercase()).as_deref() {
        Some("light") => Some(true),
        Some("dark") => Some(false),
        _ => None,
    };
    let default = slot == Theme::Light;
    let (system, app) = (parse(system), parse(app));
    (
        Modes {
            apps_light: app.unwrap_or(default),
            system_light: system.unwrap_or(default),
        },
        system.is_some() && app.is_some(),
    )
}

/// One schedule slot (day = `Theme::Light`, night = `Theme::Dark`), resolved
/// once per tick or toggle: the file that will actually be applied (after
/// the path policy and stock fallback) and the modes THAT file sets.
#[derive(Debug, Clone)]
struct Slot {
    theme: Theme,
    file: PathBuf,
    /// What the screen should read once this slot's theme is applied.
    modes: Modes,
    /// Both modes were read from the file rather than defaulted.
    keyed: bool,
    /// A configured path that was rejected in favor of the stock theme:
    /// (as written, as expanded, why).
    rejected: Option<(String, PathBuf, ThemePathIssue)>,
}

/// Whether the screen's modes can tell the two slots apart. When they can't
/// — identical modes (two themes that differ only in wallpaper), or a slot
/// whose modes are a guess — "the screen matches the schedule" is not
/// observable, so ticks never skip as in-sync and Toggle alternates by what
/// the app last applied instead of by the screen.
fn sync_decidable(day: &Slot, night: &Slot) -> bool {
    day.keyed && night.keyed && day.modes != night.modes
}

/// What the Toggle menu item should apply. Pure — unit-tested.
///
/// When the modes tell the slots apart: a screen matching the day slot goes
/// to night and vice versa; a screen matching neither (some other theme)
/// goes to whichever the apps mode is not. Otherwise — undecidable slots or
/// an unreadable screen — it flips `last_applied` (the slot this process
/// last applied successfully), else `fallback` (the slot the schedule says
/// is current; day when there's no location).
fn toggle_target(
    screen: Option<Modes>,
    day: Modes,
    night: Modes,
    decidable: bool,
    last_applied: Option<Theme>,
    fallback: Theme,
) -> Theme {
    match screen {
        Some(m) if decidable && m == day => Theme::Dark,
        Some(m) if decidable && m == night => Theme::Light,
        Some(m) if decidable => {
            if m.apps_light {
                Theme::Dark
            } else {
                Theme::Light
            }
        }
        _ => last_applied.unwrap_or(fallback).opposite(),
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
    /// The modes read right AFTER the last apply failed — so any partial
    /// effect of our own apply (e.g. a tier-3 write that flipped one value
    /// before failing on the other) is part of the baseline, never mistaken
    /// for a move. If the screen later no longer matches this, the user
    /// intervened during the retry window and the retry must stand down
    /// rather than clobber their choice.
    retry_baseline: Option<Modes>,
    /// The next-transition instant computed at the tick whose apply failed —
    /// the failure episode's own window. An intervention only cancels the
    /// retry while `now < episode_next`; past it, a transition has passed
    /// and reconciling to the schedule outranks the stand-down (otherwise an
    /// intervention right before a suspend that spans a transition would be
    /// promoted to a day-long override).
    episode_next: Option<DateTime<Utc>>,
    /// A Toggle applied since the last apply failed (`note_toggle`). The
    /// modes can't show every Toggle: a second one puts the screen back on
    /// the baseline, and with two same-mode themes or an unreadable registry
    /// not even the first is visible. Stands the retry down like an observed
    /// move, within the same window. Always false outside a failure episode.
    retry_toggled: bool,
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
            retry_toggled: false,
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
///
/// `current` is the screen's modes (None = unreadable), `target` the modes
/// of the scheduled slot's theme. `sync_decidable` false (see the fn of that
/// name) only disables the in-sync skip — equal modes then prove nothing —
/// while override preservation and the retry stand-down work as usual.
#[allow(clippy::too_many_arguments)]
fn decide_tick(
    kind: TickKind,
    current: Option<Modes>,
    target: Modes,
    now: DateTime<Utc>,
    next: DateTime<Utc>,
    clock_stepped_back: bool,
    sync_decidable: bool,
    state: &TickState,
) -> TickAction {
    // Refresh is fresh user intent: always force-apply.
    if kind == TickKind::Refresh {
        return TickAction::Apply;
    }
    // Pending-retry gate: the user intervened mid-episode — the screen moved
    // away from the failure snapshot, or a Toggle applied (which the modes
    // can't always show: a second Toggle lands back on the snapshot). For
    // the modes only an OBSERVED move counts — an unreadable reading (None)
    // on EITHER side proves nothing and must not cancel. And the stand-down
    // only applies within the failure episode's own window
    // (now < episode_next): once a transition has passed, reconciling to the
    // schedule outranks it, exactly like any override ending at the next
    // natural transition.
    let screen_moved =
        state.retry_baseline.is_some() && current.is_some() && current != state.retry_baseline;
    if state.retry_count > 0
        && (screen_moved || state.retry_toggled)
        && state.episode_next.is_some_and(|n| now < n)
    {
        return TickAction::CancelRetry;
    }
    if sync_decidable && current == Some(target) {
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

/// End the failure episode, if there is one: no retry pending, a fresh
/// budget for the next failure.
fn clear_retry_episode(state: &mut TickState) {
    state.retry_count = 0;
    state.retry_baseline = None;
    state.episode_next = None;
    state.retry_toggled = false;
}

/// Record a tick that ended in sync with the schedule (any non-Err outcome).
fn note_reconciled(state: &mut TickState, next: DateTime<Utc>) {
    state.reconciled_next = Some(next);
    clear_retry_episode(state);
}

/// Record a failed apply. `observed` must be read AFTER the failure (see
/// TickState::retry_baseline). Returns true when a quick retry should be
/// scheduled (reconciled_next is left stale); false when the budget is
/// exhausted (the episode resets so the NEXT transition window gets a fresh
/// budget, and reconciled_next is cleared so later wakes remain free retry
/// opportunities).
fn note_apply_failed(
    state: &mut TickState,
    observed: Option<Modes>,
    next_utc: DateTime<Utc>,
) -> bool {
    state.retry_count += 1;
    state.retry_baseline = observed;
    // The baseline is new, so only a Toggle after THIS failure counts (an
    // earlier one can still be set here when the retry ran past its window).
    state.retry_toggled = false;
    if state.retry_count > MAX_APPLY_RETRIES {
        clear_retry_episode(state);
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

/// Record a Toggle; `applied` is whether its apply_theme returned Ok (which
/// for tier 2 means "launched", as for a tick's own apply). This is the only
/// thing Toggle writes in TickState, and all it does is flag a pending
/// failure episode: the retry gate in `decide_tick` then stands the retry
/// down, where the modes alone show an odd number of Toggles at best. A
/// Toggle that failed is not flagged (in the ordinary case — the screen
/// still on the other slot's theme — it was aiming at the scheduled theme,
/// which the retry then delivers), and outside an episode nothing is
/// recorded — a Toggle is then an ordinary override, which the
/// modes and `reconciled_next` already handle. Pure — unit-tested.
fn note_toggle(state: &mut TickState, applied: bool) {
    if applied && state.retry_count > 0 {
        state.retry_toggled = true;
    }
}

/// How a tick's action came out — what `settle_tick` records.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TickOutcome {
    /// The tick ended in step with the schedule: apply_theme succeeded, or
    /// the tick skipped (already in sync, an override preserved, a retry
    /// stood down).
    Reconciled,
    /// apply_theme failed. `after` is the screen's modes read AFTER the
    /// failure, never the pre-tick reading (see TickState::retry_baseline).
    ApplyFailed { after: Option<Modes> },
}

/// What follows a failed apply — the ` retry=` field of the tick's log line.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RetryNote {
    /// Retry number N is scheduled.
    Attempt(u32),
    /// The budget is spent: nothing until the next transition, wake or
    /// Refresh.
    Exhausted,
}

/// What a tick arms and logs once its outcome is recorded.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Settled {
    /// Wall-clock deadline of the next tick: the next transition, or sooner
    /// for a retry.
    deadline: DateTime<Utc>,
    /// Set after a failed apply.
    retry: Option<RetryNote>,
}

/// Record a tick's outcome in `state` and say when the next tick is due.
/// Everything a tick records about its outcome goes through here; `tick`
/// then arms the returned deadline together with the clock mark. `now` is
/// read after the apply (a retry is timed from the failure, not from the
/// start of a slow apply); `next` is the schedule's next transition. Pure —
/// unit-tested.
fn settle_tick(
    state: &mut TickState,
    kind: TickKind,
    outcome: TickOutcome,
    now: DateTime<Utc>,
    next: DateTime<Utc>,
) -> Settled {
    match outcome {
        TickOutcome::Reconciled => {
            note_reconciled(state, next);
            Settled {
                deadline: next,
                retry: None,
            }
        }
        TickOutcome::ApplyFailed { after } => {
            if kind == TickKind::Refresh {
                // Fresh user intent (likely a just-fixed config): a Refresh
                // never inherits a part-used retry budget.
                clear_retry_episode(state);
            }
            // Bounded retry: tick again soon instead of waiting up to ~12 h
            // for the next transition. The retry arrives as a normal
            // Scheduled tick; the pending-retry gate in decide_tick stands
            // it down if the user changes the theme in the meantime.
            if note_apply_failed(state, after, next) {
                Settled {
                    deadline: retry_deadline(now, next),
                    retry: Some(RetryNote::Attempt(state.retry_count)),
                }
            } else {
                Settled {
                    deadline: next,
                    retry: Some(RetryNote::Exhausted),
                }
            }
        }
    }
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
fn retry_deadline(now: DateTime<Utc>, next: DateTime<Utc>) -> DateTime<Utc> {
    std::cmp::min(
        next,
        now + chrono::Duration::seconds(APPLY_RETRY_DELAY_SECS),
    )
}

#[derive(Debug, Clone)]
enum AppEvent {
    Menu(MenuId),
    Wake(WakeKind),
    /// Re-check that the tray icon is registered (attempt number). Never
    /// ticks.
    TrayProbe(u32),
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

/// A self-heal performed by `load_config_ex`.
#[derive(Debug)]
struct Healed {
    reason: &'static str,
    /// The `auto_start` the repaired config got (healed_config).
    auto_start: bool,
    /// Set when writing the healed defaults back failed (the reset then only
    /// exists in memory, and repeats next launch).
    save_err: Option<String>,
}

/// What a successful `load_config_ex` has to tell besides the config.
#[derive(Debug)]
enum LoadNote {
    /// The file was empty and was reset (so callers can log it — a heal
    /// resets settings, and must not be silent).
    Healed(Healed),
    /// First run, but config.json could not be created (the reason): the
    /// defaults exist in memory only, and nothing this session holds came
    /// from a file.
    NotCreated(String),
}

/// A loaded config, with what the session has to know about where it came
/// from.
#[derive(Debug)]
struct Loaded {
    cfg: Config,
    /// false: config.json was missing and could not be created
    /// (LoadNote::NotCreated), so `cfg` is the defaults and nothing of it
    /// came from a file. A repaired empty file counts as read: it was read,
    /// and held nothing.
    from_file: bool,
}

impl Loaded {
    /// Pure — unit-tested.
    fn new(cfg: Config, note: Option<&LoadNote>) -> Self {
        Loaded {
            cfg,
            from_file: !matches!(note, Some(LoadNote::NotCreated(_))),
        }
    }
}

/// `load_config_ex` without the note, for a PC where autostart is registered
/// (tests use this form; it never reads the registry).
#[cfg(test)]
fn load_config_at(path: &Path) -> Result<Config, String> {
    load_config_ex(path, || Some(true)).map(|(cfg, _)| cfg)
}

/// Load the config from `path`. A missing file is first-run: defaults are
/// written and returned. `Err` means the file EXISTS but could not be read or
/// parsed, or its coordinates can't be used (parse_config) — it is left
/// untouched on disk so a hand-edit typo can be fixed instead of silently
/// wiping the user's coordinates and theme paths.
///
/// `registered` says whether autostart is registered right now (None: can't
/// tell); it is asked only when an empty file is repaired (healed_config).
/// Production passes auto_start_registered, tests a constant — they never
/// read the real registry.
///
/// Logging and dialogs stay out of here because tests call this directly
/// and log_path() is exe-relative.
fn load_config_ex(
    path: &Path,
    registered: impl FnMut() -> Option<bool>,
) -> Result<(Config, Option<LoadNote>), String> {
    load_config_via(path, read_config_text, registered)
}

/// The config an empty or NUL-filled config.json is repaired with: the
/// defaults, except that `auto_start` says what is registered at that moment
/// (`registered`: whether the Run value exists; None when that can't be
/// read, which keeps the default). The file can't say what the user had
/// chosen, and the default `true` used to turn autostart back on for someone
/// who had turned it off — in that session, and through the repaired file at
/// every later start. This way a repair changes nothing about autostart.
/// The cost: an empty file on a PC where autostart isn't registered is not a
/// first run — it is repaired with `auto_start: false`. That is a file left
/// by a crash during the very first run, before autostart was registered,
/// and a config.json created empty ahead of the first start, by hand or by
/// an installer (one that holds `{}` isn't empty, and loads as the
/// defaults). Pure — unit-tested.
fn healed_config(registered: Option<bool>) -> Config {
    let default = Config::default();
    Config {
        auto_start: registered.unwrap_or(default.auto_start),
        ..default
    }
}

/// `load_config_ex` with its read step passed in, so tests can make the file
/// change between a read and the write that follows it.
///
/// Defaults are only written into a file that is still missing, or still
/// empty, at the moment of writing. When it isn't — an editor saved it in
/// between (e.g. via delete-then-rename) — the file is read again: the
/// user's file wins, and is never answered with defaults that a later
/// persist_config would write over it.
fn load_config_via(
    path: &Path,
    mut read: impl FnMut(&Path) -> std::io::Result<String>,
    mut registered: impl FnMut() -> Option<bool>,
) -> Result<(Config, Option<LoadNote>), String> {
    for _ in 0..3 {
        match read(path) {
            Ok(content) if config_content_is_empty(&content) => {
                // A crash mid-write can leave a 0-byte or NUL-filled file
                // (the latter when the size was extended but the data never
                // flushed). Nothing in it to preserve — self-heal
                // (healed_config: the defaults, autostart as it is
                // registered) instead of erroring on every launch.
                let cfg = healed_config(registered());
                let auto_start = cfg.auto_start;
                let save_err = match heal_empty_config(path, &cfg) {
                    Ok(true) => None,
                    Ok(false) => continue,
                    // Another program has it open to write it (or locked):
                    // what it will hold is unknown, so it's a file that
                    // can't be read, not an empty one — defaults returned
                    // now would later be persisted over that program's save.
                    Err(e) if matches!(e.raw_os_error(), Some(32 | 33)) => {
                        return Err(format!("config.json could not be read: {e}"));
                    }
                    Err(e) => Some(e.to_string()),
                };
                return Ok((
                    cfg,
                    Some(LoadNote::Healed(Healed {
                        reason: "empty",
                        auto_start,
                        save_err,
                    })),
                ));
            }
            Ok(content) => return parse_config(&content).map(|cfg| (cfg, None)),
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
                let cfg = Config::default();
                match create_config(path, &cfg) {
                    Ok(()) => return Ok((cfg, None)),
                    Err(e) if e.kind() == std::io::ErrorKind::AlreadyExists => continue,
                    // The folder can't be written (or the disk is full, …):
                    // this run, and every later one, starts from the
                    // defaults. Said out loud by the caller — it used to be
                    // silent.
                    Err(e) => return Ok((cfg, Some(LoadNote::NotCreated(e.to_string())))),
                }
            }
            Err(e) => return Err(format!("config.json could not be read: {e}")),
        }
    }
    // Missing when read but there when created, every time: in practice a
    // link to a file that can't be opened (create_new refuses even a
    // dangling link) — else a file that kept changing.
    if fs::symlink_metadata(path).is_ok_and(|m| m.file_type().is_symlink()) {
        return Err("config.json is a link to a file that can't be opened".to_string());
    }
    Err("config.json kept changing while it was being read".to_string())
}

/// The two coordinate keys as config.json has them. `Config` can't show
/// whether a key is there: a missing one deserializes to 0.0.
#[derive(Deserialize)]
struct CoordinateKeys {
    latitude: Option<serde::de::IgnoredAny>,
    longitude: Option<serde::de::IgnoredAny>,
}

/// config.json's text as a Config, with its coordinates checked. An Err is a
/// config error like any other: the file stays as it is and the user is
/// told. Pure — unit-tested.
fn parse_config(content: &str) -> Result<Config, String> {
    let cfg = serde_json::from_str::<Config>(content).map_err(config_json_error)?;
    let keys = serde_json::from_str::<CoordinateKeys>(content).map_err(config_json_error)?;
    check_coordinates(&cfg, &keys)?;
    Ok(cfg)
}

/// serde_json's error as a config error, minus the value it quotes for a
/// wrong type: `invalid type: string "24.71", expected f64 at line 2 column
/// 23` would put a coordinate typed in quotes into events.log. What was
/// expected and where stay.
fn config_json_error(e: serde_json::Error) -> String {
    let mut text = e.to_string();
    if text.starts_with("invalid type: ") || text.starts_with("invalid value: ") {
        if let Some(i) = text.rfind(", expected ") {
            text = format!("wrong kind of value{}", &text[i..]);
        }
    }
    format!("config.json is not valid JSON: {text}")
}

/// The file's coordinates are both there or both absent (absent, or both
/// 0.0, means "ask Windows Location"), and in range. One key without the
/// other — left out or misspelled — used to count as 0.0, and the schedule
/// was computed for a place on the equator or the Greenwich meridian.
///
/// The messages name the key, never its value: they go to events.log as
/// well as the dialog, and a slip (two values swapped, a digit too many)
/// would put the user's real coordinates there.
fn check_coordinates(cfg: &Config, keys: &CoordinateKeys) -> Result<(), String> {
    match (keys.latitude.is_some(), keys.longitude.is_some()) {
        (false, false) => return Ok(()),
        (true, true) => {}
        (has_lat, _) => {
            let (has, lacks) = if has_lat {
                ("latitude", "longitude")
            } else {
                ("longitude", "latitude")
            };
            return Err(format!(
                "config.json has \"{has}\" but no \"{lacks}\" (check the spelling). \
                 Set both, or neither to use Windows Location."
            ));
        }
    }
    let lat_ok = valid_coordinates(cfg.latitude, 0.0);
    let lon_ok = valid_coordinates(0.0, cfg.longitude);
    if lat_ok && lon_ok {
        return Ok(());
    }
    let mut problems = Vec::new();
    if !lat_ok {
        problems.push("latitude must be between -90 and 90");
    }
    if !lon_ok {
        problems.push("longitude must be between -180 and 180");
    }
    let hint = if valid_coordinates(cfg.longitude, cfg.latitude) {
        " (are the two swapped?)"
    } else if lat_ok && cfg.longitude > 180.0 && cfg.longitude <= 360.0 {
        // Perhaps written 0 to 360, east from Greenwich: v0.3.2 through
        // v0.5.1 happened to compute the right schedule for that. A
        // question, since it may as well be a typo (225 for 22.5).
        " (a 0 to 360 longitude? West is negative: subtract 360)"
    } else {
        ""
    };
    Err(format!("config.json: {}{hint}", problems.join("; ")))
}

/// First run: write `cfg` as a new config.json. create_new, not a replace:
/// a file that appeared since the read (AlreadyExists) is the user's.
fn create_config(path: &Path, cfg: &Config) -> std::io::Result<()> {
    use std::io::Write;
    let json = serde_json::to_string_pretty(cfg)?;
    let mut f = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)?;
    f.write_all(json.as_bytes())
}

/// Write `cfg` into a config.json that read as empty — if it still is. The
/// check and the write share one handle that lets other programs read the
/// file but not write, replace or delete it, so an editor's save can't land
/// between them and be overwritten. Written in place: there's nothing in the
/// file to lose. Ok(false): it has content now, or is gone — read it again.
/// Err with ERROR_SHARING_VIOLATION: another program still has it open to
/// write it.
fn heal_empty_config(path: &Path, cfg: &Config) -> std::io::Result<bool> {
    use std::io::{Read, Seek, Write};
    use std::os::windows::fs::OpenOptionsExt;
    let json = serde_json::to_string_pretty(cfg)?;
    let mut attempt = 0;
    let mut f = loop {
        match fs::OpenOptions::new()
            .read(true)
            .write(true)
            .share_mode(FILE_SHARE_READ)
            .open(path)
        {
            Ok(f) => break f,
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => return Ok(false),
            // Open elsewhere (ERROR_SHARING_VIOLATION) — an on-access AV
            // scan, or an editor mid-save: wait a moment, as save_config_at
            // does.
            Err(e) if e.raw_os_error() == Some(32) && attempt < 4 => {
                attempt += 1;
                std::thread::sleep(Duration::from_millis(100));
            }
            Err(e) => return Err(e),
        }
    };
    let mut bytes = Vec::new();
    f.read_to_end(&mut bytes)?;
    if !decode_config_text(&bytes).is_ok_and(|s| config_content_is_empty(&s)) {
        return Ok(false);
    }
    f.set_len(0)?;
    f.rewind()?;
    f.write_all(json.as_bytes())?;
    f.sync_all()?;
    Ok(true)
}

/// config.json as text, whatever encoding a Windows editor saved it in:
/// UTF-8 with or without a BOM (Notepad), or UTF-16 LE/BE with a BOM
/// (PowerShell 5.1's `>` / Out-File default is LE; Notepad offers both).
fn read_config_text(path: &Path) -> std::io::Result<String> {
    decode_config_text(&fs::read(path)?)
}

/// read_config_text's decoding, for bytes already read.
fn decode_config_text(bytes: &[u8]) -> std::io::Result<String> {
    use std::io::{Error, ErrorKind};
    let utf16 = |rest: &[u8], le: bool| -> std::io::Result<String> {
        let (pairs, odd) = rest.as_chunks::<2>();
        if !odd.is_empty() {
            return Err(Error::new(ErrorKind::InvalidData, "truncated UTF-16 text"));
        }
        let units: Vec<u16> = pairs
            .iter()
            .map(|&pair| {
                if le {
                    u16::from_le_bytes(pair)
                } else {
                    u16::from_be_bytes(pair)
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
    let rest = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes);
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

/// Serialize `cfg` over the config.json at `path` unconditionally. Only call
/// with a Config that was successfully loaded from that file this session —
/// persisting a default/fallback Config here is exactly the settings-wipe
/// bug fixed in v0.3.2 (broken files must stay on disk for the user to
/// repair). Two things keep to that: run() never persists in a session
/// whose file is broken, and persist_config_at sends a session whose file
/// could not be created to create_config instead. Nothing re-checks the
/// file at the moment of the replace: whatever was saved into it since the
/// load is replaced — after an ordinary load, and just the same after a
/// repair of an empty file, written or not (when the repair couldn't be
/// written, the file read as empty and this is what writes it).
///
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

/// config.json was missing and could not be created at the latest load
/// (LoadNote::NotCreated; `Loaded::from_file` false). While set,
/// persist_config only creates the file — it never replaces one: what the
/// session holds is the defaults, or after a Refresh the settings it had
/// loaded earlier, and a file that has appeared since is the user's.
static CONFIG_NOT_ON_DISK: AtomicBool = AtomicBool::new(false);

/// Load config.json from its real location, logging a self-heal and
/// reporting a first-run file that couldn't be created.
fn load_config_logged() -> Result<Loaded, String> {
    load_config_ex(&config_path(), auto_start_registered).map(|(cfg, note)| {
        let loaded = Loaded::new(cfg, note.as_ref());
        CONFIG_NOT_ON_DISK.store(!loaded.from_file, Ordering::SeqCst);
        match note {
            Some(LoadNote::Healed(h)) => {
                let action = if h.save_err.is_some() {
                    "reset-in-memory"
                } else {
                    "reset-to-defaults"
                };
                log_event(&format!(
                    "{} config_healed reason={} action={} auto_start={}",
                    Local::now().to_rfc3339(),
                    h.reason,
                    action,
                    h.auto_start,
                ));
                if let Some(e) = h.save_err {
                    log_event(&format!(
                        "{} config_save_err msg=\"{}\"",
                        Local::now().to_rfc3339(),
                        sanitize_log_msg(&e),
                    ));
                }
            }
            Some(LoadNote::NotCreated(e)) => report_config_not_created(&e),
            None => {}
        }
        loaded
    })
}

/// The can't-save-settings box was shown (once per run of the app).
static NOT_WRITABLE_BOX_SHOWN: AtomicBool = AtomicBool::new(false);

/// First run in a folder where config.json can't be created — a read-only
/// folder, typically: say so. It used to be silent: the app ran on the
/// defaults, asked for the location at every start, and Open Config opened
/// nothing. The log line is best effort (events.log lives in the same
/// folder), so the box is what tells the user; the session carries on. Also
/// reached from a Refresh that finds the file gone and can't create it (the
/// session then keeps the config it had — adopt_reloaded_config).
fn report_config_not_created(err: &str) {
    log_event(&format!(
        "{} config_create_err msg=\"{}\"",
        Local::now().to_rfc3339(),
        sanitize_log_msg(err),
    ));
    let folder = config_path()
        .parent()
        .map(|d| d.display().to_string())
        .unwrap_or_default();
    let body = format!(
        "WinThemeSwitcher can't write to its own folder:\n{folder}\n\n\
         config.json could not be created there ({err}), so your settings can't be \
         saved — the app starts from the defaults and looks up or asks for your \
         location at every start — and the log (events.log) most likely can't be \
         written either. For now the app keeps running: on the default settings, or \
         on the ones it had loaded before the file went missing.\n\n\
         To fix it: Quit, move win-theme-switcher.exe to a folder you can write \
         to — for example %LOCALAPPDATA%\\Programs\\WinThemeSwitcher — and start it \
         from there.",
    );
    show_warning_once(
        &NOT_WRITABLE_BOX_SHOWN,
        "WinThemeSwitcher — Settings can't be saved",
        body,
    );
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

/// Write `cfg` to config.json at `path`. A config that came from the file
/// replaces it (save_config_at). One that didn't (`not_on_disk`: the file
/// was missing and couldn't be created) only goes into a file that is still
/// missing — a file that has appeared since is the user's, and defaults
/// plus a location must not be written over it.
fn persist_config_at(path: &Path, cfg: &Config, not_on_disk: bool) -> Result<(), String> {
    if not_on_disk {
        create_config(path, cfg).map_err(|e| e.to_string())
    } else {
        save_config_at(path, cfg).map_err(|e| e.to_string())
    }
}

/// persist_config_at for the real config.json, with the failure logged
/// instead of dropped.
fn persist_config(cfg: &Config) {
    let not_on_disk = CONFIG_NOT_ON_DISK.load(Ordering::SeqCst);
    match persist_config_at(&config_path(), cfg, not_on_disk) {
        // Created after all: from here on the file holds this config.
        Ok(()) => CONFIG_NOT_ON_DISK.store(false, Ordering::SeqCst),
        Err(e) => log_event(&format!(
            "{} config_save_err msg=\"{}\"",
            Local::now().to_rfc3339(),
            sanitize_log_msg(&e),
        )),
    }
}

/// Refresh's config step, minus I/O. Adopts a freshly (re)loaded config when
/// it parsed. Returns whether the autostart registration may be re-asserted —
/// only from a config that came from disk at some point this session, never
/// from the in-memory fallback of a session that started with a broken file
/// (whose auto_start=true default must not override the file's false) — and
/// the load error to report, if any. Pure — unit-tested.
fn adopt_reloaded_config(
    load: Result<Loaded, String>,
    cfg: &mut Config,
    have_disk_cfg: &mut bool,
) -> (bool, Option<String>) {
    match load {
        Ok(Loaded {
            cfg: new_cfg,
            from_file: true,
        }) => {
            *cfg = new_cfg;
            *have_disk_cfg = true;
            (true, None)
        }
        // config.json is missing and can't be created. At a Refresh that is
        // as likely a drive or share that can't be reached right now as a
        // deleted file, and the defaults would replace the settings this
        // session holds — and turn autostart on over a file that says
        // false. Nothing was loaded, so nothing changes (through v0.5.1 the
        // defaults were adopted); load_config_logged has reported it.
        Ok(Loaded {
            from_file: false, ..
        }) => (*have_disk_cfg, None),
        // Keep the last-known-good (or fallback) config; the broken file
        // stays on disk for the user to fix.
        Err(e) => (*have_disk_cfg, Some(e)),
    }
}

/// Make the calling thread — the main thread — a single-threaded apartment:
/// IThemeManager2 needs one, and every theme apply runs here. The Windows
/// Location lookup doesn't: it has a thread of its own (keep_mta_alive).
fn ensure_com_initialized() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }
}

/// The longest a Windows Location lookup holds up its caller — startup
/// (before the tray exists) or a Refresh (the event-loop thread). Windows'
/// own limit for a position is 60 s, and its answer to "may this app use
/// the location" has none.
const LOCATION_LOOKUP_CAP: Duration = Duration::from_secs(15);

/// Why run_with_timeout came back without the work's value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NoAnswer {
    /// The limit passed first. The worker goes on; what it produces later
    /// is dropped.
    TimedOut,
    /// The worker thread couldn't be started, or ended without a value (a
    /// panic, in a build that unwinds).
    WorkerLost,
}

/// Run `work` on a thread of its own and wait for its value, at most
/// `limit`. On a timeout the caller goes on and the worker is left to
/// finish (or hang) by itself: its late value is sent into a channel
/// nobody reads any more and dropped there. So `work` must own everything
/// it touches, and must not change anything the caller would have to know
/// about — its only way to report is the value it returns.
fn run_with_timeout<T, F>(thread_name: &str, limit: Duration, work: F) -> Result<T, NoAnswer>
where
    T: Send + 'static,
    F: FnOnce() -> T + Send + 'static,
{
    // Room for the one value, so the worker's send never waits for a
    // receiver — also not for one that has given up.
    let (tx, rx) = mpsc::sync_channel::<T>(1);
    let spawned = std::thread::Builder::new()
        .name(thread_name.to_string())
        .spawn(move || {
            // Err means the caller timed out and dropped the receiver: the
            // value comes back inside the error and is dropped right here.
            let _ = tx.send(work());
        });
    if spawned.is_err() {
        return Err(NoAnswer::WorkerLost);
    }
    match rx.recv_timeout(limit) {
        Ok(value) => Ok(value),
        Err(mpsc::RecvTimeoutError::Timeout) => Err(NoAnswer::TimedOut),
        Err(mpsc::RecvTimeoutError::Disconnected) => Err(NoAnswer::WorkerLost),
    }
}

/// Why a Windows Location lookup produced no location.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LocationFailure {
    /// Windows answered "denied": Location is off, or apps (desktop apps)
    /// aren't allowed to use it.
    Denied,
    /// Windows answered "unspecified": not allowed, and no reason given.
    Unspecified,
    /// No answer within LOCATION_LOOKUP_CAP.
    TimedOut,
    /// A call failed. `stage` says which step, `hresult` is the error code
    /// (0 for the "worker" stage: the lookup thread couldn't be started or
    /// ended without an answer). Both are for the log only.
    Failed { stage: &'static str, hresult: i32 },
    /// Windows returned a reading, and usable_location refused it for this
    /// reason.
    Rejected(&'static str),
}

impl LocationFailure {
    /// The `reason=` token of the `location_lookup_failed` log line.
    fn token(self) -> &'static str {
        match self {
            LocationFailure::Denied => "access-denied",
            LocationFailure::Unspecified => "access-unspecified",
            LocationFailure::TimedOut => "timeout",
            LocationFailure::Failed { .. } => "error",
            LocationFailure::Rejected(_) => "reading-rejected",
        }
    }
}

/// The events.log line for a lookup that produced no location, without
/// its timestamp. Never a coordinate: a rejected reading is named by its
/// reason only. Pure — unit-tested.
fn location_failure_line(failure: LocationFailure, elapsed_ms: u128) -> String {
    let mut line = format!(
        "location_lookup_failed reason={} after_ms={}",
        failure.token(),
        elapsed_ms,
    );
    match failure {
        LocationFailure::Failed { stage, hresult } => {
            // {:08X} of an i32 prints its two's complement: 0x80070005.
            line.push_str(&format!(" stage={stage} hresult=0x{hresult:08X}"));
        }
        LocationFailure::Rejected(why) => line.push_str(&format!(" reading={why}")),
        _ => {}
    }
    line
}

/// Windows' answer to "may this app use the location" as the lookup's
/// next step: go on, or stop with the reason. Anything that isn't
/// "allowed" or "denied" counts as unspecified. Pure — unit-tested.
fn access_answer(status: GeolocationAccessStatus) -> Result<(), LocationFailure> {
    if status == GeolocationAccessStatus::Allowed {
        Ok(())
    } else if status == GeolocationAccessStatus::Denied {
        Err(LocationFailure::Denied)
    } else {
        Err(LocationFailure::Unspecified)
    }
}

/// From what the lookup thread brought back (or didn't) to what a caller
/// may put into a Config: the reading goes through usable_location here,
/// on the caller's thread, and nothing else comes out as coordinates.
/// Pure — unit-tested.
fn settle_location(
    answer: Result<Result<(f64, f64), LocationFailure>, NoAnswer>,
) -> Result<(f64, f64), LocationFailure> {
    match answer {
        Ok(Ok((lat, lon))) => usable_location(lat, lon).map_err(LocationFailure::Rejected),
        Ok(Err(failure)) => Err(failure),
        Err(NoAnswer::TimedOut) => Err(LocationFailure::TimedOut),
        Err(NoAnswer::WorkerLost) => Err(LocationFailure::Failed {
            stage: "worker",
            hresult: 0,
        }),
    }
}

/// Make sure the process has a multithreaded apartment (MTA), and keep it
/// for the life of the process. A thread that never initialized COM — the
/// lookup thread — is then in that apartment implicitly, which is what
/// its WinRT calls need. No per-thread CoInitializeEx: that would have to
/// be balanced by a CoUninitialize the thread never reaches when it is
/// abandoned mid-call, and the last one would tear the apartment down
/// under the activation factories the `windows` crate caches. The main
/// thread stays the STA it made itself (ensure_com_initialized).
fn keep_mta_alive() -> Result<(), HRESULT> {
    static DONE: AtomicBool = AtomicBool::new(false);
    if DONE.load(Ordering::SeqCst) {
        return Ok(());
    }
    // The cookie would undo it (CoDecrementMTAUsage); never used. Two
    // threads getting here together take two usages — harmless.
    let mut cookie = ptr::null_mut();
    let hr = unsafe { CoIncrementMTAUsage(&mut cookie) };
    if hr < 0 {
        return Err(hr);
    }
    DONE.store(true, Ordering::SeqCst);
    Ok(())
}

/// The Windows Location lookup itself: may the app use the location, and
/// if so, where is the PC. Blocks until Windows answers — the first
/// question has no limit of its own (Windows may be showing the user a
/// prompt), the position gives up after 60 s — so it runs ONLY on the
/// lookup thread (lookup_windows_location), never on the main thread. It
/// touches nothing but its own WinRT objects, so a caller that timed out
/// can leave it running. Returns the raw reading: settle_location decides
/// whether it's usable.
fn query_windows_location() -> Result<(f64, f64), LocationFailure> {
    fn failed_at(stage: &'static str) -> impl Fn(windows::core::Error) -> LocationFailure {
        move |e| LocationFailure::Failed {
            stage,
            hresult: e.code().0,
        }
    }
    keep_mta_alive().map_err(|hresult| LocationFailure::Failed {
        stage: "apartment",
        hresult,
    })?;
    let access = Geolocator::RequestAccessAsync()
        .and_then(|op| op.get())
        .map_err(failed_at("request-access"))?;
    access_answer(access)?;
    let geo = Geolocator::new().map_err(failed_at("geolocator"))?;
    let pos = geo
        .GetGeopositionAsync()
        .and_then(|op| op.get())
        .map_err(failed_at("position"))?;
    let p = pos
        .Coordinate()
        .and_then(|coord| coord.Point())
        .and_then(|point| point.Position())
        .map_err(failed_at("reading"))?;
    Ok((p.Latitude, p.Longitude))
}

/// One Windows Location lookup, held to `cap`: the location to use —
/// rounded (round_coordinates) — or why there is none, and how long it
/// took in milliseconds. The WinRT calls run on a thread of their own; a
/// result that arrives after the cap is dropped.
fn lookup_windows_location(cap: Duration) -> (Result<(f64, f64), LocationFailure>, u128) {
    let started = Instant::now();
    let answer = run_with_timeout("location", cap, query_windows_location);
    (settle_location(answer), started.elapsed().as_millis())
}

/// Where Windows Location says the PC is, or why the app has no location
/// from it. Waits at most LOCATION_LOOKUP_CAP. A failure is logged — one
/// `location_lookup_failed` line, after `location_rejected` for a reading
/// the app can't use — always without coordinates.
fn try_get_windows_location() -> Result<(f64, f64), LocationFailure> {
    let (result, elapsed_ms) = lookup_windows_location(LOCATION_LOOKUP_CAP);
    if let Err(failure) = result {
        if let LocationFailure::Rejected(reason) = failure {
            log_event(&format!(
                "{} location_rejected reason={}",
                Local::now().to_rfc3339(),
                reason,
            ));
        }
        log_event(&format!(
            "{} {}",
            Local::now().to_rfc3339(),
            location_failure_line(failure, elapsed_ms),
        ));
    }
    result
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

/// Whether a ShellExecuteW result says the launch was accepted: per its
/// reference, a value greater than 32. 0, an error code up to 32 and a
/// negative value all mean nothing was started. Pure — unit-tested.
fn launch_accepted(code: isize) -> bool {
    code > 32
}

/// Open config.json for editing: with the `.json` "open" handler if there
/// is one, or else in Notepad. Returns whether one of the two was
/// launched (launch_accepted); a failure is logged.
fn open_config_in_editor() -> bool {
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
            if launch_accepted(code) {
                return true;
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
        if !launch_accepted(h2 as isize) {
            log_event(&format!(
                "{} open_config_err code={} fallback_code={}",
                Local::now().to_rfc3339(),
                code,
                h2 as isize,
            ));
            return false;
        }
        true
    }
}

/// Ask Windows to open Settings on its Location page. Returns whether the
/// launch was accepted (launch_accepted) — the most the app can know; a
/// refusal is logged.
fn open_location_settings() -> bool {
    let verb = wide("open");
    let uri = wide("ms-settings:privacy-location");
    let code = unsafe {
        ShellExecuteW(
            ptr::null_mut(),
            verb.as_ptr(),
            uri.as_ptr(),
            ptr::null(),
            ptr::null(),
            SW_SHOWNORMAL,
        )
    } as isize;
    let accepted = launch_accepted(code);
    if !accepted {
        log_event(&format!(
            "{} open_settings_err code={}",
            Local::now().to_rfc3339(),
            code,
        ));
    }
    accepted
}

fn show_message_box(title: &str, body: &str, flags: u32) -> i32 {
    let title_w = wide(title);
    let body_w = wide(body);
    unsafe { MessageBoxW(ptr::null_mut(), body_w.as_ptr(), title_w.as_ptr(), flags) }
}

// === What the user is told when Windows Location gave no location ===
//
// The texts are built by pure functions from what actually happened, so
// each sentence can be checked against the situation it is shown in (the
// tests do): why the lookup failed, whether Settings could be opened,
// whether the tray icon exists yet — at first run these boxes come up
// BEFORE the tray icon, so they can't send the user to an icon that isn't
// there — and whether there is a config.json to type coordinates into (not
// in a session whose config.json couldn't be created). Plain words only;
// codes and API names go to events.log.

const LOCATION_BOX_TITLE: &str = "WinThemeSwitcher — Location";

/// Where the Location switches are, in Windows 11's Settings.
const LOCATION_SETTINGS_PAGE: &str = "Privacy & security > Location";

/// Why the app asks for the location at all.
const LOCATION_NEEDED: &str =
    "WinThemeSwitcher needs to know where this PC is to work out when the sun rises and sets.";

/// A Refresh shows nothing about its lookup, found or not (a dialog for
/// one that fails is v0.8.0's), so a box that sends the user to Refresh
/// says how to tell — by something the user can see.
const REFRESH_IS_SILENT: &str =
    "Refresh does not say whether it found the location: if the theme still does not \
     change at the next sunrise or sunset, it most likely did not.";

/// What stands in the way of typing coordinates when config.json could not
/// be created (the app's folder can't be written), and the way out. The
/// can't-save-settings box (report_config_not_created) has the details.
const NO_CONFIG_FILE: &str =
    "config.json could not be created in WinThemeSwitcher's folder. Move WinThemeSwitcher \
     to a folder you can write to and start it from there";

/// Where the app's icon — and with it the menu that has Refresh — is, or
/// is about to be. The caller adds what to do with it.
fn tray_icon_sentence(tray_exists: bool) -> &'static str {
    if tray_exists {
        "The WinThemeSwitcher icon is in the notification area at the end of the \
         taskbar (it may be behind the ^ arrow)."
    } else {
        "After you close this message, the WinThemeSwitcher icon appears in the \
         notification area at the end of the taskbar (it may be behind the ^ arrow)."
    }
}

/// What a lookup that gave no location brings up.
#[derive(Debug, Clone, PartialEq, Eq)]
enum LocationPrompt {
    /// A Yes/No question, after Windows said the app may not use the
    /// location: Yes opens the Location page of Settings, No goes on to
    /// typing coordinates into config.json.
    OfferSettings(String),
    /// One message: what happened and what the user can do. For a timeout,
    /// an error or an unusable reading — nothing says Location is off, so
    /// nothing offers to turn it on.
    Inform(String),
}

/// The first box after a failed lookup. "Off or not allowed" is said only
/// when Windows answered that; every other reason says what happened
/// instead. `config_on_disk` is false in a session whose config.json could
/// not be created: nothing then sends the user to type into it or to Open
/// Config. Pure — unit-tested.
fn location_prompt(
    failure: LocationFailure,
    tray_exists: bool,
    config_on_disk: bool,
) -> LocationPrompt {
    let what = match failure {
        LocationFailure::Denied | LocationFailure::Unspecified => {
            let (why, question) = if failure == LocationFailure::Denied {
                (
                    "Windows Location is turned off, or apps are not allowed to use it.",
                    "Open Windows Settings to turn it on?",
                )
            } else {
                (
                    "Windows did not let WinThemeSwitcher use this PC's location, and did \
                     not say why. Location may be turned off, or blocked for apps.",
                    "Open Windows Settings to check?",
                )
            };
            let no = if config_on_disk {
                "No: type your coordinates into config.json instead."
            } else {
                "No: go on without a location for now."
            };
            return LocationPrompt::OfferSettings(format!(
                "{why}\n\n{LOCATION_NEEDED}\n\n{question}\n\n\
                 Yes: open the Location settings.\n\
                 {no}",
            ));
        }
        // The answer that comes after the cap is dropped, so answering
        // Windows' own question isn't enough: Refresh has to ask again.
        LocationFailure::TimedOut => format!(
            "Windows did not answer within {} seconds when WinThemeSwitcher asked for \
             this PC's location. If Windows is showing a question about location \
             access, answer it, then close this message and choose Refresh (step 1 \
             below).",
            LOCATION_LOOKUP_CAP.as_secs(),
        ),
        LocationFailure::Failed { .. } => {
            "WinThemeSwitcher asked Windows for this PC's location, and the lookup failed."
                .to_string()
        }
        LocationFailure::Rejected(_) => {
            "Windows returned a location for this PC that WinThemeSwitcher cannot use.".to_string()
        }
    };
    // A reading the app refuses isn't cured by a switch in Settings.
    let check_settings = if matches!(failure, LocationFailure::Rejected(_)) {
        String::new()
    } else {
        format!(
            "\n\nIf this keeps happening, check that location is turned on in Windows \
             Settings, under {LOCATION_SETTINGS_PAGE}.",
        )
    };
    // The format is said here too: this path never shows manual_setup_text,
    // and a west or south value typed without its minus sign is a valid
    // place elsewhere — no check catches it.
    let type_them_in = if config_on_disk {
        "2. Or type your latitude and longitude into config.json, in decimal degrees \
         (south and west are negative): choose Open Config in the same menu, save the \
         file, then choose Refresh."
            .to_string()
    } else {
        format!(
            "2. Or type your latitude and longitude into config.json. Not yet, though: \
             {NO_CONFIG_FILE} first.",
        )
    };
    LocationPrompt::Inform(format!(
        "{what}\n\n{LOCATION_NEEDED} Until it has a location, it does not switch \
         themes by itself.\n\n\
         What you can do:\n\n\
         1. Try again later. {icon} Right-click it and choose Refresh. \
         {REFRESH_IS_SILENT}\n\n\
         {type_them_in}{check_settings}",
        icon = tray_icon_sentence(tray_exists),
    ))
}

/// The box after the user chose to open the Location settings — true
/// whether Settings opened or not. It points at the page and says what
/// must be on there; the switch names are given as a guide, since a PC may
/// not show all of them ("Location services" is shown to administrators
/// only) or may word them differently. Pure — unit-tested.
fn location_settings_followup(settings_opened: bool, tray_exists: bool) -> String {
    let where_to = if settings_opened {
        format!(
            "Windows Settings should now be open on its Location page. If it is not, \
             open Settings yourself and go to {LOCATION_SETTINGS_PAGE}.",
        )
    } else {
        format!(
            "Windows Settings could not be opened from here. Open Settings yourself \
             (Start > Settings) and go to {LOCATION_SETTINGS_PAGE}.",
        )
    };
    format!(
        "{where_to}\n\n\
         On that page, turn location on and let apps use it, desktop apps included. \
         The switches are called \"Location services\", \"Let apps access your \
         location\" and \"Let desktop apps access your location\", or something close \
         to that, and a PC may not show all of them. If location is turned off for \
         the whole PC, only an administrator can turn it on.\n\n\
         Then let WinThemeSwitcher look again. {icon} Right-click it and choose Refresh. \
         {REFRESH_IS_SILENT}",
        icon = tray_icon_sentence(tray_exists),
    )
}

/// The box before config.json is opened for typing coordinates. It says
/// what open_config_in_editor does next — the `.json` program, or else
/// Notepad — and what happens if neither starts (config_not_opened_text).
/// Pure — unit-tested.
fn manual_setup_text(tray_exists: bool) -> String {
    format!(
        "To set the location yourself, type your latitude and longitude into \
         config.json, in decimal degrees (south and west are negative), and save the \
         file.\n\n\
         When you close this message, WinThemeSwitcher opens config.json: in the \
         program set up for .json files, or else in Notepad. If neither can be \
         started, another message says where the file is.\n\n\
         {icon} After saving the file, right-click that icon and choose Refresh.",
        icon = tray_icon_sentence(tray_exists),
    )
}

/// The box in place of manual_setup_text when there is no config.json to
/// type into (it could not be created): nothing is opened, and the text
/// says what has to happen first. Pure — unit-tested.
fn no_config_file_text() -> String {
    format!(
        "WinThemeSwitcher has no location yet, so it does not switch themes by \
         itself.\n\n\
         The location cannot be typed in either: {NO_CONFIG_FILE}.",
    )
}

/// The box for a config.json that couldn't be opened for the user: where
/// the file is. Pure — unit-tested.
fn config_not_opened_text(path: &Path) -> String {
    format!(
        "config.json could not be opened automatically. Open it yourself, for example \
         with Notepad. The file is here:\n\n{}",
        path.display(),
    )
}

/// The "config.json could not be opened" box is already on screen.
static CONFIG_NOT_OPENED_BOX_OPEN: AtomicBool = AtomicBool::new(false);

/// First run, and Windows Location gave no location: tell the user what
/// happened and what to do. Runs before the tray icon exists, so every
/// box here blocks startup until it is closed — except the last-resort
/// "couldn't open config.json" one, which must not keep the icon the
/// previous box promised from appearing. `tray_exists` and
/// `config_on_disk` are the caller's to say (acquire_location): they
/// decide what the texts may claim.
fn explain_missing_location(failure: LocationFailure, tray_exists: bool, config_on_disk: bool) {
    match location_prompt(failure, tray_exists, config_on_disk) {
        LocationPrompt::Inform(body) => {
            show_message_box(LOCATION_BOX_TITLE, &body, MB_OK | MB_ICONWARNING);
        }
        LocationPrompt::OfferSettings(body) => {
            let yes =
                show_message_box(LOCATION_BOX_TITLE, &body, MB_YESNO | MB_ICONQUESTION) == IDYES;
            if yes {
                let opened = open_location_settings();
                show_message_box(
                    LOCATION_BOX_TITLE,
                    &location_settings_followup(opened, tray_exists),
                    MB_OK | MB_ICONINFORMATION,
                );
            } else if !config_on_disk {
                // No file to type into, and Notepad on a path that doesn't
                // exist would only ask whether to create it — in a folder
                // that can't be written.
                show_message_box(
                    "WinThemeSwitcher — Setup",
                    &no_config_file_text(),
                    MB_OK | MB_ICONWARNING,
                );
            } else {
                show_message_box(
                    "WinThemeSwitcher — Setup",
                    &manual_setup_text(tray_exists),
                    MB_OK | MB_ICONINFORMATION,
                );
                if !open_config_in_editor() {
                    show_warning_async(
                        &CONFIG_NOT_OPENED_BOX_OPEN,
                        "WinThemeSwitcher — Setup",
                        config_not_opened_text(&config_path()),
                    );
                }
            }
        }
    }
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
    let body = format!(
        "{err}\n\nThe file was left unchanged — your settings are still in it. \
         Fix the error (tray menu → Open Config), then choose Refresh.",
    );
    show_warning_async(
        &CONFIG_ERROR_BOX_OPEN,
        "WinThemeSwitcher — Config error",
        body,
    );
}

/// A warning MessageBox on a detached thread (never blocks startup or the
/// event loop); `open` keeps a second one of the same kind from stacking.
fn show_warning_async(open: &'static AtomicBool, title: &'static str, body: String) {
    if open.swap(true, Ordering::SeqCst) {
        return;
    }
    std::thread::spawn(move || {
        show_message_box(title, &body, MB_OK | MB_ICONWARNING);
        open.store(false, Ordering::SeqCst);
    });
}

/// A warning MessageBox on a detached thread, shown at most once per run of
/// the app — for a condition that can't change while the app runs (where
/// its exe is, whether its folder can be written), so saying it again at
/// every Refresh would only nag. `shown` is never reset.
fn show_warning_once(shown: &'static AtomicBool, title: &'static str, body: String) {
    if shown.swap(true, Ordering::SeqCst) {
        return;
    }
    std::thread::spawn(move || {
        show_message_box(title, &body, MB_OK | MB_ICONWARNING);
    });
}

/// A theme-path warning box is already on screen.
static THEME_PATH_BOX_OPEN: AtomicBool = AtomicBool::new(false);

/// Tell the user when theme_day / theme_night can't be used — once per
/// config load (startup, Refresh), not per apply. Each apply still logs
/// `theme_path_rejected`.
fn report_theme_path_problems(cfg: &Config) {
    let problems: Vec<String> = [Theme::Light, Theme::Dark]
        .into_iter()
        .filter_map(|theme| {
            let (file, rejected) = resolve_theme_file_checked(theme, cfg);
            let (raw, _, issue) = rejected?;
            let key = match theme {
                Theme::Light => "theme_day",
                Theme::Dark => "theme_night",
            };
            Some(format!(
                "{key} = \"{raw}\": {}.\nUsing {} instead.",
                issue.describe(),
                file.display()
            ))
        })
        .collect();
    if problems.is_empty() {
        return;
    }
    let body = format!(
        "{}\n\nFix config.json (tray menu → Open Config), then choose Refresh.",
        problems.join("\n\n")
    );
    show_warning_async(
        &THEME_PATH_BOX_OPEN,
        "WinThemeSwitcher — Theme file not used",
        body,
    );
}

/// First run (config.json has no location yet): ask Windows Location, and
/// save what it says; if it gives nothing, tell the user why and what to
/// do (explain_missing_location). `config_on_disk`: whether there is a
/// config.json at all (`Loaded::from_file`). Called from run() only, ahead
/// of the event loop and the tray icon.
fn acquire_location(cfg: &mut Config, config_on_disk: bool) {
    match try_get_windows_location() {
        Ok((lat, lon)) => {
            cfg.latitude = lat;
            cfg.longitude = lon;
            persist_config(cfg);
        }
        Err(failure) => {
            // The boxes come up before the tray icon exists, and must not
            // send the user to it as if it were there.
            let tray_exists = false;
            explain_missing_location(failure, tray_exists, config_on_disk);
        }
    }
}

/// The Personalize light/dark values now in effect (registry only — safe
/// from any thread, e.g. commit_watcher's). None if either value can't be
/// read as a 4-byte REG_DWORD; the first such failure is logged once, so a
/// permanently unreadable value shows up in events.log.
fn read_modes() -> Option<Modes> {
    static LOGGED: AtomicBool = AtomicBool::new(false);
    let apps = read_personalize_dword("AppsUseLightTheme");
    let system = read_personalize_dword("SystemUsesLightTheme");
    match (apps, system) {
        (Ok(a), Ok(s)) => Some(Modes {
            apps_light: a != 0,
            system_light: s != 0,
        }),
        (a, s) => {
            if !LOGGED.swap(true, Ordering::Relaxed) {
                let (name, err) = match a {
                    Err(e) => ("AppsUseLightTheme", e),
                    Ok(_) => ("SystemUsesLightTheme", s.err().unwrap_or_default()),
                };
                log_event(&format!(
                    "{} modes_read_err value={} err={}",
                    Local::now().to_rfc3339(),
                    name,
                    err,
                ));
            }
            None
        }
    }
}

/// An open registry key, closed on drop.
struct RegKey(HKEY);

impl Drop for RegKey {
    fn drop(&mut self) {
        unsafe {
            RegCloseKey(self.0);
        }
    }
}

/// Open the HKCU key `path` with `access`. With `create_missing`, a key that
/// isn't there is created (with any missing parents) — only then is
/// RegCreateKeyExW called, so with the key in place this does exactly what a
/// plain open does. Every profile gets the Run and Personalize keys from the
/// Default hive; they are missing only after a deliberate deletion, and
/// then autostart and the registry tier used to fail for good. Err is the
/// failing API's name and its Win32 code (the return value, not
/// GetLastError).
fn open_hkcu_key(
    path: &str,
    access: u32,
    create_missing: bool,
) -> Result<RegKey, (&'static str, u32)> {
    let subkey = wide(path);
    let mut hkey: HKEY = ptr::null_mut();
    let rc = unsafe { RegOpenKeyExW(HKEY_CURRENT_USER, subkey.as_ptr(), 0, access, &mut hkey) };
    if rc == 0 {
        return Ok(RegKey(hkey));
    }
    if rc != ERROR_FILE_NOT_FOUND || !create_missing {
        return Err(("RegOpenKeyExW", rc));
    }
    let rc = unsafe {
        RegCreateKeyExW(
            HKEY_CURRENT_USER,
            subkey.as_ptr(),
            0,
            ptr::null(),
            REG_OPTION_NON_VOLATILE,
            access,
            ptr::null(),
            &mut hkey,
            ptr::null_mut(),
        )
    };
    if rc == 0 {
        Ok(RegKey(hkey))
    } else {
        Err(("RegCreateKeyExW", rc))
    }
}

/// One HKCU Personalize DWORD. Err is a short `key=value` style reason.
fn read_personalize_dword(name: &str) -> Result<u32, String> {
    read_hkcu_dword(THEME_KEY, name)
}

/// The 4-byte REG_DWORD `name` under the HKCU key `key` (the key is a
/// parameter so tests can use a throwaway one).
fn read_hkcu_dword(key: &str, name: &str) -> Result<u32, String> {
    let value = wide(name);
    let hkey =
        open_hkcu_key(key, KEY_QUERY_VALUE, false).map_err(|(_, rc)| format!("open-rc-{rc}"))?;
    let mut data: u32 = 0;
    let mut size: u32 = 4;
    let mut kind: u32 = 0;
    let rc = unsafe {
        RegQueryValueExW(
            hkey.0,
            value.as_ptr(),
            ptr::null_mut(),
            &mut kind,
            &mut data as *mut u32 as *mut u8,
            &mut size,
        )
    };
    if rc != 0 {
        return Err(format!("query-rc-{rc}"));
    }
    if kind != REG_DWORD || size != 4 {
        return Err(format!("type-{kind}-size-{size}"));
    }
    Ok(data)
}

/// Tier 3's write of the two Personalize values. On a half write (one mode
/// flipped, the other failed) it repaints: callers skip their broadcast on
/// Err, and a changed value must not stay un-announced.
fn write_theme_registry(modes: Modes) -> Result<(), Box<dyn Error>> {
    write_modes_at(THEME_KEY, modes).map_err(|(msg, half_written)| {
        if half_written {
            broadcast_setting_change();
            poke_shell();
        }
        msg.into()
    })
}

/// Write `modes` as the two light/dark DWORDs under the HKCU key `key`
/// (Personalize; a parameter so tests can use a throwaway key), creating the
/// key if it is missing. Err is the message — it names the API and carries
/// the Win32 code — and whether the first value was written before the
/// second failed.
fn write_modes_at(key: &str, modes: Modes) -> Result<(), (String, bool)> {
    let hkey = open_hkcu_key(key, KEY_SET_VALUE, true)
        .map_err(|(api, rc)| (format!("{api} Personalize rc={rc}"), false))?;
    // Both results matter: a blocked write (AV/HIPS, ACL, hive error)
    // must surface as Err so tick's bounded retry runs, instead of being
    // logged as applied=registry while nothing changed. A half write
    // (Apps flipped, System failed) by a tick is harmless to the retry
    // gate: tick takes the failure baseline from a reading made AFTER
    // this returns. (A Toggle's half write during a pending retry does
    // read as a move and stands the retry down.)
    let mut any_ok = false;
    for (name, light) in [
        ("AppsUseLightTheme", modes.apps_light),
        ("SystemUsesLightTheme", modes.system_light),
    ] {
        let name_w = wide(name);
        let value = u32::from(light);
        let rc = unsafe {
            RegSetValueExW(
                hkey.0,
                name_w.as_ptr(),
                0,
                REG_DWORD,
                &value as *const u32 as *const u8,
                4,
            )
        };
        if rc != 0 {
            return Err((format!("RegSetValueExW {name} rc={rc}"), any_ok));
        }
        any_ok = true;
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

/// Why a configured theme_day/theme_night path isn't used.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ThemePathIssue {
    Missing,
    Directory,
    /// .themepack / .deskthemepack: a CAB whose DisplayName can't be read,
    /// and whose "open" re-installs the pack at every transition. Install it
    /// once (double-click), then point the config at the installed .theme.
    ThemePack,
    /// Any other extension: tier 2 would ShellExecute("open") it, launching
    /// whatever handler that file type has — at every transition.
    NotATheme,
    /// %LOCALAPPDATA%\Microsoft\Windows\Themes\Custom.theme — Windows'
    /// "Unsaved Theme" scratch file, rewritten on every personalization change.
    ScratchFile,
    /// A `%NAME%` survived expansion: the variable isn't defined (the
    /// environment is the one the app started with — changes need a restart).
    UnexpandedVar,
    /// Drive-relative (`C:x.theme`): would resolve against that drive's
    /// current directory, which the app doesn't control.
    NotAbsolute,
    /// On a network location (`\\server\share\…` or a network drive). Every
    /// look at such a file waits for its server, and one that doesn't answer
    /// held the event-loop thread for 20-40 s per check. Refused from the
    /// path's text and the drive's type, before anything opens it.
    NetworkPath,
}

impl ThemePathIssue {
    fn as_str(self) -> &'static str {
        match self {
            ThemePathIssue::Missing => "missing",
            ThemePathIssue::Directory => "directory",
            ThemePathIssue::ThemePack => "theme-pack",
            ThemePathIssue::NotATheme => "not-a-theme",
            ThemePathIssue::ScratchFile => "scratch-file",
            ThemePathIssue::UnexpandedVar => "unexpanded-var",
            ThemePathIssue::NotAbsolute => "not-absolute",
            ThemePathIssue::NetworkPath => "network-path",
        }
    }

    /// For the warning box.
    fn describe(self) -> &'static str {
        match self {
            ThemePathIssue::Missing => "the file doesn't exist",
            ThemePathIssue::Directory => "it is a folder, not a .theme file",
            ThemePathIssue::ThemePack => {
                "theme packs can't be used directly - double-click the pack once to \
                 install it, then point the config at the installed .theme file \
                 (under %LOCALAPPDATA%\\Microsoft\\Windows\\Themes)"
            }
            ThemePathIssue::NotATheme => "only .theme files are supported",
            ThemePathIssue::ScratchFile => {
                "Custom.theme is Windows' unsaved-theme scratch file - save the theme \
                 under a name in Settings > Personalization > Themes and use that file"
            }
            ThemePathIssue::UnexpandedVar => {
                "it uses an environment variable that isn't defined (variables are \
                 read when the app starts)"
            }
            ThemePathIssue::NotAbsolute => {
                "a drive letter without a backslash (like C:x.theme) is not supported - \
                 use a full path"
            }
            ThemePathIssue::NetworkPath => {
                "it is on a network location (a \\\\server\\share path or a network \
                 drive), and a server that doesn't answer would freeze the app - copy \
                 the .theme file to a folder on this PC and use that copy"
            }
        }
    }
}

/// Where an absolute path starts, as far as its text shows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PathRoot {
    /// A share: `\\server\share\…`, `\\?\UNC\server\share\…`, or the same
    /// through the device namespace, `\\.\UNC\…`.
    Unc,
    /// A drive letter (upper-cased): `X:\…`, `\\?\X:\…` or `\\.\X:\…`.
    Drive(char),
    /// Anything else: a volume GUID or another device name (`\\?\Volume{…}\`,
    /// `\\.\pipe\…`, `\\?\GLOBALROOT\…`), a path without a drive. Where a
    /// device name leads isn't in the text — `GLOBALROOT\Device\Mup\…` is
    /// the redirector, and so is a name someone defined to point at a share.
    Other,
}

/// Pure string work — nothing is opened. For a path in the form Win32 will
/// open (`on_network` normalizes it first). `/` counts as `\`; in a `\\?\`
/// path, where Win32 doesn't read it so, that can only turn an odd spelling
/// into a share, which is the side to err on.
fn path_root(path: &str) -> PathRoot {
    fn drive(s: &str) -> Option<char> {
        let mut chars = s.chars();
        match (chars.next(), chars.next()) {
            (Some(letter), Some(':')) if letter.is_ascii_alphabetic() => {
                Some(letter.to_ascii_uppercase())
            }
            _ => None,
        }
    }
    /// `s` without `prefix` (ASCII case-insensitive), if it starts with it.
    fn after<'a>(s: &'a str, prefix: &str) -> Option<&'a str> {
        let head = s.get(..prefix.len())?;
        head.eq_ignore_ascii_case(prefix)
            .then(|| &s[prefix.len()..])
    }
    let s = path.replace('/', "\\");
    let device = s.strip_prefix(r"\\?\").or_else(|| s.strip_prefix(r"\\.\"));
    if let Some(mut rest) = device {
        // `Global\` names the same set of device names again, any number
        // of times: `\\?\Global\UNC\…` is a share, `\\?\Global\C:\…` is C:.
        while let Some(tail) = after(rest, r"Global\") {
            rest = tail;
        }
        // `\\.\UNC\` is the redirector too (both prefixes lead to `\??\UNC`).
        if after(rest, r"UNC\").is_some() {
            return PathRoot::Unc;
        }
        return drive(rest).map_or(PathRoot::Other, PathRoot::Drive);
    }
    if s.starts_with(r"\\") {
        return PathRoot::Unc;
    }
    drive(&s).map_or(PathRoot::Other, PathRoot::Drive)
}

/// Whether drive `letter` is a network drive. `GetDriveTypeW` answers for a
/// drive letter from the process's device map, without asking the server:
/// measured 2026-11-11 on Windows 11 build 26300 at under a microsecond for a
/// local drive, a mapped drive and a drive letter linked to a server name
/// that doesn't resolve. A letter that isn't a drive is not remote.
fn drive_is_remote(letter: char) -> bool {
    let root = wide(&format!("{letter}:\\"));
    unsafe { GetDriveTypeW(root.as_ptr()) == DRIVE_REMOTE }
}

/// Whether `path` is on a network location: a share by its text, or a drive
/// letter that `drive_is_remote`. The text is read the way Win32 will open
/// it: `std::path::absolute` (GetFullPathNameW — string work, no file
/// system; a `\\?\` path comes back as given, since Win32 opens that form
/// as written) resolves `.`/`..`, repeated separators and trailing dots
/// first, so `\\.\C:\..\UNC\server\share\…` is seen as the share it opens.
/// Not seen: a local path that leads to a share through a junction or a
/// symlink, through a `subst` drive that `GetDriveTypeW` doesn't report as
/// remote, or through a device name other than a drive letter.
fn on_network(path: &Path, drive_is_remote: impl Fn(char) -> bool) -> bool {
    let as_opened = std::path::absolute(path).unwrap_or_else(|_| path.to_path_buf());
    match path_root(&as_opened.to_string_lossy()) {
        PathRoot::Unc => true,
        PathRoot::Drive(letter) => drive_is_remote(letter),
        PathRoot::Other => false,
    }
}

/// Whether an expanded path still contains a `%NAME%` token.
fn has_unexpanded_var(s: &str) -> bool {
    let mut parts = s.split('%');
    parts.next();
    // Every other segment sits between a pair of '%'.
    let inner: Vec<&str> = parts.collect();
    inner
        .iter()
        .step_by(2)
        .take(inner.len() / 2)
        .any(|name| !name.is_empty() && !name.contains(['\\', '/']))
}

/// Expand %VARS% in a configured path (ExpandEnvironmentStringsW; unknown
/// variables stay as written) and resolve a relative result against `base`
/// (the exe's folder — where config.json lives), never the process CWD.
fn expand_theme_path(raw: &str, base: &Path) -> PathBuf {
    let src = wide(raw);
    let mut buf = vec![0u16; 1024];
    let expanded = loop {
        let n =
            unsafe { ExpandEnvironmentStringsW(src.as_ptr(), buf.as_mut_ptr(), buf.len() as u32) }
                as usize;
        if n == 0 {
            break raw.to_string();
        }
        if n <= buf.len() {
            // n counts the terminating NUL.
            break String::from_utf16_lossy(&buf[..n - 1]);
        }
        buf.resize(n, 0);
    };
    let path = PathBuf::from(expanded.trim());
    if path.is_absolute() {
        path
    } else {
        base.join(path)
    }
}

/// Whether an (expanded) path may be used as a theme: an existing FILE, not
/// on a network location, with a `.theme` extension, that isn't Windows'
/// scratch Custom.theme. Ok is the path to use: as written, with the file
/// name as stored on disk (`with_stored_name`).
fn check_theme_path(path: &Path, scratch: &Path) -> Result<PathBuf, ThemePathIssue> {
    check_theme_path_with(path, scratch, drive_is_remote)
}

/// `check_theme_path`, with the network-drive question as a parameter so
/// tests can answer it.
fn check_theme_path_with(
    path: &Path,
    scratch: &Path,
    drive_is_remote: impl Fn(char) -> bool,
) -> Result<PathBuf, ThemePathIssue> {
    let prefixed = matches!(
        path.components().next(),
        Some(std::path::Component::Prefix(_))
    );
    if prefixed && !path.has_root() {
        return Err(ThemePathIssue::NotAbsolute);
    }
    // Before the first call that opens the path: everything below waits for
    // the server when the file is on a share.
    if on_network(path, drive_is_remote) {
        return Err(ThemePathIssue::NetworkPath);
    }
    if path.is_dir() {
        return Err(ThemePathIssue::Directory);
    }
    if !path.is_file() {
        // A %NAME% left after expansion means an undefined variable — said
        // only when nothing exists there, since file names may contain '%'.
        return Err(if has_unexpanded_var(&path.to_string_lossy()) {
            ThemePathIssue::UnexpandedVar
        } else {
            ThemePathIssue::Missing
        });
    }
    // The extension of the name as stored: canonicalizing expands 8.3 names
    // (`NIGHTT~1.THE`) and drops the trailing dots/spaces Win32 ignores.
    let real = fs::canonicalize(path).ok();
    let name = real
        .as_deref()
        .unwrap_or(path)
        .file_name()
        .map(|n| n.to_string_lossy().into_owned())
        .unwrap_or_default();
    let ext = name
        .trim_end_matches(['.', ' '])
        .rsplit_once('.')
        .map(|(_, e)| e.to_ascii_lowercase())
        .unwrap_or_default();
    match ext.as_str() {
        "theme" => {}
        "themepack" | "deskthemepack" => return Err(ThemePathIssue::ThemePack),
        _ => return Err(ThemePathIssue::NotATheme),
    }
    let same = match (&real, fs::canonicalize(scratch)) {
        (Some(a), Ok(b)) => *a == b,
        _ => false,
    };
    if same {
        return Err(ThemePathIssue::ScratchFile);
    }
    Ok(with_stored_name(path, real.as_deref()))
}

/// `path` with its file name as stored on disk, taken from `real` (its
/// canonical form). An 8.3 (`NIGHTT~1.THE`), trailing-dot or differently
/// cased spelling opens the same file, but a theme without a DisplayName is
/// named after its file, and tier 1 matches that name exactly. Kept as
/// written when the stored name, in the folder as written, isn't the same
/// file — a symlink to a file named differently.
fn with_stored_name(path: &Path, real: Option<&Path>) -> PathBuf {
    if let Some(name) = real.and_then(Path::file_name) {
        let respelled = path.with_file_name(name);
        if respelled.as_path() != path && fs::canonicalize(&respelled).ok().as_deref() == real {
            return respelled;
        }
    }
    path.to_path_buf()
}

fn stock_theme_file(theme: Theme) -> PathBuf {
    let leaf = match theme {
        Theme::Light => "aero.theme",
        Theme::Dark => "dark.theme",
    };
    windows_resources_dir().join("Themes").join(leaf)
}

fn scratch_theme_file() -> PathBuf {
    let local = std::env::var("LOCALAPPDATA").unwrap_or_default();
    PathBuf::from(local).join("Microsoft\\Windows\\Themes\\Custom.theme")
}

/// The .theme file to apply for `theme`: the configured one if it passes
/// `check_theme_path`, else the stock theme. The Err side of the second value
/// reports a configured path that was rejected (and its expansion), for the
/// caller to log.
fn resolve_theme_file_checked(
    theme: Theme,
    cfg: &Config,
) -> (PathBuf, Option<(String, PathBuf, ThemePathIssue)>) {
    let configured = match theme {
        Theme::Light => cfg.theme_day.as_deref(),
        Theme::Dark => cfg.theme_night.as_deref(),
    };
    let Some(raw) = configured.filter(|s| !s.trim().is_empty()) else {
        return (stock_theme_file(theme), None);
    };
    let base = config_path()
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));
    let path = expand_theme_path(raw, &base);
    match check_theme_path(&path, &scratch_theme_file()) {
        Ok(file) => (file, None),
        Err(issue) => (
            stock_theme_file(theme),
            Some((raw.to_string(), path, issue)),
        ),
    }
}

#[cfg(test)]
fn resolve_theme_file(theme: Theme, cfg: &Config) -> PathBuf {
    resolve_theme_file_checked(theme, cfg).0
}

/// Resolve a schedule slot: the file that will be applied (path policy +
/// stock fallback) and the modes read from THAT file, so the target that
/// ticks, Toggle and commit_watcher compare against is always the theme
/// actually applied — never a rejected configured file's.
fn resolve_slot(theme: Theme, cfg: &Config) -> Slot {
    let (file, rejected) = resolve_theme_file_checked(theme, cfg);
    let system = read_theme_ini(&file, "VisualStyles", "SystemMode").flatten();
    let app = read_theme_ini(&file, "VisualStyles", "AppMode").flatten();
    let (modes, keyed) = modes_from(theme, system.as_deref(), app.as_deref());
    Slot {
        theme,
        file,
        modes,
        keyed,
        rejected,
    }
}

/// Last sync-decidability state logged (0 = not yet computed, 1 = decidable,
/// 2 = undecidable), so `theme_sync` lines appear only on a change.
static SYNC_STATE_LOGGED: std::sync::atomic::AtomicU8 = std::sync::atomic::AtomicU8::new(0);

/// `sync_decidable`, logging a `theme_sync` line whenever the answer changes
/// (the first computation logs only if undecidable — the normal case is
/// silent).
fn sync_decidable_logged(day: &Slot, night: &Slot) -> bool {
    let decidable = sync_decidable(day, night);
    let state = if decidable { 1 } else { 2 };
    let prev = SYNC_STATE_LOGGED.swap(state, Ordering::Relaxed);
    if prev != state && !(prev == 0 && decidable) {
        let why = if decidable {
            "modes-differ"
        } else if day.keyed && night.keyed {
            "identical-modes"
        } else {
            "modes-not-in-theme-file"
        };
        log_event(&format!(
            "{} theme_sync decidable={} reason={} day={}{} night={}{}",
            Local::now().to_rfc3339(),
            decidable,
            why,
            modes_str(Some(day.modes)),
            if day.keyed { "" } else { "(guessed)" },
            modes_str(Some(night.modes)),
            if night.keyed { "" } else { "(guessed)" },
        ));
    }
    decidable
}

/// The slot this process last applied successfully (0 = none, 1 = day,
/// 2 = night) — Toggle's fallback when the screen's modes can't tell the
/// slots apart. Deliberately NOT in TickState: all Toggle writes there is
/// `note_toggle`'s flag on a pending retry.
static LAST_APPLIED_SLOT: std::sync::atomic::AtomicU8 = std::sync::atomic::AtomicU8::new(0);

fn note_applied_slot(theme: Theme) {
    let v = match theme {
        Theme::Light => 1,
        Theme::Dark => 2,
    };
    LAST_APPLIED_SLOT.store(v, Ordering::Relaxed);
}

fn last_applied_slot() -> Option<Theme> {
    match LAST_APPLIED_SLOT.load(Ordering::Relaxed) {
        1 => Some(Theme::Light),
        2 => Some(Theme::Dark),
        _ => None,
    }
}

/// Extra fields for a tick/toggle log line: the slot name whenever the
/// target modes alone don't identify it, and a note when sync is
/// undecidable. Empty in the common case (stock-like themes).
fn slot_note(slot: &Slot, decidable: bool) -> String {
    let mut s = String::new();
    if !decidable || slot.modes != Modes::from(slot.theme) {
        s.push_str(" slot=");
        s.push_str(slot.theme.slot_str());
    }
    if !decidable {
        s.push_str(" sync=undecidable");
    }
    s
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

/// Why the Personalize modes can't show whether applying `slot` took effect,
/// if they can't: the theme file doesn't declare them (the slot's modes are
/// a guess — forcing them over the theme would be wrong), or `pre`, the
/// screen just before the apply, already reads them. The reason is the log
/// value. Shared by tier 1's read-back and tier 2's commit watcher. Pure —
/// unit-tested.
fn unverifiable_reason(slot: &Slot, pre: Option<Modes>) -> Option<&'static str> {
    if !slot.keyed {
        Some("modes-not-in-theme-file")
    } else if pre == Some(slot.modes) {
        Some("already-matching")
    } else {
        None
    }
}

/// Watch a tier-2 apply for its registry effect; on a silent fail, force the
/// modes via tier 3. `pre` is the reading from before the apply: when the
/// modes can't show whether the apply committed (`unverifiable_reason`), it
/// only logs `commit_unverifiable`.
fn start_commit_watcher(slot: &Slot, pre: Option<Modes>) {
    let target = slot.modes;
    if let Some(reason) = unverifiable_reason(slot, pre) {
        log_event(&format!(
            "{} commit_unverifiable target={} slot={} reason={}",
            Local::now().to_rfc3339(),
            modes_str(Some(target)),
            slot.theme.slot_str(),
            reason,
        ));
        return;
    }
    let started = Instant::now();
    std::thread::spawn(move || {
        for _ in 0..25 {
            std::thread::sleep(Duration::from_millis(200));
            if read_modes() == Some(target) {
                log_event(&format!(
                    "{} commit_observed target={} after_ms={}",
                    Local::now().to_rfc3339(),
                    modes_str(Some(target)),
                    started.elapsed().as_millis(),
                ));
                return;
            }
        }
        log_event(&format!(
            "{} commit_timeout target={} actual={} after_ms={}",
            Local::now().to_rfc3339(),
            modes_str(Some(target)),
            modes_str(read_modes()),
            started.elapsed().as_millis(),
        ));
        // ShellExecute(.theme) lied about success — Settings UWP didn't actually apply.
        // Observed when the schedule fires while the user isn't interactive (sunset while
        // away, immediately after WTS_SESSION_UNLOCK, immediately after PBT_APMRESUMEAUTOMATIC).
        // Force the mode flip via direct registry write so at minimum the
        // theme's light/dark modes are correct; wallpaper won't change on
        // this path (would require IThemeManager2 — see CLAUDE.md).
        let fb_started = Instant::now();
        match write_theme_registry(target) {
            Ok(()) => {
                broadcast_setting_change();
                poke_shell();
                let mut confirmed = false;
                for _ in 0..10 {
                    std::thread::sleep(Duration::from_millis(100));
                    if read_modes() == Some(target) {
                        confirmed = true;
                        break;
                    }
                }
                log_event(&format!(
                    "{} fallback_registry target={} confirmed={} after_ms={}",
                    Local::now().to_rfc3339(),
                    modes_str(Some(target)),
                    confirmed,
                    fb_started.elapsed().as_millis(),
                ));
            }
            Err(e) => {
                log_event(&format!(
                    "{} fallback_registry_err target={} err=\"{}\"",
                    Local::now().to_rfc3339(),
                    modes_str(Some(target)),
                    sanitize_log_msg(&e.to_string()),
                ));
            }
        }
    });
}

/// Reads one value from a `.theme` (INI) file with Windows' own INI reader
/// (GetPrivateProfileStringW), not by hand, so what we see is exactly what
/// Windows reads: section and key names case-insensitive, whitespace around
/// `=` trimmed, enclosing quotes stripped, and the file decoded the way
/// Windows decodes it (UTF-16 with a BOM, otherwise the system ANSI code page
/// — system .theme files are Windows-1252, e.g. `aero.theme`'s raw `0xa9`
/// copyright byte).
///
/// None: `theme_file` isn't a readable file. Some(None): the key (or section)
/// is absent. Some(Some(v)): present, trimmed (possibly empty).
fn read_theme_ini(theme_file: &Path, section: &str, key: &str) -> Option<Option<String>> {
    // A default that can't be a real value tells "absent" from "empty".
    const ABSENT: &str = "\u{1}wts-absent\u{1}";
    // GetPrivateProfileStringW also returns the default when it can't OPEN
    // the file (sharing violation, ACL) — prove it's readable first, so a
    // transient read failure is None, never "the key is absent".
    if !theme_file.is_file() || fs::File::open(theme_file).is_err() {
        return None;
    }
    let section = wide(section);
    let key = wide(key);
    let default = wide(ABSENT);
    // Absolute: given a bare file name, GetPrivateProfileStringW looks in
    // %WINDIR%, not the directory is_file() just checked. And in the form
    // that opens at any length — std's is_file()/open() above add the
    // `\\?\` prefix themselves; without it the API can't open a long path
    // and returns the default, which would read as "absent".
    let abs = std::path::absolute(theme_file).unwrap_or_else(|_| theme_file.to_path_buf());
    let file = wide(&win32_path_arg(&abs.to_string_lossy()));
    let mut buf = vec![0u16; 1024];
    let n = unsafe {
        GetPrivateProfileStringW(
            section.as_ptr(),
            key.as_ptr(),
            default.as_ptr(),
            buf.as_mut_ptr(),
            buf.len() as u32,
            file.as_ptr(),
        )
    } as usize;
    let raw = String::from_utf16_lossy(&buf[..n.min(buf.len())]);
    if raw == ABSENT {
        return Some(None);
    }
    Some(Some(raw.trim().to_string()))
}

/// An absolute path as a Win32 file API should get it from this exe, which
/// has no longPathAware manifest: from 247 UTF-16 units (std's own
/// threshold, a little under MAX_PATH) in the `\\?\` form —
/// `\\?\UNC\server\share\...` for a share — which opens at any length;
/// shorter paths as they are. The prefix turns off Win32's own path
/// normalization (`/`, `.`, `..`, trailing dots), so `abs` must already be
/// normalized (std::path::absolute). Pure — unit-tested.
fn win32_path_arg(abs: &str) -> String {
    // std's limit counts the terminating NUL.
    const LEGACY_MAX_PATH: usize = 248;
    if abs.encode_utf16().count() + 1 < LEGACY_MAX_PATH
        || !Path::new(abs).is_absolute()
        || abs.starts_with(r"\\?\")
        || abs.starts_with(r"\\.\")
    {
        return abs.to_string();
    }
    match abs.strip_prefix(r"\\") {
        Some(share) => format!(r"\\?\UNC\{share}"),
        None => format!(r"\\?\{abs}"),
    }
}

/// The name Windows lists a `.theme` under — what tier 1 matches against
/// `ITheme::GetDisplayName`. `[Theme] DisplayName` may be a literal, OR an
/// SHLoadIndirectString resource reference such as
/// `@%SystemRoot%\System32\themeui.dll,-2060` (system themes use this — the
/// user-visible name is in a localized string table). With no DisplayName at
/// all, Windows uses the file name WITHOUT the `.theme` extension (MS Learn
/// "Theme file format"; confirmed 2026-09-28 by enumerating IThemeManager2
/// with a no-DisplayName theme in the user Themes folder, which was listed
/// under its file stem). None if the file is unreadable, the value is empty,
/// or indirect-string resolution fails.
fn resolve_theme_display_name(theme_file: &Path) -> Option<String> {
    match read_theme_ini(theme_file, "Theme", "DisplayName")? {
        None => theme_file
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned()),
        Some(v) if v.is_empty() => None,
        Some(v) if v.starts_with('@') => resolve_indirect_string(&v),
        Some(v) => Some(v),
    }
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
/// Matching is by name only (ITheme exposes no file path), so a name shared
/// with another theme file is a hazard: tier 1 could apply THAT theme —
/// e.g. a stale saved copy — and report success. For one of Windows' own
/// themes (unique resource-string names) the first match is used. For any
/// other file whose name another theme file also has, tier 1 declines (Err)
/// and tier 2 applies the configured file by its exact path — see
/// `pick_theme_index`.
///
/// SetCurrentTheme's success is then checked against the registry
/// (`tier1_readback`): a Windows build where it reports success without
/// applying would otherwise be logged as applied, recorded as in sync, and
/// kept by later wakes as if it were the user's override. Logs
/// `theme_manager2_apply` (with the read-back verdict), or
/// `theme_manager2_noop` when the modes never reached the slot's — the
/// caller then falls through to tier 2 as on an error, and logs the err
/// string of an Err.
fn apply_via_theme_manager2(slot: &Slot) -> Result<ReadBack, Box<dyn Error>> {
    let target_name = resolve_theme_display_name(&slot.file)
        .ok_or("could not resolve DisplayName from .theme file")?;
    let windows_theme = is_windows_theme(&slot.file);
    let collides =
        !windows_theme && name_taken_elsewhere(&slot.file, &target_name, &theme_search_dirs());
    let started = Instant::now();
    let (i, pre) = unsafe {
        let mgr =
            ThemeMgr::create().map_err(|hr| format!("CoCreateInstance/Init hr=0x{:08x}", hr))?;
        let n = mgr
            .count()
            .map_err(|hr| format!("GetThemeCount hr=0x{:08x}", hr))?;
        let mut matches = Vec::new();
        for i in 0..n {
            match mgr.theme_display_name(i) {
                Ok(name) if name == target_name => matches.push(i),
                Ok(_) => {}
                Err(hr) => log_event(&format!(
                    "{} theme_manager2_enum_skip i={} hr=0x{:08x}",
                    Local::now().to_rfc3339(),
                    i,
                    hr
                )),
            }
        }
        let i = pick_theme_index(&matches, windows_theme, collides, &target_name)?;
        if matches.len() > 1 {
            log_event(&format!(
                "{} theme_manager2_ambiguous display=\"{}\" matches={:?} using_idx={}",
                Local::now().to_rfc3339(),
                sanitize_log_msg(&target_name),
                matches,
                i,
            ));
        }
        // The screen right before the apply: the read-back can only tell
        // from a change.
        let pre = read_modes();
        mgr.set_current(i, THEME_APPLY_FLAG_NO_HOURGLASS)
            .map_err(|hr| format!("SetCurrentTheme i={} hr=0x{:08x}", i, hr))?;
        // mgr is released at the end of this block — right after
        // SetCurrentTheme, as before; the read-back needs no COM.
        (i, pre)
    };
    let applied_ms = started.elapsed().as_millis();
    let (verdict, readback_ms) = tier1_readback(slot, pre);
    if let ReadBack::NoOp(actual) = verdict {
        log_event(&format!(
            "{} theme_manager2_noop target={} actual={} display=\"{}\" idx={} after_ms={} readback_ms={}",
            Local::now().to_rfc3339(),
            modes_str(Some(slot.modes)),
            modes_str(Some(actual)),
            sanitize_log_msg(&target_name),
            i,
            applied_ms,
            readback_ms,
        ));
    } else {
        log_event(&format!(
            "{} theme_manager2_apply target={} display=\"{}\" idx={} after_ms={} readback={} readback_ms={}",
            Local::now().to_rfc3339(),
            modes_str(Some(slot.modes)),
            sanitize_log_msg(&target_name),
            i,
            applied_ms,
            verdict.as_str(),
            readback_ms,
        ));
    }
    Ok(verdict)
}

/// What tier 1's read-back concluded — see `readback_verdict`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReadBack {
    /// A reading matched the slot's modes: the apply took effect.
    Reached,
    /// The modes can't show it (`unverifiable_reason`, or "unreadable": no
    /// reading succeeded). SetCurrentTheme's success is trusted.
    Unverifiable(&'static str),
    /// SetCurrentTheme reported success, but the modes never reached the
    /// slot's (the last reading) — treated as a tier-1 failure.
    NoOp(Modes),
}

impl ReadBack {
    /// The `readback=` log value.
    fn as_str(self) -> &'static str {
        match self {
            ReadBack::Reached => "reached",
            ReadBack::Unverifiable(why) => why,
            ReadBack::NoOp(_) => "noop",
        }
    }
}

/// How long tier 1's read-back waits for the modes, in case Windows writes
/// them after SetCurrentTheme returns.
const READBACK_WINDOW: Duration = Duration::from_secs(2);
/// Interval between read-back readings.
const READBACK_POLL: Duration = Duration::from_millis(100);

/// Tier 1's read-back verdict from `readings`, the modes read after
/// SetCurrentTheme reported success (oldest first; polling stops at a
/// match). Only a reading that succeeded can show a no-op — an unreadable
/// registry proves nothing, as for the retry gate. Pure — unit-tested.
fn readback_verdict(slot: &Slot, pre: Option<Modes>, readings: &[Option<Modes>]) -> ReadBack {
    if let Some(reason) = unverifiable_reason(slot, pre) {
        return ReadBack::Unverifiable(reason);
    }
    if readings.contains(&Some(slot.modes)) {
        return ReadBack::Reached;
    }
    match readings.iter().rev().find_map(|r| *r) {
        Some(last) => ReadBack::NoOp(last),
        None => ReadBack::Unverifiable("unreadable"),
    }
}

/// Read the Personalize modes after SetCurrentTheme reported success until
/// they match `slot`'s or READBACK_WINDOW has passed; returns the verdict
/// and the milliseconds it took. Runs on the main thread, between tier 1
/// (which must stay on that STA thread) and the tier-2 fallback it may lead
/// to: registry reads only, no COM — and if SetCurrentTheme writes the modes
/// before returning, the first reading matches and nothing waits.
fn tier1_readback(slot: &Slot, pre: Option<Modes>) -> (ReadBack, u128) {
    let started = Instant::now();
    let mut readings = Vec::new();
    if unverifiable_reason(slot, pre).is_none() {
        loop {
            let reading = read_modes();
            readings.push(reading);
            if reading == Some(slot.modes) || started.elapsed() >= READBACK_WINDOW {
                break;
            }
            std::thread::sleep(READBACK_POLL);
        }
    }
    (
        readback_verdict(slot, pre, &readings),
        started.elapsed().as_millis(),
    )
}

/// Which enumerated theme tier 1 applies, given every index whose
/// DisplayName matched the configured file's. `taken_elsewhere`: another
/// theme file in the folders Windows lists also has this name (ignored for
/// Windows' own themes). Pure — unit-tested.
fn pick_theme_index(
    matches: &[i32],
    windows_theme: bool,
    taken_elsewhere: bool,
    name: &str,
) -> Result<i32, String> {
    if taken_elsewhere && !windows_theme {
        return Err(format!(
            "another theme file is also named \"{name}\" - give the configured theme a \
             unique name so tier 1 can't pick the wrong one"
        ));
    }
    match matches {
        [] => Err(format!("no installed theme matches DisplayName \"{name}\"")),
        [only] => Ok(*only),
        [first, ..] if windows_theme => Ok(*first),
        _ => Err(format!(
            "{} installed themes are named \"{name}\" - can't tell which is the configured file",
            matches.len()
        )),
    }
}

/// Windows' own theme root, %SystemRoot%\Resources.
fn windows_resources_dir() -> PathBuf {
    let win_dir = std::env::var("SystemRoot").unwrap_or_else(|_| "C:\\Windows".to_string());
    PathBuf::from(win_dir).join("Resources")
}

/// Whether `file` is one of Windows' own themes — directly inside a folder
/// of %SystemRoot%\Resources (Themes, Ease of Access Themes, ...), the
/// folders Windows lists.
fn is_windows_theme(file: &Path) -> bool {
    let Some(grandparent) = file.parent().and_then(Path::parent) else {
        return false;
    };
    match (
        fs::canonicalize(grandparent),
        fs::canonicalize(windows_resources_dir()),
    ) {
        (Ok(a), Ok(b)) => a == b,
        _ => false,
    }
}

/// Folders whose .theme files Windows lists in Settings (MS Learn "Theme
/// file format"): each folder directly in %SystemRoot%\Resources, and the
/// user's %LOCALAPPDATA%\Microsoft\Windows\Themes tree.
fn theme_search_dirs() -> Vec<(PathBuf, bool)> {
    let mut dirs = Vec::new();
    if let Ok(entries) = fs::read_dir(windows_resources_dir()) {
        dirs.extend(
            entries
                .filter_map(|e| e.ok().map(|e| e.path()))
                .filter(|p| p.is_dir())
                .map(|p| (p, false)),
        );
    }
    if let Some(user) = scratch_theme_file().parent() {
        dirs.push((user.to_path_buf(), true));
    }
    dirs
}

/// Whether a .theme file OTHER than `configured` in `dirs` (each
/// `(dir, recursive)`) has the DisplayName `name` — the file-stem fallback
/// included. Identity is by canonical path, so the configured file itself
/// never counts, however it was spelled.
fn name_taken_elsewhere(configured: &Path, name: &str, dirs: &[(PathBuf, bool)]) -> bool {
    fn walk(dir: &Path, recursive: bool, depth: u32, out: &mut Vec<PathBuf>) {
        let Ok(entries) = fs::read_dir(dir) else {
            return;
        };
        for p in entries.filter_map(|e| e.ok().map(|e| e.path())) {
            if p.is_dir() {
                if recursive && depth < 4 {
                    walk(&p, recursive, depth + 1, out);
                }
            } else if p
                .extension()
                .is_some_and(|x| x.eq_ignore_ascii_case("theme"))
            {
                out.push(p);
            }
        }
    }
    let me = fs::canonicalize(configured).ok();
    let mut files = Vec::new();
    for (dir, recursive) in dirs {
        walk(dir, *recursive, 0, &mut files);
    }
    files.iter().any(|f| {
        fs::canonicalize(f).ok() != me && resolve_theme_display_name(f).as_deref() == Some(name)
    })
}

/// Three-tier apply, best-to-worst:
///   1. IThemeManager2  — atomic, reliable, no Settings UWP, no AV-tripping broadcast.
///      Its success is read back; one that changed nothing falls through to (2).
///   2. ShellExecuteW(.theme) + commit_watcher — legacy. Watcher promotes to (3) on silent fail.
///   3. Direct registry write — flips light/dark mode but not wallpaper. Last resort.
///
/// Takes the slot as resolved by the caller (`resolve_slot`), so the modes
/// the caller decided against are the modes of the file applied here.
fn apply_theme(slot: &Slot) -> Result<&'static str, Box<dyn Error>> {
    let theme_file = &slot.file;
    if let Some((raw, expanded, issue)) = &slot.rejected {
        // Fell back to the stock theme — say so, or a typo / wrong file type
        // in theme_day/theme_night is invisible.
        log_event(&format!(
            "{} theme_path_rejected slot={} reason={} path=\"{}\" expanded=\"{}\" using=\"{}\"",
            Local::now().to_rfc3339(),
            slot.theme.slot_str(),
            issue.as_str(),
            sanitize_log_msg(raw),
            sanitize_log_msg(&expanded.to_string_lossy()),
            sanitize_log_msg(&theme_file.to_string_lossy()),
        ));
    }

    if theme_file.exists() {
        match apply_via_theme_manager2(slot) {
            // Reported success, but the modes never reached the target
            // (logged as theme_manager2_noop): fall through as on an error.
            Ok(ReadBack::NoOp(_)) => {}
            Ok(_) => return Ok("theme-manager2"),
            Err(e) => log_event(&format!(
                "{} theme_manager2_err target={} slot={} msg=\"{}\"",
                Local::now().to_rfc3339(),
                modes_str(Some(slot.modes)),
                slot.theme.slot_str(),
                sanitize_log_msg(&e.to_string()),
            )),
        }
    }

    let pre = read_modes();
    if theme_file.exists() && apply_theme_file(theme_file) {
        start_commit_watcher(slot, pre);
        start_settings_closer();
        std::thread::sleep(Duration::from_millis(300));
        poke_shell();
        return Ok("theme-file");
    }

    write_theme_registry(slot.modes)?;
    broadcast_setting_change();
    poke_shell();
    Ok("registry")
}

/// Make the HKCU Run value match `enable`: this exe's quoted path, or no
/// value. Writes only when the value differs (a missing value — e.g. deleted
/// by an AV quarantine — always differs, so the documented
/// re-assert-on-launch/Refresh recovery still works), and surfaces every
/// registry failure as Err instead of dropping it. Call it through
/// apply_auto_start, which first decides whether to register from here.
fn set_auto_start(enable: bool) -> Result<(), String> {
    let command = if enable {
        let exe = std::env::current_exe().map_err(|e| format!("current_exe: {e}"))?;
        Some(format!("\"{}\"", exe.to_string_lossy()))
    } else {
        None
    };
    set_run_value(RUN_KEY, APP_NAME, command.as_deref()).map(|_| ())
}

/// set_auto_start's registry work, with the key, the value name and the
/// command line as parameters — tests run it against a throwaway HKCU key,
/// never the real Run key. `Some(command)`: the REG_SZ value `name` under
/// `key` holds exactly `command` afterwards; the key is created if it is
/// missing, and nothing is written when the value already matches. `None`:
/// the value is gone afterwards; a value or a key that isn't there is
/// success, and no key is ever created to delete from. Ok says whether a
/// value was written or deleted (false: it was already as asked).
fn set_run_value(key: &str, name: &str, command: Option<&str>) -> Result<bool, String> {
    let reg_err = |(api, rc): (&str, u32)| format!("{api} Run rc={rc}");
    let name_w = wide(name);
    let access = KEY_QUERY_VALUE | KEY_SET_VALUE;
    let Some(want) = command else {
        let hkey = match open_hkcu_key(key, access, false) {
            Ok(hkey) => hkey,
            // No Run key at all → certainly no Run value to remove.
            Err((_, ERROR_FILE_NOT_FOUND)) => return Ok(false),
            Err(e) => return Err(reg_err(e)),
        };
        return match unsafe { RegDeleteValueW(hkey.0, name_w.as_ptr()) } {
            0 => Ok(true),
            ERROR_FILE_NOT_FOUND => Ok(false),
            rc => Err(format!("RegDeleteValueW Run rc={rc}")),
        };
    };
    let hkey = open_hkcu_key(key, access, true).map_err(reg_err)?;
    if unsafe { read_reg_sz(hkey.0, &name_w) }.as_deref() == Some(want) {
        return Ok(false);
    }
    let want_w = wide(want);
    let rc = unsafe {
        RegSetValueExW(
            hkey.0,
            name_w.as_ptr(),
            0,
            REG_SZ,
            want_w.as_ptr() as *const u8,
            (want_w.len() * 2) as u32,
        )
    };
    if rc == 0 {
        Ok(true)
    } else {
        Err(format!("RegSetValueExW Run rc={rc}"))
    }
}

/// Whether the HKCU key `key` has a value `name`, of any type. Some(false)
/// when the value or the whole key isn't there; None when that can't be
/// read. Read-only.
fn hkcu_value_exists(key: &str, name: &str) -> Option<bool> {
    let name_w = wide(name);
    let hkey = match open_hkcu_key(key, KEY_QUERY_VALUE, false) {
        Ok(hkey) => hkey,
        Err((_, ERROR_FILE_NOT_FOUND)) => return Some(false),
        Err(_) => return None,
    };
    let rc = unsafe {
        RegQueryValueExW(
            hkey.0,
            name_w.as_ptr(),
            ptr::null_mut(),
            ptr::null_mut(),
            ptr::null_mut(),
            ptr::null_mut(),
        )
    };
    match rc {
        0 => Some(true),
        ERROR_FILE_NOT_FOUND => Some(false),
        _ => None,
    }
}

/// Whether autostart is registered right now: the Run value exists, whatever
/// it points at. What load_config_ex asks when it repairs an empty
/// config.json (healed_config).
fn auto_start_registered() -> Option<bool> {
    hkcu_value_exists(RUN_KEY, APP_NAME)
}

/// The type and data of the value `name` under the HKCU key `key`; None
/// when it isn't there or can't be read. Read-only.
fn read_hkcu_value(key: &str, name: &str) -> Option<(u32, Vec<u8>)> {
    let name_w = wide(name);
    let hkey = open_hkcu_key(key, KEY_QUERY_VALUE, false).ok()?;
    let mut kind: u32 = 0;
    let mut size: u32 = 0;
    let rc = unsafe {
        RegQueryValueExW(
            hkey.0,
            name_w.as_ptr(),
            ptr::null_mut(),
            &mut kind,
            ptr::null_mut(),
            &mut size,
        )
    };
    if rc != 0 {
        return None;
    }
    let mut data = vec![0u8; size as usize];
    let rc = unsafe {
        RegQueryValueExW(
            hkey.0,
            name_w.as_ptr(),
            ptr::null_mut(),
            &mut kind,
            data.as_mut_ptr(),
            &mut size,
        )
    };
    // A value that grew between the two calls (ERROR_MORE_DATA) is
    // unreadable for this purpose.
    if rc != 0 {
        return None;
    }
    data.truncate(size as usize);
    Some((kind, data))
}

/// What Windows' own startup switch says about a Run entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StartupSwitch {
    On,
    /// Switched off in Settings → Apps → Startup or Task Manager → Startup
    /// apps: Windows skips the Run entry at login.
    Off,
    /// No recognizable record. Nothing is said then.
    Unknown,
}

/// Read a StartupApproved\Run value (its registry type and data). Windows
/// writes 12 bytes of REG_BINARY: a little-endian DWORD of flags, then the
/// FILETIME of the moment the entry was switched off (zero while it is on).
/// The entry is off when the lowest flag bit is set — 02 and 06 are "on",
/// 03 and 07 "off". Undocumented; taken from what Windows 11 build 26300
/// stores for entries that are on (02 for the Run entries of this account,
/// 06 for a machine-wide one) and from how startup managers read it. Any
/// other type or length is Unknown, since saying "off" for a record that
/// doesn't mean it would mislead. Pure — unit-tested.
fn startup_switch(kind: u32, data: &[u8]) -> StartupSwitch {
    if kind != REG_BINARY || data.len() != 12 {
        return StartupSwitch::Unknown;
    }
    if data[0] & 1 == 1 {
        StartupSwitch::Off
    } else {
        StartupSwitch::On
    }
}

/// Whether the switched-off line is due: the record `name` under the HKCU
/// key `key` (StartupApproved\Run; parameters so tests can use a throwaway
/// key) says Off, and `logged` hasn't been set yet — it is set here. A
/// missing, unreadable or unrecognized record never makes a line. Read-only.
fn startup_switch_note_due(key: &str, name: &str, logged: &AtomicBool) -> bool {
    read_hkcu_value(key, name)
        .is_some_and(|(kind, data)| startup_switch(kind, &data) == StartupSwitch::Off)
        && !logged.swap(true, Ordering::Relaxed)
}

/// With autostart registered: one line, once per run of the app, when the
/// user has switched the entry off in Windows' Startup settings — the
/// explanation for "auto_start is true but the app doesn't start at login".
/// The switch itself is never touched: it is the user's, and overriding it
/// is what malware does.
fn note_startup_switch() {
    static LOGGED: AtomicBool = AtomicBool::new(false);
    if startup_switch_note_due(STARTUP_APPROVED_RUN_KEY, APP_NAME, &LOGGED) {
        log_event(&format!(
            "{} autostart_disabled_in_startup_settings auto_start=true \
             note=\"switched off in Settings > Apps > Startup or Task Manager > Startup apps; \
             the app will not start at login until it is switched on there\"",
            Local::now().to_rfc3339(),
        ));
    }
}

/// A Windows path as the list of names it consists of, for comparing two
/// spellings of a folder: a `\\?\` or `\\.\` prefix is dropped (`\\?\UNC\`
/// reads as `\\`), `/` counts as `\`, repeated and trailing separators and
/// `.` vanish, and letters are lower-cased. A share's key starts with `\\`,
/// so it never equals a drive path. None for a path that doesn't start at a
/// drive letter or a share (relative, drive-relative, a volume GUID).
/// 8.3 short names are a different spelling this can't see through — the
/// caller passes the long spelling as well (path_spellings). Pure.
fn path_key(path: &str) -> Option<Vec<String>> {
    let s = path.trim().replace('/', "\\");
    let device = s.strip_prefix(r"\\?\").or_else(|| s.strip_prefix(r"\\.\"));
    let (unc, rest) = match device {
        Some(rest) => match rest.get(..4) {
            Some(head) if head.eq_ignore_ascii_case(r"UNC\") => (true, &rest[4..]),
            _ => (false, rest),
        },
        None => match s.strip_prefix(r"\\") {
            Some(rest) => (true, rest),
            None => (false, s.as_str()),
        },
    };
    if !unc {
        // `X:\…` only: `C:name` (drive-relative) is not `C:\name`.
        let b = rest.as_bytes();
        if !(b.len() >= 3 && b[0].is_ascii_alphabetic() && b[1] == b':' && b[2] == b'\\') {
            return None;
        }
    }
    let mut names: Vec<String> = rest
        .split('\\')
        .filter(|name| !name.is_empty() && *name != ".")
        .map(str::to_lowercase)
        .collect();
    if unc {
        // Server and share.
        if names.len() < 2 {
            return None;
        }
        names.insert(0, r"\\".to_string());
    }
    Some(names)
}

/// Whether the exe sits in, or anywhere below, one of the temporary
/// folders. Each side is a list of spellings of the same thing — the path
/// as given and, where it could be resolved, its long form (path_spellings)
/// — because %TEMP% is often stored in 8.3 form
/// (`C:\Users\LONGUS~1\AppData\Local\Temp`) while the exe's path is not, or
/// the other way round; one matching pair is enough. Compared by path_key,
/// name by name, so `C:\Temporary` is not inside `C:\Temp`. A temp folder
/// that is only a drive or a share root, or not absolute, is ignored: with
/// TEMP set to `D:\` every exe on that drive would count. Pure —
/// unit-tested.
fn in_temp_folder(exe_spellings: &[String], temp_spellings: &[String]) -> bool {
    let exes: Vec<Vec<String>> = exe_spellings.iter().filter_map(|s| path_key(s)).collect();
    temp_spellings
        .iter()
        .filter_map(|s| path_key(s))
        .filter(|temp| {
            let root_len = if temp[0] == r"\\" { 3 } else { 1 };
            temp.len() > root_len
        })
        .any(|temp| {
            exes.iter()
                .any(|exe| exe.len() > temp.len() && exe[..temp.len()] == temp[..])
        })
}

/// `path` as given plus, when it can be resolved, its canonical form (long
/// names, links followed, `\\?\` prefix — path_key reads both).
fn path_spellings(path: &Path) -> Vec<String> {
    let mut spellings = vec![path.to_string_lossy().into_owned()];
    if let Ok(real) = fs::canonicalize(path) {
        spellings.push(real.to_string_lossy().into_owned());
    }
    spellings
}

/// The temporary folders an exe may have been unpacked to, in every
/// spelling: %TMP% and %TEMP% (where Explorer extracts a file started from
/// inside a zip, and where archivers and browsers put theirs),
/// %LOCALAPPDATA%\Temp (their usual value, in case this process was started
/// with other ones) and %SystemRoot%\Temp. Not GetTempPathW: without TMP
/// and TEMP it answers with the profile folder, and then the recommended
/// install folder would count as temporary.
fn temp_folder_spellings() -> Vec<String> {
    temp_folders(|name| std::env::var_os(name))
        .iter()
        .flat_map(|f| path_spellings(f))
        .collect()
}

/// temp_folder_spellings' list of folders, as the environment names them
/// (`var`: a variable's value — a parameter so tests don't depend on the
/// real environment). A variable that is unset or empty contributes
/// nothing. Pure — unit-tested.
fn temp_folders(var: impl Fn(&str) -> Option<std::ffi::OsString>) -> Vec<PathBuf> {
    let var = |name: &str| var(name).filter(|v| !v.is_empty());
    let mut folders: Vec<PathBuf> = ["TMP", "TEMP"]
        .into_iter()
        .filter_map(&var)
        .map(PathBuf::from)
        .collect();
    folders.extend(
        ["LOCALAPPDATA", "SystemRoot"]
            .into_iter()
            .filter_map(&var)
            .map(|base| PathBuf::from(base).join("Temp")),
    );
    folders
}

/// Whether this exe runs from a temporary folder — typically started from
/// inside the downloaded zip. Worked out once: neither the exe's place nor
/// the environment changes while the app runs.
fn exe_in_temp_folder() -> bool {
    static IN_TEMP: OnceLock<bool> = OnceLock::new();
    *IN_TEMP.get_or_init(|| {
        std::env::current_exe()
            .is_ok_and(|exe| in_temp_folder(&path_spellings(&exe), &temp_folder_spellings()))
    })
}

/// What apply_auto_start does about the Run value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AutoStartPlan {
    /// Point the Run value at this exe.
    Register,
    /// Remove the Run value.
    Remove,
    /// `auto_start` is on, but the exe runs from a temporary folder, which
    /// Windows or the archiver deletes: the Run value is left exactly as it
    /// is (one that exists may belong to the real install), and the user is
    /// told.
    LeaveTempFolder,
}

/// Pure — unit-tested. `auto_start: false` removes the value wherever the
/// exe runs, as before.
fn auto_start_plan(enable: bool, exe_in_temp: bool) -> AutoStartPlan {
    match (enable, exe_in_temp) {
        (false, _) => AutoStartPlan::Remove,
        (true, true) => AutoStartPlan::LeaveTempFolder,
        (true, false) => AutoStartPlan::Register,
    }
}

/// The running-from-a-temporary-folder box was shown (once per run).
static TEMP_FOLDER_BOX_SHOWN: AtomicBool = AtomicBool::new(false);

/// Autostart not registered because the exe runs from a temporary folder:
/// one log line and one warning box per run of the app. The exe's path goes
/// into the box only, never into events.log (it names the user's profile).
fn report_temp_folder() {
    static LOGGED: AtomicBool = AtomicBool::new(false);
    if !LOGGED.swap(true, Ordering::Relaxed) {
        log_event(&format!(
            "{} autostart_skipped reason=temp-folder",
            Local::now().to_rfc3339(),
        ));
    }
    let folder = std::env::current_exe()
        .ok()
        .and_then(|exe| exe.parent().map(|d| d.display().to_string()))
        .unwrap_or_default();
    let body = format!(
        "WinThemeSwitcher is running from a temporary folder:\n{folder}\n\n\
         That usually means it was started from inside the downloaded zip. Such \
         folders get deleted, so the app has not set itself up to start with Windows \
         from here, and settings saved here will be lost.\n\n\
         To install it: Quit, extract the zip to a permanent folder that only your \
         account can write to — for example %LOCALAPPDATA%\\Programs\\WinThemeSwitcher \
         — and start win-theme-switcher.exe from there.",
    );
    show_warning_once(
        &TEMP_FOLDER_BOX_SHOWN,
        "WinThemeSwitcher — Running from a temporary folder",
        body,
    );
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

/// Bring the autostart registration in line with config.json's `auto_start`
/// (auto_start_plan), with every outcome that leaves the app not starting
/// at login traceable in events.log: a failed registry call, a temporary
/// folder, or the user's own switch in Windows' Startup settings.
fn apply_auto_start(enable: bool) {
    let plan = auto_start_plan(enable, enable && exe_in_temp_folder());
    match carry_out_auto_start(plan, set_auto_start) {
        AutoStartOutcome::Registered => note_startup_switch(),
        AutoStartOutcome::Removed => {}
        AutoStartOutcome::LeftAlone => report_temp_folder(),
        AutoStartOutcome::Failed { op, msg } => log_event(&format!(
            "{} autostart_err op={} msg=\"{}\"",
            Local::now().to_rfc3339(),
            op,
            sanitize_log_msg(&msg),
        )),
    }
}

/// How carrying out an AutoStartPlan ended — what apply_auto_start then
/// logs or tells the user.
#[derive(Debug, PartialEq, Eq)]
enum AutoStartOutcome {
    /// The Run value points at this exe (written now, or already so).
    Registered,
    /// There is no Run value (deleted now, or already gone).
    Removed,
    /// Nothing was done to the Run value (AutoStartPlan::LeaveTempFolder).
    LeftAlone,
    /// The registry call failed: the `op=` of the autostart_err line and
    /// its message.
    Failed { op: &'static str, msg: String },
}

/// Carry out `plan` through `set_run` — set_auto_start; a parameter so
/// tests can watch it, or point it at a throwaway key. LeaveTempFolder
/// never calls it: the Run value stays exactly as it is, neither rewritten
/// nor removed.
fn carry_out_auto_start(
    plan: AutoStartPlan,
    set_run: impl FnOnce(bool) -> Result<(), String>,
) -> AutoStartOutcome {
    let register = match plan {
        AutoStartPlan::LeaveTempFolder => return AutoStartOutcome::LeftAlone,
        AutoStartPlan::Register => true,
        AutoStartPlan::Remove => false,
    };
    match set_run(register) {
        Ok(()) if register => AutoStartOutcome::Registered,
        Ok(()) => AutoStartOutcome::Removed,
        Err(msg) => AutoStartOutcome::Failed {
            op: if register { "set" } else { "delete" },
            msg,
        },
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

const TRAY_TOOLTIP: &str = "WinThemeSwitcher";
/// Delay between tray-registration probes, and how many to make (~2 min).
const TRAY_PROBE_DELAY: Duration = Duration::from_secs(20);
const TRAY_PROBE_ATTEMPTS: u32 = 6;

/// Whether the icon is registered with the taskbar: NIM_MODIFY (what
/// set_tooltip sends) fails for an icon the taskbar doesn't have. Needed
/// because tray-icon (0.21+) ignores a failed NIM_ADD in build() and waits
/// for TaskbarCreated — fine when Explorer isn't up yet, but a taskbar that
/// is up and rejects the add would otherwise leave the app running with no
/// icon and no trace.
fn tray_registered(tray: &TrayIcon) -> bool {
    tray.set_tooltip(Some(TRAY_TOOLTIP)).is_ok()
}

fn schedule_tray_probe(attempt: u32) {
    std::thread::spawn(move || {
        std::thread::sleep(TRAY_PROBE_DELAY);
        if let Some(proxy) = EVENT_PROXY.get() {
            let _ = proxy.send_event(AppEvent::TrayProbe(attempt));
        }
    });
}

/// A tray-registration box is already on screen.
static TRAY_BOX_OPEN: AtomicBool = AtomicBool::new(false);

/// Runs on the main thread (TrayIcon is not Send). If the icon still isn't
/// registered, has tray-icon redo its own remove + add — the same thing it
/// does when Explorer restarts — and checks again; after the last attempt,
/// says so in a box instead of running invisibly.
fn handle_tray_probe(tray: &TrayIcon, attempt: u32) {
    let now = || Local::now().to_rfc3339();
    if tray_registered(tray) {
        log_event(&format!("{} tray_registered attempt={}", now(), attempt));
        return;
    }
    let msg = wide("TaskbarCreated");
    unsafe {
        let id = RegisterWindowMessageW(msg.as_ptr());
        if id != 0 {
            SendMessageW(tray.window_handle(), id, 0, 0);
        }
    }
    let ok = tray_registered(tray);
    log_event(&format!(
        "{} tray_register_retry attempt={} ok={}",
        now(),
        attempt,
        ok
    ));
    if ok {
        return;
    }
    if attempt < TRAY_PROBE_ATTEMPTS {
        schedule_tray_probe(attempt + 1);
        return;
    }
    log_event(&format!("{} tray_register_failed", now()));
    show_warning_async(
        &TRAY_BOX_OPEN,
        "WinThemeSwitcher — No tray icon",
        "WinThemeSwitcher is running and still switches the theme at sunrise and \
         sunset, but Windows didn't accept its notification-area icon, so its menu \
         isn't reachable.\n\nThe icon is added automatically when the taskbar \
         restarts: restart Windows Explorer (Task Manager), or sign out and back in."
            .to_string(),
    );
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
    let day = resolve_slot(Theme::Light, cfg);
    let night = resolve_slot(Theme::Dark, cfg);
    let decidable = sync_decidable_logged(&day, &night);
    let slot = if want == Theme::Light { &day } else { &night };
    let current = read_modes();

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

    // Decide with the PRE-tick state; settle_tick records the outcome once
    // the apply result is known (a failed apply must never advance
    // reconciled_next — see TickState).
    let action = decide_tick(
        kind,
        current,
        slot.modes,
        now_utc,
        next_utc,
        clock_stepped_back,
        decidable,
        state,
    );

    let (outcome, result) = match action {
        TickAction::Apply => match apply_theme(slot) {
            Ok(method) => {
                note_applied_slot(want);
                (TickOutcome::Reconciled, format!("applied={}", method))
            }
            // The screen is read NOW, after the failure, so a partial effect
            // of this apply never reads as the user's move.
            Err(e) => (
                TickOutcome::ApplyFailed {
                    after: read_modes(),
                },
                format!("err=\"{}\"", sanitize_log_msg(&e.to_string())),
            ),
        },
        TickAction::SkipInSync => (TickOutcome::Reconciled, "applied=skip".to_string()),
        TickAction::SkipOverride => (TickOutcome::Reconciled, "applied=skip-override".to_string()),
        TickAction::CancelRetry => (
            TickOutcome::Reconciled,
            "applied=skip-user-intervened".to_string(),
        ),
    };
    let settled = settle_tick(state, kind, outcome, Utc::now(), next_utc);
    let retry_note = match settled.retry {
        None => String::new(),
        Some(RetryNote::Attempt(n)) => format!(" retry={}", n),
        Some(RetryNote::Exhausted) => " retry=exhausted".to_string(),
    };

    // Stamp at write time, not tick start — apply_theme logs detail lines
    // (theme_manager2_apply, theme_manager2_err) mid-tick, and reusing the
    // tick-start timestamp here made this summary line sort before them.
    log_event(&format!(
        "{} cause={} current={} target={}{} {}{} next={}",
        Local::now().to_rfc3339(),
        cause,
        modes_str(current),
        modes_str(Some(slot.modes)),
        slot_note(slot, decidable),
        result,
        retry_note,
        settled.deadline.with_timezone(&Local).to_rfc3339(),
    ));

    arm(elwt, state, settled.deadline, (now_utc, now_mono));
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
/// still at the lock screen; the unlock that follows is then a no-op
/// (SkipInSync — or SkipOverride when sync is undecidable or an override is
/// being preserved).
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

/// What creating the single-instance mutex says about another copy.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum InstanceClaim {
    /// The mutex is new: this is the only copy.
    First,
    /// It existed and was opened: another copy holds it.
    Duplicate,
    /// It exists, but this process may not open it — `CreateMutexW` asks for
    /// full access to an existing mutex, and a copy started with fewer
    /// rights than the first (at low integrity, say) is refused.
    DuplicateNoAccess,
    /// Creating it failed some other way (the code is the last error). No
    /// reason to refuse to run.
    Unknown(u32),
}

impl InstanceClaim {
    /// For a claim that means another copy is running — show the notice
    /// and exit — its events.log text; `None` means carry on. The reason is
    /// added only when the other copy's mutex couldn't be opened, so the
    /// ordinary line reads as it always has.
    fn duplicate_log(self) -> Option<&'static str> {
        match self {
            InstanceClaim::Duplicate => Some("duplicate_instance action=exit"),
            InstanceClaim::DuplicateNoAccess => {
                Some("duplicate_instance action=exit reason=access-denied")
            }
            InstanceClaim::First | InstanceClaim::Unknown(_) => None,
        }
    }
}

/// Read a `CreateMutexW` result: whether a handle came back, and the last
/// error captured right after the call.
fn instance_claim(handle_is_null: bool, last_error: u32) -> InstanceClaim {
    match (handle_is_null, last_error) {
        (false, ERROR_ALREADY_EXISTS) => InstanceClaim::Duplicate,
        (false, _) => InstanceClaim::First,
        (true, ERROR_ACCESS_DENIED) => InstanceClaim::DuplicateNoAccess,
        (true, code) => InstanceClaim::Unknown(code),
    }
}

/// Whether a mutex name is taken, from an `OpenMutexW` attempt: it opened,
/// or failed for any reason but "no such name" (refused again, or the name
/// belongs to another kind of object).
fn mutex_name_taken(open_is_null: bool, last_error: u32) -> bool {
    !open_is_null || last_error != ERROR_FILE_NOT_FOUND
}

/// Settle the claim from the two calls it may take, each given as a closure
/// that returns (no handle came back, the last error captured right after):
/// `create` is `CreateMutexW`, `probe` is `OpenMutexW` for the least access.
/// Creating that is refused with access denied says the mutex exists but
/// may not be opened — or that this process may not create one at all. The
/// probe tells: only a name that isn't there fails with "not found". Then
/// either the other copy left between the two calls, and creating works
/// now, or creating is what is refused; one more round shows which.
fn settle_instance_claim(
    mut create: impl FnMut() -> (bool, u32),
    mut probe: impl FnMut() -> (bool, u32),
) -> InstanceClaim {
    for _ in 0..2 {
        let (handle_is_null, last_error) = create();
        let claim = instance_claim(handle_is_null, last_error);
        if claim != InstanceClaim::DuplicateNoAccess {
            return claim;
        }
        let (open_is_null, open_error) = probe();
        if mutex_name_taken(open_is_null, open_error) {
            return claim;
        }
    }
    // Refused twice with nothing there to open: a creation failure.
    InstanceClaim::Unknown(ERROR_ACCESS_DENIED)
}

/// `OpenMutexW` for `SYNCHRONIZE` only, closed again at once: (no handle
/// came back, the last error captured right after the call).
fn probe_mutex(name: &[u16]) -> (bool, u32) {
    unsafe {
        let probe = OpenMutexW(SYNCHRONIZATION_SYNCHRONIZE, 0, name.as_ptr());
        let last = GetLastError();
        let is_null = (probe as usize) == 0;
        if !is_null {
            CloseHandle(probe);
        }
        (is_null, last)
    }
}

/// Create the named mutex, or learn that another process has it. A handle
/// to a new or opened mutex is intentionally leaked — it must live exactly
/// as long as the process (the OS closes it at exit).
fn claim_instance_mutex(name: &str) -> InstanceClaim {
    let name = wide(name);
    settle_instance_claim(
        || unsafe {
            let handle = CreateMutexW(ptr::null(), 0, name.as_ptr());
            let last = GetLastError();
            ((handle as usize) == 0, last)
        },
        || probe_mutex(&name),
    )
}

/// Claim the per-session single-instance mutex. `Duplicate` and
/// `DuplicateNoAccess` mean another copy is running; a failure to create
/// the mutex is logged here and is no reason to refuse to run.
fn claim_single_instance() -> InstanceClaim {
    let claim = claim_instance_mutex("Local\\WinThemeSwitcher.single-instance");
    if let InstanceClaim::Unknown(code) = claim {
        log_event(&format!(
            "{} single_instance_err code={}",
            Local::now().to_rfc3339(),
            code,
        ));
    }
    claim
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
    if let Some(line) = claim_single_instance().duplicate_log() {
        // Best effort: a copy too restricted to open the first copy's mutex
        // (low integrity) can't append to events.log in an ordinary folder
        // either, and then the notice below is the only sign.
        log_event(&format!("{} {}", Local::now().to_rfc3339(), line));
        show_message_box(
            "WinThemeSwitcher",
            "WinThemeSwitcher is already running — look for its icon in the \
             notification area.",
            MB_OK | MB_ICONINFORMATION,
        );
        std::process::exit(0);
    }
    if let Err(e) = run() {
        // Fail loudly: tray-window creation, event-loop build errors, and
        // event-loop death all used to be silent exits. (A taskbar that isn't
        // up yet no longer fails build() — see tray_registered.)
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
        // A first run whose config.json couldn't be created (`from_file`
        // false) carries on like any first run, on the defaults; the user
        // has been told (load_config_logged), and the location boxes then
        // don't send them to a file that isn't there.
        Ok(Loaded { mut cfg, from_file }) => {
            if !cfg.has_location() {
                acquire_location(&mut cfg, from_file);
            }
            apply_auto_start(cfg.auto_start);
            cfg
        }
        Err(e) => {
            report_config_error(&e);
            let mut cfg = Config::default();
            // In memory only, and without the first-run boxes: the config
            // error is what the user has to deal with. A failed lookup is
            // in the log.
            if let Ok((lat, lon)) = try_get_windows_location() {
                cfg.latitude = lat;
                cfg.longitude = lon;
            }
            cfg
        }
    };
    report_theme_path_problems(&cfg);

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
        .with_tooltip(TRAY_TOOLTIP);
    if let Some(icon) = make_tray_icon() {
        tray_builder = tray_builder.with_icon(icon);
    }
    let tray = tray_builder.build()?;
    if !tray_registered(&tray) {
        // Expected at an early login (Explorer not up yet): tray-icon adds
        // the icon itself when the taskbar broadcasts TaskbarCreated.
        log_event(&format!(
            "{} tray_register_pending attempt=0",
            Local::now().to_rfc3339(),
        ));
        schedule_tray_probe(1);
    }

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
        Event::UserEvent(AppEvent::TrayProbe(attempt)) => handle_tray_probe(&tray, attempt),
        Event::UserEvent(AppEvent::Menu(id)) => {
            if id == quit_id {
                elwt.exit();
            } else if id == open_cfg_id {
                // A failure is in the log (open_config_err).
                let _ = open_config_in_editor();
            } else if id == toggle_id {
                // A deliberate manual override: applies the opposite theme
                // and intentionally does NOT tick, change what TickState
                // knows about the schedule (reconciled_next, the armed
                // deadline, the clock mark), or disturb the pending
                // WaitUntil — so the override survives lock/unlock (see
                // decide_tick) and resets at the next natural transition,
                // exactly like an override made in Settings. The one thing
                // it records is note_toggle: if a failed-apply retry is
                // pending, a Toggle that applied stands it down. The modes
                // alone show that only for an odd number of toggles — a
                // second one lands back on the retry baseline, and through
                // v0.5.1 the retry then reverted it.
                let day = resolve_slot(Theme::Light, &cfg);
                let night = resolve_slot(Theme::Dark, &cfg);
                let decidable = sync_decidable_logged(&day, &night);
                let before = read_modes();
                let fallback = if cfg.has_location() {
                    schedule(Utc::now(), cfg.latitude, cfg.longitude).0
                } else {
                    Theme::Light
                };
                let target = toggle_target(
                    before,
                    day.modes,
                    night.modes,
                    decidable,
                    last_applied_slot(),
                    fallback,
                );
                let slot = if target == Theme::Light { &day } else { &night };
                let result = apply_theme(slot);
                note_toggle(&mut state, result.is_ok());
                let outcome = match result {
                    Ok(method) => {
                        note_applied_slot(target);
                        format!("applied={}", method)
                    }
                    Err(e) => format!("err=\"{}\"", sanitize_log_msg(&e.to_string())),
                };
                log_event(&format!(
                    "{} cause=toggle current={} target={}{} {}",
                    Local::now().to_rfc3339(),
                    modes_str(before),
                    modes_str(Some(slot.modes)),
                    slot_note(slot, decidable),
                    outcome,
                ));
            } else if id == refresh_id {
                let (assert_autostart, error) =
                    adopt_reloaded_config(load_config_logged(), &mut cfg, &mut have_disk_cfg);
                match error {
                    Some(e) => report_config_error(&e),
                    None => {
                        if !cfg.has_location() {
                            // Holds the event loop for up to
                            // LOCATION_LOOKUP_CAP; a failure is logged, and
                            // the tick below logs skipped=no-location.
                            if let Ok((lat, lon)) = try_get_windows_location() {
                                cfg.latitude = lat;
                                cfg.longitude = lon;
                                persist_config(&cfg);
                            }
                        }
                        report_theme_path_problems(&cfg);
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
        // self-heal instead of erroring on every launch.
        let t = TempConfig::new("empty", Some("  \n"));
        let cfg = load_config_at(&t.0).expect("empty file must heal to defaults");
        assert!(!cfg.has_location());
        // Not the default as such: load_config_at stands for a PC where
        // autostart is registered, and a repair follows the registration
        // (repairing_an_empty_config_leaves_autostart_as_it_is_registered).
        assert!(cfg.auto_start);
        let written = fs::read_to_string(&t.0).unwrap();
        let reparsed: Config = serde_json::from_str(&written).unwrap();
        assert!(!reparsed.has_location());
    }

    #[test]
    fn config_unknown_and_missing_fields_are_defaulted() {
        let json = r#"{"theme_day": "C:\\d.theme", "some_future_field": true}"#;
        let t = TempConfig::new("partial", Some(json));
        let cfg = load_config_at(&t.0).expect("unknown/missing fields are not errors");
        // What the file does set is read (not swapped for plain defaults).
        assert_eq!(cfg.theme_day.as_deref(), Some("C:\\d.theme"));
        assert!(cfg.extra.contains_key("some_future_field"));
        assert!(
            !cfg.has_location(),
            "no coordinate keys: Windows Location is asked"
        );
        assert!(cfg.auto_start, "missing fields fall back to defaults");
        assert!(cfg.theme_night.is_none());
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json);
    }

    // --- coordinates: range, both-or-neither, saved precision (v0.6.0) ---

    #[test]
    fn out_of_range_coordinates_are_a_config_error_and_the_file_is_kept() {
        // The roadmap's case: 225 typed for 22.5 used to give a schedule
        // that was hours off, with nothing said.
        let json = r#"{"latitude": 225, "longitude": 46.7, "theme_night": "C:\\night.theme"}"#;
        let t = TempConfig::new("range-225", Some(json));
        let err = load_config_at(&t.0).expect_err("225 is not a latitude");
        let bad_lat = "config.json: latitude must be between -90 and 90";
        let bad_lon = "config.json: longitude must be between -180 and 180";
        assert_eq!(err, bad_lat);
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json, "file untouched");

        let swapped = format!("{bad_lat} (are the two swapped?)");
        let from_360 = format!("{bad_lon} (a 0 to 360 longitude? West is negative: subtract 360)");
        for (json, want) in [
            (r#"{"latitude": 24.7, "longitude": -180.5}"#, bad_lon),
            (
                r#"{"latitude": -90.0001, "longitude": 1e300}"#,
                "config.json: latitude must be between -90 and 90; \
                 longitude must be between -180 and 180",
            ),
            // Typed in the other order, a common slip: ask. Any latitude
            // that could be a longitude gets the question, typo or not.
            (r#"{"latitude": -122.33, "longitude": 47.61}"#, &swapped),
            (r#"{"latitude": 91, "longitude": 0}"#, &swapped),
            // Not when the pair makes no sense the other way round either.
            (r#"{"latitude": 181, "longitude": 47.61}"#, bad_lat),
            (r#"{"latitude": 95, "longitude": 120}"#, bad_lat),
            // A longitude written 0 to 360 (v0.5.1 happened to handle it),
            // or a digit too many: the hint is a question.
            (r#"{"latitude": 47.61, "longitude": 237.67}"#, &from_360),
            (r#"{"latitude": 47.61, "longitude": 360}"#, &from_360),
            (r#"{"latitude": 47.61, "longitude": 360.5}"#, bad_lon),
            // 0.0 for the other value doesn't make the pair "no location yet".
            (r#"{"latitude": 0.0, "longitude": -200}"#, bad_lon),
        ] {
            assert_eq!(parse_config(json).expect_err(json), want, "{json}");
        }
    }

    #[test]
    fn a_coordinate_error_never_quotes_the_coordinates() {
        // The message goes to events.log as well as the dialog. A slip must
        // not put the user's real coordinates there: here the longitude,
        // typed as the latitude, and the digits of a latitude typed 10x.
        for json in [
            r#"{"latitude": -122.33, "longitude": 47.61}"#,
            r#"{"latitude": 247.13, "longitude": 46.68}"#,
            r#"{"latitude": 24.71, "longitude": 466.8}"#,
            r#"{"latitude": 24.71}"#,
            r#"{"longitude": 46.68}"#,
        ] {
            let err = parse_config(json).expect_err(json);
            for digits in ["122", "47", "247", "13", "46", "68", "24", "71"] {
                assert!(!err.contains(digits), "{json}: {err}");
            }
        }
        // Nor does serde_json's own error for a coordinate typed in quotes
        // (it says `invalid type: string "24.713552", expected f64 at …`).
        assert_eq!(
            parse_config("{\n  \"latitude\": \"24.713552\",\n  \"longitude\": 46.675296\n}")
                .unwrap_err(),
            "config.json is not valid JSON: wrong kind of value, expected f64 at line 2 column 25"
        );
        // What was expected, and where, is still said for the other keys.
        assert_eq!(
            parse_config(r#"{"auto_start": "yes"}"#).unwrap_err(),
            "config.json is not valid JSON: wrong kind of value, expected a boolean \
             at line 1 column 20"
        );
        // Errors that quote no value are passed on as they are.
        let err = parse_config(r#"{"latitude": 1.0, "latitude": 2.0}"#).unwrap_err();
        assert!(err.contains("duplicate field `latitude`"), "{err}");
        let err = parse_config(r#"{"latitude": 24.7,}"#).unwrap_err();
        assert!(err.contains("at line 1 column 19"), "{err}");
    }

    #[test]
    fn coordinates_at_the_limits_load_and_have_a_schedule() {
        let now = utc(2026, 6, 21, 12, 0, 0);
        // Noon UTC at the June solstice: the north pole has the midnight sun,
        // the south pole the polar night, and at longitude 180 it's midnight.
        for (lat, lon, want) in [
            (90.0, 180.0, Theme::Light),
            (-90.0, -180.0, Theme::Dark),
            (0.0, 180.0, Theme::Dark),
            (0.0, -180.0, Theme::Dark),
            (90.0, 0.0, Theme::Light),
        ] {
            let json = format!(r#"{{"latitude": {lat:?}, "longitude": {lon:?}}}"#);
            let cfg = parse_config(&json).expect(&json);
            assert_eq!((cfg.latitude, cfg.longitude), (lat, lon));
            assert!(cfg.has_location());
            let (theme, next) = schedule(now, lat, lon);
            assert_eq!(theme, want, "{json}");
            assert!(next > now, "{json}");
            assert_ne!(schedule(next, lat, lon).0, theme, "{json}: a real flip");
        }
        // Whole numbers, as a hand edit may have them.
        let cfg = parse_config(r#"{"latitude": -90, "longitude": 180}"#).unwrap();
        assert_eq!((cfg.latitude, cfg.longitude), (-90.0, 180.0));
        assert!(valid_coordinates(-90.0, -180.0) && valid_coordinates(0.0, 0.0));
        for (lat, lon) in [
            (90.000001, 0.0),
            (0.0, -180.000001),
            (f64::NAN, 0.0),
            (0.0, f64::NAN),
            (f64::INFINITY, 0.0),
            (0.0, f64::NEG_INFINITY),
        ] {
            assert!(!valid_coordinates(lat, lon), "{lat} {lon}");
        }
    }

    #[test]
    fn one_coordinate_without_the_other_is_a_config_error() {
        // Left out, or misspelled (an unknown key is kept, not an error):
        // the missing one used to count as 0.0, and the schedule was
        // computed for that place.
        let missing_lon = "config.json has \"latitude\" but no \"longitude\" (check the \
                           spelling). Set both, or neither to use Windows Location.";
        let missing_lat = "config.json has \"longitude\" but no \"latitude\" (check the \
                           spelling). Set both, or neither to use Windows Location.";
        for (json, want) in [
            (r#"{"latitude": 24.7}"#, missing_lon),
            (r#"{"latitude": 24.7, "longitud": 46.7}"#, missing_lon),
            (r#"{"latitude": 24.7, "Longitude": 46.7}"#, missing_lon),
            (r#"{"longitude": 46.7, "auto_start": false}"#, missing_lat),
            (r#"{"lattitude": 24.7, "longitude": 46.7}"#, missing_lat),
            // Also when the one that's there is 0.0 or out of range.
            (r#"{"latitude": 0.0}"#, missing_lon),
            (r#"{"longitude": 999}"#, missing_lat),
        ] {
            assert_eq!(parse_config(json).expect_err(json), want, "{json}");
        }
        let json = r#"{"latitude": 24.7, "longitud": 46.7}"#;
        let t = TempConfig::new("one-key", Some(json));
        load_config_at(&t.0).expect_err("one key without the other");
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json, "file untouched");

        // Neither key, or both 0.0: no location yet, not an error.
        for json in [
            r#"{}"#,
            r#"{"auto_start": false}"#,
            r#"{"latitude": 0.0, "longitude": 0.0}"#,
            r#"{"latitude": 0, "longitude": -0.0}"#,
            // Both misspelled is "neither": Windows Location is asked.
            r#"{"lat": 24.7, "lon": 46.7}"#,
        ] {
            assert!(!parse_config(json).expect(json).has_location(), "{json}");
        }
        // One of them 0.0 with the other set is a real place.
        let cfg = parse_config(r#"{"latitude": 51.48, "longitude": 0.0}"#).unwrap();
        assert!(cfg.has_location());
    }

    #[test]
    fn a_coordinate_that_is_not_a_number_is_still_a_json_error() {
        for json in [
            r#"{"latitude": "24.7", "longitude": 46.7}"#,
            r#"{"latitude": null, "longitude": 46.7}"#,
            r#"{"latitude": 24.7, "longitude": [46.7]}"#,
            r#"{"latitude": 24.7, "longitude": 1e999}"#,
            r#"{"latitude": 24.7, "latitude": 24.8, "longitude": 46.7}"#,
            r#"[24.7, 46.7]"#,
        ] {
            let err = parse_config(json).expect_err(json);
            assert!(err.starts_with("config.json is not valid JSON: "), "{err}");
        }
    }

    #[test]
    fn refresh_keeps_the_last_good_config_when_a_coordinate_goes_out_of_range() {
        let t = TempConfig::new(
            "range-refresh",
            Some(r#"{"latitude": 24.71, "longitude": 46.68, "auto_start": false}"#),
        );
        let mut cfg = Config::default();
        let mut have_disk_cfg = false;
        let reload = || load_config_at(&t.0).map(read_from_file);
        let (_, error) = adopt_reloaded_config(reload(), &mut cfg, &mut have_disk_cfg);
        assert!(error.is_none() && have_disk_cfg);
        // The user edits the file and slips: Refresh reports it and keeps
        // switching by the coordinates it had.
        fs::write(&t.0, r#"{"latitude": 247.1, "longitude": 46.68}"#).unwrap();
        let (assert_autostart, error) =
            adopt_reloaded_config(reload(), &mut cfg, &mut have_disk_cfg);
        assert!(error.is_some_and(|e| e.contains("latitude must be between")));
        assert_eq!((cfg.latitude, cfg.longitude), (24.71, 46.68));
        assert!(assert_autostart && !cfg.auto_start);
    }

    #[test]
    fn a_windows_location_reading_is_rounded_before_it_is_kept() {
        for ((lat, lon), want) in [
            // 2 decimals below 60°, north or south.
            ((24.713_552, 46.675_296), (24.71, 46.68)),
            ((-33.868_82, 151.209_3), (-33.87, 151.21)),
            ((40.712_776, -74.005_974), (40.71, -74.01)),
            ((59.994_9, 10.752_245), (59.99, 10.75)),
            // 3 decimals from 60° on — decided by the reading, so a
            // latitude that rounds up to 60 has 2.
            ((59.999_9, 10.752_245), (60.0, 10.75)),
            ((60.0, 10.752_245), (60.0, 10.752)),
            ((69.649_205, 18.955_324), (69.649, 18.955)),
            ((-77.846_323, 166.668_235), (-77.846, 166.668)),
            // The limits stay in range.
            ((89.999_96, 179.999_96), (90.0, 180.0)),
            ((-89.999_96, -179.999_96), (-90.0, -180.0)),
        ] {
            assert_eq!(round_coordinates(lat, lon), want, "{lat} {lon}");
            assert_eq!(usable_location(lat, lon), Ok(want));
            // Rounding what was rounded changes nothing.
            assert_eq!(round_coordinates(want.0, want.1), want);
            assert!(valid_coordinates(want.0, want.1));
        }
        // Saved as typed digits, not as 24.710000000000001.
        let cfg = Config {
            latitude: round_coordinates(24.713_552, 46.675_296).0,
            longitude: round_coordinates(24.713_552, 46.675_296).1,
            ..Config::default()
        };
        let json = serde_json::to_string(&cfg).unwrap();
        assert!(
            json.contains(r#""latitude":24.71,"longitude":46.68,"#),
            "{json}"
        );
        // Just west of Greenwich: 0.0, never the "-0.0" that would be saved.
        let (lat, lon) = round_coordinates(51.507_4, -0.001_2);
        assert_eq!((lat, lon), (51.51, 0.0));
        assert!(lon.is_sign_positive());
        assert!(round_coordinates(-0.001_2, 36.82).0.is_sign_positive());
        // A reading next to 0/0 is not rounded into "no location yet".
        assert_eq!(round_coordinates(0.004, -0.003), (0.004, -0.003));
        assert_eq!(usable_location(0.004, -0.003), Ok((0.004, -0.003)));
        let (lat, lon) = round_coordinates(-0.0, 0.004);
        assert!(lat.is_sign_positive() && lon == 0.004);
        // Exactly 0/0 is no location: kept, it would read as "not set yet"
        // with nothing said.
        assert_eq!(usable_location(0.0, 0.0), Err("zero"));
        assert_eq!(usable_location(-0.0, 0.0), Err("zero"));
        // Not a place on Earth: no location, rather than a schedule for it.
        // Checked before rounding, which would bring the last two in range.
        for (lat, lon) in [
            (f64::NAN, 46.7),
            (24.7, f64::INFINITY),
            (90.5, 0.0),
            (0.0, 180.5),
            (90.000_4, 0.0),
            (10.0, 180.004),
        ] {
            assert_eq!(
                usable_location(lat, lon),
                Err("out-of-range"),
                "{lat} {lon}"
            );
        }
    }

    #[test]
    fn rounding_a_reading_moves_sunrise_and_sunset_by_seconds() {
        // What "2 decimals is plenty" rests on: readings as far from their
        // rounded values as rounding allows, through a year.
        for (lat, lon, off) in [
            (0.0, 36.82, 0.0049),
            (24.71, 46.68, 0.0049),
            (-33.87, 151.21, 0.0049),
            (45.0, -93.27, 0.0049),
            (59.99, 10.75, 0.0049),
            // 3 decimals: Reykjavik, and just below the polar model.
            (64.146, -21.942, 0.00049),
            (65.4, 25.47, 0.00049),
        ] {
            let (raw_lat, raw_lon) = (lat + off, lon - off);
            assert_eq!(round_coordinates(raw_lat, raw_lon), (lat, lon));
            for month in 1..=12 {
                for day in [1, 11, 21] {
                    let now = utc(2026, month, day, 12, 0, 0);
                    let raw = transitions_window(now, raw_lat, raw_lon);
                    let rounded = transitions_window(now, lat, lon);
                    assert_eq!(raw.len(), rounded.len(), "{lat} {lon} {now}");
                    for (a, b) in raw.iter().zip(&rounded) {
                        assert_eq!(a.1, b.1);
                        let moved = (a.0 - b.0).num_seconds().abs();
                        assert!(moved <= 10, "{lat} {lon} {now}: moved {moved} s");
                    }
                }
            }
        }
    }

    // --- the Windows Location lookup's cap, outcomes and dialogs (v0.6.0) ---

    #[test]
    fn work_that_finishes_in_time_hands_its_value_back_from_a_thread_of_its_own() {
        let caller = std::thread::current().id();
        let got = run_with_timeout("wts-test-fast", Duration::from_secs(30), move || {
            let me = std::thread::current();
            (me.id() != caller, me.name().map(str::to_string), 42)
        });
        assert_eq!(got, Ok((true, Some("wts-test-fast".to_string()), 42)));
        // A limit of zero still isn't a panic or a hang, whichever way the
        // race goes.
        let quick = run_with_timeout("wts-test-zero", Duration::ZERO, || 7);
        assert!(
            matches!(quick, Ok(7) | Err(NoAnswer::TimedOut)),
            "{quick:?}"
        );
    }

    #[test]
    fn a_timeout_frees_the_caller_and_the_late_value_is_dropped_without_a_panic() {
        use std::sync::Arc;
        /// The worker's late value. Says when it has been dropped, and
        /// whether that happened while its thread was unwinding from a
        /// panic — a worker that panicked over the closed channel
        /// (`tx.send(..).unwrap()`) drops the value too, on its way out.
        struct Late {
            dropped: Arc<AtomicBool>,
            dropped_in_panic: Arc<AtomicBool>,
        }
        impl Drop for Late {
            fn drop(&mut self) {
                // Before `dropped`, which is what the test waits for.
                self.dropped_in_panic
                    .store(std::thread::panicking(), Ordering::SeqCst);
                self.dropped.store(true, Ordering::SeqCst);
            }
        }
        let dropped = Arc::new(AtomicBool::new(false));
        let dropped_in_panic = Arc::new(AtomicBool::new(false));
        let finished = Arc::new(AtomicBool::new(false));
        // The work waits for the test's go-ahead, so "slow" doesn't depend
        // on a sleep being long enough.
        let (go, wait_for_go) = mpsc::channel::<()>();
        let (flag, panic_flag, done) =
            (dropped.clone(), dropped_in_panic.clone(), finished.clone());
        let limit = Duration::from_millis(100);
        let started = Instant::now();
        let got = run_with_timeout("wts-test-slow", limit, move || {
            let _ = wait_for_go.recv();
            done.store(true, Ordering::SeqCst);
            Late {
                dropped: flag,
                dropped_in_panic: panic_flag,
            }
        });
        let waited = started.elapsed();
        assert!(matches!(got, Err(NoAnswer::TimedOut)));
        // The caller waited for the limit and no longer: the work is
        // still blocked, and nothing has been produced or dropped yet.
        assert!(waited >= limit, "{waited:?}");
        assert!(waited < Duration::from_secs(20), "{waited:?}");
        assert!(!finished.load(Ordering::SeqCst) && !dropped.load(Ordering::SeqCst));
        // Now the work finishes, late. Its value has nowhere to go and is
        // dropped by the worker: no leak (the drop happens), and no panic
        // over the closed channel (the drop happens in a thread that isn't
        // unwinding — in the release build, which aborts on a panic, a
        // worker that panicked there would end the app whenever Windows
        // answered after the cap).
        go.send(()).unwrap();
        let deadline = Instant::now() + Duration::from_secs(20);
        while !dropped.load(Ordering::SeqCst) {
            assert!(
                Instant::now() < deadline,
                "the late value was never dropped"
            );
            std::thread::sleep(Duration::from_millis(5));
        }
        assert!(finished.load(Ordering::SeqCst));
        assert!(
            !dropped_in_panic.load(Ordering::SeqCst),
            "the worker panicked over the value nobody waited for"
        );
        // And the next call is not disturbed by the one that timed out.
        assert_eq!(
            run_with_timeout("wts-test-next", Duration::from_secs(30), || 1),
            Ok(1)
        );
    }

    #[test]
    fn a_worker_that_ends_without_a_value_is_reported_not_waited_for() {
        // The test build unwinds (the release build aborts on a panic, and
        // its panic hook logs it). resume_unwind skips the panic message.
        let started = Instant::now();
        let got: Result<u8, NoAnswer> =
            run_with_timeout("wts-test-lost", Duration::from_secs(30), || {
                std::panic::resume_unwind(Box::new("no value"))
            });
        assert_eq!(got, Err(NoAnswer::WorkerLost));
        assert!(started.elapsed() < Duration::from_secs(20));
    }

    /// One of each kind of failed lookup, with a realistic error code.
    fn location_failures() -> [LocationFailure; 7] {
        [
            LocationFailure::Denied,
            LocationFailure::Unspecified,
            LocationFailure::TimedOut,
            LocationFailure::Failed {
                stage: "position",
                hresult: 0x8007_0005_u32 as i32,
            },
            LocationFailure::Failed {
                stage: "worker",
                hresult: 0,
            },
            LocationFailure::Rejected("zero"),
            LocationFailure::Rejected("out-of-range"),
        ]
    }

    #[test]
    fn a_failed_lookup_is_logged_by_reason_and_time_never_by_coordinates() {
        let lines: Vec<String> = location_failures()
            .into_iter()
            .map(|f| location_failure_line(f, 1234))
            .collect();
        assert_eq!(
            lines,
            [
                "location_lookup_failed reason=access-denied after_ms=1234",
                "location_lookup_failed reason=access-unspecified after_ms=1234",
                "location_lookup_failed reason=timeout after_ms=1234",
                "location_lookup_failed reason=error after_ms=1234 stage=position hresult=0x80070005",
                "location_lookup_failed reason=error after_ms=1234 stage=worker hresult=0x00000000",
                "location_lookup_failed reason=reading-rejected after_ms=1234 reading=zero",
                "location_lookup_failed reason=reading-rejected after_ms=1234 reading=out-of-range",
            ]
        );
        // Nothing to quote or escape, and the line survives the log's
        // sanitizer unchanged.
        for line in &lines {
            assert_eq!(&sanitize_log_msg(line), line);
        }
        // The five reasons are told apart.
        let mut tokens: Vec<_> = location_failures().iter().map(|f| f.token()).collect();
        tokens.dedup();
        assert_eq!(
            tokens,
            [
                "access-denied",
                "access-unspecified",
                "timeout",
                "error",
                "reading-rejected"
            ]
        );
    }

    #[test]
    fn only_allowed_lets_the_lookup_go_on() {
        assert_eq!(access_answer(GeolocationAccessStatus::Allowed), Ok(()));
        assert_eq!(
            access_answer(GeolocationAccessStatus::Denied),
            Err(LocationFailure::Denied)
        );
        assert_eq!(
            access_answer(GeolocationAccessStatus::Unspecified),
            Err(LocationFailure::Unspecified)
        );
        // A value Windows doesn't have today is not taken for "allowed" —
        // nor for a denial, which would put "off or not allowed" on screen.
        assert_eq!(
            access_answer(GeolocationAccessStatus(7)),
            Err(LocationFailure::Unspecified)
        );
    }

    #[test]
    fn coordinates_leave_a_lookup_only_through_usable_location() {
        // A reading comes out exactly as usable_location returns it:
        // rounded, or refused with its reason.
        for (lat, lon) in [
            (24.713_552, 46.675_296),
            (69.649_205, 18.955_324),
            (-33.868_82, 151.209_3),
            (0.004, -0.003),
            (0.0, 0.0),
            (90.000_4, 0.0),
            (10.0, 180.004),
            (f64::NAN, 46.7),
        ] {
            assert_eq!(
                settle_location(Ok(Ok((lat, lon)))),
                usable_location(lat, lon).map_err(LocationFailure::Rejected),
                "{lat} {lon}"
            );
        }
        assert_eq!(
            settle_location(Ok(Ok((24.713_552, 46.675_296)))),
            Ok((24.71, 46.68))
        );
        assert_eq!(
            settle_location(Ok(Ok((0.0, 0.0)))),
            Err(LocationFailure::Rejected("zero"))
        );
        assert_eq!(
            settle_location(Ok(Ok((91.0, 0.0)))),
            Err(LocationFailure::Rejected("out-of-range"))
        );
        // What the lookup thread reports is passed on as it is.
        for failure in location_failures() {
            assert_eq!(settle_location(Ok(Err(failure))), Err(failure));
        }
        // No answer in time, or no answer at all.
        assert_eq!(
            settle_location(Err(NoAnswer::TimedOut)),
            Err(LocationFailure::TimedOut)
        );
        assert_eq!(
            settle_location(Err(NoAnswer::WorkerLost)),
            Err(LocationFailure::Failed {
                stage: "worker",
                hresult: 0
            })
        );
    }

    /// What every location box must keep to, whatever it is about: plain
    /// words, and no instruction to click an icon that isn't there yet.
    fn assert_plain_and_true_about_the_tray(text: &str, tray_exists: bool) {
        for jargon in [
            "WinRT",
            "HRESULT",
            "hresult",
            "API",
            "0x",
            "Geolocator",
            "stage",
            "ShellExecute",
            "ms-settings",
            "STA",
        ] {
            assert!(!text.contains(jargon), "{jargon:?} in: {text}");
        }
        // Never the old claims.
        for old in ["just opened", "opening now", "tray icon"] {
            assert!(!text.contains(old), "{old:?} in: {text}");
        }
        let lower = text.to_lowercase();
        let Some(click) = lower.find("right-click") else {
            // A box that doesn't send the user to the icon doesn't need to
            // say where it is.
            assert!(!lower.contains("icon"), "{text}");
            return;
        };
        if tray_exists {
            assert!(text.contains("icon is in the notification area"), "{text}");
            assert!(!text.contains("icon appears"), "{text}");
        } else {
            // Before the tray exists the text first says when and where
            // the icon turns up, and only then what to click.
            let appears = text
                .find("After you close this message, the WinThemeSwitcher icon appears")
                .unwrap_or_else(|| panic!("no word on when the icon appears: {text}"));
            assert!(appears < click, "{text}");
            assert!(!text.contains("icon is in"), "{text}");
        }
        assert!(text.contains("^ arrow"), "{text}");
        assert!(text.contains("Refresh"), "{text}");
    }

    #[test]
    fn the_first_location_box_says_what_actually_happened() {
        // "About 15 s", and the number on screen is the cap's.
        assert!((10..=20).contains(&LOCATION_LOOKUP_CAP.as_secs()));
        let within = format!("within {} seconds", LOCATION_LOOKUP_CAP.as_secs());
        let cases = [(false, true), (true, true), (false, false), (true, false)];
        for (tray_exists, config_on_disk) in cases {
            for failure in location_failures() {
                let prompt = location_prompt(failure, tray_exists, config_on_disk);
                let (text, offers_settings) = match &prompt {
                    LocationPrompt::OfferSettings(t) => (t.as_str(), true),
                    LocationPrompt::Inform(t) => (t.as_str(), false),
                };
                assert_plain_and_true_about_the_tray(text, tray_exists);
                let denial = matches!(
                    failure,
                    LocationFailure::Denied | LocationFailure::Unspecified
                );
                // Settings is offered only after Windows said no.
                assert_eq!(offers_settings, denial, "{failure:?}");
                // "Off or not allowed" as a fact: only for a real denial.
                assert_eq!(
                    text.contains("not allowed"),
                    failure == LocationFailure::Denied,
                    "{failure:?}: {text}"
                );
                assert_eq!(
                    text.contains("is turned off"),
                    failure == LocationFailure::Denied,
                    "{failure:?}: {text}"
                );
                // Each other reason says its own thing, and only that.
                assert_eq!(
                    text.contains(&within) && text.contains("did not answer"),
                    failure == LocationFailure::TimedOut,
                    "{failure:?}: {text}"
                );
                assert_eq!(
                    text.contains("the lookup failed"),
                    matches!(failure, LocationFailure::Failed { .. }),
                    "{failure:?}: {text}"
                );
                assert_eq!(
                    text.contains("a location for this PC that WinThemeSwitcher cannot use"),
                    matches!(failure, LocationFailure::Rejected(_)),
                    "{failure:?}: {text}"
                );
                assert_eq!(
                    text.contains("did not say why"),
                    failure == LocationFailure::Unspecified,
                    "{failure:?}: {text}"
                );
                if offers_settings {
                    // A question with both ways out named; it promises
                    // nothing about Settings having opened, and leaves the
                    // icon to the box that follows.
                    assert!(text.contains("Yes: open the Location settings."));
                    assert_eq!(
                        text.contains("No: type your coordinates into config.json"),
                        config_on_disk,
                        "{text}"
                    );
                    assert_eq!(
                        text.contains("No: go on without a location for now."),
                        !config_on_disk,
                        "{text}"
                    );
                    assert!(!text.contains("should now be open"));
                    assert!(!text.contains("Refresh"));
                } else {
                    // What to do next: Refresh later, or config.json.
                    assert!(text.contains("Try again later."), "{text}");
                    assert!(text.contains("choose Refresh"), "{text}");
                    assert!(
                        text.contains("latitude and longitude into config.json"),
                        "{text}"
                    );
                    // A Refresh shows nothing about its lookup, so the box
                    // says how the user can tell.
                    assert!(text.contains(REFRESH_IS_SILENT), "{text}");
                    // Settings is pointed at, not said to be the cause —
                    // and not at all for a reading no switch would change.
                    assert_eq!(
                        text.contains(LOCATION_SETTINGS_PAGE),
                        !matches!(failure, LocationFailure::Rejected(_)),
                        "{failure:?}: {text}"
                    );
                    // This is the only box of its path, so it has the
                    // format the manual-setup box gives.
                    assert_eq!(
                        text.contains("in decimal degrees (south and west are negative)"),
                        config_on_disk,
                        "{text}"
                    );
                    // The timeout box says what to do about Windows' own
                    // question, in order: the late answer is dropped, so
                    // answering it is not enough.
                    assert_eq!(
                        text.contains(
                            "answer it, then close this message and choose Refresh (step 1"
                        ),
                        failure == LocationFailure::TimedOut,
                        "{failure:?}: {text}"
                    );
                }
                // Nothing sends the user to a config.json that isn't there:
                // not to Open Config, not to saving it.
                assert_eq!(
                    text.contains("Open Config"),
                    config_on_disk && !offers_settings
                );
                assert_eq!(
                    text.contains("save the file"),
                    config_on_disk && !offers_settings
                );
                assert_eq!(
                    text.contains("config.json could not be created"),
                    !config_on_disk && !offers_settings,
                    "{text}"
                );
                // The details that belong in the log stay there.
                assert!(!text.contains("position") && !text.contains("out-of-range"));
                assert!(!text.contains("80070005"));
            }
        }
        // The two kinds of rejected reading, and any two errors, read the
        // same on screen.
        for config_on_disk in [true, false] {
            assert_eq!(
                location_prompt(LocationFailure::Rejected("zero"), false, config_on_disk),
                location_prompt(
                    LocationFailure::Rejected("out-of-range"),
                    false,
                    config_on_disk
                )
            );
            assert_eq!(
                location_prompt(
                    LocationFailure::Failed {
                        stage: "position",
                        hresult: 0x8007_0005_u32 as i32,
                    },
                    false,
                    config_on_disk
                ),
                location_prompt(
                    LocationFailure::Failed {
                        stage: "worker",
                        hresult: 0,
                    },
                    false,
                    config_on_disk
                )
            );
        }
        // "How to tell" is something the user can see, and claims no more
        // than "most likely".
        assert!(REFRESH_IS_SILENT.contains("next sunrise or sunset"));
        assert!(REFRESH_IS_SILENT.contains("most likely"));
    }

    #[test]
    fn only_a_shell_result_above_32_counts_as_launched() {
        // ShellExecuteW: "greater than 32" is success. The error codes it
        // documents (0, 2 file not found, 5 access denied, 31 no
        // association, 32 DLL not found) and a negative value are not.
        for code in [isize::MIN, -1, 0, 2, 5, 31, 32] {
            assert!(!launch_accepted(code), "{code}");
        }
        for code in [33, 42, isize::MAX] {
            assert!(launch_accepted(code), "{code}");
        }
    }

    #[test]
    fn the_box_after_opening_settings_is_true_whether_settings_opened_or_not() {
        for tray_exists in [false, true] {
            for opened in [false, true] {
                let text = location_settings_followup(opened, tray_exists);
                assert_plain_and_true_about_the_tray(&text, tray_exists);
                // That Settings opened is said only when the launch was
                // accepted — and even then with the way to get there by
                // hand, which is all the other case has.
                assert_eq!(text.contains("should now be open"), opened, "{text}");
                assert_eq!(text.contains("could not be opened"), !opened, "{text}");
                assert!(text.contains("open Settings yourself") || !opened, "{text}");
                assert!(text.contains("Open Settings yourself") || opened, "{text}");
                assert!(text.contains(LOCATION_SETTINGS_PAGE), "{text}");
                assert_eq!(LOCATION_SETTINGS_PAGE, "Privacy & security > Location");
                // What must be on, in words that hold if a switch is
                // missing or renamed; the names are a guide.
                assert!(text.contains("desktop apps included"), "{text}");
                assert!(text.contains("or something close to that"), "{text}");
                assert!(text.contains("may not show all of them"), "{text}");
                for switch in [
                    "\"Location services\"",
                    "\"Let apps access your location\"",
                    "\"Let desktop apps access your location\"",
                ] {
                    assert!(text.contains(switch), "{switch}: {text}");
                }
                assert!(text.contains("choose Refresh"), "{text}");
                assert!(text.ends_with(REFRESH_IS_SILENT), "{text}");
            }
        }
    }

    #[test]
    fn the_manual_setup_boxes_match_what_opening_config_json_does() {
        for tray_exists in [false, true] {
            let text = manual_setup_text(tray_exists);
            assert_plain_and_true_about_the_tray(&text, tray_exists);
            assert!(text.contains("latitude and longitude into config.json"));
            // open_config_in_editor: the .json program, or else Notepad —
            // after this box is closed, and it may fail.
            assert!(text.contains("When you close this message"), "{text}");
            assert!(text.contains("program set up for .json files"), "{text}");
            assert!(text.contains("or else in Notepad"), "{text}");
            assert!(
                text.contains("If neither can be started, another message says where"),
                "{text}"
            );
            assert!(!text.contains("is open") && !text.contains("has opened"));
            assert!(text.contains("After saving the file"), "{text}");
        }
        // The message that follows a failed open names the file.
        let path = Path::new(r"C:\Some Folder\WinThemeSwitcher\config.json");
        let text = config_not_opened_text(path);
        assert!(text.contains("could not be opened"), "{text}");
        assert!(text.ends_with(r"C:\Some Folder\WinThemeSwitcher\config.json"));
        assert_plain_and_true_about_the_tray(&text, false);
        // Without a config.json (it could not be created) the box that
        // takes the manual one's place promises no editor and no saving:
        // it says why not, and what to do first.
        let text = no_config_file_text();
        assert_plain_and_true_about_the_tray(&text, false);
        assert!(text.contains("config.json could not be created"), "{text}");
        assert!(text.contains("to a folder you can write to"), "{text}");
        for promise in ["opens config.json", "Notepad", "save", "Open Config"] {
            assert!(!text.contains(promise), "{promise:?} in: {text}");
        }
    }

    #[test]
    fn the_real_windows_location_lookup_ends_within_the_cap_whatever_windows_says() {
        // Through the same lookup thread, apartment and cap as the app. It
        // must pass with Location on, off, or absent (GitHub's runners),
        // so it asserts what holds for every outcome — and prints the
        // outcome, never a coordinate (--nocapture shows which one a
        // machine got). On a PC with Location off, Windows may show its
        // own "allow location?" prompt for the test binary.
        use windows::Win32::System::Com::CoUninitialize;
        use windows_sys::Win32::System::Com::{
            CoGetApartmentType, APTTYPE_CURRENT, APTTYPE_MAINSTA, APTTYPE_STA,
        };
        // The caller's side of the app's layout too: the app's main thread
        // is a single-threaded apartment — the first in the process, which
        // makes it the one COM sends main-threaded components to — and it
        // waits for the lookup without pumping messages. So this thread
        // becomes one, the way run() does it; if anything in Windows'
        // location code needed that apartment while called from the lookup
        // thread, the outcome here would be a timeout, as in the app. No
        // other test initializes COM, so this is normally the process's
        // first ("main") one; the printed line says which it was.
        ensure_com_initialized();
        let (mut apartment, mut qualifier) = (APTTYPE_CURRENT, 0);
        let hr = unsafe { CoGetApartmentType(&mut apartment, &mut qualifier) };
        assert!(
            hr >= 0 && (apartment == APTTYPE_MAINSTA || apartment == APTTYPE_STA),
            "the test thread is not a single-threaded apartment: hr={hr:#x} type={apartment}"
        );
        /// Leaves the apartment when the test is over, pass or fail.
        struct LeaveApartment;
        impl Drop for LeaveApartment {
            fn drop(&mut self) {
                unsafe { CoUninitialize() };
            }
        }
        let _leave = LeaveApartment;
        let caller = if apartment == APTTYPE_MAINSTA {
            "main-sta"
        } else {
            "sta"
        };
        let cap_ms = LOCATION_LOOKUP_CAP.as_millis();
        // Twice: the first lookup thread of a process creates the
        // apartment (keep_mta_alive), a later one — a Refresh after the
        // startup lookup — finds it there. Not a second time after a slow
        // first one, to keep the test short where Windows doesn't answer.
        for nth in 1..=2 {
            let (result, elapsed_ms) = lookup_windows_location(LOCATION_LOOKUP_CAP);
            assert!(elapsed_ms <= cap_ms + 3000, "took {elapsed_ms} ms");
            match result {
                Ok((lat, lon)) => {
                    println!(
                        "windows-location lookup {nth}: caller={caller} ok after_ms={elapsed_ms}"
                    );
                    // Messages without the values: a failure here must not
                    // print where the machine is.
                    assert!(valid_coordinates(lat, lon), "not a place on Earth");
                    assert!(!(lat == 0.0 && lon == 0.0), "0/0 came through");
                    assert!(
                        round_coordinates(lat, lon) == (lat, lon),
                        "the reading came through unrounded"
                    );
                }
                Err(failure) => {
                    println!(
                        "windows-location lookup {nth}: caller={caller} {}",
                        location_failure_line(failure, elapsed_ms)
                    );
                    // A timeout is reported at the cap, not before it.
                    if failure == LocationFailure::TimedOut {
                        assert!(elapsed_ms >= cap_ms, "timeout after {elapsed_ms} ms");
                    }
                    // The lookup thread started and reported.
                    assert_ne!(
                        failure,
                        LocationFailure::Failed {
                            stage: "worker",
                            hresult: 0
                        }
                    );
                }
            }
            if elapsed_ms > 3000 {
                break;
            }
        }
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

    /// "Autostart is registered", for loads where a heal may ask. Tests
    /// never read the real Run key.
    fn registered() -> Option<bool> {
        Some(true)
    }

    /// The heal a load reported.
    fn healed(note: Option<LoadNote>) -> Healed {
        match note {
            Some(LoadNote::Healed(h)) => h,
            other => panic!("expected a heal, got {other:?}"),
        }
    }

    #[test]
    fn load_config_ex_reports_heals_and_their_save_failures() {
        let valid = TempConfig::new(
            "heal-valid",
            Some("{\"latitude\": 1.0, \"longitude\": 2.0}"),
        );
        assert!(load_config_ex(&valid.0, registered).unwrap().1.is_none());
        let missing = TempConfig::new("heal-missing", None);
        assert!(
            load_config_ex(&missing.0, registered).unwrap().1.is_none(),
            "first run is not a heal"
        );
        let empty = TempConfig::new("heal-empty", Some("\0\0 \n"));
        let h = healed(load_config_ex(&empty.0, registered).unwrap().1);
        assert_eq!(h.reason, "empty");
        assert!(h.save_err.is_none());
        // Healing a file that can't be written: reported, file untouched.
        let ro = TempConfig::new("heal-readonly", Some("  "));
        let mut perms = fs::metadata(&ro.0).unwrap().permissions();
        perms.set_readonly(true);
        fs::set_permissions(&ro.0, perms.clone()).unwrap();
        let result = load_config_ex(&ro.0, registered);
        #[allow(clippy::permissions_set_readonly_false)]
        perms.set_readonly(false);
        fs::set_permissions(&ro.0, perms).unwrap();
        let h = healed(result.unwrap().1);
        assert!(
            h.save_err.is_some(),
            "the failed write-back must be reported"
        );
        assert_eq!(fs::read_to_string(&ro.0).unwrap(), "  ");
    }

    // --- config.json changing while it's loaded (v0.5.1) ---

    /// A reader whose first read returns `first` and later reads read the
    /// file: the file changed between load_config_via's read and its write.
    fn racing_read(first: std::io::Result<String>) -> impl FnMut(&Path) -> std::io::Result<String> {
        let mut first = Some(first);
        move |p| first.take().unwrap_or_else(|| read_config_text(p))
    }

    #[test]
    fn first_run_never_answers_a_file_that_just_appeared_with_defaults() {
        // Missing when read, but an editor saved it before the defaults were
        // written: create_new refuses, and the user's file must be read and
        // returned — not defaults that a later persist_config would write
        // over it.
        let json = r#"{"latitude": 24.753, "longitude": 46.765, "auto_start": false}"#;
        let t = TempConfig::new("race-create", Some(json));
        let (cfg, healed) = load_config_via(
            &t.0,
            racing_read(Err(std::io::ErrorKind::NotFound.into())),
            registered,
        )
        .expect("the user's file must load");
        assert_eq!(cfg.latitude, 24.753);
        assert!(!cfg.auto_start && healed.is_none());
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json, "file untouched");
    }

    #[test]
    fn the_empty_file_heal_never_overwrites_a_save_that_landed_first() {
        // Empty when read, but the user saved real content before the
        // defaults were written: the heal must see it and stand down.
        let json = r#"{"latitude": 24.753, "longitude": 46.765}"#;
        let t = TempConfig::new("race-heal", Some(json));
        let (cfg, healed) = load_config_via(&t.0, racing_read(Ok(String::new())), registered)
            .expect("the user's file must load");
        assert_eq!(cfg.longitude, 46.765);
        assert!(healed.is_none(), "not a heal");
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json, "file untouched");
    }

    #[test]
    fn a_config_deleted_before_its_heal_is_recreated_like_first_run() {
        // The heal was started (the registry was asked, and says autostart
        // isn't registered), but what gets written is a first run's file:
        // the default auto_start, not the repair's.
        let t = TempConfig::new("race-heal-gone", None);
        let (cfg, healed) =
            load_config_via(&t.0, racing_read(Ok("\0\0".to_string())), || Some(false)).unwrap();
        assert!(
            !cfg.has_location() && healed.is_none(),
            "first run, not a heal"
        );
        assert!(cfg.auto_start, "a first run keeps the default");
        let written: Config = serde_json::from_str(&fs::read_to_string(&t.0).unwrap()).unwrap();
        assert!(!written.has_location() && written.auto_start);
    }

    #[test]
    fn a_config_that_keeps_changing_is_an_error_and_left_alone() {
        let json = r#"{"latitude": 1.0}"#;
        let t = TempConfig::new("race-forever", Some(json));
        let err = load_config_via(
            &t.0,
            |_| Err(std::io::ErrorKind::NotFound.into()),
            registered,
        )
        .expect_err("never defaults for a file that exists");
        assert!(err.contains("kept changing"), "{err}");
        assert_eq!(fs::read_to_string(&t.0).unwrap(), json);
    }

    #[test]
    fn an_empty_config_another_program_is_writing_is_unreadable_not_reset() {
        // An editor holds the file open to write it: what it will hold is
        // unknown. Defaults returned as if read would be persisted over its
        // save later, so the load fails (the session then runs read-only)
        // and nothing is written.
        let t = TempConfig::new("heal-busy", Some(""));
        let writer = fs::OpenOptions::new().write(true).open(&t.0).unwrap();
        let result = load_config_ex(&t.0, registered);
        drop(writer);
        let err = result.expect_err("a file being written is not an empty one");
        assert!(err.contains("could not be read"), "{err}");
        assert_eq!(fs::read(&t.0).unwrap(), b"");
    }

    #[test]
    fn a_dangling_config_link_is_an_error_that_names_no_race() {
        // Missing when read, but create_new refuses the link itself: the
        // load must fail without writing through the link, and must not
        // blame a file that "kept changing". (A dangling junction already
        // fails at the read, with access denied.)
        let t = TempConfig::new("dangling-link", None);
        let target = t.0.with_extension("missing-target");
        if let Err(e) = std::os::windows::fs::symlink_file(&target, &t.0) {
            eprintln!("SKIP: can't create a symlink here ({e}) - dangling-link path not exercised");
            return;
        }
        let err = load_config_ex(&t.0, registered).expect_err("nothing to read or create");
        assert!(!err.contains("kept changing"), "{err}");
        assert!(!target.exists(), "the link's target must not be created");
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
        let (assert_auto, err) =
            adopt_reloaded_config(Ok(read_from_file(good(false))), &mut cfg, &mut have);
        assert!(assert_auto && err.is_none() && have);
        assert!(!cfg.auto_start);
        // (c) Good earlier, broken now: keep last-known-good and re-assert
        // from it (the documented AV-quarantine recovery).
        let (assert_auto, err) = adopt_reloaded_config(Err("bad".into()), &mut cfg, &mut have);
        assert!(assert_auto && err.is_some());
        assert!(!cfg.auto_start, "last-known-good retained");
    }

    /// A config that was read from its file, as load_config_logged hands it
    /// on.
    fn read_from_file(cfg: Config) -> Loaded {
        Loaded {
            cfg,
            from_file: true,
        }
    }

    #[test]
    fn refresh_never_adopts_the_defaults_of_a_config_that_could_not_be_created() {
        // config.json missing and not creatable at a Refresh — the app's
        // drive or share can't be reached right now, say. Through v0.5.1
        // the defaults replaced the session's settings, and their
        // auto_start=true registered autostart over a file that said false.
        let not_created = || -> Result<Loaded, String> {
            Ok(Loaded::new(
                Config::default(),
                Some(&LoadNote::NotCreated(
                    "the network path was not found".into(),
                )),
            ))
        };
        let mut cfg = Config {
            latitude: 24.71,
            longitude: 46.68,
            auto_start: false,
            theme_night: Some("C:\\Themes\\night.theme".into()),
            ..Config::default()
        };
        let mut have = true;
        let (assert_auto, err) = adopt_reloaded_config(not_created(), &mut cfg, &mut have);
        assert!(err.is_none(), "not a config error: the file isn't broken");
        assert_eq!((cfg.latitude, cfg.longitude), (24.71, 46.68));
        assert_eq!(cfg.theme_night.as_deref(), Some("C:\\Themes\\night.theme"));
        // Autostart is re-asserted, but from the settings that were kept.
        assert!(assert_auto && have && !cfg.auto_start);
        // A session that started with a broken file has nothing from disk:
        // it keeps its fallback, and still must not touch autostart.
        let mut cfg = Config {
            latitude: 1.0,
            longitude: 2.0,
            ..Config::default()
        };
        let mut have = false;
        let (assert_auto, err) = adopt_reloaded_config(not_created(), &mut cfg, &mut have);
        assert!(!assert_auto && err.is_none() && !have);
        assert_eq!((cfg.latitude, cfg.longitude), (1.0, 2.0));
        // The file is back: adopted as always.
        let back = Config {
            auto_start: false,
            ..Config::default()
        };
        let (assert_auto, _) = adopt_reloaded_config(Ok(read_from_file(back)), &mut cfg, &mut have);
        assert!(assert_auto && have && !cfg.auto_start && !cfg.has_location());
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

    // --- autostart: an empty config's repair, a folder that can't be
    // written, a temporary folder, the registry, Startup settings (v0.6.0) ---

    #[test]
    fn repairing_an_empty_config_leaves_autostart_as_it_is_registered() {
        // The default auto_start=true used to go into the repaired file,
        // which re-registered autostart for someone who had turned it off.
        for (i, (registered, want)) in [(Some(true), true), (Some(false), false), (None, true)]
            .into_iter()
            .enumerate()
        {
            assert_eq!(healed_config(registered).auto_start, want);
            for (kind, content) in [("empty", ""), ("nul", "\0\0\0\0")] {
                let t = TempConfig::new(&format!("heal-auto-{i}-{kind}"), Some(content));
                let (cfg, note) = load_config_ex(&t.0, || registered).unwrap();
                assert_eq!(cfg.auto_start, want, "{registered:?} {kind}");
                assert!(!cfg.has_location() && cfg.theme_day.is_none());
                let h = healed(note);
                assert_eq!(h.auto_start, want);
                assert!(h.save_err.is_none());
                // The repaired file says the same, so later starts agree.
                let written: Config =
                    serde_json::from_str(&fs::read_to_string(&t.0).unwrap()).unwrap();
                assert_eq!(written.auto_start, want);
            }
        }
    }

    #[test]
    fn a_repair_that_cannot_be_written_follows_the_same_autostart_rule() {
        let t = TempConfig::new("heal-auto-readonly", Some(" "));
        let mut perms = fs::metadata(&t.0).unwrap().permissions();
        perms.set_readonly(true);
        fs::set_permissions(&t.0, perms.clone()).unwrap();
        let result = load_config_ex(&t.0, || Some(false));
        #[allow(clippy::permissions_set_readonly_false)]
        perms.set_readonly(false);
        fs::set_permissions(&t.0, perms).unwrap();
        let (cfg, note) = result.unwrap();
        assert!(
            !cfg.auto_start,
            "healed in memory, autostart not registered"
        );
        let h = healed(note);
        assert!(!h.auto_start && h.save_err.is_some());
        assert_eq!(fs::read_to_string(&t.0).unwrap(), " ", "file untouched");
    }

    #[test]
    fn only_a_repair_asks_whether_autostart_is_registered() {
        // First run keeps the default (on), and neither it nor an ordinary
        // load reads the registry.
        fn never() -> Option<bool> {
            panic!("the registry must not be read for this load")
        }
        let missing = TempConfig::new("auto-first-run", None);
        let (cfg, note) = load_config_ex(&missing.0, never).unwrap();
        assert!(cfg.auto_start && note.is_none());
        let written: Config =
            serde_json::from_str(&fs::read_to_string(&missing.0).unwrap()).unwrap();
        assert!(written.auto_start);
        let off = TempConfig::new("auto-valid-off", Some(r#"{"auto_start": false}"#));
        assert!(!load_config_ex(&off.0, never).unwrap().0.auto_start);
        let broken = TempConfig::new("auto-broken", Some("{"));
        assert!(load_config_ex(&broken.0, never).is_err());
    }

    #[test]
    fn a_config_that_cannot_be_created_is_reported_and_runs_on_defaults() {
        // Stands in for a folder that can't be written: one that isn't
        // there. The read finds nothing, and so does the create.
        let dir = std::env::temp_dir().join(format!("wts-test-{}-no-folder", std::process::id()));
        let path = dir.join("config.json");
        let (cfg, note) = load_config_ex(&path, registered).expect("the session carries on");
        assert!(cfg.auto_start && !cfg.has_location(), "the defaults");
        match &note {
            Some(LoadNote::NotCreated(err)) => assert!(!err.is_empty()),
            other => panic!("expected NotCreated, got {other:?}"),
        }
        assert!(!dir.exists() && !path.exists(), "nothing was created");
        // What the session is told: none of this came from a file …
        let loaded = Loaded::new(cfg, note.as_ref());
        assert!(!loaded.from_file);
        // … so when the folder is there after all, with the user's file in
        // it, a save (the location, once found) leaves that file alone.
        let folder = TempDir::new("no-folder");
        assert_eq!(folder.0, dir);
        let json = r#"{"latitude": 1.0, "longitude": 2.0, "auto_start": false}"#;
        fs::write(&path, json).unwrap();
        let located = Config {
            latitude: 24.71,
            longitude: 46.68,
            ..loaded.cfg
        };
        assert!(persist_config_at(&path, &located, !loaded.from_file).is_err());
        assert_eq!(fs::read_to_string(&path).unwrap(), json);
    }

    #[test]
    fn every_other_load_counts_as_read_from_the_file() {
        // An ordinary load, a first run that wrote its file, and a repair —
        // written or not: the file was read, and a later save replaces it.
        let valid = TempConfig::new("from-file-valid", Some(r#"{"auto_start": false}"#));
        let first = TempConfig::new("from-file-first-run", None);
        let empty = TempConfig::new("from-file-empty", Some(""));
        for t in [&valid, &first, &empty] {
            let (cfg, note) = load_config_ex(&t.0, registered).unwrap();
            assert!(Loaded::new(cfg, note.as_ref()).from_file, "{:?}", t.0);
        }
        let unwritten = LoadNote::Healed(Healed {
            reason: "empty",
            auto_start: false,
            save_err: Some("access denied".into()),
        });
        assert!(Loaded::new(healed_config(Some(false)), Some(&unwritten)).from_file);
    }

    #[test]
    fn a_config_that_did_not_come_from_a_file_never_replaces_one() {
        let located = Config {
            latitude: 24.71,
            longitude: 46.68,
            ..Config::default()
        };
        // config.json couldn't be created at the load, but is there now:
        // it is the user's, and stays.
        let json = r#"{"latitude": 1.0, "longitude": 2.0, "auto_start": false}"#;
        let appeared = TempConfig::new("persist-appeared", Some(json));
        assert!(persist_config_at(&appeared.0, &located, true).is_err());
        assert_eq!(fs::read_to_string(&appeared.0).unwrap(), json);
        // Still missing: created.
        let missing = TempConfig::new("persist-missing", None);
        persist_config_at(&missing.0, &located, true).expect("created");
        assert_eq!(load_config_at(&missing.0).unwrap().latitude, 24.71);
        // A config that was loaded from the file replaces it, as always.
        persist_config_at(&appeared.0, &located, false).expect("replaced");
        assert_eq!(load_config_at(&appeared.0).unwrap().latitude, 24.71);
    }

    #[test]
    fn auto_start_plan_table() {
        use AutoStartPlan::*;
        assert_eq!(auto_start_plan(true, false), Register);
        // From a temporary folder the Run value is left alone …
        assert_eq!(auto_start_plan(true, true), LeaveTempFolder);
        // … but auto_start: false removes it wherever the exe runs.
        assert_eq!(auto_start_plan(false, false), Remove);
        assert_eq!(auto_start_plan(false, true), Remove);
    }

    #[test]
    fn path_key_reads_every_spelling_of_a_folder_alike() {
        let want: Vec<String> = ["c:", "users", "ann", "appdata", "local", "temp"]
            .map(String::from)
            .to_vec();
        for spelling in [
            r"C:\Users\Ann\AppData\Local\Temp",
            r"c:\users\ann\appdata\local\temp\",
            r"C:\Users\Ann\AppData\Local\Temp\\",
            r"C:/Users/Ann/AppData/Local/Temp/",
            r"\\?\C:\Users\Ann\AppData\Local\Temp",
            r"\\.\C:\USERS\ANN\APPDATA\LOCAL\TEMP\",
            r"C:\Users\Ann\.\AppData\\Local\Temp",
            r"  C:\Users\Ann\AppData\Local\Temp ",
        ] {
            assert_eq!(path_key(spelling).as_ref(), Some(&want), "{spelling}");
        }
        // Non-ASCII names fold too.
        assert_eq!(path_key(r"D:\ÄÖ\Ünïcode"), path_key(r"d:\äö\ünïcode\"));
        // A share is never a drive path, in either of its spellings.
        let share = path_key(r"\\Server\Share\Temp").unwrap();
        assert_eq!(share, [r"\\", "server", "share", "temp"]);
        assert_eq!(path_key(r"\\?\UNC\server\share\temp\"), Some(share.clone()));
        assert_eq!(path_key(r"\\?\unc\SERVER\SHARE\Temp"), Some(share));
        // Not a place that can be compared: relative, drive-relative,
        // rooted without a drive, a volume GUID, a server without a share.
        for not_absolute in [
            "",
            "Temp",
            r".\Temp",
            r"C:Temp",
            "C:",
            r"\Temp",
            r"\\?\Volume{4c1b02c1-d990-11dc-99ae-806e6f6e6963}\Temp",
            r"\\server",
            r"%TEMP%\x",
        ] {
            assert_eq!(path_key(not_absolute), None, "{not_absolute}");
        }
    }

    #[test]
    fn an_exe_in_a_temp_folder_is_seen_through_8_3_names_case_and_prefixes() {
        let s = |list: &[&str]| list.iter().map(|p| p.to_string()).collect::<Vec<String>>();
        // What the app gathers when %TEMP% is stored in 8.3 form: the value
        // as it is, and its resolved long form.
        let temps = s(&[
            r"C:\Users\LONGUS~1\AppData\Local\Temp",
            r"\\?\C:\Users\Long User Name\AppData\Local\Temp",
        ]);
        // Started from inside the zip (Explorer's Temp1_<zip> folder), in
        // the long spelling, the short one, other case, with a prefix.
        for exe in [
            r"C:\Users\Long User Name\AppData\Local\Temp\Temp1_wts.zip\win-theme-switcher.exe",
            r"C:\Users\LONGUS~1\AppData\Local\Temp\Temp1_wts.zip\win-theme-switcher.exe",
            r"c:\users\longus~1\appdata\local\temp\7zO4A1B\WIN-THEME-SWITCHER.EXE",
            r"\\?\C:\Users\Long User Name\AppData\Local\Temp\win-theme-switcher.exe",
            r"C:/Users/Long User Name/AppData/Local/Temp/x/y/win-theme-switcher.exe",
        ] {
            assert!(in_temp_folder(&s(&[exe]), &temps), "{exe}");
        }
        // The exe in 8.3 form while %TEMP% is long: its own long form is
        // the second spelling that matches.
        let long_temp = s(&[r"C:\Users\Long User Name\AppData\Local\Temp\"]);
        let short_exe = r"C:\Users\LONGUS~1\AppData\Local\Temp\Rar$EXa0.123\win-theme-switcher.exe";
        assert!(!in_temp_folder(&s(&[short_exe]), &long_temp));
        assert!(in_temp_folder(
            &s(&[
                short_exe,
                r"\\?\C:\Users\Long User Name\AppData\Local\Temp\Rar$EXa0.123\win-theme-switcher.exe",
            ]),
            &long_temp
        ));
        // Not temporary: the recommended folder, a sibling whose name only
        // starts like the temp folder's, the temp folder's parent, another
        // drive, and something named like the temp folder itself.
        for exe in [
            r"C:\Users\Long User Name\AppData\Local\Programs\WinThemeSwitcher\win-theme-switcher.exe",
            r"C:\Users\Long User Name\AppData\Local\Temporary\win-theme-switcher.exe",
            r"C:\Users\Long User Name\AppData\Local\win-theme-switcher.exe",
            r"D:\Users\Long User Name\AppData\Local\Temp\win-theme-switcher.exe",
            r"C:\Users\Long User Name\AppData\Local\Temp",
        ] {
            assert!(!in_temp_folder(&s(&[exe]), &temps), "{exe}");
        }
        // A temp folder that is a whole drive or share, or not absolute, is
        // ignored — everything would be "temporary" otherwise.
        let exe = s(&[r"D:\Apps\WinThemeSwitcher\win-theme-switcher.exe"]);
        for temp in [r"D:\", "D:", r"\\?\D:\", "", "Temp", r"\Apps"] {
            assert!(!in_temp_folder(&exe, &s(&[temp])), "{temp}");
        }
        assert!(in_temp_folder(&exe, &s(&[r"D:\Apps"])));
        let on_share = s(&[r"\\nas\home\tmp\win-theme-switcher.exe"]);
        assert!(!in_temp_folder(&on_share, &s(&[r"\\nas\home"])));
        assert!(in_temp_folder(&on_share, &s(&[r"\\?\UNC\NAS\Home\Tmp\"])));
        // Nothing to compare with.
        assert!(!in_temp_folder(&exe, &[]));
        assert!(!in_temp_folder(&[], &temps));
    }

    #[test]
    fn temp_folders_are_what_the_environment_names_plus_the_usual_two() {
        fn env(
            pairs: &'static [(&'static str, &'static str)],
        ) -> impl Fn(&str) -> Option<std::ffi::OsString> {
            move |name| {
                pairs
                    .iter()
                    .find(|(key, _)| *key == name)
                    .map(|(_, value)| value.into())
            }
        }
        let paths = |list: &[&str]| list.iter().map(PathBuf::from).collect::<Vec<PathBuf>>();
        assert_eq!(
            temp_folders(env(&[
                ("TMP", r"D:\Scratch"),
                ("TEMP", r"C:\Users\ANN~1\AppData\Local\Temp"),
                ("LOCALAPPDATA", r"C:\Users\Ann\AppData\Local"),
                ("SystemRoot", r"C:\Windows"),
                ("USERPROFILE", r"C:\Users\Ann"),
            ])),
            paths(&[
                r"D:\Scratch",
                r"C:\Users\ANN~1\AppData\Local\Temp",
                r"C:\Users\Ann\AppData\Local\Temp",
                r"C:\Windows\Temp",
            ])
        );
        // Each source by itself.
        assert_eq!(temp_folders(env(&[("TMP", r"E:\t")])), paths(&[r"E:\t"]));
        assert_eq!(temp_folders(env(&[("TEMP", r"E:\t")])), paths(&[r"E:\t"]));
        assert_eq!(
            temp_folders(env(&[("LOCALAPPDATA", r"C:\Users\Ann\AppData\Local")])),
            paths(&[r"C:\Users\Ann\AppData\Local\Temp"])
        );
        assert_eq!(
            temp_folders(env(&[("SystemRoot", r"C:\Windows")])),
            paths(&[r"C:\Windows\Temp"])
        );
        // Unset or empty: nothing — never the profile folder (what
        // GetTempPathW falls back to), and never a relative "Temp".
        assert!(temp_folders(env(&[("USERPROFILE", r"C:\Users\Ann")])).is_empty());
        assert!(temp_folders(env(&[
            ("TMP", ""),
            ("TEMP", ""),
            ("LOCALAPPDATA", ""),
            ("SystemRoot", ""),
        ]))
        .is_empty());
    }

    #[test]
    fn a_file_in_the_real_temp_folder_is_seen_whatever_its_spelling() {
        // One temp folder at a time (the one the test files go to), so no
        // other entry of the app's list can stand in for a spelling that
        // wasn't resolved.
        let t = TempConfig::new("in-temp", Some("x"));
        let text = |p: &Path| p.to_string_lossy().into_owned();
        let temp = std::env::temp_dir();
        assert!(in_temp_folder(&[text(&t.0)], &[text(&temp)]));
        // The file's long spelling alone, with and without the prefix: what
        // current_exe() gives when %TEMP% itself is stored in 8.3 form, as
        // on GitHub's runners (C:\Users\RUNNER~1\AppData\Local\Temp). Only
        // the resolved spelling path_spellings adds to the folder matches
        // then.
        let long = text(&fs::canonicalize(&t.0).unwrap());
        assert!(in_temp_folder(
            std::slice::from_ref(&long),
            &path_spellings(&temp)
        ));
        let bare = long.trim_start_matches(r"\\?\").to_string();
        assert!(in_temp_folder(&[bare], &path_spellings(&temp)));
        // The other way round: the file as given against the folder's long
        // spelling alone.
        let long_temp = text(&fs::canonicalize(&temp).unwrap());
        assert!(in_temp_folder(&path_spellings(&t.0), &[long_temp]));
        // And the app's own list has the folder — where TMP or TEMP names
        // it (without them the test files go to the profile folder, which
        // the app rightly doesn't count).
        if std::env::var_os("TMP").is_some() || std::env::var_os("TEMP").is_some() {
            assert!(in_temp_folder(&[text(&t.0)], &temp_folder_spellings()));
        }
        // The folder README recommends is not a temporary one. (Compared as
        // text: the tests stay out of that folder, where the app may be
        // installed.)
        if let Some(local) = std::env::var_os("LOCALAPPDATA") {
            let exe =
                PathBuf::from(local).join("Programs\\WinThemeSwitcher\\win-theme-switcher.exe");
            assert!(!in_temp_folder(&[text(&exe)], &temp_folder_spellings()));
        }
    }

    #[test]
    fn an_8_3_spelling_on_disk_is_resolved_on_either_side() {
        // A real folder with a long name and, where the volume makes them,
        // an 8.3 alias: the two spellings share no text, and only
        // path_spellings' resolved form brings them together.
        use windows_sys::Win32::Storage::FileSystem::GetShortPathNameW;
        let d = TempDir::new("a long folder name for 8.3");
        let text = |p: &Path| p.to_string_lossy().into_owned();
        fs::write(d.0.join("win-theme-switcher.exe"), b"x").unwrap();
        let long_dir = fs::canonicalize(&d.0).unwrap();
        let given = wide(&text(&d.0));
        let mut buf = [0u16; 1024];
        let n = unsafe { GetShortPathNameW(given.as_ptr(), buf.as_mut_ptr(), buf.len() as u32) }
            as usize;
        let short_dir = PathBuf::from(String::from_utf16_lossy(&buf[..n.min(buf.len())]));
        if n == 0 || n >= buf.len() || path_key(&text(&short_dir)) == path_key(&text(&long_dir)) {
            eprintln!("SKIP: no 8.3 name for {}", d.0.display());
            return;
        }
        let short_exe = short_dir.join("win-theme-switcher.exe");
        let long_exe = long_dir.join("win-theme-switcher.exe");
        // The exe in 8.3 form, the temp folder in long form.
        assert!(!in_temp_folder(&[text(&short_exe)], &[text(&long_dir)]));
        assert!(in_temp_folder(
            &path_spellings(&short_exe),
            &[text(&long_dir)]
        ));
        // The temp folder in 8.3 form (as %TEMP% often is), the exe long.
        assert!(!in_temp_folder(&[text(&long_exe)], &[text(&short_dir)]));
        assert!(in_temp_folder(
            &[text(&long_exe)],
            &path_spellings(&short_dir)
        ));
    }

    #[test]
    fn startup_switch_reads_what_windows_records() {
        use StartupSwitch::*;
        // On, as Windows 11 stores it for an entry that was never switched
        // off (02) and for a machine-wide one (06).
        let on = [0x02, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
        assert_eq!(startup_switch(REG_BINARY, &on), On);
        let on6 = [0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
        assert_eq!(startup_switch(REG_BINARY, &on6), On);
        // Off: the lowest flag bit, followed by the time it was switched off.
        let off = [
            0x03, 0, 0, 0, 0x5C, 0x1B, 0x9A, 0x2E, 0x47, 0x33, 0xDC, 0x01,
        ];
        assert_eq!(startup_switch(REG_BINARY, &off), Off);
        let off7 = [
            0x07, 0, 0, 0, 0x10, 0x72, 0xE4, 0x09, 0xA1, 0x2F, 0xDC, 0x01,
        ];
        assert_eq!(startup_switch(REG_BINARY, &off7), Off);
        // Switched off and on again: the flag is clear whatever follows.
        let on_again = [
            0x02, 0, 0, 0, 0x5C, 0x1B, 0x9A, 0x2E, 0x47, 0x33, 0xDC, 0x01,
        ];
        assert_eq!(startup_switch(REG_BINARY, &on_again), On);
        // Not the record Windows writes: nothing is said.
        assert_eq!(startup_switch(REG_BINARY, &[]), Unknown);
        assert_eq!(startup_switch(REG_BINARY, &[0x03]), Unknown);
        assert_eq!(startup_switch(REG_BINARY, &[0x03, 0, 0, 0]), Unknown);
        assert_eq!(startup_switch(REG_BINARY, &off[..11]), Unknown);
        assert_eq!(startup_switch(REG_BINARY, &[0x03; 16]), Unknown);
        assert_eq!(startup_switch(REG_DWORD, &[0x03, 0, 0, 0]), Unknown);
        assert_eq!(startup_switch(REG_SZ, &off), Unknown);
        assert_eq!(startup_switch(REG_DWORD, &off), Unknown);
    }

    /// A throwaway key `HKCU\Software\WinThemeSwitcherTest-<pid>-<n>-<name>`
    /// for the registry tests, deleted with everything below it when the
    /// test ends — also when an assertion fails. The registry tests work
    /// under such a key only: never the real Run, StartupApproved or
    /// Personalize keys. `new` doesn't create it (a missing key is a case).
    struct TestKey(String);

    impl TestKey {
        const PREFIX: &'static str = "Software\\WinThemeSwitcherTest-";

        fn new(name: &str) -> Self {
            static NEXT: std::sync::atomic::AtomicU32 = std::sync::atomic::AtomicU32::new(0);
            let key = Self(format!(
                "{}{}-{}-{name}",
                Self::PREFIX,
                std::process::id(),
                NEXT.fetch_add(1, Ordering::Relaxed),
            ));
            key.delete();
            key
        }

        /// The path of a key below this one.
        fn sub(&self, name: &str) -> String {
            format!("{}\\{name}", self.0)
        }

        fn delete(&self) {
            use windows_sys::Win32::System::Registry::{RegDeleteKeyW, RegDeleteTreeW};
            // Whatever happens to self.0, nothing else is ever deleted.
            if !self.0.starts_with(Self::PREFIX) || self.0.len() <= Self::PREFIX.len() {
                return;
            }
            let path = wide(&self.0);
            unsafe {
                RegDeleteTreeW(HKEY_CURRENT_USER, path.as_ptr());
                RegDeleteKeyW(HKEY_CURRENT_USER, path.as_ptr());
            }
        }
    }

    impl Drop for TestKey {
        fn drop(&mut self) {
            self.delete();
        }
    }

    fn key_exists(path: &str) -> bool {
        open_hkcu_key(path, KEY_QUERY_VALUE, false).is_ok()
    }

    /// Write a value of any type, creating the key.
    fn set_raw_value(key: &str, name: &str, kind: u32, data: &[u8]) {
        let hkey = open_hkcu_key(key, KEY_SET_VALUE, true).expect("test key");
        let name_w = wide(name);
        let rc = unsafe {
            RegSetValueExW(
                hkey.0,
                name_w.as_ptr(),
                0,
                kind,
                data.as_ptr(),
                data.len() as u32,
            )
        };
        assert_eq!(rc, 0);
    }

    /// A REG_SZ value's text, as set_run_value reads it.
    fn read_sz(key: &str, name: &str) -> Option<String> {
        let hkey = open_hkcu_key(key, KEY_QUERY_VALUE, false).ok()?;
        unsafe { read_reg_sz(hkey.0, &wide(name)) }
    }

    const CMD: &str = "\"C:\\Program Files\\Some App\\win-theme-switcher.exe\"";

    #[test]
    fn the_test_key_guard_removes_its_key_and_everything_below() {
        let path;
        {
            let t = TestKey::new("guard");
            path = t.0.clone();
            assert!(path.starts_with("Software\\WinThemeSwitcherTest-"));
            assert!(!key_exists(&path), "not created by new()");
            set_raw_value(&t.sub("a\\b"), "v", REG_DWORD, &1u32.to_le_bytes());
            set_raw_value(&t.0, "w", REG_SZ, &[0, 0]);
            assert!(key_exists(&t.sub("a\\b")));
        }
        assert!(!key_exists(&path), "deleted on drop");
        // And on the way out of a failed assertion.
        let caught = std::panic::catch_unwind(|| {
            let t = TestKey::new("guard-panic");
            set_raw_value(&t.0, "v", REG_SZ, &[0, 0]);
            let path = t.0.clone();
            std::panic::resume_unwind(Box::new(path));
        });
        let path = *caught.unwrap_err().downcast::<String>().unwrap();
        assert!(!key_exists(&path), "deleted when the test fails");
    }

    #[test]
    fn enabling_autostart_creates_a_missing_run_key_and_its_value() {
        // The key, and here its parent too, isn't there (deleted on
        // purpose — every profile starts with one): enable used to fail
        // with "RegOpenKeyExW Run rc=2" at every start.
        let t = TestKey::new("run-create");
        let run = t.sub("CurrentVersion\\Run");
        assert_eq!(hkcu_value_exists(&run, APP_NAME), Some(false), "no key");
        assert_eq!(
            set_run_value(&run, APP_NAME, Some(CMD)),
            Ok(true),
            "key and value created"
        );
        assert_eq!(read_sz(&run, APP_NAME).as_deref(), Some(CMD));
        assert_eq!(hkcu_value_exists(&run, APP_NAME), Some(true));
        // Stored as REG_SZ with its terminator, as Windows expects.
        let (kind, data) = read_hkcu_value(&run, APP_NAME).unwrap();
        assert_eq!(kind, REG_SZ);
        assert_eq!(data.len(), (CMD.encode_utf16().count() + 1) * 2);
        assert_eq!(&data[data.len() - 2..], [0, 0]);
    }

    #[test]
    fn enabling_autostart_writes_only_when_the_value_differs() {
        let t = TestKey::new("run-match");
        assert_eq!(set_run_value(&t.0, APP_NAME, Some(CMD)), Ok(true));
        // Other programs' entries are not this function's business.
        set_raw_value(&t.0, "Other App", REG_SZ, &[b'x', 0, 0, 0]);
        // Already what it should be: no registry write is made. What the
        // function reports is the evidence — the key's last-write time
        // can't be, since Windows leaves a key alone anyway when a value is
        // set to the data it already has (measured on build 26300).
        assert_eq!(set_run_value(&t.0, APP_NAME, Some(CMD)), Ok(false));
        assert_eq!(read_sz(&t.0, APP_NAME).as_deref(), Some(CMD));
        // Another exe (the app was moved): rewritten.
        let moved = "\"D:\\Apps\\WinThemeSwitcher\\win-theme-switcher.exe\"";
        assert_eq!(set_run_value(&t.0, APP_NAME, Some(moved)), Ok(true));
        assert_eq!(read_sz(&t.0, APP_NAME).as_deref(), Some(moved));
        assert_eq!(set_run_value(&t.0, APP_NAME, Some(moved)), Ok(false));
        // A value of another type with that name: replaced by the string.
        set_raw_value(&t.0, APP_NAME, REG_BINARY, &[1, 2, 3]);
        assert_eq!(set_run_value(&t.0, APP_NAME, Some(moved)), Ok(true));
        assert_eq!(read_sz(&t.0, APP_NAME).as_deref(), Some(moved));
        assert_eq!(read_sz(&t.0, "Other App").as_deref(), Some("x"));
    }

    #[test]
    fn from_a_temp_folder_the_run_value_is_left_exactly_as_it_is() {
        use AutoStartOutcome::*;
        // The Run value of the real install, and an exe started from
        // inside the zip that would otherwise point it at itself.
        let t = TestKey::new("run-temp");
        let unpacked =
            "\"C:\\Users\\Ann\\AppData\\Local\\Temp\\Temp1_wts.zip\\win-theme-switcher.exe\"";
        let set_run = |register: bool| {
            set_run_value(&t.0, APP_NAME, register.then_some(unpacked)).map(|_| ())
        };
        set_run_value(&t.0, APP_NAME, Some(CMD)).unwrap();
        let in_temp = auto_start_plan(true, true);
        assert_eq!(carry_out_auto_start(in_temp, set_run), LeftAlone);
        assert_eq!(
            read_sz(&t.0, APP_NAME).as_deref(),
            Some(CMD),
            "neither rewritten nor removed"
        );
        // The registry step isn't even called.
        let never = |_: bool| -> Result<(), String> { panic!("the Run value must be left alone") };
        assert_eq!(carry_out_auto_start(in_temp, never), LeftAlone);
        // No entry yet: none is made from there.
        set_run_value(&t.0, APP_NAME, None).unwrap();
        assert_eq!(carry_out_auto_start(in_temp, set_run), LeftAlone);
        assert_eq!(hkcu_value_exists(&t.0, APP_NAME), Some(false));
        // From a permanent folder it is registered …
        let installed = auto_start_plan(true, false);
        assert_eq!(carry_out_auto_start(installed, set_run), Registered);
        assert_eq!(read_sz(&t.0, APP_NAME).as_deref(), Some(unpacked));
        // … and auto_start: false removes it from anywhere.
        let off = auto_start_plan(false, true);
        assert_eq!(carry_out_auto_start(off, set_run), Removed);
        assert_eq!(hkcu_value_exists(&t.0, APP_NAME), Some(false));
        // A failed registry call comes back with the autostart_err line's
        // op and message.
        let denied = |_: bool| Err("RegSetValueExW Run rc=5".to_string());
        assert_eq!(
            carry_out_auto_start(installed, denied),
            Failed {
                op: "set",
                msg: "RegSetValueExW Run rc=5".to_string()
            }
        );
        let denied = |_: bool| Err("RegDeleteValueW Run rc=5".to_string());
        assert_eq!(
            carry_out_auto_start(off, denied),
            Failed {
                op: "delete",
                msg: "RegDeleteValueW Run rc=5".to_string()
            }
        );
    }

    #[test]
    fn disabling_autostart_removes_the_value_and_nothing_else() {
        let t = TestKey::new("run-remove");
        set_run_value(&t.0, APP_NAME, Some(CMD)).unwrap();
        set_raw_value(&t.0, "Other App", REG_SZ, &[b'x', 0, 0, 0]);
        assert_eq!(set_run_value(&t.0, APP_NAME, None), Ok(true), "removed");
        assert_eq!(hkcu_value_exists(&t.0, APP_NAME), Some(false));
        assert_eq!(read_sz(&t.0, "Other App").as_deref(), Some("x"));
        // Already gone: success, again and again.
        assert_eq!(
            set_run_value(&t.0, APP_NAME, None),
            Ok(false),
            "a missing value is success"
        );
        assert!(key_exists(&t.0), "the key itself stays");
    }

    #[test]
    fn disabling_autostart_without_a_run_key_is_success_and_creates_none() {
        let t = TestKey::new("run-remove-no-key");
        let run = t.sub("Run");
        assert_eq!(
            set_run_value(&run, APP_NAME, None),
            Ok(false),
            "a missing key is success"
        );
        assert!(!key_exists(&run) && !key_exists(&t.0), "nothing created");
        assert_eq!(hkcu_value_exists(&run, APP_NAME), Some(false));
        assert_eq!(read_hkcu_value(&run, APP_NAME), None);
    }

    #[test]
    fn the_startup_switch_is_read_from_the_registry_as_stored() {
        // A stand-in for StartupApproved\Run, with records as Windows
        // writes them — the real key is only ever read, by the app.
        let t = TestKey::new("startup-approved");
        let off = [
            0x03, 0, 0, 0, 0x5C, 0x1B, 0x9A, 0x2E, 0x47, 0x33, 0xDC, 0x01,
        ];
        set_raw_value(&t.0, APP_NAME, REG_BINARY, &off);
        set_raw_value(
            &t.0,
            "On",
            REG_BINARY,
            &[2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        );
        set_raw_value(&t.0, "Empty", REG_BINARY, &[]);
        set_raw_value(&t.0, "Text", REG_SZ, &[b'3', 0, 0, 0]);
        let switch = |name: &str| {
            read_hkcu_value(&t.0, name).map(|(kind, data)| startup_switch(kind, &data))
        };
        assert_eq!(
            read_hkcu_value(&t.0, APP_NAME),
            Some((REG_BINARY, off.to_vec()))
        );
        assert_eq!(switch(APP_NAME), Some(StartupSwitch::Off));
        assert_eq!(switch("On"), Some(StartupSwitch::On));
        assert_eq!(switch("Empty"), Some(StartupSwitch::Unknown));
        assert_eq!(switch("Text"), Some(StartupSwitch::Unknown));
        // No record: an entry may have none (it means on, like a 02 one).
        assert_eq!(switch("Missing"), None);
        // Reading changed nothing.
        assert_eq!(
            read_hkcu_value(&t.0, APP_NAME),
            Some((REG_BINARY, off.to_vec()))
        );
    }

    #[test]
    fn the_switched_off_line_is_due_once_and_only_for_a_record_that_says_off() {
        let t = TestKey::new("startup-note");
        let off = [
            0x03, 0, 0, 0, 0x5C, 0x1B, 0x9A, 0x2E, 0x47, 0x33, 0xDC, 0x01,
        ];
        set_raw_value(&t.0, "Off", REG_BINARY, &off);
        set_raw_value(
            &t.0,
            "On",
            REG_BINARY,
            &[2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        );
        set_raw_value(&t.0, "Short", REG_BINARY, &[3, 0, 0, 0]);
        set_raw_value(&t.0, "Text", REG_SZ, &[b'3', 0, 0, 0]);
        let logged = AtomicBool::new(false);
        // On, unrecognized, or no record at all: nothing to say …
        for quiet in ["On", "Short", "Text", "Missing"] {
            assert!(!startup_switch_note_due(&t.0, quiet, &logged), "{quiet}");
        }
        assert!(!startup_switch_note_due(&t.sub("NoKey"), "Off", &logged));
        // … and the one line isn't used up by it.
        assert!(!logged.load(Ordering::Relaxed));
        assert!(startup_switch_note_due(&t.0, "Off", &logged));
        assert!(
            !startup_switch_note_due(&t.0, "Off", &logged),
            "once per run of the app"
        );
        // The user's switch is only read.
        assert_eq!(
            read_hkcu_value(&t.0, "Off"),
            Some((REG_BINARY, off.to_vec()))
        );
    }

    #[test]
    fn the_registry_tier_creates_a_missing_personalize_key() {
        let t = TestKey::new("personalize");
        let key = t.sub("Themes\\Personalize");
        assert_eq!(
            read_hkcu_dword(&key, "AppsUseLightTheme"),
            Err("open-rc-2".to_string()),
            "no key yet"
        );
        let mixed = Modes {
            apps_light: true,
            system_light: false,
        };
        write_modes_at(&key, mixed).expect("key created, both values written");
        assert_eq!(read_hkcu_dword(&key, "AppsUseLightTheme"), Ok(1));
        assert_eq!(read_hkcu_dword(&key, "SystemUsesLightTheme"), Ok(0));
        // With the key there: overwritten in place.
        write_modes_at(&key, Modes::from(Theme::Dark)).unwrap();
        assert_eq!(read_hkcu_dword(&key, "AppsUseLightTheme"), Ok(0));
        assert_eq!(read_hkcu_dword(&key, "SystemUsesLightTheme"), Ok(0));
        // What read_modes rejects: a value that isn't a 4-byte DWORD.
        set_raw_value(&key, "AppsUseLightTheme", REG_SZ, &[b'1', 0, 0, 0]);
        assert_eq!(
            read_hkcu_dword(&key, "AppsUseLightTheme"),
            Err("type-1-size-4".to_string())
        );
        assert_eq!(
            read_hkcu_dword(&key, "Missing"),
            Err("query-rc-2".to_string())
        );
    }

    #[test]
    fn a_registry_key_that_cannot_be_opened_names_the_api_and_its_code() {
        // Without create_missing a missing key is the caller's to judge.
        let t = TestKey::new("open-err");
        match open_hkcu_key(&t.0, KEY_QUERY_VALUE, false) {
            Err(e) => assert_eq!(e, ("RegOpenKeyExW", ERROR_FILE_NOT_FOUND)),
            Ok(_) => panic!("the key doesn't exist"),
        }
        assert!(!key_exists(&t.0), "an open never creates");
        assert!(open_hkcu_key(&t.0, KEY_QUERY_VALUE, true).is_ok());
        assert!(key_exists(&t.0));
    }

    #[test]
    fn a_failed_registry_call_is_reported_with_the_api_and_its_code() {
        use windows_sys::Win32::Foundation::ERROR_CHILD_MUST_BE_VOLATILE;
        use windows_sys::Win32::System::Registry::REG_OPTION_VOLATILE;
        let access = KEY_QUERY_VALUE | KEY_SET_VALUE;
        let code = |r: Result<RegKey, (&'static str, u32)>| r.err().expect("must fail");
        let dark = Modes::from(Theme::Dark);
        // An open that fails with something other than "not found": a key
        // name longer than the registry allows (255 characters). Nothing is
        // created for it, and the callers pass the API and its code on.
        let t = TestKey::new("reg-err");
        let too_long = t.sub(&"k".repeat(300));
        let (api, rc) = code(open_hkcu_key(&too_long, access, true));
        assert_eq!(api, "RegOpenKeyExW");
        assert!(rc != 0 && rc != ERROR_FILE_NOT_FOUND, "rc={rc}");
        assert_eq!(
            write_modes_at(&too_long, dark),
            Err((format!("RegOpenKeyExW Personalize rc={rc}"), false))
        );
        let run_err = Err(format!("RegOpenKeyExW Run rc={rc}"));
        assert_eq!(set_run_value(&too_long, APP_NAME, Some(CMD)), run_err);
        assert_eq!(set_run_value(&too_long, APP_NAME, None), run_err);
        // A key that isn't there and can't be created: a lasting key below
        // a volatile one. That is the create's error, not the open's.
        let parent = wide(&t.sub("volatile"));
        let mut hkey: HKEY = ptr::null_mut();
        let made = unsafe {
            RegCreateKeyExW(
                HKEY_CURRENT_USER,
                parent.as_ptr(),
                0,
                ptr::null(),
                REG_OPTION_VOLATILE,
                KEY_QUERY_VALUE,
                ptr::null(),
                &mut hkey,
                ptr::null_mut(),
            )
        };
        assert_eq!(made, 0);
        drop(RegKey(hkey));
        let child = t.sub("volatile\\Run");
        assert_eq!(
            code(open_hkcu_key(&child, access, true)),
            ("RegCreateKeyExW", ERROR_CHILD_MUST_BE_VOLATILE)
        );
        let rc = ERROR_CHILD_MUST_BE_VOLATILE;
        assert_eq!(
            write_modes_at(&child, dark),
            Err((format!("RegCreateKeyExW Personalize rc={rc}"), false))
        );
        assert_eq!(
            set_run_value(&child, APP_NAME, Some(CMD)),
            Err(format!("RegCreateKeyExW Run rc={rc}"))
        );
        // Removing never creates, so there it is just a key that isn't
        // there.
        assert_eq!(set_run_value(&child, APP_NAME, None), Ok(false));
        assert!(!key_exists(&child));
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

    /// decide_tick for stock-like slots (plain light/dark modes that tell the
    /// slots apart) with a schedule `next` that doesn't invoke the
    /// backward-step frame rule (it equals the recorded transition, as on any
    /// tick with an unstepped clock). Frame-rule and mixed-modes tests call
    /// decide_tick directly.
    fn dt(
        kind: TickKind,
        current: Option<Theme>,
        target: Theme,
        now: DateTime<Utc>,
        s: &TickState,
    ) -> TickAction {
        dtm(kind, current.map(Modes::from), target.into(), now, true, s)
    }

    /// dt with explicit modes and decidability.
    fn dtm(
        kind: TickKind,
        current: Option<Modes>,
        target: Modes,
        now: DateTime<Utc>,
        decidable: bool,
        s: &TickState,
    ) -> TickAction {
        let next = s
            .reconciled_next
            .filter(|&n| n > now)
            .unwrap_or(now + chrono::Duration::hours(12));
        decide_tick(kind, current, target, now, next, false, decidable, s)
    }

    /// Shorthand for a Modes value.
    fn md(apps_light: bool, system_light: bool) -> Modes {
        Modes {
            apps_light,
            system_light,
        }
    }

    /// A resolved slot with the given modes (the file is never read).
    fn slot(theme: Theme, modes: Modes, keyed: bool) -> Slot {
        Slot {
            theme,
            file: PathBuf::from("unused.theme"),
            modes,
            keyed,
            rejected: None,
        }
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
        // read_modes() = None (registry read failed): inside a preserved
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
        let (l, d) = (Modes::from(Theme::Light), Modes::from(Theme::Dark));
        let tt = |screen| toggle_target(screen, l, d, true, None, Theme::Light);
        assert_eq!(tt(Some(l)), Theme::Dark);
        assert_eq!(tt(Some(d)), Theme::Light);
        // Unreadable screen, nothing applied yet, no location (fallback day):
        // lands on Dark, as before v0.5.0.
        assert_eq!(tt(None), Theme::Dark);
    }

    #[test]
    fn toggle_target_table_with_mixed_and_undecidable_slots() {
        // Day theme = light apps + dark taskbar; night = all dark.
        let day = md(true, false);
        let night = md(false, false);
        let tt = |screen, last| toggle_target(screen, day, night, true, last, Theme::Light);
        assert_eq!(tt(Some(day), None), Theme::Dark);
        assert_eq!(tt(Some(night), None), Theme::Light);
        // Neither slot on screen (some other theme): the apps mode decides.
        assert_eq!(tt(Some(md(true, true)), None), Theme::Dark);
        assert_eq!(tt(Some(md(false, true)), None), Theme::Light);
        // Decidable slots still ignore last_applied when the screen reads.
        assert_eq!(tt(Some(day), Some(Theme::Dark)), Theme::Dark);
        // Unreadable screen: flip what was last applied, else the fallback.
        assert_eq!(tt(None, Some(Theme::Dark)), Theme::Light);
        assert_eq!(
            toggle_target(None, day, night, true, None, Theme::Dark),
            Theme::Light
        );
    }

    #[test]
    fn toggle_alternates_when_the_slots_share_modes() {
        // Two light-mode themes differing only in wallpaper: the screen can't
        // say which is on, so two presses must go A → B → A, not B → B.
        let l = Modes::from(Theme::Light);
        let mut last = Some(Theme::Light);
        let mut seen = Vec::new();
        for _ in 0..3 {
            let t = toggle_target(Some(l), l, l, false, last, Theme::Light);
            seen.push(t);
            last = Some(t);
        }
        assert_eq!(seen, [Theme::Dark, Theme::Light, Theme::Dark]);
        // Nothing applied yet: flip the schedule's current slot.
        assert_eq!(
            toggle_target(Some(l), l, l, false, None, Theme::Dark),
            Theme::Light
        );
    }

    #[test]
    fn sync_is_decidable_only_for_keyed_distinct_slots() {
        let (l, d) = (Modes::from(Theme::Light), Modes::from(Theme::Dark));
        let day = |m, k| slot(Theme::Light, m, k);
        let night = |m, k| slot(Theme::Dark, m, k);
        assert!(sync_decidable(&day(l, true), &night(d, true)));
        assert!(sync_decidable(
            &day(md(true, false), true),
            &night(md(false, false), true)
        ));
        // Identical modes (wallpaper-only difference).
        assert!(!sync_decidable(&day(l, true), &night(l, true)));
        // A slot whose file doesn't declare its modes (high contrast).
        assert!(!sync_decidable(&day(l, true), &night(d, false)));
        assert!(!sync_decidable(&day(l, false), &night(d, true)));
    }

    #[test]
    fn modes_from_parses_both_keys_and_defaults_per_value() {
        let (l, d) = (Theme::Light, Theme::Dark);
        assert_eq!(
            modes_from(l, Some("Light"), Some("Light")),
            (md(true, true), true)
        );
        assert_eq!(
            modes_from(d, Some(" dark "), Some("LIGHT")),
            (md(true, false), true)
        );
        // Absent / empty / unknown values take the SLOT's default, each
        // value independently, and mark the slot as guessed.
        assert_eq!(modes_from(l, None, None), (md(true, true), false));
        assert_eq!(modes_from(d, None, None), (md(false, false), false));
        assert_eq!(modes_from(d, Some("Light"), None), (md(false, true), false));
        assert_eq!(
            modes_from(l, Some(""), Some("Dark")),
            (md(false, true), false)
        );
        assert_eq!(
            modes_from(d, Some("Dim"), Some("Dark")),
            (md(false, false), false)
        );
    }

    #[test]
    fn modes_log_form_has_no_separators_inside() {
        assert_eq!(modes_str(None), "unknown");
        assert_eq!(modes_str(Some(md(true, true))), "light");
        assert_eq!(modes_str(Some(md(false, false))), "dark");
        assert_eq!(modes_str(Some(md(true, false))), "apps-light/system-dark");
        assert_eq!(modes_str(Some(md(false, true))), "apps-dark/system-light");
    }

    #[test]
    fn slot_note_only_speaks_up_when_modes_dont_name_the_slot() {
        let (l, d) = (Modes::from(Theme::Light), Modes::from(Theme::Dark));
        assert_eq!(slot_note(&slot(Theme::Light, l, true), true), "");
        assert_eq!(slot_note(&slot(Theme::Dark, d, true), true), "");
        assert_eq!(
            slot_note(&slot(Theme::Light, md(true, false), true), true),
            " slot=day"
        );
        assert_eq!(
            slot_note(&slot(Theme::Dark, l, true), false),
            " slot=night sync=undecidable"
        );
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
        let now = utc(2026, 7, 4, 9, 0, 0);
        let far_next = utc(2026, 7, 4, 15, 0, 0);
        let near_next = utc(2026, 7, 4, 9, 0, 30);
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
            Some(Theme::Light.into()),
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
            Some(Theme::Light.into()),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        // The user picks Dark (via Toggle or Settings), moving the screen off
        // the failure baseline (Light). Note the modes can only show a MOVE:
        // re-picking the stuck theme (Light) in Settings is indistinguishable
        // from the failure itself, so that case retries — an inherent limit
        // of observing read_modes(). (A Toggle is recorded separately: see
        // note_toggle and the tests after reconcile_clears_retry_episode.)
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
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light.into()),
            sunrise
        ));
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
        assert!(note_apply_failed(&mut s, Some(Theme::Dark.into()), sunset));
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
            Some(Theme::Light.into()),
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
        assert!(note_apply_failed(&mut s, Some(Theme::Dark.into()), sunset));
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
                Some(Theme::Dark.into()),
                Theme::Light.into(),
                real_now,
                sunset,
                true,
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
                Some(Theme::Dark.into()),
                Theme::Light.into(),
                real_now,
                tomorrow_sunrise,
                false,
                true,
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
                    decide_tick(
                        kind,
                        Some(Theme::Dark.into()),
                        Theme::Light.into(),
                        now,
                        next,
                        true,
                        true,
                        &s
                    ),
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
                Some(Theme::Light.into()),
                Theme::Dark.into(),
                now,
                earlier_next,
                false,
                true,
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
            assert!(note_apply_failed(&mut s, Some(Theme::Dark.into()), sunset));
        }
        assert!(!note_apply_failed(&mut s, Some(Theme::Dark.into()), sunset));
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
            Some(Theme::Light.into()),
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
            Some(Theme::Light.into()),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light.into()),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light.into()),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        assert!(!note_apply_failed(
            &mut s,
            Some(Theme::Light.into()),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        assert_eq!(s.retry_count, 0);
        assert_eq!(s.retry_baseline, None);
        assert!(note_apply_failed(
            &mut s,
            Some(Theme::Light.into()),
            utc(2026, 7, 5, 2, 35, 0)
        ));
    }

    #[test]
    fn reconcile_clears_retry_episode() {
        let next = utc(2026, 7, 5, 2, 35, 0);
        let mut s = TickState::new();
        assert!(note_apply_failed(&mut s, Some(Theme::Light.into()), next));
        note_reconciled(&mut s, next);
        assert_eq!(s.retry_count, 0);
        assert_eq!(s.retry_baseline, None);
        assert_eq!(s.reconciled_next, Some(next));
    }

    // --- Toggle during a pending retry; settle_tick (v0.6.0) ---

    /// One tick as `tick()` runs it, minus the I/O: decide from the pre-tick
    /// state, then settle the outcome. `apply_fails_leaving` is what a failed
    /// apply leaves on screen (None = the apply succeeds, if there is one).
    fn run_tick(
        s: &mut TickState,
        kind: TickKind,
        current: Option<Modes>,
        target: Modes,
        now: DateTime<Utc>,
        next: DateTime<Utc>,
        apply_fails_leaving: Option<Option<Modes>>,
    ) -> (TickAction, Settled) {
        let action = decide_tick(kind, current, target, now, next, false, true, s);
        let outcome = match (action, apply_fails_leaving) {
            (TickAction::Apply, Some(after)) => TickOutcome::ApplyFailed { after },
            _ => TickOutcome::Reconciled,
        };
        (action, settle_tick(s, kind, outcome, now, next))
    }

    #[test]
    fn a_double_toggle_during_a_retry_stands_the_retry_down() {
        // Sunset: the apply to dark fails and the screen stays light. Within
        // the minute before the retry the user toggles twice, ending on
        // light again — exactly the retry baseline. That is their choice,
        // and the retry must leave it alone.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let sunrise = utc(2026, 7, 5, 2, 35, 0);
        let (light, dark): (Modes, Modes) = (Theme::Light.into(), Theme::Dark.into());
        let failed = || {
            let mut s = reconciled_at(sunset);
            let (action, settled) = run_tick(
                &mut s,
                TickKind::Scheduled,
                Some(light),
                dark,
                sunset,
                sunrise,
                Some(Some(light)),
            );
            assert_eq!(action, TickAction::Apply);
            assert_eq!(settled.retry, Some(RetryNote::Attempt(1)));
            assert_eq!(settled.deadline, sunset + mins(1));
            s
        };
        let retry_at = sunset + mins(1);

        // Control: nobody touched anything, so the retry applies.
        let s = failed();
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

        // Two toggles: the screen is back on the baseline, the retry still
        // stands down — as the retry tick, or as an unlock that gets there
        // first.
        for kind in [TickKind::Scheduled, TickKind::Wake] {
            let mut s = failed();
            note_toggle(&mut s, true);
            note_toggle(&mut s, true);
            let (action, settled) =
                run_tick(&mut s, kind, Some(light), dark, retry_at, sunrise, None);
            assert_eq!(action, TickAction::CancelRetry, "{kind:?}");
            // The episode is over and the tick waits for sunrise...
            assert_eq!(
                settled,
                Settled {
                    deadline: sunrise,
                    retry: None
                }
            );
            assert_eq!((s.retry_count, s.retry_toggled), (0, false));
            // ...and until then the toggled theme is an override like any
            // other: a later unlock preserves it.
            assert_eq!(s.reconciled_next, Some(sunrise));
            assert_eq!(
                dt(
                    TickKind::Wake,
                    Some(Theme::Light),
                    Theme::Dark,
                    sunset + mins(90),
                    &s
                ),
                TickAction::SkipOverride
            );
        }

        // One toggle is seen twice over (the screen moved, and the flag).
        let mut s = failed();
        note_toggle(&mut s, true);
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
    }

    #[test]
    fn a_toggle_stands_a_retry_down_where_the_modes_cant_show_it() {
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let sunrise = utc(2026, 7, 5, 2, 35, 0);
        let retry_at = sunset + mins(1);
        let l = md(true, true);
        // (baseline, screen at the retry, sync decidable): two light-mode
        // themes, where a toggle changes the wallpaper only; a registry that
        // couldn't be read after the failure; one that can't be read now;
        // one that couldn't be read either time.
        for (baseline, current, decidable) in [
            (Some(l), Some(l), false),
            (None, Some(l), true),
            (Some(l), None, true),
            (None, None, true),
        ] {
            let mut s = reconciled_at(sunset);
            assert!(note_apply_failed(&mut s, baseline, sunrise));
            let target = if decidable { md(false, false) } else { l };
            assert_eq!(
                dtm(
                    TickKind::Scheduled,
                    current,
                    target,
                    retry_at,
                    decidable,
                    &s
                ),
                TickAction::Apply,
                "without a toggle: {baseline:?} {current:?}"
            );
            note_toggle(&mut s, true);
            assert_eq!(
                dtm(
                    TickKind::Scheduled,
                    current,
                    target,
                    retry_at,
                    decidable,
                    &s
                ),
                TickAction::CancelRetry,
                "after a toggle: {baseline:?} {current:?}"
            );
        }
    }

    #[test]
    fn a_toggle_only_flags_the_failure_it_follows() {
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let sunrise = utc(2026, 7, 5, 2, 35, 0);
        let (light, dark): (Modes, Modes) = (Theme::Light.into(), Theme::Dark.into());
        // Everything in TickState but the flag, which a Toggle must leave
        // alone: what the state knows of the schedule, the retry episode,
        // the armed deadline and the clock mark.
        let rest = |s: &TickState| {
            (
                s.reconciled_next,
                s.retry_count,
                s.retry_baseline,
                s.episode_next,
                s.armed,
                s.mark,
            )
        };

        // No retry pending: nothing is recorded at all.
        let mut s = reconciled_at(sunset);
        s.armed = Some(sunset);
        s.mark = Some((sunset - mins(400), Instant::now()));
        let before = rest(&s);
        note_toggle(&mut s, true);
        assert!(!s.retry_toggled);
        assert_eq!(rest(&s), before);
        // A toggle BEFORE a failure doesn't stand its retry down (the flag
        // wasn't set, and the failure would have dropped it anyway).
        assert!(note_apply_failed(&mut s, Some(light), sunrise));
        assert_eq!(
            dtm(
                TickKind::Scheduled,
                Some(light),
                dark,
                sunset + mins(1),
                true,
                &s
            ),
            TickAction::Apply
        );

        // A toggle whose apply failed isn't recorded either, and the retry
        // still runs...
        let before = rest(&s);
        note_toggle(&mut s, false);
        assert!(!s.retry_toggled);
        assert_eq!(rest(&s), before);
        assert_eq!(
            dtm(
                TickKind::Scheduled,
                Some(light),
                dark,
                sunset + mins(1),
                true,
                &s
            ),
            TickAction::Apply
        );
        // ...one that applied sets the flag and nothing else, and a failed
        // one after it doesn't take the flag back.
        note_toggle(&mut s, true);
        assert!(s.retry_toggled);
        assert_eq!(rest(&s), before);
        note_toggle(&mut s, false);
        assert!(s.retry_toggled);

        // The flag stands the retry down up to the episode's window, not at
        // or past it, as for a move the modes show: a retry armed for the
        // transition itself (a failure in its last minute) reconciles to the
        // schedule. (An unreadable screen, so only the flag can decide.)
        for (at, want) in [
            (
                sunrise - chrono::Duration::milliseconds(1),
                TickAction::CancelRetry,
            ),
            (sunrise, TickAction::Apply),
        ] {
            assert_eq!(
                decide_tick(
                    TickKind::Scheduled,
                    None,
                    light,
                    at,
                    sunrise + mins(790),
                    false,
                    true,
                    &s
                ),
                want
            );
        }

        // Further past the window: the machine slept through sunrise with
        // the screen still light, which is what the morning wants. Plain
        // in-sync, not a stood-down retry.
        let next_evening = utc(2026, 7, 5, 15, 45, 40);
        let morning = sunrise + mins(200);
        let (action, settled) = run_tick(
            &mut s,
            TickKind::Wake,
            Some(light),
            light,
            morning,
            next_evening,
            None,
        );
        assert_eq!(action, TickAction::SkipInSync);
        assert_eq!(settled.deadline, next_evening);
        assert_eq!(s.reconciled_next, Some(next_evening));
        assert_eq!((s.retry_count, s.retry_toggled), (0, false));

        // Same night, but one toggle left the screen dark: the morning wake
        // applies. If THAT apply fails, the failure opens a new window (the
        // count carries on: retry 2) with a new baseline and no flag, so its
        // retry runs...
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(&mut s, Some(light), sunrise));
        note_toggle(&mut s, true);
        let (action, settled) = run_tick(
            &mut s,
            TickKind::Wake,
            Some(dark),
            light,
            morning,
            next_evening,
            Some(Some(dark)),
        );
        assert_eq!(action, TickAction::Apply);
        assert_eq!(settled.retry, Some(RetryNote::Attempt(2)));
        assert!(!s.retry_toggled, "a new failure drops the flag");
        assert_eq!(s.retry_baseline, Some(dark));
        assert_eq!(s.episode_next, Some(next_evening));
        let retry = |s: &TickState| {
            decide_tick(
                TickKind::Scheduled,
                Some(dark),
                light,
                morning + mins(1),
                next_evening,
                false,
                true,
                s,
            )
        };
        assert_eq!(retry(&s), TickAction::Apply);
        // ...unless a toggle follows THIS failure, inside the new window.
        note_toggle(&mut s, true);
        note_toggle(&mut s, true);
        assert_eq!(retry(&s), TickAction::CancelRetry);
    }

    #[test]
    fn refresh_applies_over_a_toggled_retry_and_starts_a_fresh_episode() {
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let sunrise = utc(2026, 7, 5, 2, 35, 0);
        let (light, dark): (Modes, Modes) = (Theme::Light.into(), Theme::Dark.into());
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(&mut s, Some(light), sunrise));
        assert!(note_apply_failed(&mut s, Some(light), sunrise));
        note_toggle(&mut s, true);
        // Refresh force-applies whatever the gate would say...
        let now = sunset + mins(2);
        let (action, settled) = run_tick(
            &mut s,
            TickKind::Refresh,
            Some(light),
            dark,
            now,
            sunrise,
            Some(Some(light)),
        );
        assert_eq!(action, TickAction::Apply);
        // ...and when it fails too, its episode is new: retry 1 of 3, not 3,
        // and no toggle on record.
        assert_eq!(settled.retry, Some(RetryNote::Attempt(1)));
        assert_eq!(settled.deadline, now + mins(1));
        assert_eq!((s.retry_count, s.retry_toggled), (1, false));
        assert_eq!(
            dtm(
                TickKind::Scheduled,
                Some(light),
                dark,
                now + mins(1),
                true,
                &s
            ),
            TickAction::Apply
        );
    }

    #[test]
    fn settle_tick_records_a_reconciled_tick_and_waits_for_the_transition() {
        let next = utc(2026, 7, 5, 2, 35, 0);
        let now = utc(2026, 7, 4, 15, 46, 5);
        for kind in [
            TickKind::Init,
            TickKind::Scheduled,
            TickKind::Wake,
            TickKind::Refresh,
        ] {
            // From a clean state, and from the middle of a failure episode.
            let mut clean = TickState::new();
            let mut failing = reconciled_at(now);
            assert!(note_apply_failed(&mut failing, Some(md(true, true)), next));
            note_toggle(&mut failing, true);
            for s in [&mut clean, &mut failing] {
                let settled = settle_tick(s, kind, TickOutcome::Reconciled, now, next);
                assert_eq!(
                    settled,
                    Settled {
                        deadline: next,
                        retry: None
                    }
                );
                assert_eq!(s.reconciled_next, Some(next));
                assert_eq!(
                    (
                        s.retry_count,
                        s.retry_baseline,
                        s.episode_next,
                        s.retry_toggled
                    ),
                    (0, None, None, false)
                );
            }
        }
    }

    #[test]
    fn settle_tick_schedules_bounded_retries_from_the_reading_after_the_failure() {
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let sunrise = utc(2026, 7, 5, 2, 35, 0);
        let half = md(false, true);
        let mut s = reconciled_at(sunset);
        let mut now = sunset;
        for attempt in 1..=MAX_APPLY_RETRIES {
            // The apply ran for 3 s before failing: the retry is timed from
            // the failure.
            let failed_at = now + chrono::Duration::seconds(3);
            let settled = settle_tick(
                &mut s,
                TickKind::Scheduled,
                TickOutcome::ApplyFailed { after: Some(half) },
                failed_at,
                sunrise,
            );
            assert_eq!(settled.retry, Some(RetryNote::Attempt(attempt)));
            assert_eq!(
                settled.deadline,
                failed_at + chrono::Duration::seconds(APPLY_RETRY_DELAY_SECS)
            );
            // The baseline is the reading the failure left behind, the
            // window is this tick's next transition, and a failure never
            // counts as reconciled.
            assert_eq!(s.retry_count, attempt);
            assert_eq!(s.retry_baseline, Some(half));
            assert_eq!(s.episode_next, Some(sunrise));
            assert_eq!(s.reconciled_next, Some(sunset));
            now = settled.deadline;
        }
        // One failure more: the budget is spent. Wait for sunrise, with
        // nothing left that could make a later wake "preserve" the failure.
        let settled = settle_tick(
            &mut s,
            TickKind::Scheduled,
            TickOutcome::ApplyFailed { after: Some(half) },
            now,
            sunrise,
        );
        assert_eq!(
            settled,
            Settled {
                deadline: sunrise,
                retry: Some(RetryNote::Exhausted)
            }
        );
        assert_eq!(s.reconciled_next, None);
        assert_eq!(
            (
                s.retry_count,
                s.retry_baseline,
                s.episode_next,
                s.retry_toggled
            ),
            (0, None, None, false)
        );

        // A failure less than a minute before the transition retries AT the
        // transition, never after it; an unreadable screen after it replaces
        // the baseline like any other reading.
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(&mut s, Some(half), sunrise));
        let late = sunrise - chrono::Duration::seconds(20);
        let settled = settle_tick(
            &mut s,
            TickKind::Wake,
            TickOutcome::ApplyFailed { after: None },
            late,
            sunrise,
        );
        assert_eq!(
            settled,
            Settled {
                deadline: sunrise,
                retry: Some(RetryNote::Attempt(2))
            }
        );
        assert_eq!(s.retry_baseline, None);
    }

    #[test]
    fn settle_tick_gives_a_failed_refresh_a_fresh_budget() {
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let sunrise = utc(2026, 7, 5, 2, 35, 0);
        let failed = TickOutcome::ApplyFailed {
            after: Some(md(true, true)),
        };
        for kind in [TickKind::Scheduled, TickKind::Wake, TickKind::Refresh] {
            let mut s = reconciled_at(sunset);
            for _ in 0..MAX_APPLY_RETRIES {
                assert!(note_apply_failed(&mut s, Some(md(true, true)), sunrise));
            }
            let settled = settle_tick(&mut s, kind, failed, sunset + mins(3), sunrise);
            if kind == TickKind::Refresh {
                // The user asked: three more tries, starting now.
                assert_eq!(settled.retry, Some(RetryNote::Attempt(1)));
                assert_eq!(settled.deadline, sunset + mins(4));
                assert_eq!(s.reconciled_next, Some(sunset));
            } else {
                assert_eq!(settled.retry, Some(RetryNote::Exhausted), "{kind:?}");
                assert_eq!(settled.deadline, sunrise);
            }
        }
    }

    // --- two-value modes (v0.5.0) ---

    #[test]
    fn partial_registry_write_is_retried_not_mistaken_for_the_user() {
        // THE v0.5.0 design-review blocker. Sunset, target dark/dark from
        // light/light: tiers 1-2 fail, tier 3 flips Apps then fails on
        // System. The baseline is read AFTER the failure (apps dark, system
        // light), so the retry sees no move and re-applies.
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let sunrise = utc(2026, 7, 5, 2, 35, 0);
        let mut s = reconciled_at(sunset);
        let target = md(false, false);
        let half = md(false, true);
        assert!(note_apply_failed(&mut s, Some(half), sunrise));
        let retry_at = utc(2026, 7, 4, 15, 47, 5);
        assert_eq!(
            dtm(TickKind::Scheduled, Some(half), target, retry_at, true, &s),
            TickAction::Apply
        );
        // A mixed target reached by the partial write (only Apps had to
        // change, System's write "failed" re-writing the same value): the
        // retry finds the screen in sync.
        let mixed = md(false, true);
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(&mut s, Some(mixed), sunrise));
        assert_eq!(
            dtm(TickKind::Scheduled, Some(mixed), mixed, retry_at, true, &s),
            TickAction::SkipInSync
        );
    }

    #[test]
    fn a_one_value_move_during_a_retry_is_an_intervention() {
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let mut s = reconciled_at(sunset);
        assert!(note_apply_failed(
            &mut s,
            Some(md(true, true)),
            utc(2026, 7, 5, 2, 35, 0)
        ));
        // The user switched only the apps mode in Settings.
        assert_eq!(
            dtm(
                TickKind::Scheduled,
                Some(md(false, true)),
                md(false, false),
                utc(2026, 7, 4, 15, 47, 5),
                true,
                &s
            ),
            TickAction::CancelRetry
        );
    }

    #[test]
    fn mixed_target_needs_both_values_to_match() {
        // Day theme = light apps, dark taskbar.
        let day = md(true, false);
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let s = reconciled_at(sunset);
        let noon = utc(2026, 7, 4, 9, 0, 0);
        assert_eq!(
            dtm(TickKind::Wake, Some(day), day, noon, true, &s),
            TickAction::SkipInSync
        );
        // One value off inside the window: an override, preserved...
        assert_eq!(
            dtm(TickKind::Wake, Some(md(true, true)), day, noon, true, &s),
            TickAction::SkipOverride
        );
        // ...and reconciled once a transition has passed.
        assert_eq!(
            dtm(
                TickKind::Scheduled,
                Some(md(true, true)),
                md(false, false),
                sunset,
                true,
                &s
            ),
            TickAction::Apply
        );
    }

    #[test]
    fn undecidable_sync_never_skips_as_in_sync_but_still_preserves_and_stands_down() {
        // Two light-mode themes (wallpaper-only difference): matching modes
        // prove nothing, so Init re-applies; everything else works as usual.
        let l = md(true, true);
        let sunset = utc(2026, 7, 4, 15, 46, 5);
        let sunrise = utc(2026, 7, 5, 2, 35, 0);
        let noon = utc(2026, 7, 4, 9, 0, 0);
        let mut s = TickState::new();
        assert_eq!(
            dtm(TickKind::Init, Some(l), l, noon, false, &s),
            TickAction::Apply
        );
        note_reconciled(&mut s, sunset);
        // A wake inside the window is left alone (no transition passed).
        assert_eq!(
            dtm(TickKind::Wake, Some(l), l, noon, false, &s),
            TickAction::SkipOverride
        );
        // The transition applies.
        assert_eq!(
            dtm(TickKind::Scheduled, Some(l), l, sunset, false, &s),
            TickAction::Apply
        );
        // A failed apply, then the user visibly picks a dark theme: the
        // retry stands down — the evidence is real even though sync isn't
        // decidable.
        assert!(note_apply_failed(&mut s, Some(l), sunrise));
        assert_eq!(
            dtm(
                TickKind::Scheduled,
                Some(md(false, false)),
                l,
                sunset + mins(1),
                false,
                &s
            ),
            TickAction::CancelRetry
        );
    }

    #[test]
    fn slot_modes_come_from_the_visualstyles_section_of_the_applied_file() {
        let t = TempTheme::new(
            "modes-mixed",
            b"[Theme]\r\nSystemMode=Light\r\n[VisualStyles]\r\nsystemmode = Dark\r\nAppMode=Light\r\n",
        );
        let cfg = Config {
            theme_day: Some(t.0.to_string_lossy().into_owned()),
            ..Config::default()
        };
        let slot = resolve_slot(Theme::Light, &cfg);
        assert_eq!(slot.file, t.0);
        assert_eq!(slot.modes, md(true, false));
        assert!(slot.keyed);
        assert!(slot.rejected.is_none());
    }

    #[test]
    fn slot_modes_read_utf16_files_and_default_when_keys_are_missing() {
        let text = "[VisualStyles]\r\nSystemMode=Dark\r\nAppMode=Dark\r\n";
        let mut bytes = vec![0xFF, 0xFE];
        for u in text.encode_utf16() {
            bytes.extend_from_slice(&u.to_le_bytes());
        }
        let t16 = TempTheme::new("modes-utf16", &bytes);
        let keyless = TempTheme::new("modes-keyless", b"[Theme]\r\nDisplayName=HC-like\r\n");
        let cfg = Config {
            theme_day: Some(t16.0.to_string_lossy().into_owned()),
            theme_night: Some(keyless.0.to_string_lossy().into_owned()),
            ..Config::default()
        };
        let day = resolve_slot(Theme::Light, &cfg);
        assert_eq!((day.modes, day.keyed), (md(false, false), true));
        let night = resolve_slot(Theme::Dark, &cfg);
        assert_eq!((night.modes, night.keyed), (md(false, false), false));
        assert!(!sync_decidable(&day, &night));
    }

    #[test]
    fn a_rejected_path_takes_the_stock_files_modes_not_the_configured_ones() {
        // The configured file declares light/light but is a .themepack, so
        // the stock night theme is applied — the target must be ITS modes.
        let pack =
            std::env::temp_dir().join(format!("wts-test-{}-modes.themepack", std::process::id()));
        fs::write(
            &pack,
            b"[VisualStyles]\r\nSystemMode=Light\r\nAppMode=Light\r\n",
        )
        .unwrap();
        let cfg = Config {
            theme_night: Some(pack.to_string_lossy().into_owned()),
            ..Config::default()
        };
        let slot = resolve_slot(Theme::Dark, &cfg);
        let _ = fs::remove_file(&pack);
        assert_eq!(slot.file, stock_theme_file(Theme::Dark));
        assert!(slot.rejected.is_some());
        let stock = read_theme_ini(&slot.file, "VisualStyles", "SystemMode").flatten();
        let app = read_theme_ini(&slot.file, "VisualStyles", "AppMode").flatten();
        assert_eq!(
            (slot.modes, slot.keyed),
            modes_from(Theme::Dark, stock.as_deref(), app.as_deref())
        );
    }

    #[test]
    fn stock_themes_declare_their_modes() {
        // Guarded like the indirect-string test. Values are not asserted:
        // Windows 10's default theme is mixed (dark taskbar, light apps).
        for theme in [Theme::Light, Theme::Dark] {
            let f = stock_theme_file(theme);
            if !f.exists() {
                eprintln!("SKIP: {} absent", f.display());
                continue;
            }
            let slot = resolve_slot(theme, &Config::default());
            assert!(slot.keyed, "{} lacks SystemMode/AppMode", f.display());
        }
    }

    // --- tier-1 read-back (v0.5.1) ---

    #[test]
    fn only_a_declared_change_can_be_verified_from_the_registry() {
        let night = slot(Theme::Dark, md(false, false), true);
        assert_eq!(unverifiable_reason(&night, Some(md(true, true))), None);
        assert_eq!(
            unverifiable_reason(&night, Some(md(true, false))),
            None,
            "half matching still has a change to show"
        );
        assert_eq!(
            unverifiable_reason(&night, None),
            None,
            "unreadable before: a reading after can still tell"
        );
        assert_eq!(
            unverifiable_reason(&night, Some(md(false, false))),
            Some("already-matching")
        );
        let guessed = slot(Theme::Dark, md(false, false), false);
        for pre in [None, Some(md(true, true)), Some(md(false, false))] {
            assert_eq!(
                unverifiable_reason(&guessed, pre),
                Some("modes-not-in-theme-file")
            );
        }
    }

    #[test]
    fn readback_catches_a_tier1_success_that_changed_nothing() {
        let night = slot(Theme::Dark, md(false, false), true);
        let day = Some(md(true, true));
        // The modes stayed put for the whole window: a no-op, whatever
        // SetCurrentTheme returned. The last reading is reported.
        assert_eq!(
            readback_verdict(&night, day, &[day; 3]),
            ReadBack::NoOp(md(true, true))
        );
        // Half applied (only the taskbar changed) is not the target either.
        assert_eq!(
            readback_verdict(&night, day, &[Some(md(true, false))]),
            ReadBack::NoOp(md(true, false))
        );
        // Reached at once, or once Windows wrote the values a moment later.
        assert_eq!(
            readback_verdict(&night, day, &[Some(md(false, false))]),
            ReadBack::Reached
        );
        assert_eq!(
            readback_verdict(
                &night,
                day,
                &[day, Some(md(true, false)), Some(md(false, false))]
            ),
            ReadBack::Reached
        );
        assert_eq!(
            readback_verdict(&night, None, &[Some(md(false, false))]),
            ReadBack::Reached
        );
        // A failed reading proves nothing: the last one that succeeded
        // decides, and with none at all tier 1 is trusted.
        assert_eq!(
            readback_verdict(&night, day, &[day, None]),
            ReadBack::NoOp(md(true, true))
        );
        assert_eq!(
            readback_verdict(&night, day, &[None, None]),
            ReadBack::Unverifiable("unreadable")
        );
    }

    #[test]
    fn readback_trusts_tier1_when_the_modes_cant_tell() {
        // Same-mode themes (only the wallpaper differs) and key-less themes
        // have no change to show, so no reading can make them a no-op.
        let night = slot(Theme::Dark, md(false, false), true);
        assert_eq!(
            readback_verdict(&night, Some(md(false, false)), &[]),
            ReadBack::Unverifiable("already-matching")
        );
        let keyless = slot(Theme::Dark, md(false, false), false);
        assert_eq!(
            readback_verdict(&keyless, Some(md(true, true)), &[Some(md(true, true))]),
            ReadBack::Unverifiable("modes-not-in-theme-file")
        );
    }

    #[test]
    fn readback_polls_the_live_modes_until_the_target_or_the_window_ends() {
        // Read-only against this machine's Personalize values: the screen's
        // own modes are reached at the first reading, their opposite stays a
        // no-op for the whole window, and an unverifiable slot isn't polled.
        let Some(screen) = read_modes() else {
            eprintln!("SKIP: Personalize modes unreadable - read-back loop not exercised");
            return;
        };
        let other = md(!screen.apps_light, !screen.system_light);
        let missed = slot(Theme::Light, other, true);
        let reached = tier1_readback(&slot(Theme::Light, screen, true), Some(other));
        let noop = tier1_readback(&missed, Some(screen));
        let unverifiable = tier1_readback(&missed, Some(other));
        if read_modes() != Some(screen) {
            eprintln!("SKIP: the theme changed during the test");
            return;
        }
        assert_eq!(reached.0, ReadBack::Reached);
        assert!(reached.1 < 500, "{} ms", reached.1);
        assert_eq!(noop.0, ReadBack::NoOp(screen));
        assert!(noop.1 >= READBACK_WINDOW.as_millis(), "{} ms", noop.1);
        assert_eq!(unverifiable.0, ReadBack::Unverifiable("already-matching"));
        assert!(unverifiable.1 < 100, "{} ms", unverifiable.1);
    }

    // --- tier-1 name matching (v0.5.0) ---

    #[test]
    fn pick_theme_index_table() {
        let n = "Night";
        // Windows' own theme: first match, even with duplicates or a
        // same-named custom file elsewhere.
        assert_eq!(pick_theme_index(&[3], true, false, n), Ok(3));
        assert_eq!(pick_theme_index(&[3, 9], true, true, n), Ok(3));
        assert!(pick_theme_index(&[], true, false, n).is_err());
        // Custom theme: exactly one match and a unique name, or decline.
        assert_eq!(pick_theme_index(&[4], false, false, n), Ok(4));
        assert!(pick_theme_index(&[], false, false, n).is_err());
        assert!(pick_theme_index(&[4, 7], false, false, n).is_err());
        assert!(pick_theme_index(&[4], false, true, n).is_err());
    }

    /// A temp folder of .theme files, removed on drop.
    struct TempDir(PathBuf);

    impl TempDir {
        fn new(name: &str) -> Self {
            let p = std::env::temp_dir().join(format!("wts-test-{}-{name}", std::process::id()));
            let _ = fs::remove_dir_all(&p);
            fs::create_dir_all(&p).unwrap();
            Self(p)
        }
    }

    impl Drop for TempDir {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    #[test]
    fn name_taken_elsewhere_ignores_the_configured_file_itself() {
        let d = TempDir::new("names");
        fs::create_dir_all(d.0.join("pack")).unwrap();
        let mine = d.0.join("Mine.theme");
        fs::write(&mine, b"[Theme]\r\nDisplayName=Night\r\n").unwrap();
        fs::write(d.0.join("Other.theme"), b"[Theme]\r\nDisplayName=Day\r\n").unwrap();
        let dirs = [(d.0.clone(), true)];
        // Only the configured file has the name — spelled differently too.
        assert!(!name_taken_elsewhere(&mine, "Night", &dirs));
        let spelled = PathBuf::from(format!("{}\\.\\MINE.THEME", d.0.display()));
        assert!(!name_taken_elsewhere(&spelled, "Night", &dirs));
        // A saved copy in a subfolder with the same name counts...
        fs::write(
            d.0.join("pack\\Copy.theme"),
            b"[Theme]\r\nDisplayName=Night\r\n",
        )
        .unwrap();
        assert!(name_taken_elsewhere(&mine, "Night", &dirs));
        // ...but not when subfolders aren't searched.
        assert!(!name_taken_elsewhere(
            &mine,
            "Night",
            &[(d.0.clone(), false)]
        ));
        // A no-DisplayName file is named by its stem.
        fs::write(d.0.join("Glow.theme"), b"[Theme]\r\n").unwrap();
        assert!(name_taken_elsewhere(&mine, "Glow", &dirs));
    }

    #[test]
    fn windows_theme_folders_are_recognized() {
        let root = windows_resources_dir();
        let aero = root.join("Themes\\aero.theme");
        if aero.exists() {
            assert!(is_windows_theme(&aero));
            let spelled = PathBuf::from(
                aero.to_string_lossy()
                    .to_ascii_uppercase()
                    .replace("\\THEMES\\", "\\Themes\\.\\"),
            );
            assert!(is_windows_theme(&spelled), "{spelled:?}");
        }
        let hc = root.join("Ease of Access Themes\\hcblack.theme");
        if hc.exists() {
            assert!(is_windows_theme(&hc));
        }
        let t = TempTheme::new("not-windows", b"[Theme]\r\n");
        assert!(!is_windows_theme(&t.0));
    }

    #[test]
    fn unreadable_theme_file_is_none_not_absent() {
        // GetPrivateProfileStringW returns the default when it can't open the
        // file; that must not read as "DisplayName absent" (→ file stem).
        use std::os::windows::fs::OpenOptionsExt;
        let t = TempTheme::new("locked", b"[Theme]\r\nDisplayName=Locked\r\n");
        let _lock = fs::OpenOptions::new()
            .read(true)
            .share_mode(0)
            .open(&t.0)
            .unwrap();
        assert_eq!(read_theme_ini(&t.0, "Theme", "DisplayName"), None);
        assert_eq!(resolve_theme_display_name(&t.0), None);
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
        // A DisplayName outside [Theme] doesn't count — so the theme has no
        // name of its own and Windows lists it under its file stem.
        let t = TempTheme::new(
            "wrong-section",
            b"[Control Panel\\Desktop]\r\nDisplayName=Nope\r\n[Slideshow]\r\nInterval=1\r\n",
        );
        let stem = t.0.file_stem().unwrap().to_string_lossy().into_owned();
        assert_eq!(resolve_theme_display_name(&t.0), Some(stem));
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
    fn theme_display_name_missing_key_falls_back_to_the_file_stem() {
        // What Windows does (verified by enumerating IThemeManager2 with a
        // no-DisplayName theme installed): the name is the file stem, without
        // ".theme". A missing file still has no name.
        let missing =
            std::env::temp_dir().join(format!("wts-test-{}-nonexistent.theme", std::process::id()));
        assert_eq!(resolve_theme_display_name(&missing), None);
        let t = TempTheme::new("no-name", b"[Theme]\r\nColor=1\r\n");
        let stem = t.0.file_stem().unwrap().to_string_lossy().into_owned();
        assert!(!stem.ends_with(".theme"));
        assert_eq!(resolve_theme_display_name(&t.0), Some(stem));
    }

    #[test]
    fn theme_display_name_present_but_empty_is_none() {
        let t = TempTheme::new("empty-name", b"[Theme]\r\nDisplayName=\r\n");
        assert_eq!(resolve_theme_display_name(&t.0), None);
    }

    #[test]
    fn theme_files_on_long_paths_are_read_not_taken_as_keyless() {
        // MAX_PATH (260) or more: neither this test binary nor the app has a
        // longPathAware manifest, so GetPrivateProfileStringW couldn't open
        // the plain path and returned the default — every key read as absent.
        let d = TempDir::new("long-path");
        let deep = d.0.join("a".repeat(120)).join("b".repeat(120));
        fs::create_dir_all(&deep).unwrap();
        let file = deep.join("Long Night.theme");
        assert!(file.as_os_str().len() > 260);
        fs::write(
            &file,
            b"[Theme]\r\nDisplayName=Far Away\r\n[VisualStyles]\r\nSystemMode=Dark\r\nAppMode=Light\r\n",
        )
        .unwrap();
        assert_eq!(
            resolve_theme_display_name(&file).as_deref(),
            Some("Far Away")
        );
        assert_eq!(
            read_theme_ini(&file, "VisualStyles", "AppMode"),
            Some(Some("Light".to_string()))
        );
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

    // --- theme path policy (v0.5.0) ---

    #[test]
    fn theme_paths_expand_env_vars_and_resolve_relative_to_the_exe_folder() {
        let base = PathBuf::from("C:\\Apps\\WTS");
        let windir = std::env::var("SystemRoot").unwrap();
        assert_eq!(
            expand_theme_path("%SystemRoot%\\Resources\\Themes\\dark.theme", &base),
            PathBuf::from(&windir).join("Resources\\Themes\\dark.theme")
        );
        assert_eq!(
            expand_theme_path("themes\\night.theme", &base),
            PathBuf::from("C:\\Apps\\WTS\\themes\\night.theme")
        );
        // An unknown variable is left as written (and then fails as missing).
        assert_eq!(
            expand_theme_path("%WTS_NO_SUCH_VAR%\\x.theme", &base),
            PathBuf::from("C:\\Apps\\WTS\\%WTS_NO_SUCH_VAR%\\x.theme")
        );
        assert_eq!(
            expand_theme_path("C:\\Themes\\day.theme", &base),
            PathBuf::from("C:\\Themes\\day.theme")
        );
    }

    #[test]
    fn theme_path_policy_accepts_only_theme_files() {
        let scratch = std::env::temp_dir().join("wts-test-no-scratch.theme");
        let ok = TempTheme::new("policy-ok", b"[Theme]\r\nDisplayName=Ok\r\n");
        assert_eq!(check_theme_path(&ok.0, &scratch), Ok(ok.0.clone()));
        // Extension match is case-insensitive.
        let upper =
            std::env::temp_dir().join(format!("wts-test-{}-upper.THEME", std::process::id()));
        fs::write(&upper, b"[Theme]\r\n").unwrap();
        assert_eq!(check_theme_path(&upper, &scratch), Ok(upper.clone()));
        let _ = fs::remove_file(&upper);

        let missing =
            std::env::temp_dir().join(format!("wts-test-{}-nope.theme", std::process::id()));
        assert_eq!(
            check_theme_path(&missing, &scratch),
            Err(ThemePathIssue::Missing)
        );
        assert_eq!(
            check_theme_path(&std::env::temp_dir(), &scratch),
            Err(ThemePathIssue::Directory)
        );
        for (ext, issue) in [
            ("themepack", ThemePathIssue::ThemePack),
            ("deskthemepack", ThemePathIssue::ThemePack),
            ("exe", ThemePathIssue::NotATheme),
            ("jpg", ThemePathIssue::NotATheme),
        ] {
            let p =
                std::env::temp_dir().join(format!("wts-test-{}-policy.{ext}", std::process::id()));
            fs::write(&p, b"x").unwrap();
            assert_eq!(check_theme_path(&p, &scratch), Err(issue), "{ext}");
            let _ = fs::remove_file(&p);
        }
        // Windows' scratch Custom.theme is refused even though it's a .theme
        // — however the path is spelled.
        assert_eq!(
            check_theme_path(&ok.0, &ok.0),
            Err(ThemePathIssue::ScratchFile)
        );
        let respelled = PathBuf::from(
            ok.0.to_string_lossy()
                .to_ascii_uppercase()
                .replacen('\\', "\\.\\", 1),
        );
        assert_eq!(
            check_theme_path(&respelled, &ok.0),
            Err(ThemePathIssue::ScratchFile)
        );
        // A folder named like a theme is still a folder; the pack check is
        // case-insensitive.
        let d = TempDir::new("dir.theme");
        assert_eq!(
            check_theme_path(&d.0, &scratch),
            Err(ThemePathIssue::Directory)
        );
        let pack =
            std::env::temp_dir().join(format!("wts-test-{}-x.DeskThemePack", std::process::id()));
        fs::write(&pack, b"x").unwrap();
        assert_eq!(
            check_theme_path(&pack, &scratch),
            Err(ThemePathIssue::ThemePack)
        );
        let _ = fs::remove_file(&pack);
        // Undefined variables and drive-relative paths are named as such.
        assert_eq!(
            check_theme_path(Path::new("C:\\Apps\\%WTS_NO_SUCH_VAR%\\x.theme"), &scratch),
            Err(ThemePathIssue::UnexpandedVar)
        );
        assert_eq!(
            check_theme_path(Path::new("C:x.theme"), &scratch),
            Err(ThemePathIssue::NotAbsolute)
        );
        // ...but an existing file whose name has a %...% pair is just a file.
        let pct = TempTheme::new("50% Dark 50%", b"[Theme]\r\n");
        assert_eq!(check_theme_path(&pct.0, &scratch), Ok(pct.0.clone()));
        // Win32 ignores trailing dots: "x.theme." is the file x.theme — and
        // is used by that name.
        let dotted = PathBuf::from(format!("{}.", ok.0.display()));
        assert_eq!(check_theme_path(&dotted, &scratch), Ok(ok.0.clone()));
    }

    #[test]
    fn a_respelled_theme_path_takes_the_file_name_stored_on_disk() {
        // Win32 opens a trailing-dot, differently cased or 8.3 spelling as
        // the same file, but a theme without a DisplayName is named after
        // its file, and tier 1 matches that name exactly.
        use windows_sys::Win32::Storage::FileSystem::GetShortPathNameW;
        let scratch = std::env::temp_dir().join("wts-test-no-scratch.theme");
        let t = TempTheme::new("Stored Name", b"[Theme]\r\nColor=1\r\n");
        let stored = t.0.file_name().unwrap();
        let shouted =
            t.0.with_file_name(stored.to_string_lossy().to_ascii_uppercase());
        let dotted = PathBuf::from(format!("{}.", t.0.display()));
        for typed in [&shouted, &dotted] {
            assert_eq!(
                check_theme_path(typed, &scratch),
                Ok(t.0.clone()),
                "{typed:?}"
            );
        }
        // The slot carries the stored name, so the file-name fallback names
        // the theme the way Windows lists it.
        let cfg = Config {
            theme_night: Some(shouted.to_string_lossy().into_owned()),
            ..Config::default()
        };
        let slot = resolve_slot(Theme::Dark, &cfg);
        assert_eq!(slot.file, t.0);
        assert_eq!(
            resolve_theme_display_name(&slot.file),
            Some(t.0.file_stem().unwrap().to_string_lossy().into_owned())
        );
        // An 8.3 alias, where the volume makes them.
        let long = wide(&t.0.to_string_lossy());
        let mut buf = [0u16; 1024];
        let n = unsafe { GetShortPathNameW(long.as_ptr(), buf.as_mut_ptr(), buf.len() as u32) }
            as usize;
        let short = PathBuf::from(String::from_utf16_lossy(&buf[..n.min(buf.len())]));
        if n == 0 || n >= buf.len() || short.file_name() == Some(stored) {
            eprintln!("SKIP: no 8.3 name for {}", t.0.display());
        } else {
            assert_eq!(
                check_theme_path(&short, &scratch),
                Ok(short.with_file_name(stored))
            );
        }
    }

    #[test]
    fn a_stored_name_that_isnt_the_same_file_keeps_the_path_as_written() {
        // A symlink `links\Night.theme` -> `targets\Target.theme`, with an
        // unrelated `links\Target.theme` beside the link: the canonical name
        // names a different file there, so the link's own spelling stays.
        // (Simulated with the target's real canonical path — creating a
        // symlink needs a privilege tests can't count on.)
        let d = TempDir::new("stored-name-guard");
        fs::create_dir_all(d.0.join("links")).unwrap();
        fs::create_dir_all(d.0.join("targets")).unwrap();
        let decoy = d.0.join("links\\Target.theme");
        let target = d.0.join("targets\\Target.theme");
        fs::write(&decoy, b"[Theme]\r\nDisplayName=Decoy\r\n").unwrap();
        fs::write(&target, b"[Theme]\r\nDisplayName=Target\r\n").unwrap();
        let link = d.0.join("links\\Night.theme");
        let real = fs::canonicalize(&target).unwrap();
        assert_eq!(with_stored_name(&link, Some(&real)), link);
        assert_eq!(with_stored_name(&link, None), link);
        // The same comparison does respell a path to the very file it names.
        let typed = d.0.join("targets\\target.THEME");
        assert_eq!(with_stored_name(&typed, Some(&real)), target);
    }

    #[test]
    fn long_paths_get_the_verbatim_form_for_win32_file_apis() {
        let short = r"C:\Users\me\Themes\Night.theme";
        assert_eq!(win32_path_arg(short), short);
        let deep = "d".repeat(250);
        let local = format!(r"C:\{deep}\Night.theme");
        assert_eq!(win32_path_arg(&local), format!(r"\\?\{local}"));
        let share = format!(r"\\server\share\{deep}\Night.theme");
        assert_eq!(
            win32_path_arg(&share),
            format!(r"\\?\UNC\server\share\{deep}\Night.theme")
        );
        let verbatim = format!(r"\\?\{local}");
        assert_eq!(win32_path_arg(&verbatim), verbatim);
        // Not absolute: a prefix can't fix that.
        let relative = format!(r"{deep}\Night.theme");
        assert_eq!(win32_path_arg(&relative), relative);
        // The threshold counts UTF-16 units: 246 stay plain, 247 don't.
        let of_len = |n: usize| format!(r"C:\{}", "x".repeat(n - 3));
        assert_eq!(win32_path_arg(&of_len(246)), of_len(246));
        assert_eq!(
            win32_path_arg(&of_len(247)),
            format!(r"\\?\{}", of_len(247))
        );
    }

    #[test]
    fn unexpanded_var_detection() {
        assert!(has_unexpanded_var("C:\\%FOO%\\x.theme"));
        assert!(has_unexpanded_var("%A%"));
        assert!(!has_unexpanded_var("C:\\100%\\x.theme"));
        assert!(!has_unexpanded_var("C:\\a%b\\c%d.theme"));
        assert!(!has_unexpanded_var("C:\\plain\\x.theme"));
        assert!(!has_unexpanded_var("%%"));
    }

    #[test]
    fn rejected_theme_paths_fall_back_to_the_stock_theme_and_say_why() {
        let pack =
            std::env::temp_dir().join(format!("wts-test-{}-night.themepack", std::process::id()));
        fs::write(&pack, b"MSCF").unwrap();
        let cfg = Config {
            theme_night: Some(pack.to_string_lossy().into_owned()),
            ..Config::default()
        };
        let (file, rejected) = resolve_theme_file_checked(Theme::Dark, &cfg);
        let _ = fs::remove_file(&pack);
        assert!(file.ends_with("dark.theme"), "got {file:?}");
        let (_, _, issue) = rejected.expect("rejection reported");
        assert_eq!(issue, ThemePathIssue::ThemePack);
    }

    // --- themes on network locations (v0.6.0) ---

    #[test]
    fn path_root_table() {
        use PathRoot::{Drive, Other, Unc};
        for (path, want) in [
            // Shares, in every spelling Win32 takes for one.
            (r"\\server\share\Night.theme", Unc),
            (r"\\server.example.com\share\a\b\Night.theme", Unc),
            (r"\\192.0.2.7\c$\Night.theme", Unc),
            ("//server/share/Night.theme", Unc),
            (r"\/server\share\Night.theme", Unc),
            (r"\\?\UNC\server\share\Night.theme", Unc),
            (r"\\?\unc\server\share\Night.theme", Unc),
            (r"\\.\UNC\server\share\Night.theme", Unc),
            ("//?/UNC/server/share/Night.theme", Unc),
            // `Global\` behind a device prefix names the same device names
            // again, once or more, in any letter case.
            (r"\\?\Global\UNC\server\share\Night.theme", Unc),
            (r"\\.\Global\UNC\server\share\Night.theme", Unc),
            (r"\\?\GLOBAL\global\unc\server\share\Night.theme", Unc),
            // Local-looking UNC paths are shares too: they go through the
            // redirector all the same.
            (r"\\localhost\c$\Night.theme", Unc),
            (r"\\wsl.localhost\Ubuntu\home\Night.theme", Unc),
            // A drive letter, plain or behind a device prefix; whether it is
            // a network drive is asked of Windows, not of the text.
            (r"C:\Themes\Night.theme", Drive('C')),
            (r"h:\Themes\Night.theme", Drive('H')),
            ("C:/Themes/Night.theme", Drive('C')),
            (r"\\?\C:\Themes\Night.theme", Drive('C')),
            (r"\\?\z:\Night.theme", Drive('Z')),
            (r"\\.\D:\Night.theme", Drive('D')),
            (r"C:Night.theme", Drive('C')),
            (r"\\?\Global\C:\Themes\Night.theme", Drive('C')),
            (r"\\.\Global\Global\h:\Night.theme", Drive('H')),
            // Other device names are not taken for shares: a volume GUID is
            // a local volume, and most of the rest aren't files a theme
            // could be.
            (
                r"\\?\Volume{4c1b02c1-d990-11dc-99ae-806e6f6e6963}\Night.theme",
                Other,
            ),
            (r"\\.\pipe\Night.theme", Other),
            (r"\\.\PhysicalDrive0", Other),
            (r"\\?\UNC", Other),
            (r"\\?\UNCLE\Night.theme", Other),
            (r"\\?\Global", Other),
            (r"\\?\Globals\UNC\server\share\Night.theme", Other),
            // A known gap, pinned here so it isn't mistaken for a decision
            // that these are local: GLOBALROOT is the root of every device
            // name, and this one IS the redirector. Where a device name
            // leads isn't in the text (README, Known limitations).
            (r"\\?\GLOBALROOT\Device\Mup\server\share\Night.theme", Other),
            (r"\\?\GLOBALROOT\Device\HarddiskVolume3\Night.theme", Other),
            // No drive at all (never reaches the policy: a relative path is
            // joined to the exe's folder first).
            (r"\Themes\Night.theme", Other),
            (r"Themes\Night.theme", Other),
            (r"1:\Night.theme", Other),
            ("é:\\Night.theme", Other),
            ("", Other),
        ] {
            assert_eq!(path_root(path), want, "{path:?}");
        }
    }

    #[test]
    fn network_drives_are_told_by_drive_type() {
        use windows_sys::Win32::Storage::FileSystem::{GetLogicalDrives, QueryDosDeviceW};
        // The drive Windows is on is not a network drive.
        let system = std::env::var("SystemRoot").unwrap();
        let system = system.chars().next().unwrap().to_ascii_uppercase();
        assert!(!drive_is_remote(system), "{system}:");
        // Neither is a letter that isn't a drive.
        let present = unsafe { GetLogicalDrives() };
        let is_drive = |l: &char| present & (1 << (*l as u32 - 'A' as u32)) != 0;
        match ('A'..='Z').find(|l| !is_drive(l)) {
            Some(absent) => assert!(!drive_is_remote(absent), "{absent}:"),
            None => eprintln!("SKIP: every drive letter is in use"),
        }
        // A mapped network drive is one — where this PC has any. What a
        // drive letter stands for is read from Windows' table of device
        // names (nothing is opened): `net use` links the letter to the
        // redirector. Elsewhere (CI) nothing here can show a "yes".
        let target = |letter: char| {
            let mut buf = [0u16; 1024];
            let name = wide(&format!("{letter}:"));
            let n = unsafe { QueryDosDeviceW(name.as_ptr(), buf.as_mut_ptr(), buf.len() as u32) };
            let text = &buf[..n as usize];
            let end = text.iter().position(|&c| c == 0).unwrap_or(text.len());
            String::from_utf16_lossy(&text[..end])
        };
        let mapped: Vec<char> = ('A'..='Z')
            .filter(is_drive)
            .filter(|l| target(*l).starts_with(r"\Device\LanmanRedirector\"))
            .collect();
        if mapped.is_empty() {
            eprintln!("SKIP: no mapped network drive on this PC");
        }
        let scratch = std::env::temp_dir().join("wts-test-no-scratch.theme");
        for letter in mapped {
            assert!(drive_is_remote(letter), "{letter}:");
            assert!(drive_is_remote(letter.to_ascii_lowercase()), "{letter}:");
            // And the path check asks the real question: a theme on that
            // drive is refused, not looked for.
            let theme = PathBuf::from(format!(r"{letter}:\wts-test-no-such-folder\Night.theme"));
            assert_eq!(
                check_theme_path(&theme, &scratch),
                Err(ThemePathIssue::NetworkPath),
                "{theme:?}"
            );
        }
        // Every letter, lower case included, in far less time than one
        // unreachable server costs. This can only fail on a PC that has a
        // mapped drive whose server is slow to answer, if the question
        // ever starts waiting for it.
        let started = Instant::now();
        for letter in ('A'..='Z').chain('a'..='z') {
            let _ = drive_is_remote(letter);
        }
        let took = started.elapsed();
        assert!(
            took < Duration::from_secs(1),
            "52 drive letters took {took:?}"
        );
        // The drive question reaches the answer for a path through its root.
        let asked = std::cell::Cell::new(None);
        let ask = |l| {
            asked.set(Some(l));
            true
        };
        assert!(on_network(Path::new(r"\\?\n:\Themes\x.theme"), ask));
        assert_eq!(asked.get(), Some('N'));
        assert!(!on_network(Path::new(r"N:\Themes\x.theme"), |_| false));
        // A share is one whatever the drives are; another device name never.
        assert!(on_network(Path::new(r"\\server\share\x.theme"), |_| false));
        assert!(!on_network(Path::new(r"\\.\pipe\x.theme"), |_| true));
    }

    #[test]
    fn a_network_path_is_read_the_way_win32_opens_it() {
        // Win32 resolves `.`, `..`, repeated separators and a trailing dot
        // in every form but `\\?\` before it opens a path, and behind `\\.\`
        // a `..` climbs out of the drive letter: each of these opens
        // `\\.\UNC\server\share\x.theme` (seen with GetFullPathNameW, and
        // with a lookup of a host name that can't resolve, 2026-11-11).
        // No drive is a network drive here, so only the text decides.
        for share in [
            r"\\.\C:\..\UNC\server\share\x.theme",
            "//./C:/../UNC/server/share/x.theme",
            r"\\.\C:\Themes\..\..\UNC\server\share\x.theme",
            r"\\.\..\UNC\server\share\x.theme",
            r"\\.\\UNC\server\share\x.theme",
            r"\\.\.\UNC\server\share\x.theme",
            r"\\.\UNC.\server\share\x.theme",
            r"\\.\C:\..\Global\UNC\server\share\x.theme",
            // `//?/` with forward slashes is not the as-written form.
            "//?/C:/../UNC/server/share/x.theme",
        ] {
            assert!(on_network(Path::new(share), |_| false), "{share:?}");
        }
        // The drive asked about is the one the path ends up on.
        for (path, drive) in [
            (r"\\.\C:\..\N:\Themes\x.theme", 'N'),
            (r"C:\Themes\..\x.theme", 'C'),
            (r"c:/Themes/./x.theme", 'C'),
            // `\\?\` is opened as written: `..` is a name there, on C:.
            (r"\\?\C:\..\UNC\server\share\x.theme", 'C'),
        ] {
            let asked = std::cell::Cell::new(None);
            let remote = on_network(Path::new(path), |l| {
                asked.set(Some(l));
                false
            });
            assert!(!remote, "{path:?}");
            assert_eq!(asked.get(), Some(drive), "{path:?}");
        }
        // Not recognised (README, Known limitations): a device name that
        // leads to the redirector.
        assert!(!on_network(
            Path::new(r"\\?\GLOBALROOT\Device\Mup\server\share\x.theme"),
            |_| true
        ));
        // Through the path check: refused as a share, nothing looked up.
        let scratch = std::env::temp_dir().join("wts-test-no-scratch.theme");
        let host = format!("wts-no-such-host-{}-dots.invalid", std::process::id());
        for share in [
            format!(r"\\.\C:\..\UNC\{host}\share\Night.theme"),
            format!(r"\\?\Global\UNC\{host}\share\Night.theme"),
        ] {
            assert_eq!(
                check_theme_path(Path::new(&share), &scratch),
                Err(ThemePathIssue::NetworkPath),
                "{share:?}"
            );
        }
    }

    #[test]
    fn theme_paths_on_a_network_location_are_refused_before_anything_opens_them() {
        let scratch = std::env::temp_dir().join("wts-test-no-scratch.theme");
        // A share on a host name that can't resolve (.invalid is reserved),
        // a fresh name each time so no lookup result is cached. Reading the
        // text takes microseconds; asking for the host took 120-145 ms per
        // name on the PC this was written on. The fastest of four is
        // compared, so one stalled thread can't fail the test. (How long a
        // failed lookup takes depends on the PC's name resolver; the cases
        // with a drive that answers "network", further down, pin the order
        // without a clock.)
        let mut fastest = Duration::MAX;
        for i in 0..4 {
            let host = format!("wts-no-such-host-{}-{i}.invalid", std::process::id());
            let share = PathBuf::from(format!(r"\\{host}\share\Night.theme"));
            let started = Instant::now();
            let verdict = check_theme_path(&share, &scratch);
            fastest = fastest.min(started.elapsed());
            assert_eq!(verdict, Err(ThemePathIssue::NetworkPath), "{share:?}");
        }
        assert!(
            fastest < Duration::from_millis(20),
            "a share was looked up before being refused: {fastest:?}"
        );
        // The same through the config: the stock theme is used, and the
        // rejection carries the reason, the text as configured and its
        // expansion.
        let raw = r"\\wts-no-such-host.invalid\share\Night.theme";
        let cfg = Config {
            theme_night: Some(raw.to_string()),
            ..Config::default()
        };
        let (file, rejected) = resolve_theme_file_checked(Theme::Dark, &cfg);
        assert_eq!(file, stock_theme_file(Theme::Dark));
        let (as_configured, expanded, issue) = rejected.expect("rejection reported");
        assert_eq!(as_configured, raw);
        assert_eq!(expanded, PathBuf::from(raw));
        assert_eq!(issue, ThemePathIssue::NetworkPath);
        assert_eq!(issue.as_str(), "network-path");
        // The dialog text says what to do about it.
        assert!(issue.describe().contains("copy the .theme file"));
        // The other spellings of a share, and a folder, a pack or an
        // undefined variable on one: the verdict is the network location,
        // not what a look at the path would have found.
        for share in [
            r"\\?\UNC\wts-no-such-host.invalid\share\Night.theme",
            r"\\.\UNC\wts-no-such-host.invalid\share\Night.theme",
            "//wts-no-such-host.invalid/share/Night.theme",
            r"\\wts-no-such-host.invalid\share",
            r"\\wts-no-such-host.invalid\share\Night.themepack",
            r"\\wts-no-such-host.invalid\share\%WTS_NO_SUCH_VAR%\Night.theme",
        ] {
            assert_eq!(
                check_theme_path(Path::new(share), &scratch),
                Err(ThemePathIssue::NetworkPath),
                "{share:?}"
            );
        }
        // An exe that itself runs from a share: its relative theme paths
        // are on the share too.
        let beside_exe = expand_theme_path(
            r"themes\Night.theme",
            Path::new(r"\\wts-no-such-host.invalid\apps\WTS"),
        );
        assert_eq!(
            beside_exe,
            PathBuf::from(r"\\wts-no-such-host.invalid\apps\WTS\themes\Night.theme")
        );
        assert_eq!(
            check_theme_path(&beside_exe, &scratch),
            Err(ThemePathIssue::NetworkPath)
        );
        // A network drive: the same file is refused when its drive is one,
        // and used when it isn't. (The drive answer is a parameter; no test
        // can count on a mapped drive.)
        let ok = TempTheme::new("network-drive", b"[Theme]\r\nDisplayName=Net\r\n");
        let PathRoot::Drive(temp_drive) = path_root(&ok.0.to_string_lossy()) else {
            panic!("temp dir has no drive letter: {:?}", ok.0);
        };
        assert_eq!(
            check_theme_path_with(&ok.0, &scratch, |l| l == temp_drive),
            Err(ThemePathIssue::NetworkPath)
        );
        assert_eq!(
            check_theme_path_with(&ok.0, &scratch, |_| false),
            Ok(ok.0.clone())
        );
        let verbatim = PathBuf::from(format!(r"\\?\{}", ok.0.display()));
        assert_eq!(
            check_theme_path_with(&verbatim, &scratch, |l| l == temp_drive),
            Err(ThemePathIssue::NetworkPath)
        );
        // The order, without a clock: on a drive that answers "network", a
        // folder that exists and a file that doesn't are both refused as a
        // network path. A check that looked first would say `directory`
        // and `missing`.
        let folder = std::env::temp_dir();
        let absent = folder.join(format!("wts-test-{}-absent.theme", std::process::id()));
        assert!(folder.is_dir() && !absent.exists());
        for path in [&folder, &absent] {
            assert_eq!(
                check_theme_path_with(path, &scratch, |_| true),
                Err(ThemePathIssue::NetworkPath),
                "{path:?}"
            );
        }
        assert_eq!(
            check_theme_path_with(&folder, &scratch, |_| false),
            Err(ThemePathIssue::Directory)
        );
        assert_eq!(
            check_theme_path_with(&absent, &scratch, |_| false),
            Err(ThemePathIssue::Missing)
        );
        // A drive-relative path is still named as that, drive type unasked.
        assert_eq!(
            check_theme_path_with(Path::new("N:x.theme"), &scratch, |_| {
                panic!("the drive was asked about")
            }),
            Err(ThemePathIssue::NotAbsolute)
        );
    }

    // --- single instance (v0.6.0) ---

    #[test]
    fn single_instance_claim_table() {
        use InstanceClaim::{Duplicate, DuplicateNoAccess, First, Unknown};
        const ERROR_INVALID_HANDLE: u32 = 6;
        for (handle_is_null, last_error, want, what) in [
            (false, 0, First, "created"),
            (
                false,
                ERROR_ALREADY_EXISTS,
                Duplicate,
                "opened another copy's",
            ),
            (
                true,
                ERROR_ACCESS_DENIED,
                DuplicateNoAccess,
                "another copy's, and it can't be opened",
            ),
            (
                true,
                ERROR_INVALID_HANDLE,
                Unknown(ERROR_INVALID_HANDLE),
                "the name belongs to another kind of object",
            ),
            (true, 0, Unknown(0), "no handle and no error"),
            // Only the handle makes a claim or a duplicate.
            (
                true,
                ERROR_ALREADY_EXISTS,
                Unknown(ERROR_ALREADY_EXISTS),
                "exists, but no handle",
            ),
            (
                false,
                ERROR_ACCESS_DENIED,
                First,
                "a stale error beside a handle",
            ),
        ] {
            assert_eq!(instance_claim(handle_is_null, last_error), want, "{what}");
        }
        // The second look at a refused mutex: taken unless the name is gone.
        for (open_is_null, last_error, taken, what) in [
            (false, 0, true, "opened"),
            (true, ERROR_ACCESS_DENIED, true, "refused again"),
            (true, ERROR_INVALID_HANDLE, true, "another kind of object"),
            (true, ERROR_FILE_NOT_FOUND, false, "no such name"),
            (
                false,
                ERROR_FILE_NOT_FOUND,
                true,
                "a stale error beside a handle",
            ),
        ] {
            assert_eq!(mutex_name_taken(open_is_null, last_error), taken, "{what}");
        }
        // What main() does with a claim: the exact log text for the two
        // that end this copy (the ordinary one unchanged since v0.4.0),
        // nothing for the two that let it run.
        assert_eq!(First.duplicate_log(), None);
        assert_eq!(Unknown(ERROR_ACCESS_DENIED).duplicate_log(), None);
        assert_eq!(
            Duplicate.duplicate_log(),
            Some("duplicate_instance action=exit")
        );
        assert_eq!(
            DuplicateNoAccess.duplicate_log(),
            Some("duplicate_instance action=exit reason=access-denied")
        );
    }

    #[test]
    fn single_instance_claim_sequences() {
        use InstanceClaim::{Duplicate, DuplicateNoAccess, First, Unknown};
        /// One call's result: (no handle came back, the last error).
        type Call = (bool, u32);
        const DENIED: Call = (true, ERROR_ACCESS_DENIED);
        const NOT_FOUND: Call = (true, ERROR_FILE_NOT_FOUND);
        const OPENED: Call = (false, 0);
        const CREATED: Call = (false, 0);
        const EXISTS: Call = (false, ERROR_ALREADY_EXISTS);
        // Each row: what CreateMutexW returns call by call, what the
        // OpenMutexW probe returns call by call, the claim. Every scripted
        // result must be used, and no call may go beyond the script.
        let rows: [(&[Call], &[Call], InstanceClaim, &str); 10] = [
            (&[CREATED], &[], First, "created: no probe"),
            (&[EXISTS], &[], Duplicate, "opened: no probe"),
            (&[(true, 6)], &[], Unknown(6), "another failure: no probe"),
            (
                &[DENIED],
                &[OPENED],
                DuplicateNoAccess,
                "refused, and the mutex is there (a copy with fewer rights)",
            ),
            (
                &[DENIED],
                &[DENIED],
                DuplicateNoAccess,
                "refused, and refused again for the least access",
            ),
            (
                &[DENIED, CREATED],
                &[NOT_FOUND],
                First,
                "the other copy left between the two calls: this one holds the mutex now",
            ),
            (
                &[DENIED, EXISTS],
                &[NOT_FOUND],
                Duplicate,
                "it left, and a third copy got there first",
            ),
            (
                &[DENIED, DENIED],
                &[NOT_FOUND, OPENED],
                DuplicateNoAccess,
                "it left, and another copy this one can't open took the name",
            ),
            (
                &[DENIED, DENIED],
                &[NOT_FOUND, NOT_FOUND],
                Unknown(ERROR_ACCESS_DENIED),
                "nothing there, twice: creating is what is refused, and the app runs",
            ),
            (
                &[DENIED, (true, 6)],
                &[NOT_FOUND],
                Unknown(6),
                "the second round fails another way",
            ),
        ];
        for (creates, probes, want, what) in rows {
            let (created, probed) = (std::cell::Cell::new(0), std::cell::Cell::new(0));
            let next = |script: &[Call], at: &std::cell::Cell<usize>| {
                let i = at.get();
                at.set(i + 1);
                *script
                    .get(i)
                    .unwrap_or_else(|| panic!("{what}: call {} is one too many", i + 1))
            };
            let got = settle_instance_claim(|| next(creates, &created), || next(probes, &probed));
            assert_eq!(got, want, "{what}");
            assert_eq!(created.get(), creates.len(), "{what}: creates");
            assert_eq!(probed.get(), probes.len(), "{what}: probes");
        }
    }

    #[test]
    fn a_mutex_this_process_may_not_open_counts_as_another_copy() {
        // Against Windows itself, on throwaway names (never the app's own).
        use windows_sys::Win32::Security::{
            AddAccessAllowedAce, InitializeAcl, InitializeSecurityDescriptor,
            SetSecurityDescriptorDacl, ACL, ACL_REVISION, SECURITY_ATTRIBUTES, SECURITY_DESCRIPTOR,
            SID, SID_IDENTIFIER_AUTHORITY,
        };
        let name = |tag: &str| format!("Local\\WinThemeSwitcherTest-{}-{tag}", std::process::id());
        // A name nobody made can't be opened, and Windows says so with
        // "not found" (2) — the one answer that means nothing is there.
        assert_eq!(probe_mutex(&wide(&name("never-made"))), (true, 2));
        // The first claim creates the mutex, the second finds it.
        assert_eq!(claim_instance_mutex(&name("plain")), InstanceClaim::First);
        assert_eq!(
            claim_instance_mutex(&name("plain")),
            InstanceClaim::Duplicate
        );
        assert!(!probe_mutex(&wide(&name("plain"))).0);
        // Make a mutex whose access list is given: no entry at all, or one
        // that allows Everyone (S-1-1-0) the rights in `everyone`.
        let create_with = |tag: &str, everyone: Option<u32>| {
            // Room for the list's header and one entry.
            let mut acl_buf = [0u32; 16];
            let acl = acl_buf.as_mut_ptr() as *mut ACL;
            let mut world = SID {
                Revision: 1,
                SubAuthorityCount: 1,
                IdentifierAuthority: SID_IDENTIFIER_AUTHORITY {
                    Value: [0, 0, 0, 0, 0, 1],
                },
                SubAuthority: [0],
            };
            let mut sd: SECURITY_DESCRIPTOR = unsafe { std::mem::zeroed() };
            let sd_ptr = &mut sd as *mut SECURITY_DESCRIPTOR as *mut c_void;
            unsafe {
                let size = std::mem::size_of_val(&acl_buf) as u32;
                assert_ne!(InitializeAcl(acl, size, ACL_REVISION), 0);
                if let Some(rights) = everyone {
                    let sid = &mut world as *mut SID as *mut c_void;
                    assert_ne!(AddAccessAllowedAce(acl, ACL_REVISION, rights, sid), 0);
                }
                assert_ne!(InitializeSecurityDescriptor(sd_ptr, 1), 0);
                assert_ne!(SetSecurityDescriptorDacl(sd_ptr, 1, acl, 0), 0);
                let attrs = SECURITY_ATTRIBUTES {
                    nLength: std::mem::size_of::<SECURITY_ATTRIBUTES>() as u32,
                    lpSecurityDescriptor: sd_ptr,
                    bInheritHandle: 0,
                };
                let holder = CreateMutexW(&attrs, 0, wide(&name(tag)).as_ptr());
                assert_ne!(holder as usize, 0, "creating the {tag} mutex");
                holder
            }
        };
        // What a copy started with fewer rights than the first sees: the
        // full access CreateMutexW asks for is refused (no handle, access
        // denied), while the mutex can still be opened to wait on.
        let holder = create_with("wait-only", Some(SYNCHRONIZATION_SYNCHRONIZE));
        assert_eq!(
            claim_instance_mutex(&name("wait-only")),
            InstanceClaim::DuplicateNoAccess
        );
        assert!(!probe_mutex(&wide(&name("wait-only"))).0);
        unsafe { CloseHandle(holder) };
        // A mutex whose access list lets nobody in: refused both ways, and
        // still another copy's.
        let holder = create_with("locked", None);
        assert_eq!(
            probe_mutex(&wide(&name("locked"))),
            (true, ERROR_ACCESS_DENIED)
        );
        assert_eq!(
            claim_instance_mutex(&name("locked")),
            InstanceClaim::DuplicateNoAccess
        );
        // Once its holder is gone the name is free again.
        unsafe { CloseHandle(holder) };
        assert_eq!(claim_instance_mutex(&name("locked")), InstanceClaim::First);
    }
}
