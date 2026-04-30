# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Windows tray app (Rust) that swaps the full Windows **theme** (wallpaper + colors + light/dark mode) at local sunrise/sunset — macOS's auto-theme behavior, on Win11. Primary apply path is the `IThemeManager2` COM interface (the same one the Settings UWP wraps internally) for atomic, in-process theme apply; a two-tier fallback (legacy `ShellExecute(.theme)` → registry-only DWORD toggle) handles the case where the COM interface errors. ~330 KB single-exe, signed Authenticode, no installer.

## Source tree vs deployed binary — read first

The source tree (`C:\Users\atef\Documents\Projects\WinThemeSwitcher\`) is kept for future tweaks. **The actually-running binary lives elsewhere**:

```
C:\Tools\WinThemeSwitcher\
├── win-theme-switcher.exe   ← auto-starts at login
└── config.json              ← user's Riyadh coords, auto_start: true
```

`HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WinThemeSwitcher` points at `"C:\Tools\WinThemeSwitcher\win-theme-switcher.exe"`. After any rebuild, copy the fresh binary over the `C:\Tools\` one — otherwise the next login still launches the old build:

```powershell
Get-Process -Name win-theme-switcher -EA SilentlyContinue | Stop-Process -Force
Copy-Item `
  "C:\Users\atef\Documents\Projects\WinThemeSwitcher\target\release\win-theme-switcher.exe" `
  "C:\Tools\WinThemeSwitcher\win-theme-switcher.exe" -Force
Start-Process "C:\Tools\WinThemeSwitcher\win-theme-switcher.exe"
```

The `win-theme-switcher.exe` at the repo root is the pre-Cargo Manus build (2.6 MB, 60s polling, no theme-file support). Historical; ignore.

## Kaspersky false positive — critical context

**Resolved as of the IThemeManager2 + code-signing migration** — both root causes of the AV friction were addressed simultaneously. The signed binary (`CN=WinThemeSwitcher Self-Signed` cert in `Cert:\CurrentUser\Root` and `…\TrustedPublisher`) collapses the Authenticode-trust signal, and tier-1 theme apply via `IThemeManager2::SetCurrentTheme` removes the `HWND_BROADCAST WM_SETTINGCHANGE` + direct `WM_THEMECHANGED` signals that previously tripped behavior heuristics. **Sign every release build** (see Build section above) — unsigned builds will resurrect the issue. Everything below is preserved as historical context for unsigned-build scenarios; the current signed build should not need any of it.

### Historical: pre-signing trust setup

The unsigned binary tripped `VHO:Trojan.Win32.Agent.gen` (Rust exe with no Authenticode signature + `HKCU\Run` writes + `HWND_BROADCAST` of `WM_SETTINGCHANGE` + direct `WM_THEMECHANGED` to `Shell_TrayWnd` + WinRT Geolocation = every AV heuristic signal). Plain **path-based exclusions were insufficient** — Kaspersky's Behavior Detection quarantined regardless. The pre-signing workaround was a **Trusted Applications rule** (Kaspersky Settings → Security → Threats and Exclusions → Specify trusted applications) with all checkboxes ticked: Do not scan opened files, Do not monitor application activity, Do not inherit restrictions, Do not monitor child application activity, Allow interaction with Kaspersky interface.

Rules are **path-based**, so two currently exist:
1. `C:\Tools\WinThemeSwitcher\win-theme-switcher.exe` (the deployed binary — stable).
2. `C:\Users\atef\Documents\Projects\WinThemeSwitcher\target\release\win-theme-switcher.exe` (the build output — rewritten by each `cargo build`).

If a future rebuild gets quarantined anyway (the hash changes and Kaspersky occasionally re-evaluates): drop a 0-byte placeholder at the path first (`Set-Content -Path ... -Value "" -Encoding Byte -Force`), re-add the trust rule while the placeholder exists, then rebuild. Same trick works for new deployment paths.

### When the trust rule is not enough (KSN cloud verdict)

Trusted Applications rules cover File Anti-Virus + Behavior Detection but **not Kaspersky Security Network (KSN) cloud reputation**. KSN can independently flag a fresh hash and pop a hostile two-button modal — *"Disinfect and restart"* / *"Try to disinfect without computer restart"* — with no Skip / Esc / X dismiss option. Both buttons quarantine. Adding a Threats and Exclusions entry mid-modal does **not** clear the in-progress verdict — Kaspersky finishes quarantining anyway, and even subsequent rebuilds can be flagged by `svchost.exe` (the indexer-style scanner running as `NT AUTHORITY\NETWORK SERVICE`) before the popup re-appears.

The reliable workaround for a deploy session is to **right-click the tray K → Pause protection → 15 minutes**, then immediately copy + launch within that window. Once the process is loaded into memory it survives even after protection resumes (Windows holds the file handle; on-disk re-detection won't kill the running PID). The autostart re-creation in `set_auto_start(true)` runs each launch, so even if Kaspersky deletes the `HKCU\Run` value during a quarantine event, the next launch restores it.

**Both long-term fixes are now in place** — see the resolution note at the top of this section. Code-signing landed via `New-SelfSignedCertificate` + signtool (cert in Root + TrustedPublisher). Tier-1 theme apply landed via the `IThemeManager2` COM interface (CLSID `{9324da94-50ec-4a14-a770-e90ca03e7c8f}`). The legacy paths and the trust-rule + KSN documentation in this section are kept for the contingency where the cert expires, the signtool step is skipped, or someone strips the signature.

## Build

Rust toolchain is at `%USERPROFILE%\.cargo\bin\` (rustup, not on global PATH). Use the full path:

```powershell
& "$env:USERPROFILE\.cargo\bin\cargo.exe" build --release `
    --manifest-path "C:\Users\atef\Documents\Projects\WinThemeSwitcher\Cargo.toml"
```

Default toolchain is `stable-x86_64-pc-windows-msvc` (MSVC Build Tools required; the GNU toolchain's bundled linker/dlltool was broken on this machine). Release profile: `opt-level = "z"`, `lto = true`, `codegen-units = 1`, `panic = "abort"`, `strip = true`. Output ~330 KB. No `build.rs` — `windows-sys` and `windows` self-link.

### Sign every release build

A self-signed Authenticode cert (`CN=WinThemeSwitcher Self-Signed`) is installed in `Cert:\CurrentUser\My`, `Cert:\CurrentUser\Root`, and `Cert:\CurrentUser\TrustedPublisher`. Exported pfx at `%LOCALAPPDATA%\WinThemeSwitcher\signing\winthemeswitcher-signing.pfx` (password: `wts-local-signing` — local-only, not a secret worth protecting). Sign every fresh build before deploy:

```powershell
& "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe" sign `
    /f "$env:LOCALAPPDATA\WinThemeSwitcher\signing\winthemeswitcher-signing.pfx" `
    /p "wts-local-signing" /fd SHA256 `
    "C:\Users\atef\Documents\Projects\WinThemeSwitcher\target\release\win-theme-switcher.exe"
```

This collapses the Kaspersky heuristic signal — signed builds pass without tripping the AV-pause dance documented below. If the cert ever needs regenerating: `New-SelfSignedCertificate -Type CodeSigning -Subject "CN=WinThemeSwitcher Self-Signed" -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(10)` and re-add to Root + TrustedPublisher stores. Signature has no countersigned timestamp — it expires when the cert does (10 years out).

## Architecture — `src/main.rs`

Single file, ~1100 lines, event-driven, no polling. Logs every state transition to `events.log` next to the exe (rotated to `events.log.old` past 256 KB).

### 1. Theme apply — three-tier fallback in `apply_theme`

Tiered worst-case-degradation: each tier is more invasive but less reliable than the one above. `apply_theme` walks them top-down, returning a `&'static str` tag for the tier that succeeded (logged in the `applied=` field of the `cause=...` line).

#### Tier 1: `IThemeManager2` (preferred — `applied=theme-manager2`)

Undocumented-but-stable COM interface in `themeui.dll` that the Settings UWP itself wraps. CLSID `{9324da94-50ec-4a14-a770-e90ca03e7c8f}`, IID `{c1e8c83e-845d-4d95-81db-e283fdffc000}`. Vtable layout in the `IThemeManager2Vtbl` struct at the top of `main.rs`.

Flow (`apply_via_theme_manager2`):
1. Resolve the target `.theme` file's `[Theme]\nDisplayName=...` value. System themes use SHLoadIndirectString-style refs (`@%SystemRoot%\System32\themeui.dll,-2060`); literal strings work too. `resolve_theme_display_name` parses the INI, `resolve_indirect_string` calls `SHLoadIndirectString`. For `dark.theme` → `"Windows (dark)"`; for `aero.theme` → `"Windows (light)"`.
2. `CoCreateInstance(CLSID_THEME_MANAGER2)` + `Init(0)`.
3. Enumerate via `GetThemeCount` + `GetTheme(i)` + `ITheme::GetDisplayName(&BSTR)` until a name match. Free each BSTR with `SysFreeString`. **Don't cache the index across launches** — enumeration order is not stable.
4. `SetCurrentTheme(NULL, idx, apply_now=1, apply_flags=NO_HOURGLASS, pack_flags=0)`. This is the only tier-1 call that applies; it does the WM_THEMECHANGED + WM_SETTINGCHANGE broadcasts internally.

Why this is the primary path: ShellExecuteW(`.theme`) silently fails when the user isn't actively interactive (post-WTS_SESSION_UNLOCK, ResumeTimeReached while away, scheduled while no foreground UI). The UWP activation pipeline swallows the apply request — Settings flashes briefly but never commits. `IThemeManager2` is in-process, has no UI dependency, and is what every serious tool uses (AutoDarkMode, wtheme, etc.). Apply latency is ~200 ms vs. the ~5 s poll-then-fail of the legacy path.

**STA threading is mandatory** for this interface ("Shell crap is always STA" per AutoDarkMode source). The main thread already calls `CoInitializeEx(None, COINIT_APARTMENTTHREADED)` at startup; tier-1 apply runs from winit event handlers on that same thread, which is correct. **Never call from a worker thread** without CoInitializeEx(STA) on it first — you'll get RPC_E_WRONG_THREAD or silent corruption.

#### Tier 2: `ShellExecuteW(.theme)` + `commit_watcher` (legacy — `applied=theme-file`)

Fires only if tier 1 errors out (logged as `theme_manager2_err msg="…"`). Same as the original implementation: `ShellExecuteW("open", <.theme path>, ..., SW_HIDE)` to launch the Themes UWP, plus `start_settings_closer` thread to `PostMessage(WM_CLOSE)` the Settings window once it appears, plus a 300 ms sleep + `poke_shell` (taskbar repaint).

**`commit_watcher` is the safety net for tier 2's silent-fail mode**: spawns a thread that polls `current_theme()` every 200 ms for 5 s. If the registry never matches the target → logs `commit_timeout target=…` and **falls through to tier 3 from inside the watcher thread** — writes the registry directly, broadcasts, pokes shell, polls again to confirm, logs `fallback_registry target=… confirmed=true after_ms=…`. Without this, tier 2's silent-fail leaves the user stuck (e.g. sunset fires, ShellExecute reports success, registry stays light, no recovery).

If tier 1 is healthy this path is rarely entered. It exists as backup in case future Windows builds break the COM interface.

#### Tier 3: registry-only (last resort — `applied=registry`)

`write_theme_registry` writes `AppsUseLightTheme` + `SystemUsesLightTheme`, broadcasts `WM_SETTINGCHANGE("ImmersiveColorSet")` to `HWND_BROADCAST`, calls `poke_shell`. **Flips light/dark mode but not wallpaper.** Hit when the `.theme` file is missing entirely, or when reached as the commit_watcher fallback.

**`poke_shell`** sends `WM_THEMECHANGED` + targeted `WM_SETTINGCHANGE("ImmersiveColorSet")` to `Shell_TrayWnd` and `Shell_SecondaryTrayWnd`, then `DwmFlush()`. Required for tiers 2 and 3 — `IThemeManager2::SetCurrentTheme` does the broadcast internally so tier 1 doesn't need it. If future Win versions add new taskbar window classes, extend the list.

**Theme file resolution** (`resolve_theme_file`): if `config.theme_day` / `theme_night` is a valid path, use it; otherwise fall back to system defaults at `%SystemRoot%\Resources\Themes\aero.theme` (light) / `dark.theme` (dark). Custom user themes work with tier 1 only if they're already registered with Windows (i.e. installed via Settings → Themes). Otherwise tier 1 errors with `no installed theme matches DisplayName "…"` and tier 2 takes over.

### 2. Event loop — only tick on specific events

The run closure must **not** call `tick()` on every event. An earlier version did, and the app fought the user's manual theme changes: they'd set Dark in Settings → Windows broadcasts `WM_SETTINGCHANGE` → winit delivers an event → our closure called `tick()` → saw `current != target`, flipped back to Light → user saw "Settings won't stay on Dark". Current behavior only ticks on:

- `Event::NewEvents(StartCause::Init)` — first event after launch.
- `Event::NewEvents(StartCause::ResumeTimeReached { .. })` — scheduled sunrise/sunset fired.
- `Event::UserEvent(AppEvent::Menu(refresh_id))` — user clicked Refresh.

Everything else is `_ => {}`. This matches macOS behavior: manual overrides persist until the next natural transition. `ControlFlow::WaitUntil(deadline)` is set once per tick and sticks across unrelated events (no need to re-set on WaitCancelled).

**State-aware apply**: `tick` calls `apply_theme` only if `current_theme() != target`. **Refresh bypasses this check** and always force-applies — needed so config edits (e.g., user points `theme_night` at a new file) take effect without waiting for the next transition.

### 3. Location (WinRT Geolocation)

`try_get_windows_location()` uses the `windows` crate's `Geolocator::RequestAccessAsync().get()` → `GetGeopositionAsync().get()`. Blocking, but <1 s with a cached location. Called **before** `event_loop.run`, so the tray icon doesn't appear until location is known.

On failure (service off / permission denied): `ask_enable_location()` MessageBox (Yes/No). Yes → `ShellExecute("ms-settings:privacy-location")` + info MessageBox telling user to enable and click Refresh. No → `show_manual_setup_prompt` opens `config.json` in Notepad. Refresh handler silently retries WinRT if location is still empty, so enabling Location Services and clicking Refresh unblocks without a restart.

**COM init matters**: `ensure_com_initialized` → `CoInitializeEx(None, COINIT_APARTMENTTHREADED)` runs first in `main`. WinRT silently fails on an uninitialized thread.

### 4. Config

Exe-relative path (`current_exe().parent().join("config.json")`, never CWD). `#[serde(default)]` at struct level makes missing/unknown fields safe.

```rust
struct Config {
    latitude: f64,
    longitude: f64,
    auto_start: bool,
    theme_day:   Option<String>,   // .theme path, or None → aero.theme
    theme_night: Option<String>,   // .theme path, or None → dark.theme
}
```

`has_location()` returns false when both coords are `0.0` (null-island sentinel used for first-run detection).

`set_auto_start(true)` writes `HKCU\...\Run\WinThemeSwitcher` with the current exe path (quoted). `set_auto_start(false)` calls `RegDeleteValueW` — both directions work.

### 5. Wake on session unlock / power resume

`ControlFlow::WaitUntil` uses `Instant`, which is monotonic and pauses across system suspend. Before this listener existed, a sunrise transition scheduled at, say, 6 AM would never fire if the machine was asleep through it: after wake at 8 AM, the runtime still saw the deadline as ~22 hours away (24 − sleep duration). The user had to click Refresh to recover.

`start_wake_listener` spawns a worker thread that creates a hidden message-only window (`HWND_MESSAGE`) and registers two notifications against it:

- `WTSRegisterSessionNotification(hwnd, NOTIFY_FOR_THIS_SESSION)` → delivers `WM_WTSSESSION_CHANGE`. We act on `WTS_SESSION_UNLOCK` (the user re-authenticated after Win+L or wake-from-sleep with a lock screen).
- `PowerRegisterSuspendResumeNotification(DEVICE_NOTIFY_WINDOW_HANDLE, hwnd, ...)` → delivers `WM_POWERBROADCAST`. We act on `PBT_APMRESUMEAUTOMATIC` (resume from sleep without a lock screen — covers machines that don't require re-auth on resume).

Both routes call `proxy.send_event(AppEvent::Wake)` via a process-wide `OnceLock<EventLoopProxy<AppEvent>>` (the WindowProc is `extern "system"` and can't capture). The main event loop handles `AppEvent::Wake` exactly like a scheduled transition — calls `tick()`, which is state-aware (no-op if `current == target`). Idempotent across both events firing in sequence.

`WTS_SESSION_UNLOCK` (0x8) and `DEVICE_NOTIFY_WINDOW_HANDLE` (0x0) are defined as local consts because windows-sys 0.59 doesn't re-export them under their expected modules. Values are stable Win32 ABI; safe to inline.

**Why this doesn't resurrect the manual-override-fight bug**: neither event fires when the user changes theme in Settings — `WM_WTSSESSION_CHANGE` is session lifecycle only, `WM_POWERBROADCAST` is power state only. So ticking on these is safe.

**Deliberate side effect**: a manual override that diverges from the schedule (e.g., user picks Dark mid-day when the schedule says Light) does *not* survive a lock/unlock or wake-from-sleep — `tick()` snaps back to the scheduled theme on `AppEvent::Wake`. This was tested explicitly in the v0.2.0 cycle and accepted as acceptable behavior. Preserving overrides across session events would require tracking the last theme *we* applied and only re-applying on Wake when `last_applied != target` (so a missed transition still reconciles, but a user override doesn't get clobbered). Not implemented; revisit if it becomes annoying.

### 6. Tray + menu

Menu: Open Config, Refresh, separator, Quit. Menu events flow through `MenuEvent::set_event_handler` → `EventLoopProxy::send_event(AppEvent::Menu(id))` so clicks wake the event loop even when it's on a 12-hour WaitUntil.

Tray icon is generated in `make_tray_icon`: 32×32 RGBA, half orange (sun) + half dark-blue (moon). Procedural because `tray-icon`'s default placeholder is near-invisible on both taskbar modes; `with_icon` is required for the icon to actually show.

## Dependencies (`Cargo.toml`)

- `chrono`, `sun-times` — sunrise/sunset math.
- `serde` + `serde_json` — config persistence.
- `tray-icon`, `winit` — tray + event loop. Menu types come from `muda` (re-exported under `tray_icon::menu`).
- `windows-sys` (features: `Win32_Foundation`, `Win32_System_Com`, `Win32_System_LibraryLoader`, `Win32_System_Power`, `Win32_System_RemoteDesktop`, `Win32_System_Registry`, `Win32_UI_WindowsAndMessaging`, `Win32_UI_Shell`, `Win32_Graphics_Dwm`) — raw Win32 FFI. `Win32_System_Com` is for `CoCreateInstance` + `CLSCTX_INPROC_SERVER` (IThemeManager2). `SysFreeString` lives in `Win32_Foundation` in windows-sys 0.59 (not `Win32_System_Ole` as you might expect).
- `windows` (features: `Devices_Geolocation`, `Foundation`, `Win32_System_Com`) — WinRT Geolocator + `CoInitializeEx` for the main thread's STA. Kept separate from `windows-sys` because the `windows` crate's typed bindings make Geolocator usable; raw `windows-sys` is fine for everything else.

## Invariants — don't break these

- **`tick()` scope**: only Init / ResumeTimeReached / Refresh / `AppEvent::Wake` (session unlock + power resume). Adding a callsite for any *other* trigger — especially anything that fires on `WM_SETTINGCHANGE` — resurrects the manual-override-fight bug. The wake events are safe specifically because they don't fire when the user changes the theme in Settings.
- **STA thread for IThemeManager2**: `ensure_com_initialized` runs `CoInitializeEx(None, COINIT_APARTMENTTHREADED)` first in `main`. All theme apply runs on that thread. Don't spawn worker threads to call `IThemeManager2` methods — they need their own `CoInitializeEx(STA)` and proper marshaling.
- **`poke_shell` after tier-2 / tier-3 apply only**: tier 1 (`IThemeManager2::SetCurrentTheme`) does the broadcast internally — calling `poke_shell` after it is wasted work and re-introduces the AV-tripping `HWND_BROADCAST WM_SETTINGCHANGE` signal that tier 1 was supposed to eliminate. Keep `poke_shell` for the legacy paths only; don't add it to tier 1.
- **Refresh forces apply** (bypasses state check); scheduled transitions respect it (no-op if already matching). Don't invert.
- **Free BSTRs from `ITheme::GetDisplayName` with `SysFreeString`** — not `CoTaskMemFree`, and definitely don't leak. The wtheme reference treats this strictly.
- **Vtable order in `IThemeManager2Vtbl`**: every method's slot index must match the COM ABI. Wrong order = calling the wrong method (silently catastrophic). The struct only declares the slots we call; trailing slots can be omitted but never reordered. Reference: namazso C# gist + wtheme C header (linked in main.rs comments).
- **`ensure_com_initialized` before any WinRT call**: otherwise Geolocator returns errors silently.
- **UTF-16 + NUL**: all Win32 wide strings go through `wide()` which appends the null terminator. Never pass a bare `&str` to a `*W` API.
- **HWND null check**: `(hwnd as usize) == 0` — robust to `windows-sys` flipping between `*mut c_void` and `isize`.
- **Windowed subsystem** (`#![windows_subsystem = "windows"]`): no console, `println!` goes nowhere. For diagnostics, write to a file or `OutputDebugStringW`.
