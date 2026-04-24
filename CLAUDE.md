# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Windows tray app (Rust) that swaps the full Windows **theme** (wallpaper + colors + light/dark mode) at local sunrise/sunset — macOS's auto-theme behavior, on Win11. Differentiators vs community alternatives like Auto Dark Mode: (a) aggressive shell-poke so the Win11 taskbar actually repaints, and (b) full `.theme` file swap so the wallpaper changes too — not just the DWORD toggle.

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

The binary trips `VHO:Trojan.Win32.Agent.gen` (unsigned Rust exe + `HKCU\Run` writes + `HWND_BROADCAST` of `WM_SETTINGCHANGE` + direct `WM_THEMECHANGED` to `Shell_TrayWnd` + WinRT Geolocation = every AV heuristic signal). Plain **path-based exclusions are insufficient** — Kaspersky's Behavior Detection quarantines regardless. The working setup is a **Trusted Applications rule** (Kaspersky Settings → Security → Threats and Exclusions → Specify trusted applications) with all checkboxes ticked: Do not scan opened files, Do not monitor application activity, Do not inherit restrictions, Do not monitor child application activity, Allow interaction with Kaspersky interface.

Rules are **path-based**, so two currently exist:
1. `C:\Tools\WinThemeSwitcher\win-theme-switcher.exe` (the deployed binary — stable).
2. `C:\Users\atef\Documents\Projects\WinThemeSwitcher\target\release\win-theme-switcher.exe` (the build output — rewritten by each `cargo build`).

If a future rebuild gets quarantined anyway (the hash changes and Kaspersky occasionally re-evaluates): drop a 0-byte placeholder at the path first (`Set-Content -Path ... -Value "" -Encoding Byte -Force`), re-add the trust rule while the placeholder exists, then rebuild. Same trick works for new deployment paths.

## Build

Rust toolchain is at `%USERPROFILE%\.cargo\bin\` (rustup, not on global PATH). Use the full path:

```powershell
& "$env:USERPROFILE\.cargo\bin\cargo.exe" build --release `
    --manifest-path "C:\Users\atef\Documents\Projects\WinThemeSwitcher\Cargo.toml"
```

Default toolchain is `stable-x86_64-pc-windows-msvc` (MSVC Build Tools required; the GNU toolchain's bundled linker/dlltool was broken on this machine). Release profile: `opt-level = "z"`, `lto = true`, `codegen-units = 1`, `panic = "abort"`, `strip = true`. Output ~310 KB. No `build.rs` — `windows-sys` and `windows` self-link.

## Architecture — `src/main.rs`

Single file, ~560 lines, event-driven, no polling.

### 1. Theme apply (`apply_theme` → `apply_theme_file` → `start_settings_closer` → `poke_shell`)

Preferred path: `ShellExecuteW("open", <.theme path>, ..., SW_HIDE)`. Windows' theme engine applies wallpaper + colors + mode atomically.

**Theme file resolution** (`resolve_theme_file`): if `config.theme_day` / `theme_night` is a valid path, use it; otherwise fall back to system defaults at `%SystemRoot%\Resources\Themes\aero.theme` (light) / `dark.theme` (dark).

**Settings flash workaround**: `ShellExecute` on a `.theme` file pops the Settings UWP app regardless of `SW_HIDE` (UWP apps manage their own visibility). `start_settings_closer` spawns a detached thread that loops 14× over ~2 s, `FindWindowW`-ing for class `ApplicationFrameWindow` with titles `"Settings"`, `"Themes"`, `"Personalization"` and `PostMessage(WM_CLOSE)` on each. Net effect: Settings flashes for ~300 ms then disappears. Full silence would require parsing the `.theme` file and applying each key individually (`SystemParametersInfo(SPI_SETDESKWALLPAPER, ...)` for wallpaper, DWORD writes for mode, accent color, etc.) — not implemented; filed under polish.

**Fallback path**: if the theme file is missing or ShellExecute fails, drop to DWORD-only mode (`write_theme_registry` writing `AppsUseLightTheme` + `SystemUsesLightTheme` 0/1) + broadcast + poke.

**`poke_shell` is load-bearing**: after any apply (theme file or DWORD), sends `WM_THEMECHANGED` + targeted `WM_SETTINGCHANGE("ImmersiveColorSet")` to `Shell_TrayWnd` and `Shell_SecondaryTrayWnd`, then `DwmFlush()`. This is the trick that makes the Win11 taskbar repaint reliably; **don't remove it**. If future Win versions add new taskbar window classes, extend the list here.

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

### 5. Tray + menu

Menu: Open Config, Refresh, separator, Quit. Menu events flow through `MenuEvent::set_event_handler` → `EventLoopProxy::send_event(AppEvent::Menu(id))` so clicks wake the event loop even when it's on a 12-hour WaitUntil.

Tray icon is generated in `make_tray_icon`: 32×32 RGBA, half orange (sun) + half dark-blue (moon). Procedural because `tray-icon`'s default placeholder is near-invisible on both taskbar modes; `with_icon` is required for the icon to actually show.

## Dependencies (`Cargo.toml`)

- `chrono`, `sun-times` — sunrise/sunset math.
- `serde` + `serde_json` — config persistence.
- `tray-icon`, `winit` — tray + event loop. Menu types come from `muda` (re-exported under `tray_icon::menu`).
- `windows-sys` (features: `Win32_Foundation`, `Win32_System_Registry`, `Win32_UI_WindowsAndMessaging`, `Win32_UI_Shell`, `Win32_Graphics_Dwm`) — raw Win32 FFI.
- `windows` (features: `Devices_Geolocation`, `Foundation`, `Win32_System_Com`) — WinRT Geolocator + COM init. Feature-gated to keep compile time manageable.

## Invariants — don't break these

- **`tick()` scope**: only Init / ResumeTimeReached / Refresh. Adding a callsite elsewhere resurrects the manual-override-fight bug.
- **`poke_shell` after every apply**: the entire reason this app exists over Auto Dark Mode.
- **Refresh forces apply** (bypasses state check); scheduled transitions respect it (no-op if already matching). Don't invert.
- **`ensure_com_initialized` before any WinRT call**: otherwise Geolocator returns errors silently.
- **UTF-16 + NUL**: all Win32 wide strings go through `wide()` which appends the null terminator. Never pass a bare `&str` to a `*W` API.
- **HWND null check**: `(hwnd as usize) == 0` — robust to `windows-sys` flipping between `*mut c_void` and `isize`.
- **Windowed subsystem** (`#![windows_subsystem = "windows"]`): no console, `println!` goes nowhere. For diagnostics, write to a file or `OutputDebugStringW`.
