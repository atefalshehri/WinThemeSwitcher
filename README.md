# WinThemeSwitcher

Automatically swap between two Windows 11 **themes** at local sunrise and sunset — the macOS automatic dark mode experience, on Windows 11. Full theme swap (wallpaper + colors + light/dark mode), not just a DWORD toggle.

> **Alpha software.** Your antivirus will likely flag the binary as a false positive — see [Antivirus false positives](#antivirus-false-positives) for why and how to allowlist it. No code signing yet.

## Why

Windows has no native auto-theme feature. The main community alternative, [Auto Dark Mode](https://github.com/AutoDarkMode/Windows-Auto-Night-Mode), works well for most of the UI but frequently leaves the **taskbar** and some tray icons stuck on the previous theme — a known Win11 issue with how `WM_SETTINGCHANGE` propagates to Explorer.

This app's differentiator: after applying a `.theme` file, it sends `WM_THEMECHANGED` **directly to `Shell_TrayWnd`** (and `Shell_SecondaryTrayWnd` for multi-monitor setups) and calls `DwmFlush()`. That forces the taskbar to repaint reliably without the "restart Explorer" hammer.

Design priorities:
- **Tiny binary** (~310 KB).
- **Near-zero CPU** — event-driven, sleeps on a kernel timer between sunrise and sunset; no polling.
- **Full theme swap** — wallpaper, accent colors, and mode change together via Windows `.theme` files.
- **Respects manual overrides** — if you change the theme in Settings between transitions, the app won't fight you. It re-applies only at the next natural transition.

## Install

1. Download `win-theme-switcher.exe` (or the `.zip`) from the [latest release](../../releases).
2. Put it somewhere permanent, e.g. `C:\Tools\WinThemeSwitcher\`.
3. **Allowlist it in your antivirus** — see [below](#antivirus-false-positives). If you skip this, your AV will probably delete it before you can run it.
4. Double-click to run. A half-orange / half-dark-blue circle appears in your notification area.

### First run

The app tries to read your coordinates via **Windows Location** (requires Settings → Privacy & security → Location → "Let apps access your location" + "Let desktop apps access your location" enabled).

- If Location is available: coords are saved to `config.json` silently.
- If Location is off or denied: a dialog asks if you'd like to enable it. **Yes** opens the Location settings page. **No** falls back to manual entry — your `config.json` opens in the default editor; fill in `latitude` and `longitude`, save, and right-click the tray icon → **Refresh**.

From that point the app sleeps until the next sunrise or sunset, applies the corresponding theme, and goes back to sleep.

## Configuration

`config.json` lives next to the exe (not in your user profile — deliberately, so the app works from wherever you put it).

```json
{
  "latitude": 40.7128,
  "longitude": -74.0060,
  "auto_start": true,
  "theme_day": null,
  "theme_night": null
}
```

| Field | Default | Meaning |
|---|---|---|
| `latitude` / `longitude` | from Windows Location, else `0.0` | Decimal degrees. `0.0, 0.0` is the "unconfigured" sentinel — triggers the first-run location flow. |
| `auto_start` | `true` | When `true`, the app registers itself at `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` on every launch so it starts at login. When `false`, it removes that entry. |
| `theme_day` | `null` → `%SystemRoot%\Resources\Themes\aero.theme` | Path to the `.theme` file applied after sunrise. |
| `theme_night` | `null` → `%SystemRoot%\Resources\Themes\dark.theme` | Path to the `.theme` file applied after sunset. |

Windows `.theme` files are in:
- `%LocalAppData%\Microsoft\Windows\Themes\*.theme` — themes you saved via Settings → Personalization → Themes → "Save theme."
- `C:\Windows\Resources\Themes\*.theme` — system-provided (`aero`, `dark`, `spotlight`, `themeA–D`).

**JSON path escaping**: double the backslashes.

```json
"theme_night": "C:\\Users\\you\\AppData\\Local\\Microsoft\\Windows\\Themes\\Custom.theme"
```

After editing the config, right-click the tray icon → **Refresh**. No restart needed.

## Tray menu

- **Open Config** — opens `config.json` in your default editor.
- **Refresh** — re-reads config, retries Windows Location if coords are still unset, and force-applies the correct theme.
- **Quit** — exits. The auto-start registry entry persists; set `auto_start: false` and relaunch once if you want to remove it.

## Antivirus false positives

Kaspersky flags this binary as `VHO:Trojan.Win32.Agent.gen`. Microsoft Defender, Bitdefender, Avast, and Norton have similar generic detections. **The binary is not malicious** — the full source is in this repo (~560 lines in `src/main.rs`).

Why it looks suspicious to heuristics:
- Unsigned executable from an unknown publisher
- Writes to `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` (auto-start)
- Writes `HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize` DWORDs
- Broadcasts `WM_SETTINGCHANGE` to every top-level window (`HWND_BROADCAST`)
- Sends `WM_THEMECHANGED` directly to shell windows
- Uses WinRT Geolocation
- Runs hidden (no console, `windows_subsystem = "windows"`)
- Tiny binary, aggressive release-profile stripping

Every one of those is a signal real malware also produces. The only thing separating this app from malware is *intent*, which a heuristic can't read.

### How to allowlist

**Windows Defender / SmartScreen**
Settings → Privacy & security → Windows Security → Virus & threat protection → Manage settings → Exclusions → Add an exclusion → File → point at the exe.

**Kaspersky** (all tiers)
Settings → Security → Threats and Exclusions → **Specify trusted applications** → Add → point at the exe. Tick **all** exclusion checkboxes: Do not scan opened files, Do not monitor application activity, Do not inherit restrictions, Do not monitor child application activity, Allow interaction with Kaspersky interface. A plain path-based Exclusion is **not enough** — Kaspersky's Behavior Detection still quarantines without the Trusted Applications rule.

**Other AVs**
Look for "Exceptions," "Allowed applications," "Whitelist," or "Trusted applications" in the settings. Scope the rule to the full exe path.

After allowlisting, if the app gets quarantined on first run anyway, you may need to restore it from the AV's quarantine (or re-download and drop in place). Most AV rules are path-based, so future updates at the same path will usually survive.

Long-term, code signing will reduce these flags — see [Roadmap](#roadmap).

## How it works

- **Scheduling.** Uses the [`sun-times`](https://crates.io/crates/sun-times) crate to compute today's sunrise and sunset locally (no network). Sets `winit`'s `ControlFlow::WaitUntil(next_transition)` — the process blocks on a kernel object until the deadline or a menu event; **zero CPU** between transitions.
- **Theme apply.** `ShellExecuteW("open", <.theme file>)` hands the file to Windows' theme engine, which applies wallpaper + colors + mode atomically. Falls back to a DWORD-only Light/Dark toggle if the theme file is missing.
- **Settings flash workaround.** Opening a `.theme` file via the shell briefly pops the Windows Settings app open. A detached thread finds `ApplicationFrameWindow` windows titled "Settings" / "Themes" / "Personalization" and posts `WM_CLOSE` to each for ~2 seconds. Net result: Settings flashes for a few hundred milliseconds and closes itself.
- **Taskbar fix.** After the apply, the app broadcasts `WM_SETTINGCHANGE("ImmersiveColorSet")`, then sends `WM_THEMECHANGED` + a targeted `WM_SETTINGCHANGE` to `Shell_TrayWnd` and `Shell_SecondaryTrayWnd`, then calls `DwmFlush()`. This is what makes the taskbar repaint reliably on Win11.
- **Override respect.** The event loop only re-evaluates on `StartCause::Init`, `StartCause::ResumeTimeReached`, or a user-driven Refresh. Ambient events (including the `WM_SETTINGCHANGE` Windows fires when *you* change a theme manually) don't trigger a re-apply.

Full architecture details in [CLAUDE.md](CLAUDE.md).

## Building from source

Requirements: Rust `stable-x86_64-pc-windows-msvc` + Visual Studio Build Tools with the C++ workload.

```powershell
winget install Rustlang.Rustup
winget install Microsoft.VisualStudio.2022.BuildTools --override "--add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"
git clone https://github.com/<your-username>/WinThemeSwitcher.git
cd WinThemeSwitcher
cargo build --release
# Output: target\release\win-theme-switcher.exe
```

Release profile is tuned for size (`opt-level = "z"`, `lto = true`, `strip = true`, `panic = "abort"`). Final binary ≈310 KB.

No `build.rs` — `windows-sys` and `windows` self-link.

## Uninstall

1. Right-click tray → Quit.
2. Delete the install folder.
3. Remove the auto-start entry:
   ```cmd
   reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v WinThemeSwitcher /f
   ```

No installer, no uninstaller — it's a single-exe tool by design.

## Roadmap

Rough priority order:

- **Code signing** via [SignPath.io's free OSS program](https://signpath.io/foss) — will reduce SmartScreen prompts and AV flags for signed releases.
- **Submit FP reports** to Kaspersky / Microsoft Defender / Bitdefender.
- **Silent theme apply** — parse `.theme` files and apply wallpaper + DWORDs + accent color directly via `SystemParametersInfo` / registry, removing the Settings flash entirely.
- **Log file** at `%LocalAppData%\WinThemeSwitcher\log.txt` for diagnosing silent failures.
- **Migrate to `winit::application::ApplicationHandler`** (current code uses the deprecated `EventLoop::run` callback API).
- **winget package** submission for `winget install WinThemeSwitcher`.
- **Test matrix**: Windows 10, multiple timezones, multi-monitor, HiDPI, non-English locales.

### Not planned

- GUI configuration — `config.json` + Refresh is the UX.
- Custom wake times — sunrise/sunset is the whole point.
- Cross-platform — Windows only; macOS already has this natively.

## Contributing

This is an alpha / personal-use tool first, a distributable app second. Issues and PRs welcome, but please open an issue to discuss larger changes before sending a patch. Focus areas where help is most useful are in the [Roadmap](#roadmap).

## License

MIT — see [LICENSE](LICENSE).

## Acknowledgments

- [`sun-times`](https://crates.io/crates/sun-times) — local sunrise/sunset math.
- [`tray-icon`](https://crates.io/crates/tray-icon) + [`winit`](https://crates.io/crates/winit) — tray icon and event loop.
- [`windows-sys`](https://crates.io/crates/windows-sys) + [`windows`](https://crates.io/crates/windows) — official Microsoft Rust bindings.
- The [Auto Dark Mode](https://github.com/AutoDarkMode/Windows-Auto-Night-Mode) project for pioneering this space on Windows and making clear exactly where the Win11 taskbar quirk lives.
