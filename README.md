# WinThemeSwitcher

Automatically swap between two Windows 11 **themes** at local sunrise and sunset — the macOS automatic dark mode experience, on Windows 11. Full theme swap (wallpaper + colors + light/dark mode), not just a DWORD toggle.

> **Alpha software.** Releases since v0.3.0 are Authenticode-signed with a self-signed code-signing certificate (`CN=WinThemeSwitcher Self-Signed`). On first install you'll need to trust the certificate — see [Verifying the signature](#verifying-the-signature). Some AVs may still flag the binary because the publisher cert isn't from a recognized CA — see [Antivirus false positives](#antivirus-false-positives).

## Why

Windows has no native auto-theme feature. The standard approach — toggling `AppsUseLightTheme` / `SystemUsesLightTheme` and broadcasting `WM_SETTINGCHANGE` — frequently leaves the **taskbar** and some tray icons stuck on the previous theme, a known Win11 issue with how the broadcast propagates to Explorer.

This app handles that: after applying a `.theme` file, it sends `WM_THEMECHANGED` **directly to `Shell_TrayWnd`** (and `Shell_SecondaryTrayWnd` for multi-monitor setups) and calls `DwmFlush()`. That forces the taskbar to repaint reliably without the "restart Explorer" hammer.

Design priorities:
- **Tiny binary** (~330 KB).
- **Near-zero CPU** — event-driven, sleeps on a kernel timer between sunrise and sunset; no polling.
- **Full theme swap** — wallpaper, accent colors, and mode change together via Windows `.theme` files. Primary apply path is the `IThemeManager2` COM interface (the same one the Settings UWP wraps internally), with a two-tier fallback if it ever errors.
- **Respects manual overrides** — changing the theme in Settings sticks until the next natural transition; the app won't fight ambient setting-change broadcasts.
- **Catches up after sleep / lock** — if your machine is suspended through a sunrise (or you lock overnight), the theme reconciles to the schedule the moment you log back in. No need to click Refresh.
- **Diagnostic log** at `events.log` next to the exe, recording every transition with cause, target, applied tier, and timing (rotated past 256 KB).

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

## Verifying the signature

Releases since v0.3.0 are signed with a self-signed Authenticode certificate (`CN=WinThemeSwitcher Self-Signed`, RSA-2048, SHA-256, valid through 2036). Because it's not from a recognized CA, Windows / SmartScreen / your AV will still treat it as untrusted on first install — but the signature lets you verify the file was actually built from the source in this repo and hasn't been tampered with in transit.

Inspect the signature in PowerShell:

```powershell
Get-AuthenticodeSignature .\win-theme-switcher.exe | Format-List Status, SignerCertificate, StatusMessage
```

Status should be `UnknownError` or `NotTrusted` until you install the cert in your Trusted Root + Trusted Publisher stores. The publisher cert is bundled in the release zip as `WinThemeSwitcher-publisher.cer`. Install it once with:

```powershell
Import-Certificate -FilePath .\WinThemeSwitcher-publisher.cer -CertStoreLocation Cert:\CurrentUser\Root
Import-Certificate -FilePath .\WinThemeSwitcher-publisher.cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher
```

After that, `Get-AuthenticodeSignature` should return `Status: Valid`, and most AVs will stop flagging the file. (Self-signed certs can't suppress SmartScreen entirely — that needs a paid CA cert.)

## Antivirus false positives

Even with a valid signature, some AVs may still flag this binary because the publisher cert isn't from a recognized CA. Kaspersky historically flagged unsigned builds as `VHO:Trojan.Win32.Agent.gen`; Microsoft Defender, Bitdefender, Avast, and Norton have similar generic detections. **The binary is not malicious** — the full source is in this repo (~1100 lines in `src/main.rs`).

Why it looks suspicious to heuristics:
- Self-signed publisher (not a recognized CA)
- Writes to `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` (auto-start)
- Writes `HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize` DWORDs
- Tier 2/3 fallback paths broadcast `WM_SETTINGCHANGE` to every top-level window (`HWND_BROADCAST`) and send `WM_THEMECHANGED` directly to shell windows. Tier 1 (the primary path) avoids both — `IThemeManager2::SetCurrentTheme` does the broadcast internally.
- Uses WinRT Geolocation
- Runs hidden (no console, `windows_subsystem = "windows"`)
- Aggressive release-profile stripping

Every one of those is a signal real malware also produces. The only thing separating this app from malware is *intent*, which a heuristic can't read.

### How to allowlist

**Windows Defender / SmartScreen**
Settings → Privacy & security → Windows Security → Virus & threat protection → Manage settings → Exclusions → Add an exclusion → File → point at the exe.

**Kaspersky** (all tiers)
Settings → Security → Threats and Exclusions → **Specify trusted applications** → Add → point at the exe. Tick **all** exclusion checkboxes: Do not scan opened files, Do not monitor application activity, Do not inherit restrictions, Do not monitor child application activity, Allow interaction with Kaspersky interface. A plain path-based Exclusion is **not enough** — Kaspersky's Behavior Detection still quarantines without the Trusted Applications rule.

**Other AVs**
Look for "Exceptions," "Allowed applications," "Whitelist," or "Trusted applications" in the settings. Scope the rule to the full exe path.

After allowlisting, if the app gets quarantined on first run anyway, you may need to restore it from the AV's quarantine (or re-download and drop in place). Most AV rules are path-based, so future updates at the same path will usually survive.

## How it works

- **Scheduling.** Uses the [`sun-times`](https://crates.io/crates/sun-times) crate to compute today's sunrise and sunset locally (no network). Sets `winit`'s `ControlFlow::WaitUntil(next_transition)` — the process blocks on a kernel object until the deadline or a menu event; **zero CPU** between transitions.
- **Theme apply.** Three-tier fallback in `apply_theme`:
  1. **`IThemeManager2`** (primary) — undocumented-but-stable COM interface in `themeui.dll` (CLSID `9324da94-50ec-4a14-a770-e90ca03e7c8f`), the same API the Settings UWP wraps internally. `SetCurrentTheme(idx)` applies wallpaper + colors + mode atomically and broadcasts `WM_THEMECHANGED` itself. ~200 ms latency, no Settings flash, no manual broadcast needed.
  2. **`ShellExecuteW(.theme)`** + Settings closer thread — legacy path, kept as backup if Microsoft removes the COM interface. A `commit_watcher` polls the registry for 5 s after this fires; if the apply silently failed (observed when the schedule fires while no foreground UI is active), it auto-promotes to tier 3.
  3. **Direct registry write** — last resort. Writes `AppsUseLightTheme` / `SystemUsesLightTheme`, broadcasts `WM_SETTINGCHANGE("ImmersiveColorSet")`, sends `WM_THEMECHANGED` to `Shell_TrayWnd` and `Shell_SecondaryTrayWnd`, calls `DwmFlush()`. Flips light/dark mode but **not the wallpaper** (would need a `.theme` file apply for that).
- **Taskbar fix.** Tier 1's `SetCurrentTheme` handles taskbar repaint internally. Tiers 2 and 3 send `WM_THEMECHANGED` + targeted `WM_SETTINGCHANGE("ImmersiveColorSet")` to `Shell_TrayWnd` and `Shell_SecondaryTrayWnd` plus `DwmFlush()` — the trick that makes the Win11 taskbar repaint reliably on the legacy paths.
- **Override respect.** The event loop only re-evaluates on `StartCause::Init`, `StartCause::ResumeTimeReached`, a user-driven Refresh, or a session-resume signal (see *Wake recovery*). Ambient events (including the `WM_SETTINGCHANGE` Windows fires when *you* change a theme manually) don't trigger a re-apply.
- **Wake recovery.** A worker thread owns a hidden message-only window registered for `WTSRegisterSessionNotification` (delivers `WM_WTSSESSION_CHANGE` → `WTS_SESSION_UNLOCK`) and `PowerRegisterSuspendResumeNotification` (delivers `WM_POWERBROADCAST` → `PBT_APMRESUMEAUTOMATIC`). On either event the listener dispatches into the main event loop, which reconciles to the scheduled theme. This works around the fact that `winit`'s `WaitUntil` deadline uses a monotonic clock that pauses across system suspend — without this listener, a sunrise scheduled for 6 AM would never fire if the machine slept through it. The trade-off: a manual override that diverges from the schedule (e.g., Dark at noon when the schedule says Light) snaps back to the schedule on lock/unlock or wake; it persists otherwise.
- **Diagnostic log.** Every transition writes a line to `events.log` next to the exe (rotated to `events.log.old` past 256 KB). Format is space-separated key=value: `<rfc3339-timestamp> cause=<init|resume-time|wake-unlock|wake-power|refresh> current=<light|dark> target=<light|dark> applied=<theme-manager2|theme-file|registry|skip> next=<next-transition>`. Useful for diagnosing silent failures.

Full architecture details in [CLAUDE.md](CLAUDE.md).

## Building from source

Requirements: Rust `stable-x86_64-pc-windows-msvc` + Visual Studio Build Tools with the C++ workload.

```powershell
winget install Rustlang.Rustup
winget install Microsoft.VisualStudio.2022.BuildTools --override "--add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"
git clone https://github.com/atefalshehri/WinThemeSwitcher.git
cd WinThemeSwitcher
cargo build --release
# Output: target\release\win-theme-switcher.exe
```

Release profile is tuned for size (`opt-level = "z"`, `lto = true`, `strip = true`, `panic = "abort"`). Final binary ≈330 KB.

No `build.rs` — `windows-sys` and `windows` self-link.

### Signing

Release builds are signed with a self-signed Authenticode certificate. To sign your own builds:

```powershell
# One-time: generate a code-signing cert and install it to your Trusted Root + Trusted Publisher stores
$cert = New-SelfSignedCertificate -Type CodeSigning -Subject "CN=WinThemeSwitcher Self-Signed" `
    -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
    -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears(10)
foreach ($s in @("Root", "TrustedPublisher")) {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($s, "CurrentUser")
    $store.Open("ReadWrite"); $store.Add($cert); $store.Close()
}
$pwd = ConvertTo-SecureString "wts-local-signing" -Force -AsPlainText
Export-PfxCertificate -Cert "Cert:\CurrentUser\My\$($cert.Thumbprint)" `
    -FilePath "$env:LOCALAPPDATA\WinThemeSwitcher\signing\winthemeswitcher-signing.pfx" -Password $pwd

# Per-build: sign after each cargo build --release
& "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe" sign `
    /f "$env:LOCALAPPDATA\WinThemeSwitcher\signing\winthemeswitcher-signing.pfx" `
    /p "wts-local-signing" /fd SHA256 .\target\release\win-theme-switcher.exe
```

The signature has no countersigned timestamp, so it expires when the cert does.

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

- **CA-signed certificate** via [SignPath.io's free OSS program](https://signpath.io/foss) to replace the self-signed cert — would eliminate SmartScreen prompts and most remaining AV flags. (Self-signed already lands in v0.3.0.)
- **Submit FP reports** to Kaspersky / Microsoft Defender / Bitdefender.
- **Migrate to `winit::application::ApplicationHandler`** (current code uses the deprecated `EventLoop::run` callback API).
- **winget package** submission for `winget install WinThemeSwitcher`.
- **In-app update check** against GitHub Releases (passive notification only — no auto-download).
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
