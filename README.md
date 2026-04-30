# WinThemeSwitcher

Automatically swap between two Windows 11 **themes** at local sunrise and sunset — macOS's auto-theme behavior, on Windows 11. Full theme swap (wallpaper + colors + light/dark mode), not just a DWORD toggle.

> **Alpha software, signed builds.** Releases since v0.3.0 are Authenticode-signed with a self-signed cert (`CN=WinThemeSwitcher Self-Signed`). On first install you'll need to trust the cert — see [Verifying the signature](#verifying-the-signature). Architecture details and build/sign workflow are in [CLAUDE.md](CLAUDE.md).

## Features

- **Tiny binary** (~330 KB), zero CPU between transitions — event-driven, sleeps on a kernel timer until the next sunrise/sunset or a tray click.
- **Reliable theme apply** via the `IThemeManager2` COM interface — the same API the Settings UWP wraps internally. Atomic, in-process, ~200 ms latency, no Settings flash. Two-tier fallback if it ever errors.
- **Catches up after sleep / lock.** A scheduled sunrise that fires while you're suspended reconciles the moment you log back in.
- **Respects manual overrides** — changing theme in Settings sticks until the next natural transition; the app won't fight ambient setting-change broadcasts.
- **Diagnostic log** at `events.log` next to the exe (rotated past 256 KB) — every transition recorded with cause, target, applied tier, and timing.

## Install

1. Download `win-theme-switcher-vX.Y.Z-windows-x64.zip` from the [latest release](../../releases). Extract to e.g. `C:\Tools\WinThemeSwitcher\`.
2. **Trust the publisher cert** (one-time, recommended — the zip includes `WinThemeSwitcher-publisher.cer`):
   ```powershell
   Import-Certificate -FilePath .\WinThemeSwitcher-publisher.cer -CertStoreLocation Cert:\CurrentUser\Root
   Import-Certificate -FilePath .\WinThemeSwitcher-publisher.cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher
   ```
3. Run `win-theme-switcher.exe`. A half-orange / half-dark-blue circle appears in your notification area.

On first launch the app reads your coordinates via Windows Location. If Location is off or denied, a dialog asks if you want to enable it (opens Settings) or fall back to manual entry (opens `config.json` in your editor). After editing config, right-click tray → **Refresh**.

## Configuration

`config.json` lives next to the exe (deliberately, so the app works from wherever you put it).

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
| `latitude` / `longitude` | from Windows Location, else `0.0` | Decimal degrees. `0.0, 0.0` triggers the first-run location flow. |
| `auto_start` | `true` | When `true`, registers `HKCU\...\Run\WinThemeSwitcher` on every launch. When `false`, removes it. |
| `theme_day` | `null` → `%SystemRoot%\Resources\Themes\aero.theme` | Path to the `.theme` applied after sunrise. |
| `theme_night` | `null` → `%SystemRoot%\Resources\Themes\dark.theme` | Path to the `.theme` applied after sunset. |

Custom `.theme` paths must use double backslashes in JSON: `"C:\\Users\\you\\AppData\\Local\\Microsoft\\Windows\\Themes\\Custom.theme"`. They also need to be already registered with Windows (i.e. installed once via Settings → Personalization → Themes) for the primary apply path to find them.

After editing config, right-click tray → **Refresh**. No restart needed.

## Tray menu

- **Open Config** — opens `config.json` in your default editor.
- **Refresh** — re-reads config, retries Windows Location if needed, force-applies the correct theme.
- **Quit** — exits. The auto-start entry persists; set `auto_start: false` and relaunch once to remove it.

## Verifying the signature

After importing the publisher cert (Install step 2), check the binary:

```powershell
Get-AuthenticodeSignature .\win-theme-switcher.exe | Format-List Status, SignerCertificate
```

Should return `Status: Valid` and `Signer: CN=WinThemeSwitcher Self-Signed`. The signature lets you verify the file was actually built from the source in this repo and hasn't been tampered with in transit. Self-signed certs can't fully suppress SmartScreen — that needs a CA-signed cert (filed under [Roadmap](#roadmap)).

## Antivirus false positives

Even with a valid signature, some AVs may still flag the binary because the publisher cert isn't from a recognized CA. **The binary is not malicious** — full source is in this repo. Most AVs accept a path-based exclusion; **Kaspersky** specifically needs a "Trusted application" rule (Settings → Threats and Exclusions → *Specify trusted applications* → tick all five checkboxes) because its Behavior Detection ignores plain exclusions. If your AV quarantines the file anyway, restore it and add the rule before re-running.

## How it works

`apply_theme` is a three-tier fallback:

1. **`IThemeManager2`** (primary) — the undocumented-but-stable COM interface in `themeui.dll` that the Settings UWP wraps internally. Atomic, in-process apply. `SetCurrentTheme(idx)` does the `WM_THEMECHANGED` + `WM_SETTINGCHANGE` broadcasts itself.
2. **`ShellExecuteW(.theme)` + commit watcher** — legacy backup if the COM interface ever errors. A 5 s watcher polls the registry to detect silent failures and promotes to tier 3.
3. **Direct registry write** — last resort. Flips light/dark mode but not wallpaper.

Sunrise/sunset times come from the [`sun-times`](https://crates.io/crates/sun-times) crate (no network). The event loop blocks on `winit`'s `WaitUntil(next_transition)` between transitions — zero CPU. A worker thread catches `WTSRegisterSessionNotification` (unlock) and `PowerRegisterSuspendResumeNotification` (resume from sleep) so transitions reconcile after long suspends.

Full architecture, threading invariants, and the reasoning behind each tier are in [CLAUDE.md](CLAUDE.md).

## Building from source

Requirements: Rust `stable-x86_64-pc-windows-msvc` + Visual Studio Build Tools with the C++ workload.

```powershell
winget install Rustlang.Rustup
winget install Microsoft.VisualStudio.2022.BuildTools --override "--add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"
git clone https://github.com/atefalshehri/WinThemeSwitcher.git
cd WinThemeSwitcher
cargo build --release
# Output: target\release\win-theme-switcher.exe (~330 KB)
```

Release profile is tuned for size (`opt-level = "z"`, `lto = true`, `strip = true`, `panic = "abort"`). No `build.rs` — `windows-sys` and `windows` self-link.

Release builds should be Authenticode-signed before deploy. See [CLAUDE.md → Sign every release build](CLAUDE.md#sign-every-release-build) for the cert-generation + signtool commands.

## Uninstall

1. Right-click tray → Quit.
2. Delete the install folder.
3. Remove auto-start:
   ```cmd
   reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v WinThemeSwitcher /f
   ```

No installer, no uninstaller — it's a single-exe tool by design.

## Roadmap

- **CA-signed certificate** via [SignPath.io's free OSS program](https://signpath.io/foss) — would eliminate SmartScreen prompts and the remaining AV flags.
- **Submit FP reports** to Kaspersky / Microsoft Defender / Bitdefender.
- **Migrate to `winit::application::ApplicationHandler`** (current code uses the deprecated `EventLoop::run` callback API).
- **In-app update check** against GitHub Releases (passive notification only — no auto-download).
- **Auto-detect coordinates when Location gets enabled** — currently `try_get_windows_location()` only retries on Init or Refresh. Subscribing to `Geolocator::StatusChanged` (or polling the access status briefly after the user clicks Yes on the enable-Location prompt) would let the app pick up coords automatically the moment Location is allowed, removing the "now click Refresh" step from first-run UX.
- **winget package** submission for `winget install WinThemeSwitcher`.
- **Test matrix**: Windows 10, multi-monitor, HiDPI, non-English locales.

**Not planned**: GUI configuration (`config.json` + Refresh is the UX), custom wake times (sunrise/sunset is the whole point), cross-platform (Windows only — macOS already has this natively).

## Contributing

Alpha / personal-use tool first, distributable second. Issues and PRs welcome — please open an issue to discuss larger changes before sending a patch. Focus areas: see [Roadmap](#roadmap).

## License

MIT — see [LICENSE](LICENSE).

## Acknowledgments

- [`sun-times`](https://crates.io/crates/sun-times) — local sunrise/sunset math.
- [`tray-icon`](https://crates.io/crates/tray-icon) + [`winit`](https://crates.io/crates/winit) — tray icon and event loop.
- [`windows-sys`](https://crates.io/crates/windows-sys) + [`windows`](https://crates.io/crates/windows) — official Microsoft Rust bindings.
