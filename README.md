# WinThemeSwitcher

Automatically swap between two Windows 11 **themes** at local sunrise and sunset — macOS's auto-theme behavior, on Windows 11. Full theme swap (wallpaper + colors + light/dark mode), not just a DWORD toggle.

> **Alpha software, signed builds.** Releases since v0.3.0 are Authenticode-signed with a self-signed cert (`CN=WinThemeSwitcher Self-Signed`). On first install you'll need to trust the cert — see [Verifying the signature](#verifying-the-signature). Architecture details and build/sign workflow are in [CLAUDE.md](CLAUDE.md).

## Features

- **Tiny binary** (~330 KB), zero CPU between transitions — event-driven, sleeps on a kernel timer until the next sunrise/sunset or a tray click.
- **Reliable theme apply** via the `IThemeManager2` COM interface — the same API the Settings UWP wraps internally. Atomic, in-process, ~200 ms latency, no Settings flash. Two-tier fallback if it ever errors.
- **Catches up after sleep / lock.** A scheduled sunrise that fires while you're suspended reconciles the moment you log back in.
- **Respects manual overrides** — changing theme in Settings sticks until the next natural transition; the app won't fight ambient setting-change broadcasts. Known gap: a lock/unlock or wake-from-sleep currently snaps back to the schedule (fix on the [Roadmap](#roadmap)).
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
| `auto_start` | `true` | When `true`, registers `HKCU\...\Run\WinThemeSwitcher`; when `false`, removes the entry. Applied on every launch and on Refresh. |
| `theme_day` | `null` → `%SystemRoot%\Resources\Themes\aero.theme` | Path to the `.theme` applied after sunrise. |
| `theme_night` | `null` → `%SystemRoot%\Resources\Themes\dark.theme` | Path to the `.theme` applied after sunset. |

Custom `.theme` paths must use double backslashes in JSON: `"C:\\Users\\you\\AppData\\Local\\Microsoft\\Windows\\Themes\\Custom.theme"`. They also need to be already registered with Windows (i.e. installed once via Settings → Personalization → Themes) for the primary apply path to find them.

After editing config, right-click tray → **Refresh**. No restart needed.

## Tray menu

- **Open Config** — opens `config.json` in your default editor.
- **Refresh** — re-reads config, retries Windows Location if needed, force-applies the correct theme.
- **Quit** — exits. The auto-start entry persists; set `auto_start: false` and click Refresh (or relaunch) once to remove it.

## Verifying the signature

After importing the publisher cert (Install step 2), check the binary:

```powershell
Get-AuthenticodeSignature .\win-theme-switcher.exe | Format-List Status, SignerCertificate
```

Should return `Status: Valid` and `Signer: CN=WinThemeSwitcher Self-Signed`. The signature lets you verify the file was signed by this project's publisher key and hasn't been tampered with since signing. Self-signed certs can't suppress SmartScreen — a CA-signed cert reduces prompts over time as reputation accrues (filed under [Roadmap](#roadmap)).

## Antivirus false positives

Even with a valid signature, some AVs may still flag the binary because the publisher cert isn't from a recognized CA. **The binary is not malicious** — full source is in this repo. Signed releases (v0.3.0+) have passed Kaspersky in testing without any special handling. If you build from source and run *unsigned*, most AVs accept a path-based exclusion, but **Kaspersky** needs a "Trusted application" rule (Settings → Threats and Exclusions → *Specify trusted applications* → tick all five checkboxes) because its Behavior Detection ignores plain exclusions. If your AV quarantines the file anyway, restore it and add the rule before re-running.

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

Ordered by priority. The correctness items come first — they are confirmed bugs in the shipped code (July 2026 audit), not features.

### Correctness fixes

- **Fix sunrise/sunset day-bracketing.** `target_theme`/`next_transition` pass the *local* calendar date to the `sun-times` crate, which expects a UTC date and keys events to the solar day. In UTC+13/+14 timezones (Samoa, Tonga, Chatham Islands) every computed sunrise/sunset lands on the wrong local day and the app is stuck permanently dark; near-midnight sunsets (Reykjavik in June) skew by minutes. Fix: compute transitions for local dates D−1/D/D+1 as UTC instants and pick the bracketing pair. Pairs with fixing the polar fallback (decide by solar altitude instead of fixed 06:00/18:00) and removing two `.unwrap()`s that can abort the process silently.
- **Never overwrite `config.json` on a parse error.** Today a JSON typo (e.g. a single backslash in a theme path) plus Refresh silently resets the file to defaults, wiping coordinates and custom theme paths — via the exact hand-edit flow the first-run dialogs instruct. Keep the file, log the error, and tell the user when Refresh hits it.
- **Preserve manual overrides across lock/unlock and resume.** A theme picked manually in Settings is currently snapped back to schedule on the next Win+L unlock or wake-from-sleep. Fix: track the last *scheduled* target on every tick and skip the wake-time re-apply when it hasn't changed (missed transitions still reconcile). Once that lands, add a **"Toggle theme" tray item** — an action, not GUI configuration.

### Foundation

- **Unit tests + CI gate.** The repo ships zero tests. The scheduling math, config parsing, and `.theme`-name resolution are pure functions — cover them (UTC+14, polar, and Riyadh fixtures) and make `cargo test` a hard CI gate. Replaces the old "manual test matrix" item: Windows 10 hit end-of-support in Oct 2025, and the multi-monitor/HiDPI surface is one 32×32 tray icon.
- **Fail loudly instead of silently.** A panic hook that writes to `events.log` before abort (`panic = "abort"` + windowed subsystem currently means zero-trace death), a MessageBox + log line when startup fails (e.g. tray creation racing the taskbar at login), logging on wake-listener registration failures, a single-instance mutex, and a bounded retry when a theme apply fails (today the next attempt can be ~12 h away).

### Release & distribution (dependency chain, in order)

1. **Harden the release process now**: add a timestamp countersignature to signing (`/tr http://timestamp.digicert.com /td SHA256` — without it, signatures die with the cert), script the build → sign → verify → upload flow so the v0.3.0 "shipped unsigned for months" failure can't recur, and stop marking every release `prerelease` (it breaks `/releases/latest` and winget automation).
2. **CA-signed releases** via [SignPath Foundation's free OSS program](https://signpath.org) — signing moves into CI, which also permanently eliminates the manual signed-asset swap. Reduces SmartScreen prompts over time via cert reputation (no cert eliminates them outright). Azure Trusted Signing is not an option: individual validation is US/Canada-only.
3. **Submit the first CA-signed binary to Microsoft Defender** — this gates winget, whose validation pipeline runs AV scans — and to Kaspersky. Repeat only if a specific release gets flagged.
4. **winget package** (`InstallerType: portable`), only after steps 1–2 make asset hashes final at publish time. A personal Scoop bucket may come earlier: Scoop's `persist` mechanism fits the exe-relative config model better than winget's symlink layout.

### UX polish

- **Live tray tooltip** — "Dark until 06:12", "Location needed — click Refresh", or a degraded-apply warning — plus an **"Open Log"** menu item and a MessageBox when a user-initiated Refresh fails (scheduled ticks stay silent-to-log).
- **First-run without the Refresh step** — retry Windows Location a few times right after the enable-Location dialog is dismissed, instead of a `Geolocator::StatusChanged` subscription. Bigger optional variant: re-query location on unlock/resume so a traveler's coordinates don't go stale (today they are never re-read once set).
- **Optional sunrise/sunset offsets** (`offset_sunrise_min` / `offset_sunset_min` in config) — still sun-anchored, so compatible with "no custom times".

### Maintenance notes (not scheduled)

- **winit `ApplicationHandler` migration** — only when bumping to winit 0.31 (the pinned 0.30 merely deprecates `EventLoop::run`; nothing forces this today). Worth evaluating at that point: dropping winit for a plain Win32 message loop — the pattern already exists in the wake listener.

**Not planned**: GUI configuration (`config.json` + Refresh is the UX), custom wake times (sunrise/sunset is the whole point; offsets from them are fine), cross-platform (Windows only — macOS already has this natively), in-app update check (it would be the binary's only network call and re-adds the autorun+beacon AV-heuristic surface the `IThemeManager2` migration removed — winget/Scoop handle upgrades), pause/snooze toggle (a manual override already pauses until the next transition), and ADM-style scripting/hotkeys/battery rules (out of scope for a ~330 KB tray tool).

## Contributing

Alpha / personal-use tool first, distributable second. Issues and PRs welcome — please open an issue to discuss larger changes before sending a patch. Focus areas: see [Roadmap](#roadmap).

## License

MIT — see [LICENSE](LICENSE).

## Acknowledgments

- [`sun-times`](https://crates.io/crates/sun-times) — local sunrise/sunset math.
- [`chrono`](https://crates.io/crates/chrono) — date/time handling; [`serde`](https://crates.io/crates/serde) + [`serde_json`](https://crates.io/crates/serde_json) — config persistence.
- [`tray-icon`](https://crates.io/crates/tray-icon) + [`winit`](https://crates.io/crates/winit) — tray icon and event loop.
- [`windows-sys`](https://crates.io/crates/windows-sys) + [`windows`](https://crates.io/crates/windows) — official Microsoft Rust bindings.
