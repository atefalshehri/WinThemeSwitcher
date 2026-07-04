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

### Release plan

Versioning before 1.0: a **patch** (0.x.y → 0.x.y+1) is pure bug fixes; a **minor** (0.x → 0.x+1) is anything that adds surface (menu items, config fields) or changes behavior. v1.0 is a gate, not a feature drop. Numbers are assigned at ship time — the 0.6/0.7 order can swap if the SignPath approval wait stalls, and regression patches (e.g. 0.4.1) slot in anywhere.

| Version | Type | Contents |
|---|---|---|
| **v0.3.2** | patch | **Shipped 2026-07-04.** Sunrise/sunset day-bracketing fix + scheduling tests + CI gate; `config.json` never overwritten on a parse error (error is logged + shown in a non-blocking dialog, empty file self-heals, autostart setting survives a broken file) |
| **v0.4.0** | minor | Preserve manual overrides across lock/unlock/resume (changes documented behavior); "Toggle theme" tray item; fail-loudly bundle (panic hook, startup MessageBox, wake-listener logging, single-instance mutex, bounded apply retry); remaining unit tests (config parsing, `.theme`-name resolution) |
| **v0.5.0** | minor | Scripted build → sign → verify → upload; first release not marked prerelease (fixes `/releases/latest`); candidate point to launch the Scoop bucket |
| **v0.6.0** | minor | SignPath CA signing in CI (ends the manual signed-asset swap); submit the first CA-signed binary to Defender + Kaspersky |
| **v0.7.0** | minor | Live tray tooltip; "Open Log" menu item; Refresh-failure MessageBox; first-run location retry; `offset_sunrise_min` / `offset_sunset_min` |
| **v1.0.0** | gate | Cut when winget accepts the package, the Correctness and Foundation sections below are empty, and a full release cycle has shipped CA-signed with no AV flags — i.e. a stranger can install it without reading the Kaspersky section |

### Correctness fixes

- **Preserve manual overrides across lock/unlock and resume.** A theme picked manually in Settings is currently snapped back to schedule on the next Win+L unlock or wake-from-sleep. Fix: track the last *scheduled* target on every tick and skip the wake-time re-apply when it hasn't changed (missed transitions still reconcile). Once that lands, add a **"Toggle theme" tray item** — an action, not GUI configuration.

### Foundation

- **Extend the unit tests.** The scheduling math (UTC+13, polar, and midnight-sunset fixtures) and config loading (parse-error preservation, first-run, empty-file heal) are covered, and `cargo test` is a hard CI gate since 2026-07-04. Still to cover: `.theme`-name resolution. Replaces the old "manual test matrix" item: Windows 10 hit end-of-support in Oct 2025, and the multi-monitor/HiDPI surface is one 32×32 tray icon.
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
