# WinThemeSwitcher

Automatically swap between two Windows 11 **themes** at local sunrise and sunset — macOS's auto-theme behavior, on Windows 11. Full theme swap (wallpaper + colors + light/dark mode), not just a DWORD toggle.

> **Signed releases since v0.3.0** (v0.1.0/v0.2.0 are unsigned). Binaries are Authenticode-signed with a self-signed publisher cert (`CN=WinThemeSwitcher Self-Signed`, thumbprint `40E0D1EB58DAC255EB37E9D64FF34448E3D33D12`), and every signed release carries an **RFC 3161 DigiCert timestamp** so the signature outlives the cert's 2036 expiration (v0.3.0/v0.3.1 were retro-timestamped on 2026-07-04). On first install you'll trust the cert — see [Install](#install) step 2. Architecture details and the build/sign workflow are in [CLAUDE.md](CLAUDE.md).

## Features

- **Tiny binary** (~395 KB, nothing else to install) and effectively idle between transitions — event-driven, sleeps on a kernel timer until the next sunrise/sunset, waking only briefly every 10 minutes to check that the system clock hasn't been changed.
- **Reliable theme apply** via the `IThemeManager2` COM interface — the same API the Settings UWP wraps internally. Atomic, in-process, ~200 ms latency, no Settings flash. Two-tier fallback if it ever errors.
- **Catches up after sleep / lock.** A sunrise or sunset that passes while you're suspended or locked reconciles the moment you're back.
- **Follows clock corrections.** If Windows corrects its clock (e.g. booting 3 hours off after another OS on a dual-boot machine), the schedule re-evaluates right away when Windows announces the change, and within 10 minutes at the latest — instead of switching hours late.
- **Respects manual overrides** — changing theme in Settings (or via the tray's **Toggle Theme**) sticks until the next natural sunrise/sunset transition, surviving lock/unlock and sleep/resume. The app only steps in when a transition actually passed while you were away.
- **Recovers from failed applies** — a transition whose apply errors is retried up to 3 times a minute apart (instead of silently waiting for the next transition), and stands down if you change the theme yourself in the meantime.
- **Diagnostic log** at `events.log` next to the exe (rotated past 256 KB) — every transition recorded with cause, target, applied tier, and timing.

## Install

1. Download `win-theme-switcher-vX.Y.Z-windows-x64.zip` from the [latest release](../../releases/latest). Extract it to a folder only your account can write to, e.g. `%LOCALAPPDATA%\Programs\WinThemeSwitcher\` — the app registers itself to run at every login, so a shared folder like `C:\Tools\` would let any other local account swap the exe.
2. **Trust the publisher cert** (one-time, recommended). First check the `.cer` in the zip is the real one — compare against the thumbprint on **this GitHub page**, not a copy inside the download:
   ```powershell
   (Get-PfxCertificate .\WinThemeSwitcher-publisher.cer).Thumbprint   # must be 40E0D1EB58DAC255EB37E9D64FF34448E3D33D12
   Import-Certificate -FilePath .\WinThemeSwitcher-publisher.cer -CertStoreLocation Cert:\CurrentUser\Root
   ```
   Windows shows a security warning when adding a root certificate — that's expected; confirm it.
3. Run `win-theme-switcher.exe`. SmartScreen may say *"Windows protected your PC"* for a new self-signed app — choose **More info → Run anyway**. A half-orange / half-dark-blue circle appears in the notification area; on Windows 11 new icons often start in the overflow (**^**) menu — drag it onto the taskbar to keep it visible.

On first launch the app reads your coordinates via Windows Location. If Location is off or denied, a dialog asks if you want to enable it (opens Settings) or fall back to manual entry (opens `config.json` in an editor). After editing config, right-click tray → **Refresh**.

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

**Custom themes.** Save the theme you want in Settings → Personalization → Themes → **Save**, give it a **unique name**, and point the config at the saved file, e.g. `"%LOCALAPPDATA%\\Microsoft\\Windows\\Themes\\MyNight.theme"` (JSON needs double backslashes).

- Environment variables like `%LOCALAPPDATA%` are expanded (as they were when the app started), and a relative path is resolved against the folder the exe is in.
- Only `.theme` files are used. A theme pack (`.themepack` / `.deskthemepack`) must be installed once by double-clicking it; then use the `.theme` file it installed under `%LOCALAPPDATA%\Microsoft\Windows\Themes\`. `Custom.theme` is refused: it's Windows' unsaved-theme scratch file, rewritten whenever you change any personalization setting.
- A path that can't be used (missing, not a `.theme`, undefined variable, …) falls back to the stock theme; the app says why in a dialog at startup and on Refresh, and logs `theme_path_rejected` at each apply.
- The fast apply path finds themes by display name, so the name must be unique: if another theme file has the same name, the app applies your file the slower way (by path), which briefly opens Settings. (Themes inside `%SystemRoot%\Resources`, Windows' own, are always matched by name — first match wins — so give a copy of one of them a new name.)
- Light/dark modes come from each theme file (`[VisualStyles]` `SystemMode` / `AppMode`), so a theme with light apps and a dark taskbar works. If your two themes have the *same* modes (they differ only in wallpaper), or a theme doesn't declare them (high-contrast themes), the app can't tell from Windows' settings which theme is on: it re-applies at startup, and **Toggle Theme** alternates between the two.

The file may be saved as UTF-8 (with or without BOM) or UTF-16 with a BOM (LE or BE); keys the app doesn't recognize are kept (their order may change when the app rewrites the file). After editing config, right-click tray → **Refresh**. No restart needed.

## Tray menu

- **Toggle Theme** — switches to the other theme (day ↔ night) right now, as a manual override: it sticks (including across lock/unlock and sleep) until the next natural sunrise/sunset transition.
- **Open Config** — opens `config.json` with your `.json` editor, or Notepad if none is set up.
- **Refresh** — re-reads config, retries Windows Location if needed, force-applies the correct theme.
- **Quit** — exits. The auto-start entry persists; set `auto_start: false` and click Refresh (or relaunch) once to remove it.

If two copies are launched, the second shows a notice and exits (single-instance). Fatal startup errors show a dialog and are recorded in `events.log`, as are panics. If the taskbar isn't ready at login, the icon appears as soon as it is; if Windows keeps refusing the icon for about two minutes, a dialog says so (theme switching keeps working without it).

## Verifying the signature

```powershell
Get-AuthenticodeSignature .\win-theme-switcher.exe | Format-List Status, SignerCertificate
```

After [Install](#install) step 2 this shows `Status : Valid` with `SignerCertificate : [Subject] CN=WinThemeSwitcher Self-Signed`. (Before the cert is trusted, `Status` is `UnknownError` — the chain ends in a root Windows doesn't know yet.) A valid signature means the file was signed with this project's key and hasn't been modified since. Self-signed certs can't suppress SmartScreen — that's cloud reputation, which a CA-signed cert builds over time (v0.6.0 row in [Roadmap](#roadmap)).

## Antivirus false positives

With the publisher cert trusted, most AVs accept the signed binary without further action.

**Kaspersky** scores cumulative behavior — `HKCU\Run` persistence, WinRT Geolocation, COM activation of `themeui.dll` — and may still flag a self-signed build. Since v0.3.0 the signed build has run without special handling on the maintainer's machine, but if Kaspersky flags or quarantines it for you: restore it and add a **Trusted application** rule (Settings → Security → Threats and Exclusions → *Specify trusted applications* → tick all five: Do not scan opened files, Do not monitor application activity, Do not inherit restrictions, Do not monitor child application activity, Allow interaction with Kaspersky interface). **The binary is not malicious** — full source is in this repo. CA signing (planned for v0.6.0 via SignPath Foundation's free OSS program) should make this section unnecessary.

## How it works

`apply_theme` is a three-tier fallback:

1. **`IThemeManager2`** (primary, since v0.3.0) — the undocumented-but-stable COM interface in `themeui.dll` that the Settings UWP wraps internally. Atomic, in-process apply. `SetCurrentTheme(idx)` does the `WM_THEMECHANGED` + `WM_SETTINGCHANGE` broadcasts itself.
2. **`ShellExecuteW(.theme)` + commit watcher** — legacy backup if the COM interface ever errors. A 5 s watcher polls the registry to detect silent failures and promotes to tier 3 (when the theme's light/dark modes make the change visible there).
3. **Direct registry write** — last resort. Sets the theme's light/dark modes but not its wallpaper.

Sunrise/sunset times come from the [`sun-times`](https://crates.io/crates/sun-times) crate (no network); above 65.5° latitude, where a day can have no sunrise or sunset, the sun's computed altitude decides instead. The event loop sleeps on `winit`'s `WaitUntil` until the next transition (at most 10 minutes at a time, to notice system-clock corrections). At startup the app registers a resume-from-sleep callback (`PowerRegisterSuspendResumeNotification`), and a worker thread registers for session-unlock notifications (`WTSRegisterSessionNotification`), so the schedule reconciles when you're back.

Full architecture, threading invariants, and the reasoning behind each tier are in [CLAUDE.md](CLAUDE.md).

## Building from source

Requirements: Rust `stable-x86_64-pc-windows-msvc` + Visual Studio Build Tools with the C++ workload.

```powershell
winget install Rustlang.Rustup
winget install Microsoft.VisualStudio.2022.BuildTools --override "--add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"
git clone https://github.com/atefalshehri/WinThemeSwitcher.git
cd WinThemeSwitcher
cargo build --release     # → target\release\win-theme-switcher.exe
cargo test
```

Run cargo from the repo folder: `.cargo\config.toml` (picked up from the current directory) links the Visual C++ runtime statically, so the exe has no `VCRUNTIME140.dll` dependency. Some antivirus products flag freshly built *unsigned* executables; if yours does, exclude the `target\` folder or sign the output with your own code-signing cert.

**Maintainer release builds** use `scripts\build.ps1` (build → verify linker hardening → sign with the project cert + RFC 3161 timestamp → deploy), `scripts\test.ps1` (signs the test binary before running it), and `scripts\release.ps1` (the only thing that writes GitHub releases: CI gate → signed build → verified draft → tag + publish) — all need the project's private signing key, so they only work on the maintainer's machine. Details in [CLAUDE.md → Build](CLAUDE.md#build).

Release profile is tuned for size (`opt-level = "z"`, `lto = true`, `strip = true`, `panic = "abort"`). No `build.rs` — `windows-sys` and `windows` self-link.

## Uninstall

1. Right-click tray → Quit.
2. Delete the install folder.
3. Remove auto-start:
   ```cmd
   reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v WinThemeSwitcher /f
   ```
4. Remove the publisher cert you trusted at install:
   ```powershell
   Remove-Item Cert:\CurrentUser\Root\40E0D1EB58DAC255EB37E9D64FF34448E3D33D12
   Remove-Item Cert:\CurrentUser\TrustedPublisher\40E0D1EB58DAC255EB37E9D64FF34448E3D33D12 -ErrorAction SilentlyContinue   # older install instructions also added it here
   ```

No installer, no uninstaller — it's a single-exe tool by design.

## Roadmap

Ordered by priority. The September 2026 whole-project audit's bug findings shipped in v0.4.1 and its custom-theme follow-ups in v0.5.0; the rest are listed under [Foundation](#foundation) and planned for the next minor, alongside CA signing.

### Release plan

Versioning before 1.0: a **patch** (0.x.y → 0.x.y+1) is pure bug fixes; a **minor** (0.x → 0.x+1) is anything that adds surface (menu items, config fields) or changes behavior. v1.0 is a gate, not a feature drop. Numbers are assigned at ship time — the 0.6/0.7 order can swap if the SignPath approval wait stalls, and patch releases slot in anywhere.

| Version | Type | Contents |
|---|---|---|
| **v0.3.2** | patch | **Shipped 2026-07-04.** Sunrise/sunset day-bracketing fix (wrong solar day in UTC+13/+14, missed post-midnight sunsets) + scheduling tests + CI gate; `config.json` never overwritten on a parse error (error is logged + shown in a non-blocking dialog, empty file self-heals, autostart setting survives a broken file); `auto_start: false` now actually removes the Run entry |
| **v0.4.0** | minor | **Shipped 2026-09-01.** Preserve manual overrides across lock/unlock/resume (time-based reconciliation — overrides survive any-length sleeps; missed transitions still reconcile); "Toggle Theme" tray item; fail-loudly bundle (panic hook, fatal-error MessageBox, wake-listener logging + WTS registration retry, single-instance mutex, bounded apply retry with user-intervention stand-down); `.theme`-name resolution + tick-decision unit tests (40 total); `scripts\test.ps1` + `scripts\build.ps1` (sign test and release binaries before their first execution — unsigned fresh builds were being locked by Kaspersky's cloud scanner, which is what stalled this release for weeks) |
| **v0.4.1** | patch | **Shipped 2026-09-27.** Fixes from a whole-project audit: follow system-clock corrections (dual-boot machines booted hours off and switched at the wrong time); resume-from-sleep notification actually registers (the `power_register code=87` log line — it had never worked since v0.2.0); no more `VCRUNTIME140.dll` dependency (the exe failed to start on PCs without the VC++ redistributable) plus DLL-search hardening; polar latitudes (polar-day onset stuck Dark for months; now one astronomy model above 65.5°, no phantom or flickering transitions); early timer fires can't undo an override; failed registry writes are detected and retried; Refresh in a broken-config session no longer re-enables autostart; Open Config falls back to Notepad; `config.json` saved atomically (hard links kept), UTF-16/BOM files accepted, unknown keys preserved; `.theme` DisplayName read with Windows' own INI reader (non-English/UTF-16 names match); tray-event memory leak; `scripts\test.ps1` never reported a failure (GUI-subsystem test binary) — fixed, and both scripts hardened; CI actions on Node 24 and a linker-hardening check; docs corrected (83 tests) |
| **v0.5.0** | minor | **Shipped 2026-09-28.** Release pipeline: `scripts\release.ps1` is the only writer of GitHub releases (CI gate on the exact commit → signed build → private draft → download-and-verify every asset → tag + publish); the tag-triggered `release.yml` that could overwrite signed assets with unsigned CI builds is gone; CI actions pinned to commit SHAs with read-only permissions. **First non-prerelease release** — `/releases/latest` works. Custom themes: light/dark modes read from each theme file (mixed apps/taskbar themes, same-mode theme pairs, and high-contrast themes handled); tier 1 won't apply a different theme that shares the name; `%VAR%` and exe-relative paths; only `.theme` files accepted, with a dialog saying why a path isn't used; a theme without a `DisplayName` is found by its file name. `tray-icon` 0.25: no more exit at login when the taskbar isn't ready, and a missing icon is detected, retried, and reported (106 tests). |
| **v0.6.0** | minor | SignPath Foundation CA signing in CI; submit the first CA-signed binary to Microsoft Defender + Kaspersky for reputation seeding. **Retires the §Antivirus false positives section** if CA-signed builds pass without it. Also unlocks the winget submission (§Release & distribution §3) — winget's validation pipeline runs its own AV scans, and a CA-signed + Defender-submitted binary is what gets through. Plus the remaining audit follow-ups under [Foundation](#foundation). Candidate point to launch the personal Scoop bucket (`persist` fits the exe-relative config model better than winget's symlink layout — see §Release & distribution §4). |
| **v0.7.0** | minor | Live tray tooltip ("Dark until 06:12", "Location needed — click Refresh", or a degraded-apply warning); "Open Log" menu item; MessageBox when a user-initiated Refresh fails (scheduled ticks stay silent-to-log); first-run location retry without requiring Refresh; `offset_sunrise_min` / `offset_sunset_min` config fields (sun-anchored, so still compatible with "no custom times") |
| **v1.0.0** | gate | Cut when **all** of: winget accepts the package; §Foundation has no open items and there are no known unfixed correctness bugs; a full release cycle has shipped CA-signed with no AV flags; **§Antivirus false positives is deletable**; and §Contributing says "Stable for personal use and distribution". Net effect: a stranger can install WinThemeSwitcher without reading the Kaspersky section. |

### Correctness fixes

A history of bugs that shipped in a release and the version that fixed them.

- **v0.5.0 — Custom themes with mixed modes looked already applied.** The "already in sync" check read only the taskbar's light/dark value, so a day theme with a dark taskbar (or any two themes whose taskbar modes matched) made a sunset look done. Both values are now compared against what each theme file declares; when they can't tell the two themes apart, the app no longer skips as "in sync". Coverage: `mixed_target_needs_both_values_to_match`, `partial_registry_write_is_retried_not_mistaken_for_the_user`, `undecidable_sync_never_skips_as_in_sync_but_still_preserves_and_stands_down`.
- **v0.5.0 — Tier 1 could apply the wrong theme.** Themes are matched by display name; a custom theme sharing its name with another theme (e.g. an edited copy of a saved theme) applied that other theme and reported success. Now such a theme is applied by its exact path. Coverage: `pick_theme_index_table`, `name_taken_elsewhere_ignores_the_configured_file_itself`.
- **v0.5.0 — App could exit at login.** `tray-icon` 0.19 failed to build the tray when the taskbar wasn't ready yet, which ended the app with a fatal-error dialog. `tray-icon` 0.25 waits for the taskbar; the app also detects an icon Windows refused and retries it.
- **v0.4.1 — Wrong theme after a system-clock correction.** The wait for the next transition was armed once from the wall clock and never re-derived, so when Windows stepped its clock (on a dual-boot machine: booting 3 h behind, then time sync) the decision made from the wrong time stood and the next switch fired up to hours late. Now every wake compares the wall clock against the monotonic clock and re-evaluates on a step; a 10-minute heartbeat guarantees a check even if no clock-change message arrives. Coverage: `clock_step_forward_after_skewed_boot_ticks_immediately`, `backward_step_across_a_transition_reapplies_not_preserves`.
- **v0.4.1 — Resume-from-sleep hook never registered (since v0.2.0).** `PowerRegisterSuspendResumeNotification` only accepts callback registration; called with a window handle it failed with `ERROR_INVALID_PARAMETER` (87) on every launch. Unlock events masked it. Now registers a callback (`DEVICE_NOTIFY_CALLBACK`), as documented.
- **v0.4.1 — Exe needed the VC++ redistributable.** Dynamically linked `VCRUNTIME140.dll` isn't part of Windows; on a clean PC the app didn't start at all. Now linked statically (Microsoft's hybrid-CRT pattern).
- **v0.4.1 — Polar-day onset stuck on Dark.** After the last sunset before the midnight sun, `sun_times` reports no following sunrise, so the app stayed Dark for ~2 months. Above 65.5° latitude the schedule now follows the sun's computed altitude alone, so polar-day/night boundaries switch at the real sunrise/sunset and there are no phantom or flickering transitions. Coverage: `tromso_after_last_sunset_before_midnight_sun_is_not_dark_for_months`, `schedule_contract_holds_around_polar_edges`.
- **v0.4.0 — Manual overrides reverted on unlock.** v0.2.0–v0.3.2 re-applied the schedule on every session unlock whenever the screen differed, reverting a theme picked in Settings. v0.4.0's time-based `decide_tick` preserves it until a transition actually passes. Coverage: `wake_before_next_transition_preserves_override`, `wake_after_even_number_of_missed_transitions_reconciles`.
- **v0.4.0 — Failed applies were never retried.** A failed apply waited for the next transition, unlock, or Refresh (up to ~12 h). v0.4.0 adds a bounded 3 × 60 s retry that stands down if the user intervenes. Coverage: `failed_apply_then_wake_must_reapply_not_preserve`, `retry_budget_is_bounded_and_resets_per_episode`.
- **v0.3.2 — `config.json` parse errors wiped settings.** A parse error (e.g. a single-backslash path) silently overwrote the file with defaults, losing coordinates and theme paths. Coverage: `config_parse_error_is_reported_and_file_kept`.
- **v0.3.2 — Wrong solar day** in UTC+13/+14 (permanently dark) and missed post-midnight sunsets. Coverage: `apia_noon_is_light`, `reykjavik_june_sunset_crosses_midnight`.

### Foundation

- **Audit follow-ups (September 2026) that change behavior — planned for v0.6.0** (the custom-theme, `tray-icon` and Actions-pinning items shipped in v0.5.0):
  - the tier-2 Settings-window closer can close a Settings window you had open (and never matches on non-English Windows) and closes before the theme commits;
  - Windows Location lookup blocks startup/Refresh for up to a minute when slow;
  - autostart is registered from wherever the exe runs (even a temp folder);
  - coordinates aren't range-checked (a typo silently becomes 0.0);
  - saved coordinates keep full sensor precision;
  - an application manifest (themed dialogs, DPI-aware) and a version resource;
  - a double Toggle during a pending apply retry is reverted by the retry (the stand-down only sees an odd number of toggles);
  - tier 2's settle sleep + taskbar poke and tier 3's system-wide broadcast run on the event-loop thread (the tray can freeze for seconds on legacy applies);
  - first-run location dialogs block before the tray icon exists, and their text refers to it.
- **Remaining known gaps** (accepted, documented in CLAUDE.md): the bounded apply retry covers total apply failure only — a tier-2 `ShellExecute` silent-fail is recovered by `commit_watcher`'s registry fallback instead, and can't be detected at all when the target's light/dark modes already match the screen or the theme file doesn't declare them; a Toggle within ~5 s of a tier-2 apply can be reverted by that apply's still-running commit watcher (unreachable while tier 1 is healthy); and with two same-mode themes, a theme changed in Settings is invisible to the app — Toggle alternates by what the app last applied, and a pending retry can't see it to stand down.

### Release & distribution (dependency chain, in order)

1. **Harden the release process:**
   - **1a. Done (v0.3.x, automated in v0.4.0)**: RFC 3161 timestamp countersignature on every signed build (`/tr http://timestamp.digicert.com /td SHA256` — without it, signatures die when the cert expires in 2036), added 2026-07-04; `scripts\build.ps1` builds, signs, verifies, and deploys in one step.
   - **1b. Done in v0.5.0**: `scripts\release.ps1` is the only writer of releases — CI must have passed on the exact commit; the signed build goes into a private draft; every asset is downloaded back and its hash, zip contents and signature verified; only then is the tag pushed and the release published. The tag-triggered `release.yml` (which re-ran into unsigned assets) was deleted. Replaces the manual build → `gh release upload` ritual that let the v0.3.0 "shipped unsigned for months" failure happen.
   - **1c. Done in v0.4.1**: the recurring `wake_listener_err stage=power_register code=87` was a wrong API — fixed (see [Correctness fixes](#correctness-fixes)).
   - **1d. Done in v0.5.0**: releases are no longer marked *prerelease*; `/releases/latest` works, which unblocks winget automation.
2. **CA-signed releases** via [SignPath Foundation's free OSS program](https://signpath.org) — signing moves into CI, so releases no longer depend on the maintainer's machine and local key. Reduces SmartScreen prompts over time via cert reputation (no cert eliminates them outright). Azure Trusted Signing is not an option: individual validation is US/Canada-only.
3. **Submit the first CA-signed binary to Microsoft Defender** — this gates winget, whose validation pipeline runs AV scans — and to Kaspersky. Repeat only if a specific release gets flagged.
4. **winget package** (`InstallerType: portable`), only after steps 1–2 make asset hashes final at publish time. A personal Scoop bucket may come earlier: Scoop's `persist` mechanism fits the exe-relative config model better than winget's symlink layout.

### UX polish

- **Live tray tooltip** — "Dark until 06:12", "Location needed — click Refresh", or a degraded-apply warning — plus an **"Open Log"** menu item and a MessageBox when a user-initiated Refresh fails (scheduled ticks stay silent-to-log).
- **First-run without the Refresh step** — retry Windows Location a few times right after the enable-Location dialog is dismissed, instead of a `Geolocator::StatusChanged` subscription. Bigger optional variant: re-query location on unlock/resume so a traveler's coordinates don't go stale (today they are never re-read once set).
- **Optional sunrise/sunset offsets** (`offset_sunrise_min` / `offset_sunset_min` in config) — still sun-anchored, so compatible with "no custom times".

### Maintenance notes (not scheduled)

- **winit `ApplicationHandler` migration** — only when bumping to winit 0.31 (the pinned 0.30 merely deprecates `EventLoop::run`, allowed at the call site; nothing forces this today). Worth evaluating at that point: dropping winit for a plain Win32 message loop — the pattern already exists in the wake listener.

**Not planned**: GUI configuration (`config.json` + Refresh is the UX), custom wake times (sunrise/sunset is the whole point; offsets from them are fine), cross-platform (Windows only — macOS already has this natively), in-app update check (it would re-add the autorun+beacon AV-heuristic surface the `IThemeManager2` migration removed — winget/Scoop handle upgrades), pause/snooze toggle (a manual override already pauses until the next transition), and ADM-style scripting/hotkeys/battery rules (out of scope for a small tray tool).

## Contributing

**Status: stable for personal use.** The tray/apply/scheduling core is mature (106 unit tests, 0 compiler/clippy warnings, the deployed binary verified daily for months). Working toward distribution polish — winget submission (gated on v0.6.0's CA-signed cert), a Scoop bucket, and retiring the §Antivirus false positives section. Issues and PRs welcome — please open an issue to discuss larger changes before sending a patch. Focus areas: see [Roadmap](#roadmap).

## License

MIT — see [LICENSE](LICENSE).

## Acknowledgments

- [`sun-times`](https://crates.io/crates/sun-times) — local sunrise/sunset math.
- [`chrono`](https://crates.io/crates/chrono) — date/time handling; [`serde`](https://crates.io/crates/serde) + [`serde_json`](https://crates.io/crates/serde_json) — config persistence.
- [`tray-icon`](https://crates.io/crates/tray-icon) + [`winit`](https://crates.io/crates/winit) — tray icon and event loop.
- [`windows-sys`](https://crates.io/crates/windows-sys) + [`windows`](https://crates.io/crates/windows) — official Microsoft Rust bindings.
