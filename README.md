# WinThemeSwitcher

Automatically swap between two Windows 11 **themes** at local sunrise and sunset — macOS's auto-theme behavior, on Windows 11. Full theme swap (wallpaper + colors + light/dark mode), not just a DWORD toggle.

> **Signed releases since v0.3.0** (v0.1.0/v0.2.0 are unsigned). Binaries are Authenticode-signed with a self-signed publisher cert (`CN=WinThemeSwitcher Self-Signed`, thumbprint `40E0D1EB58DAC255EB37E9D64FF34448E3D33D12`), and every signed release carries an **RFC 3161 DigiCert timestamp** so the signature outlives the cert's 2036 expiration (v0.3.0/v0.3.1 were retro-timestamped on 2026-07-04). On first install you'll trust the cert — see [Install](#install) step 2. Architecture details and the build/sign workflow are in [CLAUDE.md](CLAUDE.md).

## Features

- **Tiny binary** (~470 KB, nothing else to install) and effectively idle between transitions — event-driven, sleeps on a kernel timer until the next sunrise/sunset, waking only briefly every 10 minutes to check that the system clock hasn't been changed.
- **Reliable theme apply** via the `IThemeManager2` COM interface — the same API the Settings UWP wraps internally. Atomic, in-process, ~200 ms latency, no Settings flash. Two-tier fallback if it ever errors or doesn't take effect.
- **Catches up after sleep / lock.** A sunrise or sunset that passes while you're suspended or locked reconciles the moment you're back.
- **Follows clock corrections.** If Windows corrects its clock (e.g. booting 3 hours off after another OS on a dual-boot machine), the schedule re-evaluates right away when Windows announces the change, and within 10 minutes at the latest — instead of switching hours late.
- **Respects manual overrides** — changing theme in Settings (or via the tray's **Toggle Theme**) sticks until the next natural sunrise/sunset transition, surviving lock/unlock and sleep/resume but not a restart, sign-out or reboot: when the app starts, it applies the scheduled theme. Otherwise the app only steps in when a transition actually passed while you were away.
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

   **Smart App Control** is a separate check. When it's **On** (Windows Security → App & browser control → Smart App Control settings), Windows blocks this self-signed build outright: there's no *Run anyway*, no way to allow a single app, and the certificate from step 2 doesn't count. To use the app, turn Smart App Control off (whether it can be turned back on later depends on your Windows build), or wait for the CA-signed builds planned for v0.7.0. In *Evaluation* mode it blocks nothing yet, but Windows may later switch it on by itself, and from then on the app is blocked, including its start at login.

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

- **Toggle Theme** — switches to the other theme (day ↔ night) right now, as a manual override: it sticks (including across lock/unlock and sleep) until the next natural sunrise/sunset transition or the next app start.
- **Open Config** — opens `config.json` with your `.json` editor, or Notepad if none is set up.
- **Refresh** — re-reads config, retries Windows Location if needed, force-applies the correct theme.
- **Quit** — exits. The auto-start entry persists; set `auto_start: false` and click Refresh (or relaunch) once to remove it.

If two copies are launched, the second shows a notice and exits (single-instance). Fatal startup errors show a dialog and are recorded in `events.log`, as are panics. If the taskbar isn't ready at login, the icon appears as soon as it is; if Windows keeps refusing the icon for about two minutes, a dialog says so (theme switching keeps working without it).

## Verifying the signature

```powershell
Get-AuthenticodeSignature .\win-theme-switcher.exe | Format-List Status, SignerCertificate
```

After [Install](#install) step 2 this shows `Status : Valid` with `SignerCertificate : [Subject] CN=WinThemeSwitcher Self-Signed`. (Before the cert is trusted, `Status` is `UnknownError` — the chain ends in a root Windows doesn't know yet.) A valid signature means the file was signed with this project's key and hasn't been modified since. Self-signed certs can't suppress SmartScreen — that's cloud reputation, which a CA-signed cert builds over time (v0.7.0 row in [Roadmap](#roadmap)). Nor does a self-signed cert get past Smart App Control, which goes by cloud reputation or a signature from a CA in Microsoft's Trusted Root Program, never a root you added yourself ([Install](#install) step 3).

Releases are **immutable** on GitHub (v0.3.2 and later): their files and tag can't be changed, and GitHub attaches a release attestation. With [GitHub CLI](https://cli.github.com/) **2.93 or later** (earlier versions leak your login token during this check — [GHSA-8xvp-7hj6-mcj9](https://github.com/cli/cli/security/advisories/GHSA-8xvp-7hj6-mcj9)), signed in with `gh auth login`, you can confirm that a downloaded file is exactly the one published (this checks the file's contents, not its signature):

```powershell
gh release verify-asset vX.Y.Z .\win-theme-switcher.exe -R atefalshehri/WinThemeSwitcher
```

## Antivirus false positives

With the publisher cert trusted, most AVs accept the signed binary without further action.

**Kaspersky** scores cumulative behavior — `HKCU\Run` persistence, WinRT Geolocation, COM activation of `themeui.dll` — and may still flag a self-signed build. Since v0.3.0 the signed build has run without special handling on the maintainer's machine, but if Kaspersky flags or quarantines it for you: restore it and add a **Trusted application** rule (Settings → Security → Threats and Exclusions → *Specify trusted applications* → tick all five: Do not scan opened files, Do not monitor application activity, Do not inherit restrictions, Do not monitor child application activity, Allow interaction with Kaspersky interface). **The binary is not malicious** — full source is in this repo. CA signing (planned for v0.7.0, with a Certum Open Source certificate) should make this section unnecessary.

## How it works

`apply_theme` is a three-tier fallback:

1. **`IThemeManager2`** (primary, since v0.3.0) — the undocumented-but-stable COM interface in `themeui.dll` that the Settings UWP wraps internally. Atomic, in-process apply. `SetCurrentTheme(idx)` does the `WM_THEMECHANGED` + `WM_SETTINGCHANGE` broadcasts itself. Its success is checked (since v0.5.1): if Windows' light/dark values don't reach the theme's within 2 s, the apply counts as failed and tier 2 takes over. The check needs a visible change, so it's skipped when the theme doesn't declare its modes or the screen already shows them.
2. **`ShellExecuteW(.theme)` + commit watcher** — the slower path, which briefly opens Settings: the backup when tier 1 errors or its apply doesn't take effect, and the path every switch takes for a custom theme that Settings → Personalization → Themes doesn't list, or whose name another theme file shares ([Configuration](#configuration)). A 5 s watcher polls the registry to detect silent failures and promotes to tier 3 (when the theme's light/dark modes make the change visible there).
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

Run cargo from the repo folder: `.cargo\config.toml` (picked up from the current directory) links the Visual C++ runtime statically, so the exe has no `VCRUNTIME140.dll` dependency. Some antivirus products flag freshly built *unsigned* executables; if yours does, exclude the `target\` folder or sign the output with your own code-signing cert. Smart App Control can't be handled that way: when it's on, it blocks the build scripts and tests that cargo compiles and runs, so building needs it off.

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

Ordered by priority; re-evaluated 2026-10-31. The September 2026 whole-project audit's bug findings shipped in v0.4.1 and its custom-theme follow-ups in v0.5.0, and v0.5.1 fixes bugs found in an October 2026 review. What's left of both is under [Foundation](#foundation) — behavior fixes in v0.6.0, then the tier-2 fixes as a patch — and, for tooling and tests, under [Maintenance notes](#maintenance-notes). CA signing follows in v0.7.0.

### Release plan

Versioning before 1.0: a **patch** (0.x.y → 0.x.y+1) is pure bug fixes; a **minor** (0.x → 0.x+1) is anything that adds surface (menu items, config fields) or changes behavior. v1.0 is a gate, not a feature drop. Numbers are assigned at ship time, and patch releases slot in anywhere.

| Version | Type | Contents |
|---|---|---|
| **v0.3.2** | patch | **Shipped 2026-07-04.** Sunrise/sunset day-bracketing fix (wrong solar day in UTC+13/+14, missed post-midnight sunsets) + scheduling tests + CI gate; `config.json` never overwritten on a parse error (error is logged + shown in a non-blocking dialog, empty file self-heals, autostart setting survives a broken file); `auto_start: false` now actually removes the Run entry |
| **v0.4.0** | minor | **Shipped 2026-09-01.** Preserve manual overrides across lock/unlock/resume (time-based reconciliation — overrides survive any-length sleeps; missed transitions still reconcile); "Toggle Theme" tray item; fail-loudly bundle (panic hook, fatal-error MessageBox, wake-listener logging + WTS registration retry, single-instance mutex, bounded apply retry with user-intervention stand-down); `.theme`-name resolution + tick-decision unit tests (40 total); `scripts\test.ps1` + `scripts\build.ps1` (sign test and release binaries before their first execution — unsigned fresh builds were being locked by Kaspersky's cloud scanner, which is what stalled this release for weeks) |
| **v0.4.1** | patch | **Shipped 2026-09-27.** Fixes from a whole-project audit: follow system-clock corrections (dual-boot machines booted hours off and switched at the wrong time); resume-from-sleep notification actually registers (the `power_register code=87` log line — it had never worked since v0.2.0); no more `VCRUNTIME140.dll` dependency (the exe failed to start on PCs without the VC++ redistributable) plus DLL-search hardening; polar latitudes (polar-day onset stuck Dark for months; now one astronomy model above 65.5°, no phantom or flickering transitions); early timer fires can't undo an override; failed registry writes are detected and retried; Refresh in a broken-config session no longer re-enables autostart; Open Config falls back to Notepad; `config.json` saved atomically (hard links kept), UTF-16/BOM files accepted, unknown keys preserved; `.theme` DisplayName read with Windows' own INI reader (non-English/UTF-16 names match); tray-event memory leak; `scripts\test.ps1` never reported a failure (GUI-subsystem test binary) — fixed, and both scripts hardened; CI actions on Node 24 and a linker-hardening check; docs corrected (83 tests) |
| **v0.5.0** | minor | **Shipped 2026-09-28.** Release pipeline: `scripts\release.ps1` is the only writer of GitHub releases (CI gate on the exact commit → signed build → private draft → download-and-verify every asset → tag + publish); the tag-triggered `release.yml` that could overwrite signed assets with unsigned CI builds is gone; CI actions pinned to commit SHAs with read-only permissions. **First non-prerelease release** — `/releases/latest` works. Custom themes: light/dark modes read from each theme file (mixed apps/taskbar themes, same-mode theme pairs, and high-contrast themes handled); tier 1 won't apply a different theme that shares the name; `%VAR%` and exe-relative paths; only `.theme` files accepted, with a dialog saying why a path isn't used; a theme without a `DisplayName` is found by its file name. `tray-icon` 0.25: no more exit at login when the taskbar isn't ready, and a missing icon is detected, retried, and reported (106 tests). |
| **v0.5.1** | patch | **Shipped 2026-10-31.** Fixes from an October 2026 review: a tier-1 apply that reports success but changes nothing is caught (the light/dark values are read back) and handed to the fallback tiers, instead of being logged as applied and then kept as if it were your override; a `config.json` saved at the moment the app creates or repairs it is read, never overwritten with defaults; `.theme` files on paths of 260+ characters are read correctly (their name and modes used to read as missing); a theme path typed with an 8.3 name, a trailing dot or other letter case is used under the file's real name. Clippy-clean on Rust 1.99 (120 tests). |
| **v0.6.0** | minor | Behavior fixes under [Foundation](#foundation): coordinates range-checked and saved at sensible precision; a double Toggle no longer reverted by a pending retry; no second copy when the first runs elevated; an application manifest (themed, DPI-aware dialogs), a version resource and the version in `events.log`; autostart refused from a temp folder and left as it was when an empty `config.json` is repaired, a read-only exe folder reported, missing registry keys created, and an entry disabled in Startup settings logged; the Windows Location lookup capped at about 15 s, and first-run dialogs that say what actually happened; themes on network shares refused with a reason; third-party license notices in the release. Starts with the `windows-sys` 0.61 update. |
| **v0.6.1** | patch | The tier-2 path, which every switch takes for a custom theme Settings doesn't list or whose name another theme file shares: Settings is recognized by its process, and only a Settings window the switch opened is closed, after the theme commits; tier 2's and tier 3's waits and broadcasts move off the event-loop thread, and the taskbar poke reaches every taskbar on three or more monitors. Becomes v0.7.1 if v0.7.0 ships first. |
| **v0.7.0** | minor | CA signing with a [Certum Open Source](https://www.certum.eu/en/code-signing-certificates/) certificate (chosen 2026-10-31), still signed on the maintainer's machine by `scripts\release.ps1`. It lets the exe run with Smart App Control On; SmartScreen reputation still grows only with downloads. Retires the self-signed root ([Install](#install) step 2 and the `.cer` asset), adds a Privacy section, and the first CA-signed build is tested on a fresh Windows 11 VM with Smart App Control On before it's published. After publishing: the winget package and the personal Scoop bucket (§Release & distribution §4). |
| **v0.8.0** | minor | First run without waiting: the tray appears first, Windows Location is looked up in the background, and the theme applies as soon as Location is turned on, without Refresh. A tray that reports status: live tooltip ("Dark until 06:12", "Location needed", or a degraded-apply warning), an "Open Log" menu item, and a dialog when a Refresh or Toggle you clicked fails (scheduled ticks stay silent-to-log). Also the winit 0.31 migration and the `windows` 0.62 and `tray-icon` 0.26 updates, since startup is rebuilt anyway. |
| **v1.0.0** | gate | Cut when the two latest releases both: (a) are CA-signed, show `Valid` on a PC with no imported roots, and start on a fresh Windows 11 VM with Smart App Control On; (b) get no detection from Microsoft Defender, Kaspersky or winget's installer scan, at publish and 7 days later, and run under Kaspersky's default settings with no trusted-application rule (a detection by any other VirusTotal engine gets a false-positive report to its vendor and doesn't block v1.0); (c) have a merged winget manifest, with install, upgrade (after Quit) and uninstall passed in Windows Sandbox and `config.json` kept; and (d) leave [Foundation](#foundation) with nothing planned and no known correctness bug unfixed ([Known limitations](#known-limitations) holds only by-design gaps, each with its reason). A SmartScreen "unrecognized app" prompt that names the verified publisher isn't a blocker. The v1.0 commit deletes §Antivirus false positives and sets §Contributing to "Stable for personal use and distribution". Net effect: a stranger can install WinThemeSwitcher without reading the Kaspersky section. |
| *after v1.0* | — | Not scheduled: optional sunrise/sunset offsets, re-reading location on wake, and applying a theme by path through `IThemeManager2` ([UX polish](#ux-polish)). |

### Correctness fixes

A history of bugs that shipped in a release and the version that fixed them.

- **v0.5.1 — A tier-1 apply that silently did nothing would have counted as success.** The app trusted `IThemeManager2::SetCurrentTheme` when it reported success. On a Windows build where it succeeds without applying, the log would have said `applied=theme-manager2`, the tick would have been recorded as in sync, and later unlocks would have kept the wrong theme as if you'd chosen it. No released Windows build has done this so far. The light/dark values are now read back for up to 2 s; if they never reach the theme's, `theme_manager2_noop` is logged and the fallback tiers apply the theme. Coverage: `readback_catches_a_tier1_success_that_changed_nothing`, `readback_trusts_tier1_when_the_modes_cant_tell`.
- **v0.5.1 — `config.json` saved at the wrong moment could be overwritten with defaults.** If the file appeared between the app finding it missing and writing the defaults (an editor saving via delete-and-rename), the app ran on the defaults as if it had read them, and its next save, after finding your location, wrote them over your file. Repairing an empty file had the same window. Now the file is read again instead, the repair writes only while the file is still empty, and a file another program is still writing is reported as unreadable rather than reset. Coverage: `first_run_never_answers_a_file_that_just_appeared_with_defaults`, `the_empty_file_heal_never_overwrites_a_save_that_landed_first`, `an_empty_config_another_program_is_writing_is_unreadable_not_reset`.
- **v0.5.1 — Long theme paths read as empty.** With a path of 260 or more characters, Windows' INI reader couldn't open the `.theme` file (the exe isn't long-path aware), so the theme's name and light/dark modes read as missing: the wrong modes for a mixed theme, and tier 1 couldn't find it by name. Coverage: `theme_files_on_long_paths_are_read_not_taken_as_keyless`.
- **v0.5.1 — Theme paths typed with an 8.3 name, a trailing dot or other letter case.** Windows opens `NIGHTT~1.THE` or `Night.theme.` as `Night.theme`, but a theme without a `DisplayName` is named after its file, so tier 1 looked for the typed spelling and the theme went to the slower path that briefly opens Settings. The file's stored name is used now. Coverage: `a_respelled_theme_path_takes_the_file_name_stored_on_disk`.
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

- **Planned for v0.6.0 — behavior fixes** (the September 2026 audit's remaining follow-ups and the October 2026 review's; the audit's custom-theme, `tray-icon` and Actions-pinning items shipped in v0.5.0):
  - coordinates aren't range-checked: an out-of-range value (225 for 22.5) gives a schedule that's hours off, and an out-of-range latitude can swap day and night; a single missing or misspelled coordinate key counts as 0.0 (a value that isn't a number is already reported as a config error, and a file with no coordinates at all keeps meaning "ask Windows Location");
  - saved coordinates keep full sensor precision (2 decimals is plenty; 3 from 60° on, as a margin: from about 65.7°, where polar day and night begin, a small change can move their start by tens of minutes);
  - a double Toggle during a pending apply retry is reverted by the retry (the stand-down only sees an odd number of toggles); the fix starts by making the step from a tick's apply outcome to its saved state and next deadline a pure, tested function;
  - if the first copy runs elevated, a second copy can still start (the single-instance check can't open the elevated copy's lock, and carries on);
  - there's no application manifest (themed dialogs; DPI awareness for the dialogs shown before the tray exists) and no version resource, and the running version is never logged;
  - autostart is registered from wherever the exe runs (even a temp folder); a read-only exe folder fails silently; the Run and Personalize registry keys are opened but never created when missing; an entry disabled in Startup settings goes unexplained; repairing an empty or NUL-filled `config.json` turns autostart back on even if you had turned it off;
  - a slow Windows Location lookup blocks startup and Refresh for up to a minute (v0.6.0 caps the wait; v0.8.0 moves the lookup off the startup path);
  - the first-run location dialogs can say untrue things (that Settings opened when it didn't; "off or not allowed" after a timeout), and they refer to a tray icon that doesn't exist yet;
  - a theme on an unreachable network share freezes the tray for 20–40 s each time its path is checked;
  - the release ships without third-party license notices.
- **Planned for v0.6.1 — the tier-2 path** (pure bug fixes; tier 2 runs whenever tier 1 can't apply the theme, and for a custom theme Settings doesn't list, or whose name another theme file shares, that's every switch):
  - the Settings-window closer can close a Settings window you had open, never matches on non-English Windows (so Settings stays open after each switch), and closes before the theme commits;
  - tier 2's settle sleep and taskbar poke and tier 3's system-wide broadcast run on the event-loop thread (the tray can freeze for seconds on legacy applies), and with three or more monitors the taskbar poke reaches only two of the taskbars.

### Known limitations

Accepted by design (details in CLAUDE.md); they don't block v1.0:

- the bounded apply retry covers total apply failure only — a tier-2 `ShellExecute` silent-fail is recovered by `commit_watcher`'s registry fallback instead, and can't be detected at all when the target's light/dark modes already match the screen or the theme file doesn't declare them;
- a Toggle within ~5 s of a tier-2 apply can be reverted by that apply's still-running commit watcher (it takes a Toggle inside the watcher's 5 s; an optional part of v0.6.1 would close it);
- with two same-mode themes, a theme changed in Settings is invisible to the app — Toggle alternates by what the app last applied, and a pending retry can't see it to stand down;
- the fast path (tier 1) finds a theme by its display name, since Windows' theme interface doesn't expose file paths, so it applies a custom theme only when Settings → Personalization → Themes lists it under a unique name; any other custom theme takes the slower path, which briefly opens Settings (Windows' own themes are always matched by name, first match wins);
- when tier 1 reports success but the light/dark values don't follow, the app waits 2 s to be sure before falling back — that wait is what catches a tier-1 apply that changed nothing — and the tray doesn't respond during it;
- at polar latitudes, an always-on PC can go weeks with no transition, lock, sleep or Refresh, long enough for ordinary clock drift to pass the 60 s the app treats as a clock correction; each time that happens, it costs one extra, harmless check.

### Release & distribution (dependency chain, in order)

1. **Harden the release process:**
   - **1a. Done (v0.3.x, automated in v0.4.0)**: RFC 3161 timestamp countersignature on every signed build (`/tr http://timestamp.digicert.com /td SHA256` — without it, signatures die when the cert expires in 2036), added 2026-07-04; `scripts\build.ps1` builds, signs, verifies, and deploys in one step.
   - **1b. Done in v0.5.0**: `scripts\release.ps1` is the only writer of releases — CI must have passed on the exact commit; the signed build goes into a private draft; every asset is downloaded back and its hash, zip contents and signature verified; only then is the tag pushed and the release published. The tag-triggered `release.yml` (which re-ran into unsigned assets) was deleted. Replaces the manual build → `gh release upload` ritual that let the v0.3.0 "shipped unsigned for months" failure happen.
   - **1c. Done in v0.4.1**: the recurring `wake_listener_err stage=power_register code=87` was a wrong API — fixed (see [Correctness fixes](#correctness-fixes)).
   - **1d. Done in v0.5.0**: releases are no longer marked *prerelease*; `/releases/latest` works, which unblocks winget automation.
   - **1e. Done 2026-09-28**: GitHub *immutable releases* enabled — a published release's files and tag are locked and carry a release attestation (`gh release verify-asset`); `scripts\release.ps1` requires the setting and verifies the attestation after publishing. Releases published before the setting was enabled became immutable when next edited: v0.3.2–v0.5.0 were converted the same day after their files were re-verified (v0.3.2's and v0.4.1's stale "unsigned" notes corrected); v0.1.0–v0.3.1 were left as they are.
   - **1f. Done 2026-10-31**: CI's format and clippy checks became hard gates, and clippy now checks the test code too; a `cargo-deny` job checks every dependency against the RustSec advisory database and a license allow-list; `scripts\release.ps1` refuses to start unless the publisher certificate is trusted on the machine.
2. **CA-signed releases** with a [Certum Open Source](https://www.certum.eu/en/code-signing-certificates/) code-signing certificate (chosen 2026-10-31), still signed on the maintainer's machine by `scripts\release.ps1`. What it buys is Smart App Control: a signature from a CA in Microsoft's Trusted Root Program lets the exe run with it On. SmartScreen keeps warning until downloads build reputation — no certificate skips that, and there's no way to submit a file for it — but the prompt names a verified publisher. The certificate is issued to the maintainer by name and lasts a year. For the first CA-signed release:
   - development and test builds are signed with a new self-signed certificate that only the maintainer's PC trusts (no PIN or one-time code per build), and the current publisher key is retired;
   - `release.ps1` checks the certificate's subject, issuer, chain and timestamp instead of one pinned thumbprint (the self-signed releases from v0.5.0 on re-verify with their own tag's script, which needs the old root still trusted; earlier ones are checked by hand);
   - new releases drop the `.cer` asset and [Install](#install) step 2, and their notes tell existing users how to remove the old root;
   - a Privacy section in this README: what Windows Location sends to Microsoft, where the coordinates are kept, and how to avoid Location entirely;
   - the first CA-signed exe runs on a fresh Windows 11 VM with Smart App Control On before it's published;
   - not chosen: SignPath Foundation (free, but its terms ask for "a certain verifiable reputation" this project doesn't have yet, and it would move the release build into CI) and Azure Artifact Signing (formerly Trusted Signing; individual developers must be in the US or Canada).
3. **Antivirus submissions only when a release gets flagged** — to Microsoft (Defender) or Kaspersky. A submission corrects a false positive; it doesn't build reputation.
4. **winget package** after the first CA-signed release: the release zip with `NestedInstallerType: portable` and `ArchiveBinariesDependOnPath`, a privacy URL, and install, upgrade and uninstall tested in Windows Sandbox; asset hashes are final at publish time (steps 1b and 1e). It's submitted from the maintainer's machine: a winget workflow in a file of its own, or one triggered by a release, would fail `release.ps1`'s workflow allowlist. winget doesn't require a signature, but users with Smart App Control On couldn't run a self-signed build. The personal **Scoop bucket** launches at the same time: its manifest persists `config.json` across upgrades (with a `pre_install` step that creates it empty — otherwise Scoop makes a folder of that name) and leaves `events.log` out.

### UX polish

- **v0.8.0 — first run without waiting:** the tray appears first and Windows Location is looked up in the background; the theme applies as soon as Location is turned on, without the Refresh step (Windows Location is retried a few times after the enable-Location dialog is dismissed, rather than subscribing to `Geolocator::StatusChanged`). Candidate: ask for Location only on the launch that creates `config.json`, and let the tooltip say "Location needed" after that.
- **v0.8.0 — a tray that reports status:** live tooltip ("Dark until 06:12", "Location needed", or a degraded-apply warning), an **"Open Log"** menu item, and a dialog when a Refresh or Toggle you clicked fails (scheduled ticks stay silent-to-log). Config errors and unusable theme paths already get a dialog; for failed applies, tier 2's commit watcher first has to report back to the tray, since its failures only reach `events.log` today.
- **After v1.0, not scheduled:** optional sunrise/sunset offsets (`offset_sunrise_min` / `offset_sunset_min` — still sun-anchored, so compatible with "no custom times", but they change the schedule's guarantees at polar latitudes); re-reading location on unlock/resume so a traveler's coordinates don't go stale (today they're never re-read once set, and a location call on every wake is the kind of activity antivirus heuristics notice); applying a theme by path through `IThemeManager2`, so a theme Settings doesn't list could skip the slow path (undocumented interface — investigate first).

### Maintenance notes

- **winit `ApplicationHandler` migration** — planned with v0.8.0, which rebuilds startup anyway (winit 0.31 is still in beta; the pinned 0.30 only deprecates `EventLoop::run`, allowed at the call site). Worth evaluating then: dropping winit for a plain Win32 message loop — the pattern already exists in the wake listener.
- **Dependency updates** — `windows-sys` 0.61 at the start of v0.6.0 (it touches every Win32 call); `windows` 0.62 and `tray-icon` 0.26 with v0.8.0's location and tooltip work.
- **Not tied to a release** — declare `rust-version` (1.91); Dependabot for the pinned Actions and crates; parse and lint `scripts\*.ps1` in CI; `release.ps1` refuses a CI result older than about a week (the dependency audit sees advisories as of the push); `build.ps1` and `test.ps1` check the signing certificate before compiling and also look for cargo on PATH, and `test.ps1` fails when a filter matches no test and refuses `RUSTFLAGS` like `build.ps1`; keep the build machine's profile path out of the exe; tests for `set_auto_start` (against a throwaway registry key), log rotation, a half-written registry pair, and schedule-contract sweeps at 65.4° and at a mid-latitude; tests that skip silently on CI fail there instead, once a CI run shows the runner has what they need; an Upgrading section in this README.

**Not planned**: GUI configuration (`config.json` + Refresh is the UX), custom wake times (sunrise/sunset is the whole point; offsets from them are fine), cross-platform (Windows only — macOS already has this natively), in-app update check (it would add the app's first network call, itself an AV signal; upgrades will come through winget and Scoop, and until then it's Quit and replace the exe), pause/snooze toggle (a manual override already pauses until the next transition or the next app start), and ADM-style scripting/hotkeys/battery rules (out of scope for a small tray tool).

## Contributing

**Status: stable for personal use.** The tray/apply/scheduling core is mature (120 unit tests; CI fails on any compiler or clippy warning, tests included, on unformatted code, and on a dependency with a known vulnerability or a license other than MIT, Apache-2.0 or Unicode-3.0; in daily use on the maintainer's PC). Working toward distribution — CA signing (v0.7.0), then a winget package and a Scoop bucket, and retiring the §Antivirus false positives section. Issues and PRs welcome — please open an issue to discuss larger changes before sending a patch. Focus areas: see [Roadmap](#roadmap).

## License

MIT — see [LICENSE](LICENSE).

## Acknowledgments

- [`sun-times`](https://crates.io/crates/sun-times) — local sunrise/sunset math.
- [`chrono`](https://crates.io/crates/chrono) — date/time handling; [`serde`](https://crates.io/crates/serde) + [`serde_json`](https://crates.io/crates/serde_json) — config persistence.
- [`tray-icon`](https://crates.io/crates/tray-icon) + [`winit`](https://crates.io/crates/winit) — tray icon and event loop.
- [`windows-sys`](https://crates.io/crates/windows-sys) + [`windows`](https://crates.io/crates/windows) — official Microsoft Rust bindings.
