# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Windows tray app (Rust) that swaps the full Windows **theme** (wallpaper + colors + light/dark mode) at local sunrise/sunset — macOS's auto-theme behavior, on Win11. Primary apply path is the `IThemeManager2` COM interface (the same one the Settings UWP wraps internally) for atomic, in-process theme apply; a two-tier fallback (legacy `ShellExecute(.theme)` → registry-only DWORD toggle) handles the case where the COM interface errors. ~395 KB single-exe (no VC++ redistributable needed — only the OS-provided UCRT), signed Authenticode, no installer.

Roadmap, per-version release plan, and the patch-vs-minor versioning rules live in README.md → Roadmap. v0.4.1 (audit bug-fix patch) shipped 2026-09-27; next up is **v0.5.0**: release-pipeline automation (`scripts\release.ps1` wrapping `build.ps1` + `gh release upload --clobber` + `gh release edit --prerelease=false`) plus the behavior-changing audit follow-ups listed in README → Roadmap → Foundation. Tag `v0.4.0` = `0c63d0e` (release commit `5f88d4e` + `scripts\build.ps1`).

## Source tree vs deployed binary — read first

The source tree (`C:\Users\atef\Documents\Projects\WinThemeSwitcher\`) is kept for future tweaks. **The actually-running binary lives elsewhere**:

```
C:\Tools\WinThemeSwitcher\
├── win-theme-switcher.exe   ← auto-starts at login
├── config.json              ← user's Riyadh coords, auto_start: true
└── events.log               ← diagnostic log (see "Reading events.log")
```

`HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WinThemeSwitcher` points at `"C:\Tools\WinThemeSwitcher\win-theme-switcher.exe"`. Redeploy with `scripts\build.ps1` (Build section) — it stops the running app, copies, verifies the copy by hash, and relaunches. Note `C:\Tools\WinThemeSwitcher` inherits `Authenticated Users:(M)` from `C:\` (any local account could swap the exe that auto-starts at login) — the README now recommends a per-user folder for new installs; tightening this folder's ACL or moving the install is the user's call and hasn't been done.

**Never relaunch the tray app with a plain `Start-Process` from an agent shell.** A child process inherits the shell's job object and, if the shell runs inside a packaged app's container, its package identity — and then its HKCU writes (the Run value, the theme registry values) go to a private copy-on-write hive instead of the real registry. The Claude desktop app's container has virtualized `%LOCALAPPDATA%` writes in the past (see "Sign every release build"); on 2026-09-27 the agent shell had no package identity, but don't rely on it. `build.ps1` relaunches via `explorer.exe "<exe>"`, which hands the launch to the user's shell (same context as the logon launch). Do the same for any manual relaunch.

(The pre-Cargo Manus prototype exe that used to sit at the repo root has been deleted; `/win-theme-switcher.exe` stays in `.gitignore` so it can't be re-committed.)

## Kaspersky false positive — critical context

**Signed builds run without special handling** (since the v0.3.0 IThemeManager2 + code-signing migration). The signed binary (`CN=WinThemeSwitcher Self-Signed` cert trusted via `Cert:\CurrentUser\Root`) collapses the Authenticode-trust signal, and tier-1 theme apply via `IThemeManager2::SetCurrentTheme` removes the `HWND_BROADCAST WM_SETTINGCHANGE` + direct `WM_THEMECHANGED` signals that previously tripped behavior heuristics. **Sign every release build** — unsigned builds resurrect the issue. The README tells *other* users a Trusted-application rule may still be needed for a self-signed build; on this machine it hasn't been since v0.3.0. Everything under "Historical" below is context for unsigned-build scenarios.

**Dev/test builds (2026-08-03):** KSN flagged fresh unsigned *test* binaries in `target\debug\deps\` (`VHO:Trojan.Win32.Convagent.gen`) — the trust rules are path-based and don't cover them, and KSN's scanner locks each fresh exe faster than a post-build signtool can run. Two-layer fix in place: (1) a Kaspersky **exclusion on the whole `target\` folder**, added by the user; (2) `scripts\test.ps1` builds the test binary, signs it from the cert store, and only then executes it — **run tests via this script, not bare `cargo test`**.

**Web-traffic popups during agent work:** Kaspersky's encrypted-connection scanning can't inspect the Claude app's HTTPS connections and may pop "The website … may not display correctly — Add to exclusions / Disconnect" when an agent fetches documentation. It's harmless; *Disconnect* just costs the agent that page.

### Historical: pre-signing trust setup

The unsigned binary tripped `VHO:Trojan.Win32.Agent.gen` (Rust exe with no Authenticode signature + `HKCU\Run` writes + `HWND_BROADCAST` of `WM_SETTINGCHANGE` + direct `WM_THEMECHANGED` to `Shell_TrayWnd` + WinRT Geolocation = every AV heuristic signal). Plain **path-based exclusions were insufficient** — Kaspersky's Behavior Detection quarantined regardless. The pre-signing workaround was a **Trusted Applications rule** (Kaspersky Settings → Security → Threats and Exclusions → Specify trusted applications) with all checkboxes ticked: Do not scan opened files, Do not monitor application activity, Do not inherit restrictions, Do not monitor child application activity, Allow interaction with Kaspersky interface.

Rules are **path-based**, so two currently exist:
1. `C:\Tools\WinThemeSwitcher\win-theme-switcher.exe` (the deployed binary — stable).
2. `C:\Users\atef\Documents\Projects\WinThemeSwitcher\target\release\win-theme-switcher.exe` (the build output — rewritten by each build).

If a future rebuild gets quarantined anyway (the hash changes and Kaspersky occasionally re-evaluates): drop a 0-byte placeholder at the path first (`Set-Content -Path ... -Value "" -Encoding Byte -Force`), re-add the trust rule while the placeholder exists, then rebuild. Same trick works for new deployment paths.

### When the trust rule is not enough (KSN cloud verdict)

Trusted Applications rules cover File Anti-Virus + Behavior Detection but **not Kaspersky Security Network (KSN) cloud reputation**. KSN can independently flag a fresh hash and pop a hostile two-button modal — *"Disinfect and restart"* / *"Try to disinfect without computer restart"* — with no Skip / Esc / X dismiss option. Both buttons quarantine (choose the no-restart one; the file is a disposable build artifact). Adding a Threats and Exclusions entry mid-modal does **not** clear the in-progress verdict — Kaspersky finishes quarantining anyway, and even subsequent rebuilds can be flagged by `svchost.exe` (the indexer-style scanner running as `NT AUTHORITY\NETWORK SERVICE`) before the popup re-appears.

The reliable workaround for a deploy session is to **right-click the tray K → Pause protection → 15 minutes**, then immediately copy + launch within that window. Once the process is loaded into memory it survives even after protection resumes. `set_auto_start(true)` re-creates the `HKCU\Run` value on each launch/Refresh if it's missing, so a quarantine event that deletes it self-repairs.

## Build

**Always build via `scripts\build.ps1`** — it builds, verifies the linker hardening, Authenticode-signs (with an RFC 3161 DigiCert timestamp), verifies the signature, and deploys in one step. Bare `cargo build --release` produces a fresh-hash unsigned Windows PE; KSN flags first-seen unsigned exes with `VHO:Trojan.Win32.Convagent.gen` on this machine, and the only fix is signing before the binary ever executes. The same caveat applies to `cargo test` — use `scripts\test.ps1`.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\atef\Documents\Projects\WinThemeSwitcher\scripts\build.ps1"            # build + sign + deploy + relaunch
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\atef\Documents\Projects\WinThemeSwitcher\scripts\build.ps1" -SkipCopy   # build + sign only, leave C:\Tools\ untouched
```

`build.ps1` exit codes: 0 ok · 1 cargo missing/build failed · 2 exe path not found · 3 no signtool · 4 signing/verification failed · 5 deploy dir missing · 6 copy failed (previous exe relaunched) · 7 linker hardening missing · 8 deployed but relaunch not seen. It builds with `--locked`, which refuses if `Cargo.lock` is out of date with `Cargo.toml` — after a version bump run `scripts\test.ps1` (or `cargo update -w`) first, then commit `Cargo.lock` with the bump. Both scripts `Push-Location` to the repo and restore the caller's location on exit.

**Linker hardening — `.cargo\config.toml`** (v0.4.1): hybrid CRT (`+crt-static` plus `/NODEFAULTLIB:libucrt.lib /DEFAULTLIB:ucrt.lib` — vcruntime linked statically, the OS-provided UCRT dynamically, so no `VCRUNTIME140.dll` dependency for ~21 KB) and `/DEPENDENTLOADFLAG:0x800` (the exe's static DLL imports resolve from System32 only; `main()` additionally calls `SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32)` for runtime bare-name `LoadLibrary` calls — together, a DLL planted next to the exe isn't loaded). Two gotchas, both enforced by `build.ps1` (exit 7): cargo finds `.cargo\config.toml` from the **current directory**, not `--manifest-path` (both scripts `Push-Location` to the repo first — `RUSTFLAGS` is refused up front with exit 7, and `check-hardening.ps1`'s PE check, also exit 7, is the backstop for both); and a set `RUSTFLAGS`/`CARGO_ENCODED_RUSTFLAGS` **replaces** the config's rustflags. `scripts\check-hardening.ps1 -Exe <path>` then checks the PE itself — no `VCRUNTIME140` import, and `DependentLoadFlags` (load-config directory, WORD at +0x4E) = 0x800 — called by `build.ps1` and by both GitHub workflows, so losing the config can't pass CI.

For raw cargo (debugging a compile error only — never produces a deployable exe; run it from the repo directory):

```powershell
& "$env:USERPROFILE\.cargo\bin\cargo.exe" check --tests `
    --manifest-path "C:\Users\atef\Documents\Projects\WinThemeSwitcher\Cargo.toml"
```

Default toolchain is `stable-x86_64-pc-windows-msvc` (MSVC Build Tools required; the GNU toolchain's bundled linker/dlltool was broken on this machine). Release profile: `opt-level = "z"`, `lto = true`, `codegen-units = 1`, `panic = "abort"`, `strip = true`. Output ~395 KB. No `build.rs` — `windows-sys` and `windows` self-link.

> **For Claude Code / Fable 5 / any LLM coding agent working in this repo**: do not invoke `cargo build` or `cargo test` directly (`cargo check` / `cargo clippy` / `cargo fmt` are fine — they don't build the app or test executables; the dependency build scripts they compile and run live under the Kaspersky-excluded `target\` like everything else). Always call `scripts\build.ps1` (for a release) or `scripts\test.ps1` (for tests). The wrappers exist *because* agents forget to sign.

### Test and lint

**Run the suite via the signing wrapper** (bare `cargo test` produces an unsigned fresh-hash exe that Kaspersky/KSN may lock or quarantine):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\atef\Documents\Projects\WinThemeSwitcher\scripts\test.ps1"                  # full suite
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\atef\Documents\Projects\WinThemeSwitcher\scripts\test.ps1" riyadh           # name filter
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\atef\Documents\Projects\WinThemeSwitcher\scripts\test.ps1" -Filter "--x"   # pass a harness flag (must use -Filter; PowerShell eats a bare --x)
```

`test.ps1` exits with the harness's code (101 = test failure), 1 on build/launch failure, 3 without signtool, 4 if a test binary isn't validly signed afterwards (checked with `Get-AuthenticodeSignature`, not just signtool's exit code) — it **never runs an unsigned test binary**. Before v0.4.1 it always exited 0: `#![windows_subsystem = "windows"]` makes the *test* binary a GUI-subsystem exe too, and PowerShell doesn't wait for (or record the exit code of) a GUI exe unless its output is piped — the script now pipes through `ForEach-Object`.

Lint (advisory in CI):

```powershell
$cargo = "$env:USERPROFILE\.cargo\bin\cargo.exe"
$manifest = "C:\Users\atef\Documents\Projects\WinThemeSwitcher\Cargo.toml"
& $cargo fmt --manifest-path $manifest --check
& $cargo clippy --release --locked --manifest-path $manifest -- -W clippy::all
```

Both are clean as of v0.4.1 (the winit `EventLoop::run` deprecation is `#[allow]`ed at the call site; see README → Maintenance notes).

### Sign every release build

Done automatically by `scripts\build.ps1`. Manual sign — for rebuilding a tagged release's assets or troubleshooting — uses the cert from the store:

```powershell
& "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe" sign `
    /n "WinThemeSwitcher Self-Signed" /fd SHA256 `
    /tr http://timestamp.digicert.com /td SHA256 `
    "C:\Users\atef\Documents\Projects\WinThemeSwitcher\target\release\win-theme-switcher.exe"
```

Cert: `CN=WinThemeSwitcher Self-Signed`, thumbprint `40E0D1EB58DAC255EB37E9D64FF34448E3D33D12`, in `Cert:\CurrentUser\My` (with its exportable private key) and `Cert:\CurrentUser\Root`; expires 2036-04-28. Backup: the key was re-exported under a new password on 2026-09-27 and lives in the maintainer's password manager, not on disk. The previous pfx backup (under `%LOCALAPPDATA%\WinThemeSwitcher\signing\`, partly MSIX-virtualized into the Claude app's package folder), whose password appeared in this repo's history before v0.4.1, was deleted. Never write a key backup or its password into the repo.

If the cert ever needs regenerating: `New-SelfSignedCertificate -Type CodeSigning -Subject "CN=WinThemeSwitcher Self-Signed" -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(10)`, re-add it to the Root store, update the thumbprint in README + here, and replace `WinThemeSwitcher-publisher.cer`. The `/tr` timestamp countersignature (added 2026-07-04, free DigiCert TSA, needs network at sign time) keeps signatures valid after the cert expires. The v0.3.0 and v0.3.1 published assets were retro-timestamped in place the same day; v0.1.0/v0.2.0 predate signing entirely.

### CI and releases (`.github/workflows/`)

- `ci.yml` — on push to main/master and on every PR, `windows-latest`: `cargo build --release --locked`, then `scripts/check-hardening.ps1` on the release exe (gate), then `cargo test --locked` (**hard gate**), then `cargo fmt --check` and `cargo clippy --release --locked -- -W clippy::all` (`continue-on-error`, advisory). `.cargo\config.toml` applies (cargo runs from the checkout root). Actions: `checkout@v7`, `cache@v6`, `dtolnay/rust-toolchain@stable` — the cache key includes the rustc version and hashes `Cargo.lock` + `.cargo/config.toml`. Tests live in `mod tests` at the bottom of `main.rs` (83; ~5 s — `[profile.test] opt-level = 2` because the polar sweep evaluates `schedule` tens of thousands of times): scheduling math (Apia/UTC+13, Riyadh, Reykjavik midnight-sunset, Tromsø polar boundaries with exact expected instants, a Pevek day straddling UTC midnight, a schedule-contract sweep around six polar edges incl. McMurdo, , a check that sun_times still decides below 65.5°, threshold/southern-hemisphere pins, and start-time determinism of short polar segments), solar altitude, config loading (parse-error preservation, first-run, empty/NUL heal, BOM/UTF-16 LE+BE/truncated, unknown-key round-trip, atomic save, read-only target, hard-link write-through), `decide_tick`/override/retry/frame-rule sequences (incl. "earlier next without a clock step still preserves"), clock-step `plan_wake`, `toggle_target`, `sanitize_log_msg`, `.theme` DisplayName parsing (via `GetPrivateProfileStringW`, incl. UTF-16 and quoted values), `resolve_theme_file`, and Refresh's config adoption. Windows-only (they call Win32 APIs and read `%SystemRoot%\Resources\Themes`); otherwise machine-independent — instants are built in UTC and file tests use unique per-test temp files. The system-theme test returns early (printing SKIP, visible only with `--nocapture` or `--show-output`) if the stock `.theme` files are absent.
- `release.yml` — on tag push `v*` (or manual dispatch): builds on GitHub runners and attaches a zip (exe + README + LICENSE + publisher `.cer`) plus the bare exe and `WinThemeSwitcher-publisher.cer` to a **prerelease** (`prerelease: true` is hardcoded; `softprops/action-gh-release@v3`; runs `check-hardening.ps1`; its cache restores from main's CI cache). **CI binaries are unsigned** — the signing key exists only on this machine, so every tagged release needs a manual post-tag step: build locally from the tag (`build.ps1 -SkipCopy`), zip (exe + the tag's README + LICENSE + `.cer`), then replace the workflow's assets with `gh release upload <tag> <files> --clobber`, and verify by downloading. Don't skip it: v0.3.0 originally shipped unsigned CI builds because this step was missed. **Don't re-run the Release workflow on an existing tag** — it overwrites the signed assets with unsigned ones (v0.5.0 fixes this).

## Architecture — `src/main.rs`

Single file, ~3850 lines (~2450 code + ~1400 `mod tests`), event-driven. Logs every state transition to `events.log` next to the exe (rotated to `events.log.old` past 256 KB; writes serialized by `LOG_LOCK`, one `write_all` per line).

### 1. Theme apply — three-tier fallback in `apply_theme`

Tiered worst-case-degradation: each tier is more invasive but less reliable than the one above. `apply_theme` walks them top-down, returning a `&'static str` tag for the tier that succeeded (logged in the `applied=` field of the `cause=...` line). If a configured `theme_day`/`theme_night` path doesn't exist, `resolve_theme_file` falls back to the stock theme and `apply_theme` logs `theme_path_missing`.

#### Tier 1: `IThemeManager2` (preferred — `applied=theme-manager2`)

Undocumented-but-stable COM interface in `themeui.dll` that the Settings UWP itself wraps. CLSID `{9324da94-50ec-4a14-a770-e90ca03e7c8f}`, IID `{c1e8c83e-845d-4d95-81db-e283fdffc000}`. Vtable layout in the `IThemeManager2Vtbl` struct at the top of `main.rs`.

Flow (`apply_via_theme_manager2`):
1. Resolve the target `.theme` file's `[Theme]` → `DisplayName`. `resolve_theme_display_name` reads it with **`GetPrivateProfileStringW`** (v0.4.1 — Windows' own INI parser, so key/section case, whitespace, quotes, and ANSI-vs-UTF-16 decoding match what Windows itself reads; a hand parser mis-decoded non-English names). System themes use SHLoadIndirectString-style refs (`@%SystemRoot%\System32\themeui.dll,-2060`), resolved by `resolve_indirect_string`. For `dark.theme` → `"Windows (dark)"`; for `aero.theme` → `"Windows (light)"`.
2. `CoCreateInstance(CLSID_THEME_MANAGER2)` + `Init(0)`.
3. Enumerate via `GetThemeCount` + `GetTheme(i)` + `ITheme::GetDisplayName(&BSTR)` until a name match. Free each BSTR with `SysFreeString`. **Don't cache the index across launches** — enumeration order is not stable.
4. `SetCurrentTheme(NULL, idx, apply_now=1, apply_flags=NO_HOURGLASS, pack_flags=0)`. This is the only tier-1 call that applies; it does the WM_THEMECHANGED + WM_SETTINGCHANGE broadcasts internally.

Why this is the primary path: ShellExecuteW(`.theme`) silently fails when the user isn't actively interactive (post-WTS_SESSION_UNLOCK, scheduled while away, no foreground UI). The UWP activation pipeline swallows the apply request — Settings flashes briefly but never commits. `IThemeManager2` is in-process, has no UI dependency, and is what every serious tool uses (AutoDarkMode, wtheme, etc.). Apply latency is ~200 ms vs. the ~5 s poll-then-fail of the legacy path.

**STA threading is mandatory** for this interface ("Shell crap is always STA" per AutoDarkMode source). The main thread already calls `CoInitializeEx(None, COINIT_APARTMENTTHREADED)` at startup; tier-1 apply runs from winit event handlers on that same thread, which is correct. **Never call from a worker thread** without CoInitializeEx(STA) on it first — you'll get RPC_E_WRONG_THREAD or silent corruption.

#### Tier 2: `ShellExecuteW(.theme)` + `commit_watcher` (legacy — `applied=theme-file`)

Fires only if tier 1 errors out (logged as `theme_manager2_err target=… msg="…"`). `ShellExecuteW("open", <.theme path>, ..., SW_HIDE)` launches the Themes UWP, plus a `start_settings_closer` thread to `PostMessage(WM_CLOSE)` the Settings window once it appears, plus a 300 ms sleep + `poke_shell` (taskbar repaint). (Known issues deferred to v0.5.0: the closer matches English titles only, can close a Settings window the user had open, and posts WM_CLOSE before the theme commits.)

**`commit_watcher` is the safety net for tier 2's silent-fail mode**: spawns a thread that polls `current_theme()` every 200 ms for 5 s. If the registry never matches the target → logs `commit_timeout target=…` and **falls through to tier 3 from inside the watcher thread** — writes the registry directly, broadcasts, pokes shell, polls again to confirm, logs `fallback_registry target=… confirmed=true after_ms=…`.

If tier 1 is healthy this path is rarely entered. It exists as backup in case future Windows builds break the COM interface.

#### Tier 3: registry-only (last resort — `applied=registry`)

`write_theme_registry` writes `AppsUseLightTheme` + `SystemUsesLightTheme`, then the caller broadcasts `WM_SETTINGCHANGE("ImmersiveColorSet")` to `HWND_BROADCAST` and calls `poke_shell`. **Flips light/dark mode but not wallpaper.** Both `RegSetValueExW` results are checked (v0.4.1): `AppsUseLightTheme` is written first and `SystemUsesLightTheme` — the value `current_theme()` reads — only after it succeeds, stopping at the first failure; any failure returns `Err` (so tick's bounded retry runs), and the only possible half state reads as not-applied (the reverse would look like a user intervention to the retry gate). A half-written pair still broadcasts before returning.

**`poke_shell`** sends `WM_THEMECHANGED` + targeted `WM_SETTINGCHANGE("ImmersiveColorSet")` to `Shell_TrayWnd` and `Shell_SecondaryTrayWnd`, then `DwmFlush()`. Required for tiers 2 and 3 — tier 1 broadcasts internally. If future Win versions add new taskbar window classes, extend the list.

**Theme file resolution** (`resolve_theme_file`): if `config.theme_day` / `theme_night` is an existing path, use it; otherwise fall back to `%SystemRoot%\Resources\Themes\aero.theme` (light) / `dark.theme` (dark). Custom user themes work with tier 1 only if they're already registered with Windows (installed via Settings → Themes). Otherwise tier 1 errors with `no installed theme matches DisplayName "…"` and tier 2 takes over.

### 2. Event loop — only tick on specific events

The run closure must **not** call `tick()` on arbitrary events: a theme change in Settings broadcasts `WM_SETTINGCHANGE`, and a tick on it would see `current != target` and revert the user's choice. `tick()` runs only on:

- `Event::NewEvents(StartCause::Init)` — first event after launch.
- `Event::NewEvents(StartCause::ResumeTimeReached { .. })` — **only when `plan_wake` says so**: the armed deadline is due, or a clock step was detected (below). A heartbeat wake that is neither just re-arms.
- `Event::UserEvent(AppEvent::Menu(refresh_id))` — user clicked Refresh.
- `Event::UserEvent(AppEvent::Wake(_))` — session unlock / power resume (§5; safe because these never fire on a Settings theme change).

`Event::NewEvents(StartCause::WaitCancelled { .. })` (any other wake: tray/menu input, broadcasts) never ticks; it re-runs `plan_wake` and either re-arms or, if the plan is Tick (a clock step, or the wall-clock deadline already due), arms `Instant::now()` so the next ResumeTimeReached ticks. Everything else is `_ => {}` (the Menu arm also handles Toggle Theme, Open Config and Quit, which don't tick).

**Wall clock vs. monotonic clock (v0.4.1).** The schedule is wall-clock, winit's `WaitUntil` is a monotonic `Instant`, and the mapping breaks when the system clock is *stepped* — on this machine every boot after an Ubuntu session (Ubuntu keeps the RTC in UTC, so Windows starts 3 h behind until w32time steps it forward; Kernel-General event 1 shows the +3 h jump). Before v0.4.1 the Init tick decided from the wrong time and the deadline fired up to 3 h late. Now `tick()` records `TickState.armed` (wall deadline, via `arm()`) and `TickState.mark` — the (wall, mono) readings the tick *decided* from, so a step during a slow apply is still detected; `plan_wake(armed, mark, now_wall, now_mono)` — pure, unit-tested — returns `Tick{step_ms}` when `|wall advance − mono advance| > 60 s` (`CLOCK_STEP_MS`) or the deadline is due, else `Arm(min(deadline, now + HEARTBEAT))`. `HEARTBEAT` is 10 min, so a step is caught even if the `WM_TIMECHANGE` broadcast (which reaches winit's hidden top-level window and wakes the loop immediately) never arrives. The tick that consumes a step — the `cause=clock-jump` tick, or a wake/Refresh that got there first — logs `clock_jump step_s=±N consumed_by=<cause>` (pure helper: `detected_clock_step`). A suspend may also register as a "step" if the monotonic clock paused during sleep — harmless (one extra state-aware tick).

**State-aware apply**: `tick` decides via the pure `decide_tick(kind, current, target, now, next, clock_stepped_back, &TickState)` → `Apply` / `SkipInSync` / `SkipOverride` / `CancelRetry`, called with the **pre-tick** state; outcomes are recorded after the apply result is known (`note_reconciled` on any non-Err outcome, `note_apply_failed` on Err). **Refresh always applies** (config edits take effect immediately) and resets the retry budget first. **Bounded apply retry (v0.4.0)**: a failed apply reschedules to `min(next_transition, now + 60 s)` for up to 3 consecutive attempts (log field ` retry=N`, then ` retry=exhausted`; budget resets per episode, and exhaustion clears `reconciled_next` so later wakes still retry instead of "preserving" the failure). The retry arrives as a normal Scheduled tick; `retry_baseline` (the theme observed at failure) gates it — if the screen moved off the baseline within the episode's window (`now < episode_next`), the user intervened and the retry stands down (`applied=skip-user-intervened`); an unreadable reading on either side never counts as a move. The retry only covers *total* apply failure (all three tiers, realistically a failed registry write); a tier-2 ShellExecute silent-fail recovers via commit_watcher's registry fallback instead.

Scheduling math: `schedule(now_utc) -> (Theme, next_utc)` is the single source of truth (unit-tested). It collects sunrise/sunset instants for the **UTC** dates D−1..D+1 via the `sun-times` crate, sorts them as instants, and picks state-after-last-event ≤ now / first-event > now. **Never pass a local date to `sun_times`** — it takes a UTC date and keys events to the solar day; passing a local date made UTC+13/+14 permanently dark and skipped post-midnight sunsets (fixed v0.3.2). **Polar latitudes (v0.4.1):** at `|lat| >= POLAR_MODEL_LAT` (65.5° — whole days without sunrise/sunset start at ~65.73°) the schedule uses **only** the solar-altitude model: current from `altitude_theme` (`solar_altitude_deg` — local implementation, the crate's `altitude` has math bugs — vs. the −0.833° civil threshold), next from `next_altitude_flip` (`next_altitude_crossing` in 48 h chunks up to 200 days; probes at least 15 s apart — so whether a short polar "day" is found doesn't depend on the start time — and further apart when the sun is far from the threshold, since the altitude changes ≤ 15.04°/h × |cos lat|; bisected to 1 ms and returned on the far side of the crossing). Mixing `sun_times` in at polar latitudes is what broke before: after the last sunset before the midnight sun it reports no following sunrise (~68 days of Dark in Tromsø), and a design that switched between the two models per UTC date made them disagree at midnights — phantom transitions that reverted overrides, and minute-long flickers (both caught in the v0.4.1 reviews). One model per location makes the **contract** hold by construction: the theme stays `current` for every instant in [now, next), and `next` is a real flip — unit-tested by `schedule_contract_holds_around_polar_edges` (Tromsø, Pevek, McMurdo). Below 65.5° (Reykjavik included) `sun_times` alone decides, unchanged. Everything is pure math on UTC instants; `tick` converts to `Local` only for logging.

### 3. Location (WinRT Geolocation)

`try_get_windows_location()` uses the `windows` crate's `Geolocator::RequestAccessAsync().get()` → `GetGeopositionAsync().get()`. Blocking, <1 s with a cached location (can take much longer cold — deferred to v0.5.0). Called **before** `event_loop.run`, so the tray icon doesn't appear until location is known.

On failure (service off / permission denied): `ask_enable_location()` MessageBox (Yes/No). Yes → `ShellExecute("ms-settings:privacy-location")` + info MessageBox telling user to enable and click Refresh. No → `show_manual_setup_prompt` opens `config.json` via `open_config_in_editor`: the `.json` "open" handler if one exists (`AssocQueryStringW(ASSOCF_INIT_IGNOREUNKNOWN, ASSOCSTR_FRIENDLYAPPNAME)` — the flag makes a missing association fail instead of resolving to the "How do you want to open this file?" picker, which is what happens on this machine; FRIENDLYAPPNAME also covers Store-app handlers), else `System32\notepad.exe` by full path (`GetSystemDirectoryW` — a bare name would be searched for starting in the exe's own folder). Refresh retries WinRT if location is still empty.

**COM init matters**: `ensure_com_initialized` → `CoInitializeEx(None, COINIT_APARTMENTTHREADED)` runs first in `run()`. WinRT silently fails on an uninitialized thread.

### 4. Config

Exe-relative path (`current_exe().parent().join("config.json")`, never CWD). `#[serde(default)]` at struct level makes missing fields safe; `#[serde(flatten)] extra` keeps unknown keys so a save writes them back (v0.4.1).

**Broken files are never overwritten** (v0.3.2; unit-tested). `load_config_ex(path) -> Result<(Config, Option<Healed>), String>` (`Healed { reason, save_err }`) (`load_config_at` drops the heal flag): the file is decoded by `read_config_text` (UTF-8 with/without BOM, or UTF-16 LE/BE with BOM — PowerShell 5.1's `>` default is LE; odd-length UTF-16 is an error, not a heal); missing file → first run, defaults written via `create_new` (so a file appearing in a delete-then-rename save window wins — never route this through rename, which replaces); content that is only whitespace/NUL/BOM → self-heals to defaults (crash leftover; `load_config_logged` logs `config_healed action=reset-to-defaults`, or `action=reset-in-memory` + `config_save_err` if writing the defaults back failed); anything else unparsable → `Err`, file untouched. Errors go through `report_config_error`: a `config_error` log line plus a MessageBox on a **detached thread** (an `AtomicBool` prevents stacking). Broken-config session policy: startup runs read-only against the file (in-memory defaults + best-effort WinRT coords, no `save_config`, **no autostart change** — the fallback `auto_start: true` must not override a broken file's `false`). Refresh goes through the pure `adopt_reloaded_config`: a good reload is adopted; a failed one keeps the last-known-good config; autostart is re-asserted only if a config has come from disk at some point this session (`have_disk_cfg`) — unit-tested. `save_config` must only ever be called with a config successfully loaded from disk this session.

`save_config_at` is atomic (v0.4.1): write `config.json.tmp`, `sync_all`, rename over the target (retrying briefly on access-denied/sharing-violation from an AV scan); on failure the temp file is removed and `persist_config` logs `config_save_err`. Exception: if `config.json` is a symlink/junction (a name-surrogate reparse point — OneDrive/cloud placeholders don't count) or a hard link with other names (`is_linked_file` — e.g. a package manager's persisted config), it's written in place, since a rename would detach it.

```rust
struct Config {
    latitude: f64,
    longitude: f64,
    auto_start: bool,
    theme_day:   Option<String>,   // .theme path, or None → aero.theme
    theme_night: Option<String>,   // .theme path, or None → dark.theme
    #[serde(flatten)] extra: serde_json::Map<String, serde_json::Value>,  // unknown keys, preserved
}
```

`has_location()` returns false when both coords are `0.0` (null-island sentinel used for first-run detection).

`set_auto_start(enable)` makes the `HKCU\...\Run\WinThemeSwitcher` value match: enable reads the current value and writes the quoted exe path only if it differs (a missing value always differs — that's the AV-quarantine recovery); disable deletes it (not-found counts as success). Registry failures return `Err`; `apply_auto_start` logs them as `autostart_err`.

### 5. Wake on session unlock / power resume

The wake ticks exist to reconcile the moment the user is back (a transition may have passed while locked or asleep) and to apply the override-preservation rule; the clock-step handling in §2 covers wall-clock jumps. (The original v0.2.0 rationale — "`Instant` pauses across suspend, so a deadline would fire ~22 h late" — is likely wrong on Windows, where `Instant` is QPC-based and winit fires ResumeTimeReached on the first wake past the deadline; the observed misses predated tier 1, when the ShellExecute path silently failed in exactly those contexts.)

`start_wake_listener` registers the power callback, then spawns a worker thread that creates a hidden message-only window (`HWND_MESSAGE`) for the session notification:

- `WTSRegisterSessionNotification(hwnd, NOTIFY_FOR_THIS_SESSION)` → `WM_WTSSESSION_CHANGE`; we act on `WTS_SESSION_UNLOCK`. Up to 3 attempts 2 s apart (terminal services may not be up at logon autostart).
- `register_power_resume()` (v0.4.1, called before the window thread starts): `PowerRegisterSuspendResumeNotification(DEVICE_NOTIFY_CALLBACK, &DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS{power_callback}, ..)` — the parameters are leaked (must outlive the registration). `power_callback` runs on a system thread and sends `Wake(Power)` on `PBT_APMRESUMEAUTOMATIC` (sent on every resume; `PBT_APMRESUMESUSPEND` would only duplicate it). **Before v0.4.1 this passed `DEVICE_NOTIFY_WINDOW_HANDLE` + the hwnd — the function accepts ONLY `DEVICE_NOTIFY_CALLBACK`, so it failed with 87 (`ERROR_INVALID_PARAMETER`) on every launch and the resume hook never fired; unlock events masked it.** The callback form is also independent of whether a message-only window would receive `WM_POWERBROADCAST` (undocumented). A remaining `stage=power_register` error is real; its code is the function's return value (not `GetLastError`). This wake can arrive while the session is still at the lock screen; the later unlock is then a `SkipInSync`. **Still unverified on real hardware** — after deploying, sleep with sign-in-on-wake set to Never, resume, and confirm a `cause=wake-power` line.

Both routes call `proxy.send_event(AppEvent::Wake(WakeKind::Unlock | WakeKind::Power))` via a process-wide `OnceLock<EventLoopProxy<AppEvent>>` (the WindowProc is `extern "system"` and can't capture). The main loop ticks with `TickKind::Wake` (`cause=wake-unlock` / `cause=wake-power`). Idempotent across both events firing in sequence.

`WTS_SESSION_UNLOCK` (0x8) is a local `WPARAM` const for direct comparison in the WindowProc; `PBT_APMRESUMEAUTOMATIC` and `DEVICE_NOTIFY_CALLBACK` are imported from `windows_sys::Win32::UI::WindowsAndMessaging` (not RemoteDesktop/Power, where you'd look first).

**Why this doesn't resurrect the manual-override-fight bug**: neither event fires when the user changes theme in Settings.

**Manual-override preservation (v0.4.0)**: a manual override (Settings or Toggle Theme) *survives* lock/unlock and wake-from-sleep. The rule is time-based: `TickState.reconciled_next` records the next-transition instant from the last tick that ended in sync; a Wake tick — or a Scheduled tick that runs before `reconciled_next`, which since `plan_wake` never ticks early can only be a clock-jump tick — skips the re-apply only when no transition has passed (`now < reconciled_next`, log `applied=skip-override`) and no retry is pending. v0.4.1 adds a **frame rule**: if a backward clock step was *observed* (`clock_step_ms` against the pre-tick mark < −60 s) and the schedule's upcoming transition is now more than 60 s *earlier* than `reconciled_next`, the clock was stepped back across a transition — the divergence is the app's own stale apply, so it re-applies. (Gated on an observed step because `next` can also move earlier legitimately, e.g. a polar-season crossing coming within the 48 h search.) A FAILED apply never advances `reconciled_next` (a retry leaves it stale, an exhausted budget clears it) so later wakes re-apply (a free retry). Overrides reset at the next natural transition and do not survive a process restart.

### 6. Tray + menu

Menu: Toggle Theme, Open Config, Refresh, separator, Quit. Menu events flow through `MenuEvent::set_event_handler` → `EventLoopProxy::send_event(AppEvent::Menu(id))` so clicks wake the loop. `TrayIconEvent::set_event_handler` is set to a no-op (v0.4.1): without a handler tray-icon queues every mouse event over the icon in an unbounded channel nothing drains.

**Toggle Theme** applies `toggle_target(current_theme())` (opposite of what's on screen; unreadable → Dark) via `apply_theme` directly — it does **not** call `tick()`, does not touch `TickState`, and does not disturb the pending wait. That's what makes it a manual override. Known accepted race: a toggle within ~5 s of a *tier-2* apply can be reverted by that apply's still-running commit_watcher (unreachable while tier 1 is healthy).

Tray icon is generated in `make_tray_icon`: 32×32 RGBA, half orange (sun) + half dark-blue (moon). Procedural because `tray-icon`'s default placeholder is near-invisible on both taskbar modes; `with_icon` is required for the icon to show.

### 7. Fail-loudly plumbing (v0.4.0)

`main()` is a thin wrapper: `install_panic_hook()` (writes `panic at=file:line:col msg="…"` via `log_event_from_panic`, which only `try_lock`s `LOG_LOCK` — a panic while holding it must not deadlock instead of aborting) → `claim_single_instance()` (`CreateMutexW("Local\\WinThemeSwitcher.single-instance")`; on `ERROR_ALREADY_EXISTS` → log `duplicate_instance`, info MessageBox, exit 0; handle intentionally leaked; mutex-creation *failure* logs and continues) → `run()`. Any `Err` from `run()` — tray creation racing the taskbar at login, event-loop build/death — logs `fatal_error msg="…"` and shows a blocking MessageBox before exit 1. All quoted `msg="…"`/`path="…"`/`display="…"` fields flow through `sanitize_log_msg` (quotes→apostrophes, newlines→spaces).

### Reading events.log

Main line: `<rfc3339> cause=<c> current=<t> target=<t> applied=<a> next=<rfc3339>`, or on a failed apply `… target=<t> err="…" retry=N|exhausted next=<rfc3339>`, plus `cause=toggle current=… target=… applied=…` for Toggle and `cause=<c> skipped=no-location` without coordinates.

- `cause=`: `init`, `resume-time` (deadline due), `clock-jump` (clock step detected — preceded by `clock_jump step_s=±N consumed_by=clock-jump`), `refresh`, `wake-unlock`, `wake-power`, `toggle`.
- `applied=`: `theme-manager2` (tier 1, healthy), `theme-file` (tier 2), `registry` (tier 3), `skip` (already on target), `skip-override` (manual override preserved), `skip-user-intervened` (retry stood down).
- Companion lines: `theme_manager2_apply` (tier-1 success: display name, index, `after_ms`), `theme_manager2_err`, `theme_manager2_enum_skip`, `theme_path_missing`, `settings_closed`, `commit_observed`, `commit_timeout` → `fallback_registry` / `fallback_registry_err`, `config_error`, `config_healed`, `config_save_err`, `autostart_err`, `open_config_err`, `wake_listener_err stage=register_class|create_window|power_register code=N` / `stage=wts_register attempt=N code=N`, `wake_listener_wts_ok attempt=N`, `duplicate_instance`, `single_instance_err`, `fatal_error`, `panic`.
- Healthy: `applied=theme-manager2` (its `theme_manager2_apply` line shows `after_ms` ~150–700), or `applied=skip`. Recurring `theme_manager2_err` + `commit_timeout` means tier 1 broke (e.g. a Windows update changed the COM interface).

## Dependencies (`Cargo.toml`)

- `chrono`, `sun-times` — sunrise/sunset math.
- `serde` + `serde_json` — config persistence.
- `tray-icon` (0.19), `winit` (0.30) — tray + event loop. Menu types come from `muda` (re-exported under `tray_icon::menu`). A tray-icon upgrade is deferred to v0.5.0.
- `windows-sys` 0.59 (features: `Win32_Foundation`, `Win32_Security`, `Win32_Storage_FileSystem`, `Win32_System_Com`, `Win32_System_LibraryLoader`, `Win32_System_Power`, `Win32_System_RemoteDesktop`, `Win32_System_Registry`, `Win32_System_SystemInformation`, `Win32_System_Threading`, `Win32_System_WindowsProgramming`, `Win32_UI_WindowsAndMessaging`, `Win32_UI_Shell`, `Win32_Graphics_Dwm`) — raw Win32 FFI. `SysFreeString` lives in `Win32_Foundation` (not `Win32_System_Ole`). `CreateMutexW` needs BOTH `Win32_System_Threading` *and* `Win32_Security` (cfg-gated on the latter because of its `SECURITY_ATTRIBUTES` parameter). `PowerRegisterSuspendResumeNotification` + `DEVICE_NOTIFY_SUBSCRIBE_PARAMETERS` are in `Win32::System::Power` (the flag type is from `Win32_UI_WindowsAndMessaging`). `GetPrivateProfileStringW` needs `Win32_System_WindowsProgramming`; `GetFileInformationByHandle` (hard-link check) `Win32_Storage_FileSystem`; `GetSystemDirectoryW` `Win32_System_SystemInformation`.
- `windows` (features: `Devices_Geolocation`, `Foundation`, `Win32_System_Com`) — WinRT Geolocator + `CoInitializeEx` for the main thread's STA.
- `.cargo\config.toml` — the hybrid-CRT + `/DEPENDENTLOADFLAG:0x800` linker flags (Build section).

## Invariants — don't break these

- **`tick()` scope**: only Init / ResumeTimeReached-when-`plan_wake`-says-Tick / Refresh / `AppEvent::Wake`. Adding a callsite for any *other* trigger — especially anything that fires on `WM_SETTINGCHANGE` — resurrects the manual-override-fight bug. `WaitCancelled` must never tick (it may only set the control flow).
- **Only `tick()` writes `TickState.armed`/`mark`** (via `arm()`, or clearing them without location). The clock-step check measures from the mark; re-marking anywhere else (e.g. on a heartbeat or a WaitCancelled) would hide a step.
- **STA thread for IThemeManager2**: `ensure_com_initialized` runs `CoInitializeEx(None, COINIT_APARTMENTTHREADED)` first in `run()`. All theme apply runs on that thread.
- **`poke_shell` after tier-2 / tier-3 apply only** — tier 1 broadcasts internally; adding it back re-introduces the AV-tripping `HWND_BROADCAST WM_SETTINGCHANGE` signal.
- **Refresh forces apply** (bypasses state check); scheduled transitions respect it. Don't invert.
- **`decide_tick` reads the PRE-tick state**; `note_reconciled`/`note_apply_failed` run after the apply outcome is known. A failed apply must never advance `reconciled_next` (retry: leave it stale; exhaustion: clear it). Never "simplify" by assigning state before/regardless of the outcome.
- **Toggle Theme never ticks and never touches `TickState`** — it's a manual override by construction.
- **Capture `GetLastError()` into a local immediately** after the failing Win32 call — `Local::now()`, `log_event`, and `format!` all make Win32 calls that clobber it.
- **`SetDefaultDllDirectories` stays the first call in `main()`**, and everything the app loads must be a System32 DLL (or COM/WinRT-activated by full path).
- **Run cargo from the repo directory, never with `RUSTFLAGS` set** (or `.cargo\config.toml` silently doesn't apply). Run tests via `scripts\test.ps1`, release builds via `scripts\build.ps1`.
- **Never relaunch the deployed app as a child of an agent shell** (MSIX container → virtualized HKCU); use `explorer.exe "<exe>"` as `build.ps1` does.
- **Free BSTRs from `ITheme::GetDisplayName` with `SysFreeString`** — not `CoTaskMemFree`, and never leak.
- **Vtable order in `IThemeManager2Vtbl`**: every method's slot index must match the COM ABI. The struct declares every slot up through `set_current_theme` — uncalled interior slots are `_`-prefixed placeholders that are **mandatory padding, never removable**; only trailing slots after the last called method may be omitted, and nothing may ever be reordered.
- **`ensure_com_initialized` before any WinRT call**: otherwise Geolocator returns errors silently.
- **UTF-16 + NUL**: all Win32 wide strings go through `wide()`, which appends the terminator. Never pass a bare `&str` to a `*W` API.
- **HWND null check**: `(hwnd as usize) == 0` — robust to `windows-sys` flipping between `*mut c_void` and `isize`.
- **Windowed subsystem** (`#![windows_subsystem = "windows"]`): no console, `println!` goes nowhere — and it applies to the test binary too (see `test.ps1`). For diagnostics, write to `events.log`.
