# Build a release exe, verify its linker hardening, Authenticode-sign it with
# an RFC 3161 countersignature, and deploy it.
#
# Why this exists: bare `cargo build --release` produces a fresh-hash
# unsigned Windows PE. Kaspersky KSN flags first-seen unsigned hashes with
# VHO:Trojan.Win32.Convagent.gen on this machine, and the only mitigation
# is to sign the binary before its first execution. The signing step must
# run AFTER cargo produces the exe and BEFORE anything executes it (the only
# execution here is relaunching the deployed copy, after its signature has
# been verified).
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build.ps1 -SkipCopy
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build.ps1 -DeployDir <folder>
#
# Parameters:
#   -SkipCopy   build + sign in place, don't overwrite the deployed exe
#               (useful when verifying a build without disturbing the
#               running install).
#   -DeployDir  install folder, default %LOCALAPPDATA%\Programs\WinThemeSwitcher:
#               the per-user folder README recommends, since other standard
#               accounts can't replace the exe that HKCU\Run starts at login (a
#               folder like C:\Tools inherits Modify for every local account
#               from C:\). It must already exist -- for a first install, create
#               it once from a normal (unpackaged) shell. Deploying to a
#               different folder MOVES the install: every running instance is
#               stopped, and the copy started there (no config.json, so
#               defaults with auto_start on) points HKCU\Run at itself. Copy
#               config.json across when moving; to try a build, use -SkipCopy,
#               not a scratch -DeployDir.
#
# Exit codes:
#   0  build + sign + (optional) deploy + relaunch all succeeded
#   1  cargo missing or cargo build failed
#   2  produced exe path could not be determined from cargo output
#   3  signtool not found (caller should install Windows SDK)
#   4  signing failed: signtool never succeeded (AV holding the file, cert
#      missing, timestamp server unreachable), or the result is not a Valid,
#      timestamped signature
#   5  deploy directory missing (create it once for a first install)
#   6  deploy copy failed or the copy differs from the signed build (the
#      stopped instance is relaunched if its exe is still validly signed)
#   7  linker hardening missing: RUSTFLAGS/CARGO_ENCODED_RUSTFLAGS overrides
#      .cargo\config.toml, or the exe still imports VCRUNTIME140 or lacks
#      DependentLoadFlags 0x800 (see scripts\check-hardening.ps1)
#   8  deployed, but the relaunched app was not seen running
#   9  refused to deploy into AppData: this shell has a package identity (an
#      MSIX app's container), or that couldn't be checked
# 5 and 9 are checked before building, so a deploy that can't happen fails fast.
param(
    [switch]$SkipCopy,
    # Never empty: an empty path would resolve to the current directory. The
    # default comes from the known folder, so it needs no environment variable.
    [ValidateNotNullOrEmpty()]
    [string]$DeployDir = (Join-Path ([Environment]::GetFolderPath("LocalApplicationData")) "Programs\WinThemeSwitcher")
)

# PowerShell 5.1 wraps native stderr in ErrorRecords when redirected; check
# $LASTEXITCODE explicitly instead of using Stop on $ErrorActionPreference.
$ErrorActionPreference = "Continue"
$repo = Split-Path -Parent $PSScriptRoot
$cargo = "$env:USERPROFILE\.cargo\bin\cargo.exe"
$manifest = Join-Path $repo "Cargo.toml"
# A relative -DeployDir means relative to the caller's location: resolve it
# before the Push-Location below.
$DeployDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DeployDir)

# Cargo finds .cargo\config.toml by walking up from the CURRENT directory,
# not from --manifest-path - run from the repo so the linker flags apply no
# matter where this script is invoked from. Quit restores the caller's
# location (matters when the script is run in-session with `&`).
Push-Location -LiteralPath $repo
function Quit([int]$code) {
    Pop-Location
    exit $code
}

if (-not (Test-Path $cargo)) {
    Write-Error "cargo not found at $cargo -- install Rust via rustup"
    Quit 1
}
# Either variable REPLACES [target.*].rustflags from .cargo\config.toml.
if ($env:RUSTFLAGS -or $env:CARGO_ENCODED_RUSTFLAGS) {
    Write-Error "RUSTFLAGS / CARGO_ENCODED_RUSTFLAGS is set and would override .cargo\config.toml -- unset it"
    Quit 7
}

# Deploy preconditions, checked before the build so a deploy that can't
# happen fails fast instead of after it.
if (-not $SkipCopy) {
    $deployExe = Join-Path $DeployDir "win-theme-switcher.exe"

    # A packaged (MSIX) host app -- an agent's terminal, say -- may run this
    # shell with AppData writes virtualized: files and folders it creates there
    # (and any private copy that already exists) can live in the package's
    # private store, visible only from inside it. The copy below could then
    # "succeed" and hash-match from in here while the real folder never gets
    # it. So deploying into AppData needs proof that this process has no
    # package identity.
    $inAppData = @([Environment]::GetFolderPath("LocalApplicationData"), [Environment]::GetFolderPath("ApplicationData")) |
        Where-Object { $_ } |
        Where-Object { ($DeployDir.TrimEnd('\') + '\').StartsWith($_.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase) }
    if ($inAppData) {
        $pkgRc = $null
        try {
            if (-not ('WtsBuild.Package' -as [type])) {
                Add-Type -Namespace WtsBuild -Name Package -MemberDefinition '
                    [DllImport("kernel32.dll")]
                    public static extern int GetCurrentPackageFullName(ref int length, System.IntPtr name);'
            }
            $len = 0
            $pkgRc = [WtsBuild.Package]::GetCurrentPackageFullName([ref]$len, [IntPtr]::Zero)
        } catch { }
        # 15700 = APPMODEL_ERROR_NO_PACKAGE: no identity, so no virtualization.
        if ($pkgRc -ne 15700) {
            $why = if ($null -eq $pkgRc) { "the package-identity check failed" } else { "this shell runs inside a packaged app's container (code $pkgRc)" }
            Write-Error "refusing to deploy into $DeployDir -- $why, and AppData writes may be virtualized there; run this from a normal PowerShell window"
            Quit 9
        }
    }

    if (-not (Test-Path -LiteralPath $DeployDir -PathType Container)) {
        Write-Error "deploy directory $DeployDir doesn't exist or isn't a folder -- for a first install, create it from a normal (unpackaged) PowerShell window and re-run"
        Quit 5
    }
}

# Locate signtool (any Windows Kits version).
$signtool = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin\10.0.*\x64\signtool.exe" -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
if (-not $signtool) {
    Write-Error "signtool.exe not found under C:\Program Files (x86)\Windows Kits\10\bin\10.0.*\x64\ -- install Windows SDK"
    Quit 3
}

# Build with --message-format=json so we can locate the produced exe
# programmatically. --locked refuses if Cargo.lock is out of date with
# Cargo.toml: after a version bump run scripts\test.ps1 (or `cargo update -w`)
# first, then commit Cargo.lock with the bump.
$msgs = & $cargo build --release --locked --manifest-path $manifest --message-format=json
if ($LASTEXITCODE -ne 0) {
    # Re-run humanly for a readable compile error; always fail, even if the
    # re-run happens to succeed.
    & $cargo build --release --locked --manifest-path $manifest
    Quit 1
}
$exe = $msgs | ForEach-Object { $_ | ConvertFrom-Json } |
    Where-Object { $_.reason -eq "compiler-artifact" -and -not $_.profile.test -and $_.executable } |
    Select-Object -Last 1 -ExpandProperty executable
if (-not $exe) {
    Write-Error "could not determine release executable path from cargo output"
    Quit 2
}

# Verify the linker hardening landed in the PE itself. Run in-process (a
# script's `exit N` sets $LASTEXITCODE) with a seeded failure, so a check
# that can't run can never read as a pass.
$global:LASTEXITCODE = 1
& (Join-Path $PSScriptRoot "check-hardening.ps1") -Exe $exe
if ($LASTEXITCODE -ne 0) { Quit 7 }

# --- Sign ---------------------------------------------------------------------
# /tr + /td request an RFC 3161 timestamp from DigiCert -- without it, the
# Authenticode signature dies when the cert expires in 2036. AV may briefly
# hold the fresh file; retry up to 20 times (~10 s of wall time, plenty for
# the on-access scanner to release its handle).
$signed = $false
$signOut = $null
for ($i = 1; $i -le 20; $i++) {
    # Seed a failure: if signtool can't even launch, $LASTEXITCODE would
    # otherwise still hold cargo's 0.
    $global:LASTEXITCODE = 1
    $signOut = & $signtool sign /n "WinThemeSwitcher Self-Signed" /fd SHA256 `
        /tr http://timestamp.digicert.com /td SHA256 $exe 2>&1
    if ($LASTEXITCODE -eq 0) { $signed = $true; break }
    Start-Sleep -Milliseconds 500
}
if (-not $signed) {
    Write-Warning "release exe could not be signed (locked, cert missing, or timestamp server unreachable) -- refusing to copy"
    $signOut | Select-Object -Last 5 | ForEach-Object { Write-Warning "  signtool: $_" }
    Quit 4
}

# Trust signtool's exit code only as far as it goes: confirm the result is a
# Valid signature carrying the RFC 3161 countersignature.
$sig = Get-AuthenticodeSignature -FilePath $exe
if ($sig.Status -ne "Valid" -or -not $sig.TimeStamperCertificate) {
    Write-Warning "signature check failed: Status=$($sig.Status) Timestamped=$([bool]$sig.TimeStamperCertificate) -- refusing to copy"
    Quit 4
}

Write-Output "signed: $exe (Valid, timestamped)"

if ($SkipCopy) {
    Write-Output "SkipCopy set -- leaving the deployed binary in $DeployDir untouched"
    Quit 0
}

# --- Deploy -------------------------------------------------------------------
# (-DeployDir was checked before the build: it exists, and an AppData target
# was proven safe to write.)

# Relaunch through explorer.exe, NOT as a child of this shell: the tray app
# must run in the user's normal desktop context (the same one the HKCU\Run
# logon launch uses) - not inside this shell's job object, elevation, or an
# app container (e.g. an MSIX-packaged terminal/agent host, where HKCU writes
# are virtualized into a private hive and the Run value / theme registry
# writes would never reach the real registry). explorer.exe hands the launch
# to the running shell and returns immediately.
function Start-Deployed([string]$path) {
    Start-Process -FilePath "$env:WINDIR\explorer.exe" -ArgumentList "`"$path`""
    # A fresh build's first start can take several seconds while antivirus and
    # reputation checks inspect the new hash (~10 s seen), so allow 20 s.
    for ($i = 0; $i -lt 80; $i++) {
        Start-Sleep -Milliseconds 250
        $p = Get-Process -Name win-theme-switcher -ErrorAction SilentlyContinue |
            Where-Object { $_.Path -eq $path }
        if ($p) { return $true }
    }
    return $false
}

$running = Get-Process -Name win-theme-switcher -ErrorAction SilentlyContinue
# Remember where the stopped instance ran from: if the copy fails, that's the
# exe to bring back, and it needn't live in -DeployDir.
$prevExe = @($running | ForEach-Object { $_.Path } | Where-Object { $_ }) | Select-Object -First 1
if (-not $prevExe) { $prevExe = $deployExe }
if ($running) {
    $running | Stop-Process -Force
    # Wait for the processes to actually exit (and release the file) rather
    # than guessing with a fixed sleep.
    $running | Wait-Process -Timeout 10 -ErrorAction SilentlyContinue
}

$copyErr = $null
Copy-Item -LiteralPath $exe -Destination $deployExe -Force -ErrorVariable copyErr
$srcHash = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
$dstHash = (Get-FileHash -LiteralPath $deployExe -Algorithm SHA256 -ErrorAction SilentlyContinue).Hash
if ($copyErr -or $srcHash -ne $dstHash) {
    Write-Warning "deploy failed: copy error='$copyErr' src=$srcHash dst=$dstHash"
    # Don't leave the app down: relaunch the stopped instance's exe, if it's
    # still a validly signed exe.
    if ((Test-Path -LiteralPath $prevExe) -and (Get-AuthenticodeSignature -LiteralPath $prevExe).Status -eq "Valid") {
        if (Start-Deployed $prevExe) { Write-Warning "relaunched the previous exe: $prevExe" }
        else { Write-Warning "could not relaunch $prevExe" }
    } else {
        Write-Warning "no validly signed previous exe to relaunch -- start the app yourself"
    }
    Quit 6
}
Write-Output "deployed: $deployExe ($dstHash)"

if (Start-Deployed $deployExe) {
    Write-Output "relaunched: $deployExe"
    Quit 0
}
Write-Warning "deployed, but the relaunched process was not seen within 20 s -- Smart App Control blocks this self-signed exe when it's On; otherwise start it manually or sign out/in"
Quit 8
