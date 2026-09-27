# Runs the unit tests with every test binary Authenticode-signed BEFORE its
# first execution. Unsigned fresh-hash binaries in target\debug\deps trip
# Kaspersky File AV / KSN on this machine (the Trusted Applications rules
# only cover the two release paths); signing collapses the heuristic signal
# the same way it does for release builds. Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test.ps1 [testname-filter]
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test.ps1 -Filter "--harness-flag"
#
# Exit codes: 0 only if every test binary ran and passed; the test harness's
# own code (101) for failures; 1 build failure / launch failure; 3 signtool
# not found; 4 a test binary could not be signed. Unsigned binaries are never
# executed - that is the whole point of this wrapper.
param([string]$Filter = "")

# NOT "Stop": native tools (cargo, signtool) report progress on stderr, which
# PowerShell 5.1 wraps in ErrorRecords when redirected; exit codes are checked
# explicitly instead.
$ErrorActionPreference = "Continue"
$repo = Split-Path -Parent $PSScriptRoot
$cargo = "$env:USERPROFILE\.cargo\bin\cargo.exe"
$manifest = Join-Path $repo "Cargo.toml"

# Cargo finds .cargo\config.toml from the CURRENT directory, not from
# --manifest-path - run from the repo so tests build with the same linker
# flags as the release. Quit restores the caller's location.
Push-Location -LiteralPath $repo
function Quit([int]$code) {
    Pop-Location
    exit $code
}

if (-not (Test-Path $cargo)) {
    Write-Error "cargo not found at $cargo -- install Rust via rustup"
    Quit 1
}

# Locate signtool (any Windows Kits version).
$signtool = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin\10.0.*\x64\signtool.exe" -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
if (-not $signtool) {
    Write-Error "signtool.exe not found under C:\Program Files (x86)\Windows Kits\10\bin\10.0.*\x64\ -- install the Windows SDK (refusing to run unsigned test binaries)"
    Quit 3
}

# Build without running; parse the produced executable paths from JSON output.
$msgs = & $cargo test --no-run --manifest-path $manifest --message-format=json
if ($LASTEXITCODE -ne 0) {
    # Re-run humanly for a readable compile error; always fail, even if the
    # re-run happens to succeed.
    & $cargo test --no-run --manifest-path $manifest
    Quit 1
}
$exes = @($msgs | ForEach-Object { $_ | ConvertFrom-Json } |
    Where-Object { $_.reason -eq "compiler-artifact" -and $_.profile.test -and $_.executable } |
    Select-Object -ExpandProperty executable -Unique)
if ($exes.Count -eq 0) { Write-Error "could not determine any test executable path"; Quit 1 }

$worst = 0
foreach ($exe in $exes) {
    # Sign before first execution. AV may briefly hold the fresh file; retry.
    $out = $null
    foreach ($i in 1..20) {
        # Seed a failure: if signtool can't even launch, $LASTEXITCODE would
        # otherwise still hold cargo's 0.
        $global:LASTEXITCODE = 1
        $out = & $signtool sign /n "WinThemeSwitcher Self-Signed" /fd SHA256 $exe 2>&1
        if ($LASTEXITCODE -eq 0) { break }
        Start-Sleep -Milliseconds 500
    }
    # Trust the file, not the exit code.
    if ((Get-AuthenticodeSignature -LiteralPath $exe).Status -ne "Valid") {
        Write-Error "test binary is not validly signed (locked, cert missing, or signtool failed) - refusing to run it: $exe"
        $out | Select-Object -Last 3 | ForEach-Object { Write-Warning "  signtool: $_" }
        Quit 4
    }

    # Two PowerShell 5.1 traps, both of which made this script exit 0 no
    # matter what the tests did:
    #  1. #![windows_subsystem = "windows"] makes the TEST binary a GUI-
    #     subsystem exe too, and PowerShell does not wait for a GUI exe (nor
    #     set $LASTEXITCODE) unless its output is piped. Piping through
    #     ForEach-Object makes it wait, streams the output, and records the
    #     real exit code.
    #  2. A binary that cannot launch at all (e.g. blocked by AV) leaves
    #     $LASTEXITCODE untouched - seed a failure and catch the launch
    #     exception so that can never read as a pass.
    $global:LASTEXITCODE = 1
    try {
        if ($Filter) {
            & $exe $Filter 2>&1 | ForEach-Object { "$_" }
        } else {
            & $exe 2>&1 | ForEach-Object { "$_" }
        }
        $code = $LASTEXITCODE
    } catch {
        Write-Error "failed to launch test binary ${exe}: $_"
        $code = 1
    }
    # First non-zero wins (exit codes can be negative NTSTATUS values, so
    # never aggregate with a numeric max).
    if ($code -ne 0 -and $worst -eq 0) { $worst = $code }
}
Quit $worst
