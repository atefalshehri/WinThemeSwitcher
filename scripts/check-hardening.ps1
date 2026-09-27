# Verifies that the .cargo\config.toml linker hardening actually reached an
# exe: no VCRUNTIME140 import (hybrid CRT) and DependentLoadFlags = 0x800
# (static imports resolve from System32 only). Used by scripts\build.ps1 and
# by both GitHub workflows, so losing the config can't slip through CI.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\check-hardening.ps1 -Exe <path-to-exe>
# Exit codes: 0 ok, 7 hardening missing, 2 file missing/unreadable or not a
# well-formed PE32+ image.
param([Parameter(Mandatory = $true)][string]$Exe)

$ErrorActionPreference = "Continue"
if (-not (Test-Path -LiteralPath $Exe -PathType Leaf)) {
    Write-Error "check-hardening: $Exe not found"
    exit 2
}

# DependentLoadFlags lives in the load-config directory (data directory 10),
# a WORD at offset 0x4E of IMAGE_LOAD_CONFIG_DIRECTORY64. Returns $null for
# anything that isn't a well-formed PE32+ image; 0 when the load-config
# directory (or the field) is absent.
function Get-DependentLoadFlags([byte[]]$b) {
    if ($b.Length -lt 0x40 -or $b[0] -ne 0x4D -or $b[1] -ne 0x5A) { return $null }   # "MZ"
    $pe = [BitConverter]::ToInt32($b, 0x3C)
    if ($pe -lt 0 -or $pe + 24 + 112 + 11 * 8 -gt $b.Length) { return $null }
    if ([BitConverter]::ToUInt32($b, $pe) -ne 0x00004550) { return $null }           # "PE\0\0"
    $opt = $pe + 24
    if ([BitConverter]::ToUInt16($b, $opt) -ne 0x20B) { return $null }              # PE32+ only
    $numSections = [BitConverter]::ToUInt16($b, $pe + 6)
    $optSize = [BitConverter]::ToUInt16($b, $pe + 20)
    if ($optSize -lt 112 + 11 * 8) { return $null }
    if ([BitConverter]::ToUInt32($b, $opt + 108) -lt 11) { return 0 }               # NumberOfRvaAndSizes
    $lcRva = [BitConverter]::ToUInt32($b, $opt + 112 + 10 * 8)
    if ($lcRva -eq 0) { return 0 }
    $sec = $opt + $optSize
    for ($i = 0; $i -lt $numSections; $i++) {
        $s = $sec + $i * 40
        if ($s + 40 -gt $b.Length) { return $null }
        $va = [BitConverter]::ToUInt32($b, $s + 12)
        $rawSize = [BitConverter]::ToUInt32($b, $s + 16)
        $raw = [BitConverter]::ToUInt32($b, $s + 20)
        if ($lcRva -ge $va -and ($lcRva - $va) -lt $rawSize) {
            $lc = [long]$raw + ($lcRva - $va)
            if ($lc + 0x50 -gt $b.Length) { return $null }
            # The struct's own Size must cover the field (older linkers emit
            # shorter structs without DependentLoadFlags).
            if ([BitConverter]::ToUInt32($b, [int]$lc) -lt 0x50) { return 0 }
            return [BitConverter]::ToUInt16($b, [int]($lc + 0x4E))
        }
    }
    return $null
}

try {
    $bytes = [IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Exe).ProviderPath)
} catch {
    Write-Error "check-hardening: cannot read ${Exe}: $_"
    exit 2
}
$dlf = Get-DependentLoadFlags $bytes
if ($null -eq $dlf) {
    Write-Error "check-hardening: $Exe is not a well-formed PE32+ image"
    exit 2
}
$text = [Text.Encoding]::ASCII.GetString($bytes)
if ($text -match '(?i)vcruntime140') {
    Write-Error "check-hardening: $Exe imports VCRUNTIME140 -- the .cargo\config.toml hybrid-CRT flags did not apply (was cargo run outside the repo directory, or with RUSTFLAGS set?)"
    exit 7
}
if ($dlf -ne 0x800) {
    Write-Error ("check-hardening: {0} DependentLoadFlags = 0x{1:X} (expected 0x800) -- .cargo\config.toml flags did not apply" -f $Exe, $dlf)
    exit 7
}
Write-Output "hardening ok: no VCRUNTIME140 import, DependentLoadFlags=0x800 ($Exe)"
exit 0
