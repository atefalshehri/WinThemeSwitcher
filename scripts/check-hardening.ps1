# Verifies what the build is meant to put into the exe, on the exe itself:
#  - the .cargo\config.toml linker hardening: no VCRUNTIME140 import (hybrid
#    CRT) and DependentLoadFlags = 0x800 (static imports resolve from System32
#    only - except comctl32.dll, which the manifest redirects to Windows'
#    side-by-side store, %SystemRoot%\WinSxS; see app.manifest);
#  - what build.rs embeds: the application manifest (RT_MANIFEST, id 1) with
#    its asInvoker, Common-Controls 6 and per-monitor-v2 statements and
#    without longPathAware or activeCodePage, and the version resource
#    (RT_VERSION, id 1);
#  - the version: the manifest's and the version resource's must be the one
#    in Cargo.toml, so an exe with a stale embedded version can't ship.
# Used by scripts\build.ps1 and by the CI workflow (ci.yml), so losing the
# config or the resources can't slip through CI.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\check-hardening.ps1 -Exe <path-to-exe>
# -CargoToml names the Cargo.toml to compare the version with (default: the
# repository's, next to this script's folder).
# Exit codes: 0 ok, 7 hardening, manifest or version resource missing or
# wrong, 2 file missing/unreadable or not a well-formed PE32+ image (the exe),
# or no package version found (Cargo.toml).
param(
    [Parameter(Mandatory = $true)][string]$Exe,
    [string]$CargoToml = ""
)

$ErrorActionPreference = "Continue"
# ($PSScriptRoot is still empty while Windows PowerShell 5.1 binds parameter
# defaults for a script started with -File.)
if (-not $CargoToml) { $CargoToml = Join-Path (Split-Path -Parent $PSScriptRoot) "Cargo.toml" }
if (-not (Test-Path -LiteralPath $Exe -PathType Leaf)) {
    Write-Error "check-hardening: $Exe not found"
    exit 2
}

# The [package] version in a Cargo.toml, or $null. Reads what this
# repository's Cargo.toml uses, not all of TOML: a [section] line (a comment
# may follow it) and version = "X" or 'X'. Anything else gives $null, which
# stops the check - it can't pass by misreading.
function Get-PackageVersion([string]$path) {
    $inPackage = $false
    foreach ($line in [IO.File]::ReadAllLines($path)) {
        if ($line -match '^\s*\[([^\]]+)\]\s*(#.*)?$') { $inPackage = ($Matches[1].Trim() -eq 'package'); continue }
        if ($inPackage -and $line -match '^\s*version\s*=\s*("([^"]+)"|''([^'']+)'')') { return "$($Matches[2])$($Matches[3])" }
    }
    return $null
}

# File offset of an RVA in a well-formed PE32+ image (see
# Get-DependentLoadFlags), or -1 when no section's raw data holds it.
function Get-RvaOffset([byte[]]$b, [long]$rva) {
    $pe = [BitConverter]::ToInt32($b, 0x3C)
    $numSections = [BitConverter]::ToUInt16($b, $pe + 6)
    $sec = $pe + 24 + [BitConverter]::ToUInt16($b, $pe + 20)
    for ($i = 0; $i -lt $numSections; $i++) {
        $s = $sec + $i * 40
        if ($s + 40 -gt $b.Length) { return -1 }
        $va = [BitConverter]::ToUInt32($b, $s + 12)
        $rawSize = [BitConverter]::ToUInt32($b, $s + 16)
        $raw = [BitConverter]::ToUInt32($b, $s + 20)
        if ($rva -ge $va -and ($rva - $va) -lt $rawSize) { return [long]$raw + ($rva - $va) }
    }
    return -1
}

# In the resource directory table at file offset $dir: the OffsetToData of
# the entry with numeric id $want (any entry when $want is negative), or
# $null. Entries with a string name have the top bit set and never match.
function Find-ResourceEntry([byte[]]$b, [long]$dir, [long]$want) {
    if ($dir -lt 0 -or $dir + 16 -gt $b.Length) { return $null }
    $count = [BitConverter]::ToUInt16($b, [int]$dir + 12) + [BitConverter]::ToUInt16($b, [int]$dir + 14)
    for ($i = 0; $i -lt $count; $i++) {
        $e = $dir + 16 + $i * 8
        if ($e + 8 -gt $b.Length) { return $null }
        if ($want -lt 0 -or [BitConverter]::ToUInt32($b, [int]$e) -eq $want) {
            return [long][BitConverter]::ToUInt32($b, [int]$e + 4)
        }
    }
    return $null
}

# The bytes of the resource with numeric type $type and numeric id $id (its
# first language), or $null when the image has none. The resource directory
# is data directory 2: a table per level (type, id, language), offsets
# relative to its start, the top bit marking "another table"; the last level
# points at an entry holding the data's RVA and size.
function Get-ResourceData([byte[]]$b, [int]$type, [int]$id) {
    $table = 2147483648   # 0x80000000
    $opt = [BitConverter]::ToInt32($b, 0x3C) + 24
    if ([BitConverter]::ToUInt32($b, $opt + 108) -lt 3) { return $null }
    $rva = [BitConverter]::ToUInt32($b, $opt + 112 + 2 * 8)
    if ($rva -eq 0) { return $null }
    $root = Get-RvaOffset $b $rva
    if ($root -lt 0) { return $null }
    $t = Find-ResourceEntry $b $root $type
    if ($null -eq $t -or $t -lt $table) { return $null }
    $n = Find-ResourceEntry $b ($root + $t - $table) $id
    if ($null -eq $n -or $n -lt $table) { return $null }
    $l = Find-ResourceEntry $b ($root + $n - $table) -1
    if ($null -eq $l -or $l -ge $table) { return $null }
    $entry = $root + $l
    if ($entry + 16 -gt $b.Length) { return $null }
    $size = [BitConverter]::ToUInt32($b, [int]$entry + 4)
    $at = Get-RvaOffset $b ([BitConverter]::ToUInt32($b, [int]$entry))
    if ($at -lt 0 -or $at + $size -gt $b.Length) { return $null }
    $data = New-Object byte[] $size
    [Array]::Copy($b, $at, $data, 0, $size)
    return , $data
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

# --- What build.rs embeds, and the version -----------------------------------
$version = $null
try {
    $version = Get-PackageVersion (Resolve-Path -LiteralPath $CargoToml -ErrorAction Stop).ProviderPath
} catch {
    Write-Error "check-hardening: cannot read ${CargoToml}: $_"
    exit 2
}
$parts = [regex]::Match("$version", '^(\d+)\.(\d+)\.(\d+)')
if (-not $parts.Success) {
    Write-Error "check-hardening: no [package] version found in $CargoToml"
    exit 2
}
# The four-part form in the manifest and the version resource's fixed part.
$numeric = "{0}.{1}.{2}.0" -f [int]$parts.Groups[1].Value, [int]$parts.Groups[2].Value, [int]$parts.Groups[3].Value
$why = "was the exe built from this repository, with its build.rs?"

$manifestBytes = Get-ResourceData $bytes 24 1   # RT_MANIFEST
if ($null -eq $manifestBytes) {
    Write-Error "check-hardening: $Exe has no application manifest (RT_MANIFEST, id 1) -- $why"
    exit 7
}
if ($null -eq (Get-ResourceData $bytes 16 1)) {   # RT_VERSION
    Write-Error "check-hardening: $Exe has no version resource (RT_VERSION, id 1) -- $why"
    exit 7
}

# The version resource, read the way Explorer and Task Manager read it.
$vi = [Diagnostics.FileVersionInfo]::GetVersionInfo((Resolve-Path -LiteralPath $Exe).ProviderPath)
$fileNumeric = "{0}.{1}.{2}.{3}" -f $vi.FileMajorPart, $vi.FileMinorPart, $vi.FileBuildPart, $vi.FilePrivatePart
$productNumeric = "{0}.{1}.{2}.{3}" -f $vi.ProductMajorPart, $vi.ProductMinorPart, $vi.ProductBuildPart, $vi.ProductPrivatePart
if ($vi.FileVersion -cne $version -or $vi.ProductVersion -cne $version -or $fileNumeric -ne $numeric -or $productNumeric -ne $numeric) {
    Write-Error ("check-hardening: {0} says file version '{1}' ({2}) and product version '{3}' ({4}), but Cargo.toml says {5} -- a stale build? Build again from the repo directory" -f $Exe, $vi.FileVersion, $fileNumeric, $vi.ProductVersion, $productNumeric, $version)
    exit 7
}
if ($vi.FileDescription -cne "WinThemeSwitcher") {
    Write-Error "check-hardening: $Exe has file description '$($vi.FileDescription)' (expected 'WinThemeSwitcher', the name Task Manager and Startup settings show)"
    exit 7
}

# The manifest.
$manifest = [Text.Encoding]::UTF8.GetString($manifestBytes)
$wanted = [ordered]@{
    "requestedExecutionLevel asInvoker"      = '<requestedExecutionLevel\s+level="asInvoker"\s+uiAccess="false"\s*/>'
    "the Common-Controls 6.0.0.0 dependency" = '<assemblyIdentity\s[^>]*name="Microsoft\.Windows\.Common-Controls"\s+version="6\.0\.0\.0"'
    "dpiAwareness PerMonitorV2"              = '>\s*PerMonitorV2\s*</dpiAwareness>'
    "assemblyIdentity version $numeric"      = '<assemblyIdentity\s[^>]*name="WinThemeSwitcher"\s+version="' + [regex]::Escape($numeric) + '"'
}
foreach ($what in $wanted.Keys) {
    if ($manifest -notmatch $wanted[$what]) {
        Write-Error "check-hardening: the manifest in $Exe lacks $what (Cargo.toml version: $version) -- see app.manifest and build.rs"
        exit 7
    }
}
# Settings that change how Windows APIs behave must stay out (app.manifest
# says why); build.rs strips comments, so any mention is the setting itself.
if ($manifest -match '(?i)longPathAware|activeCodePage') {
    Write-Error "check-hardening: the manifest in $Exe sets $($Matches[0]) -- it must not (see app.manifest)"
    exit 7
}
# The Common-Controls dependency must not name a language: with language="*"
# Windows looks for the assembly in a folder named for the UI language next
# to the exe first, and loads a comctl32.dll planted there (app.manifest).
if ($manifest -match '(?i)<assemblyIdentity\s[^>]*name="Microsoft\.Windows\.Common-Controls"[^>]*\slanguage\s*=') {
    Write-Error "check-hardening: the Common-Controls dependency in $Exe's manifest has a language attribute -- it must not (see app.manifest)"
    exit 7
}

Write-Output "hardening ok: no VCRUNTIME140 import, DependentLoadFlags=0x800, manifest (asInvoker, Common-Controls 6, PerMonitorV2), version $version = Cargo.toml ($Exe)"
exit 0
