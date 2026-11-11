# Generate THIRD-PARTY-NOTICES.txt (repo root): the license and notice texts
# of everything third-party that goes into win-theme-switcher.exe. The MIT and
# Apache-2.0 licenses of the crates it is built from require their texts and
# copyright notices to accompany binary copies; scripts\release.ps1 puts this
# file into the release zip, and refuses to release while it is stale.
#
# Usage (from anywhere):
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\notices.ps1          # regenerate the file
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\notices.ps1 -Check   # compare only, change nothing
#
# Parameters:
#   -Check  generate in memory and compare with the file: exit 44, naming the
#           first line that differs, unless they are identical up to line
#           endings (git's autocrlf decides those in a checkout). CI and
#           release.ps1 run this.
#   -Path   the file to write or check, default THIRD-PARTY-NOTICES.txt in the
#           repo (relative to the caller's location; for trying things out).
#
# After adding or updating a dependency: run this script and commit the file
# together with Cargo.lock.
#
# What goes in:
#   * every crate in the resolved dependency graph of the Windows build:
#     cargo tree --locked --target x86_64-pc-windows-msvc --edges normal,build
#     (normal and build dependencies, proc-macro crates included; no
#     dev-dependencies; not the app itself). cargo tree with a --format string
#     gives one "name vX.Y.Z|license|repository" line per crate, so no JSON
#     has to be parsed (cargo metadata prints one large JSON document, and
#     Windows PowerShell 5.1's ConvertFrom-Json is unreliable on those).
#   * per crate: name, version, the license expression and repository as its
#     manifest declares them, and the text of every file in its package - in
#     its root or in any folder below, where a crate keeps the license of
#     code it has copied in (tracing-core: src/spin/LICENSE) - whose name
#     starts with LICENSE, LICENCE, UNLICENSE, COPYING, COPYRIGHT or NOTICE
#     (any letter case; Rust source files, *.rs, are never taken for one),
#     read from the cargo registry ($CARGO_HOME or %USERPROFILE%\.cargo,
#     registry\src\*\<name>-<version>; cargo fetch --locked is run once if a
#     crate is not there). Files below the root are listed with their path.
#   * a crate that ships no such file gets its declared authors and the
#     standard text of each license it names ($standardText below): no
#     copyright holder is made up for it.
#   * the Rust standard library, as fixed text (the toolchain is not read):
#     see $rustLines below, and re-check it after a Rust update.
#
# The output depends only on Cargo.lock and the crates' own files, so the same
# commit gives the same bytes on any machine: crates sorted by name, then
# version (ordinal; the version's three numbers numerically); no dates, paths,
# user names or tool versions; UTF-8 without BOM, LF line endings. Each text
# is decoded (UTF-8, or UTF-16 with a BOM, else ISO-8859-1), its line endings
# become LF, trailing spaces and tabs are cut from every line, and blank lines
# at its start and end are dropped. Texts that are then identical are printed
# once and referred to by number - so every crate keeps its own copyright
# lines: a text with another copyright line is another text.
#
# Exit codes (disjoint from build.ps1's 1-9 and release.ps1's 20-31):
#    0  written, or already up to date; with -Check: the file is current
#   40  cargo not found, or cargo tree failed (Cargo.lock out of date with
#       Cargo.toml, or a crate could not be downloaded)
#   41  cargo tree printed something unexpected, or a dependency that is not a
#       plain crates.io crate (path, git, another registry): its notices need
#       a decision by hand, then an extension of this script
#   42  a crate's source is not in the cargo registry, even after cargo fetch,
#       or its files could not be listed
#   43  a license text problem: a crate has a folder named for license texts
#       (LICENSES and the like) holding a file this script would not take
#       for one, a file is empty or binary, or a crate without license files
#       names a license that has no standard text in this script
#   44  -Check: the file is missing or differs from what would be generated
#   45  the file could not be written
param(
    [switch]$Check,
    [string]$Path = ""
)

# NOT "Stop": cargo reports progress on stderr, which PowerShell 5.1 wraps in
# ErrorRecords when redirected; exit codes are checked explicitly instead.
$ErrorActionPreference = "Continue"
$repo = Split-Path -Parent $PSScriptRoot
$target = "x86_64-pc-windows-msvc"
$fileName = "THIRD-PARTY-NOTICES.txt"
# A relative -Path means relative to the caller's location: resolve it before
# the Push-Location below.
if ($Path) {
    $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
} else {
    $Path = Join-Path $repo $fileName
}

# cargo reads Cargo.lock (and .cargo\config.toml) from the current directory.
# Quit restores the caller's location (matters when run in-session with `&`,
# as release.ps1 does).
Push-Location -LiteralPath $repo
function Quit([int]$code) {
    Pop-Location
    exit $code
}
function Fail([int]$code, [string]$msg) {
    [Console]::Error.WriteLine("notices: FAILED (exit $code): $msg")
    Quit $code
}

# --- Fixed texts ----------------------------------------------------------------
# The standard license texts, for the Rust standard library and for crates
# whose package has no license file. Both are the texts the Rust project and
# many crates ship (the MIT text without a copyright line; the Apache text
# without the "how to apply" appendix), so they usually coincide with a text
# from the graph and are not printed twice. Keep this file pure ASCII:
# PowerShell 5.1 reads a script without a BOM in the system's ANSI code page.
$standardText = New-Object 'System.Collections.Generic.Dictionary[string,string]'
$standardText["Apache-2.0"] = @'
                              Apache License
                        Version 2.0, January 2004
                     http://www.apache.org/licenses/

TERMS AND CONDITIONS FOR USE, REPRODUCTION, AND DISTRIBUTION

1. Definitions.

   "License" shall mean the terms and conditions for use, reproduction,
   and distribution as defined by Sections 1 through 9 of this document.

   "Licensor" shall mean the copyright owner or entity authorized by
   the copyright owner that is granting the License.

   "Legal Entity" shall mean the union of the acting entity and all
   other entities that control, are controlled by, or are under common
   control with that entity. For the purposes of this definition,
   "control" means (i) the power, direct or indirect, to cause the
   direction or management of such entity, whether by contract or
   otherwise, or (ii) ownership of fifty percent (50%) or more of the
   outstanding shares, or (iii) beneficial ownership of such entity.

   "You" (or "Your") shall mean an individual or Legal Entity
   exercising permissions granted by this License.

   "Source" form shall mean the preferred form for making modifications,
   including but not limited to software source code, documentation
   source, and configuration files.

   "Object" form shall mean any form resulting from mechanical
   transformation or translation of a Source form, including but
   not limited to compiled object code, generated documentation,
   and conversions to other media types.

   "Work" shall mean the work of authorship, whether in Source or
   Object form, made available under the License, as indicated by a
   copyright notice that is included in or attached to the work
   (an example is provided in the Appendix below).

   "Derivative Works" shall mean any work, whether in Source or Object
   form, that is based on (or derived from) the Work and for which the
   editorial revisions, annotations, elaborations, or other modifications
   represent, as a whole, an original work of authorship. For the purposes
   of this License, Derivative Works shall not include works that remain
   separable from, or merely link (or bind by name) to the interfaces of,
   the Work and Derivative Works thereof.

   "Contribution" shall mean any work of authorship, including
   the original version of the Work and any modifications or additions
   to that Work or Derivative Works thereof, that is intentionally
   submitted to Licensor for inclusion in the Work by the copyright owner
   or by an individual or Legal Entity authorized to submit on behalf of
   the copyright owner. For the purposes of this definition, "submitted"
   means any form of electronic, verbal, or written communication sent
   to the Licensor or its representatives, including but not limited to
   communication on electronic mailing lists, source code control systems,
   and issue tracking systems that are managed by, or on behalf of, the
   Licensor for the purpose of discussing and improving the Work, but
   excluding communication that is conspicuously marked or otherwise
   designated in writing by the copyright owner as "Not a Contribution."

   "Contributor" shall mean Licensor and any individual or Legal Entity
   on behalf of whom a Contribution has been received by Licensor and
   subsequently incorporated within the Work.

2. Grant of Copyright License. Subject to the terms and conditions of
   this License, each Contributor hereby grants to You a perpetual,
   worldwide, non-exclusive, no-charge, royalty-free, irrevocable
   copyright license to reproduce, prepare Derivative Works of,
   publicly display, publicly perform, sublicense, and distribute the
   Work and such Derivative Works in Source or Object form.

3. Grant of Patent License. Subject to the terms and conditions of
   this License, each Contributor hereby grants to You a perpetual,
   worldwide, non-exclusive, no-charge, royalty-free, irrevocable
   (except as stated in this section) patent license to make, have made,
   use, offer to sell, sell, import, and otherwise transfer the Work,
   where such license applies only to those patent claims licensable
   by such Contributor that are necessarily infringed by their
   Contribution(s) alone or by combination of their Contribution(s)
   with the Work to which such Contribution(s) was submitted. If You
   institute patent litigation against any entity (including a
   cross-claim or counterclaim in a lawsuit) alleging that the Work
   or a Contribution incorporated within the Work constitutes direct
   or contributory patent infringement, then any patent licenses
   granted to You under this License for that Work shall terminate
   as of the date such litigation is filed.

4. Redistribution. You may reproduce and distribute copies of the
   Work or Derivative Works thereof in any medium, with or without
   modifications, and in Source or Object form, provided that You
   meet the following conditions:

   (a) You must give any other recipients of the Work or
       Derivative Works a copy of this License; and

   (b) You must cause any modified files to carry prominent notices
       stating that You changed the files; and

   (c) You must retain, in the Source form of any Derivative Works
       that You distribute, all copyright, patent, trademark, and
       attribution notices from the Source form of the Work,
       excluding those notices that do not pertain to any part of
       the Derivative Works; and

   (d) If the Work includes a "NOTICE" text file as part of its
       distribution, then any Derivative Works that You distribute must
       include a readable copy of the attribution notices contained
       within such NOTICE file, excluding those notices that do not
       pertain to any part of the Derivative Works, in at least one
       of the following places: within a NOTICE text file distributed
       as part of the Derivative Works; within the Source form or
       documentation, if provided along with the Derivative Works; or,
       within a display generated by the Derivative Works, if and
       wherever such third-party notices normally appear. The contents
       of the NOTICE file are for informational purposes only and
       do not modify the License. You may add Your own attribution
       notices within Derivative Works that You distribute, alongside
       or as an addendum to the NOTICE text from the Work, provided
       that such additional attribution notices cannot be construed
       as modifying the License.

   You may add Your own copyright statement to Your modifications and
   may provide additional or different license terms and conditions
   for use, reproduction, or distribution of Your modifications, or
   for any such Derivative Works as a whole, provided Your use,
   reproduction, and distribution of the Work otherwise complies with
   the conditions stated in this License.

5. Submission of Contributions. Unless You explicitly state otherwise,
   any Contribution intentionally submitted for inclusion in the Work
   by You to the Licensor shall be under the terms and conditions of
   this License, without any additional terms or conditions.
   Notwithstanding the above, nothing herein shall supersede or modify
   the terms of any separate license agreement you may have executed
   with Licensor regarding such Contributions.

6. Trademarks. This License does not grant permission to use the trade
   names, trademarks, service marks, or product names of the Licensor,
   except as required for reasonable and customary use in describing the
   origin of the Work and reproducing the content of the NOTICE file.

7. Disclaimer of Warranty. Unless required by applicable law or
   agreed to in writing, Licensor provides the Work (and each
   Contributor provides its Contributions) on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or
   implied, including, without limitation, any warranties or conditions
   of TITLE, NON-INFRINGEMENT, MERCHANTABILITY, or FITNESS FOR A
   PARTICULAR PURPOSE. You are solely responsible for determining the
   appropriateness of using or redistributing the Work and assume any
   risks associated with Your exercise of permissions under this License.

8. Limitation of Liability. In no event and under no legal theory,
   whether in tort (including negligence), contract, or otherwise,
   unless required by applicable law (such as deliberate and grossly
   negligent acts) or agreed to in writing, shall any Contributor be
   liable to You for damages, including any direct, indirect, special,
   incidental, or consequential damages of any character arising as a
   result of this License or out of the use or inability to use the
   Work (including but not limited to damages for loss of goodwill,
   work stoppage, computer failure or malfunction, or any and all
   other commercial damages or losses), even if such Contributor
   has been advised of the possibility of such damages.

9. Accepting Warranty or Additional Liability. While redistributing
   the Work or Derivative Works thereof, You may choose to offer,
   and charge a fee for, acceptance of support, warranty, indemnity,
   or other liability obligations and/or rights consistent with this
   License. However, in accepting such obligations, You may act only
   on Your own behalf and on Your sole responsibility, not on behalf
   of any other Contributor, and only if You agree to indemnify,
   defend, and hold each Contributor harmless for any liability
   incurred by, or claims asserted against, such Contributor by reason
   of your accepting any such warranty or additional liability.

END OF TERMS AND CONDITIONS
'@
$standardText["MIT"] = @'
Permission is hereby granted, free of charge, to any
person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the
Software without restriction, including without
limitation the rights to use, copy, modify, merge,
publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software
is furnished to do so, subject to the following
conditions:

The above copyright notice and this permission notice
shall be included in all copies or substantial portions
of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF
ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT
SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR
IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
DEALINGS IN THE SOFTWARE.
'@

# The Rust standard library's entry. Fixed text, so the output never depends
# on the installed toolchain - which also means nothing updates it: after a
# Rust update, compare it with the toolchain's own copyright document,
# <toolchain>\share\doc\rust\COPYRIGHT-library.html (written from it as of
# Rust 1.99.1, 2026-11-11). That document covers everything the Rust project
# builds for the library, tools included; listed here are the parts linked
# into a x86_64-pc-windows-msvc exe that carry a notice of their own: its
# "In-tree files" exceptions for core's Unicode tables, backtrace and mpmc,
# and the crates.io crates found as rlibs in the toolchain's
# lib\rustlib\x86_64-pc-windows-msvc\lib that std links (cfg-if, hashbrown,
# rustc-demangle, windows-link; not getopts and rustc-literal-escaper, which
# only test and proc_macro use). Copyright lines are copied from that
# document as they stand; no crate versions, which change with every Rust
# release. The Unicode text is the toolchain's share\doc\rust\licenses\
# Unicode-3.0.txt; its copyright sign is put in below to keep this file ASCII.
$rustTitle = "Rust standard library and runtime"
$rustLines = @(
    "  License:   MIT OR Apache-2.0",
    "  Copyright: The Rust Project Developers (see https://thanks.rust-lang.org)",
    "  Source:    https://github.com/rust-lang/rust"
)
$rustNote = @(
    "  Statically linked into the exe: std, core, alloc and the crates the",
    "  standard library is itself built from. This entry is fixed text (the",
    "  toolchain is not read to make this file), written from the copyright",
    "  notices that come with the Rust toolchain, in",
    "  share/doc/rust/COPYRIGHT-library.html. Parts of the library with a",
    "  notice of their own, each under MIT OR Apache-2.0 (texts above) unless",
    "  it says otherwise:",
    "    core's Unicode tables (library/core/src/unicode)",
    "      License:   Unicode-3.0",
    "      Copyright: 1991-2024 Unicode, Inc",
    "      Text:      Unicode-3.0 [{UNICODE}]",
    "    backtrace (library/backtrace)",
    "      Copyright: 2014 Alex Crichton",
    "    std::sync::mpmc (library/std/src/sync/mpmc)",
    "      Copyright: 2019 The Crossbeam Project Developers",
    "    cfg-if (https://crates.io/crates/cfg-if)",
    "      Copyright (c) 2014 Alex Crichton",
    "    hashbrown (https://crates.io/crates/hashbrown)",
    "      Copyright (c) 2016 Amanieu d'Antras",
    "    rustc-demangle (https://crates.io/crates/rustc-demangle)",
    "      Copyright (c) 2014 Alex Crichton",
    "    windows-link (https://crates.io/crates/windows-link)",
    "      Copyright (c) Microsoft Corporation."
)
$rustUnicodeText = @'
UNICODE LICENSE V3

COPYRIGHT AND PERMISSION NOTICE

Copyright {COPYRIGHT SIGN} 1991-2024 Unicode, Inc.

NOTICE TO USER: Carefully read the following legal agreement. BY
DOWNLOADING, INSTALLING, COPYING OR OTHERWISE USING DATA FILES, AND/OR
SOFTWARE, YOU UNEQUIVOCALLY ACCEPT, AND AGREE TO BE BOUND BY, ALL OF THE
TERMS AND CONDITIONS OF THIS AGREEMENT. IF YOU DO NOT AGREE, DO NOT
DOWNLOAD, INSTALL, COPY, DISTRIBUTE OR USE THE DATA FILES OR SOFTWARE.

Permission is hereby granted, free of charge, to any person obtaining a
copy of data files and any associated documentation (the "Data Files") or
software and any associated documentation (the "Software") to deal in the
Data Files or Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, and/or sell
copies of the Data Files or Software, and to permit persons to whom the
Data Files or Software are furnished to do so, provided that either (a)
this copyright and permission notice appear with all copies of the Data
Files or Software, or (b) this copyright and permission notice appear in
associated Documentation.

THE DATA FILES AND SOFTWARE ARE PROVIDED "AS IS", WITHOUT WARRANTY OF ANY
KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT OF
THIRD PARTY RIGHTS.

IN NO EVENT SHALL THE COPYRIGHT HOLDER OR HOLDERS INCLUDED IN THIS NOTICE
BE LIABLE FOR ANY CLAIM, OR ANY SPECIAL INDIRECT OR CONSEQUENTIAL DAMAGES,
OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION,
ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THE DATA
FILES OR SOFTWARE.

Except as contained in this notice, the name of a copyright holder shall
not be used in advertising or otherwise to promote the sale, use or other
dealings in these Data Files or Software without prior written
authorization of the copyright holder.
'@
$rustUnicodeText = $rustUnicodeText.Replace("{COPYRIGHT SIGN}", [string][char]0xA9)

# --- Helpers --------------------------------------------------------------------
$ordinal = [StringComparer]::Ordinal
$utf8Strict = New-Object Text.UTF8Encoding($false, $true)
$utf8NoBom = New-Object Text.UTF8Encoding($false)
# Nothing here may depend on the machine's culture. PowerShell's -match and
# -split ignore case by the CURRENT culture: under a Turkish locale "I" and
# "i" are not each other's case, so -match would not find "license" in
# "LICENSE", and '[^\x20-\x7E]' matches the I of "MIT" (its lower case there
# is a non-ASCII letter). So: -cmatch / -csplit, or a [regex] with
# CultureInvariant, and dictionaries with ordinal keys (not @{}).
$rxOptions = [Text.RegularExpressions.RegexOptions]"IgnoreCase, CultureInvariant"
$rxNoticeFile = [regex]::new('^(LICEN[SC]E|UNLICEN[SC]E|COPYING|COPYRIGHT|NOTICE)', $rxOptions)
$rxNoticeDir = [regex]::new('^(LICEN[SC]ES?|COPYING|NOTICES?)$', $rxOptions)
$rxPackage = [regex]::new('^([A-Za-z0-9_-]+) v([0-9A-Za-z.+-]+)( \(proc-macro\))?( \(\*\))?$')

# Run cargo. Returns its stdout lines as strings; sets $script:rc (seeded, so
# a launch failure can't read as 0) and $script:errText (its stderr).
function Invoke-Cargo {
    $global:LASTEXITCODE = 1
    $all = & $cargo @args 2>&1
    $script:rc = $LASTEXITCODE
    $script:errText = (@($all | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } |
                ForEach-Object { "$_" }) -join "`n")
    return , @($all | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] } |
            ForEach-Object { "$_" })
}

# A file's text: UTF-8 (with or without BOM), UTF-16 with a BOM, and
# ISO-8859-1 for anything that is not valid UTF-8.
function Read-TextFile([string]$file) {
    $b = [IO.File]::ReadAllBytes($file)
    $t = $null
    if ($b.Length -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE) {
        $t = [Text.Encoding]::Unicode.GetString($b, 2, $b.Length - 2)
    } elseif ($b.Length -ge 2 -and $b[0] -eq 0xFE -and $b[1] -eq 0xFF) {
        $t = [Text.Encoding]::BigEndianUnicode.GetString($b, 2, $b.Length - 2)
    } else {
        try { $t = $utf8Strict.GetString($b) } catch { $t = $null }
        if ($null -eq $t) { $t = [Text.Encoding]::GetEncoding(28591).GetString($b) }
    }
    return $t
}

# The normal form every text is stored and compared in (see the header).
function Format-Text([string]$t) {
    $t = $t.Replace("`r`n", "`n").Replace("`r", "`n")
    if ($t.Length -gt 0 -and $t[0] -eq [char]0xFEFF) { $t = $t.Substring(1) }
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($l in $t.Split([char]10)) { $lines.Add($l.TrimEnd([char]32, [char]9)) }
    while ($lines.Count -gt 0 -and $lines[0].Length -eq 0) { $lines.RemoveAt(0) }
    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1].Length -eq 0) { $lines.RemoveAt($lines.Count - 1) }
    return [string]::Join("`n", $lines.ToArray())
}

# Texts are numbered in the order they are first used; a text already seen
# keeps its number and gets one more user.
$textNumber = New-Object 'System.Collections.Generic.Dictionary[string,int]'
$textBodies = New-Object 'System.Collections.Generic.List[string]'
$textUsers = New-Object 'System.Collections.Generic.List[object]'
function Add-Text([string]$body, [string]$user) {
    $n = 0
    if (-not $textNumber.TryGetValue($body, [ref]$n)) {
        $textBodies.Add($body)
        $textUsers.Add((New-Object 'System.Collections.Generic.List[string]'))
        $n = $textBodies.Count
        $textNumber.Add($body, $n)
    }
    $textUsers[$n - 1].Add($user)
    return $n
}

# name + version -> a key that sorts by name, then by the version's numbers.
# A space sorts below every character a crate name can have.
function Get-SortKey([string]$name, [string]$version) {
    $v = $version
    if ($version -cmatch '^([0-9]{1,9})\.([0-9]{1,9})\.([0-9]{1,9})(.*)$') {
        $v = $Matches[1].PadLeft(9, [char]48) + "." + $Matches[2].PadLeft(9, [char]48) + "." +
        $Matches[3].PadLeft(9, [char]48) + $Matches[4]
    }
    return "$name $v"
}

function Find-CrateDir([string]$name, [string]$version) {
    if (-not (Test-Path -LiteralPath $registrySrc -PathType Container)) { return $null }
    [string[]]$indexes = [IO.Directory]::GetDirectories($registrySrc)
    [Array]::Sort($indexes, $ordinal)
    foreach ($i in $indexes) {
        $d = Join-Path $i "$name-$version"
        # .cargo-ok is what cargo leaves once a package is fully unpacked.
        if (Test-Path -LiteralPath (Join-Path $d ".cargo-ok") -PathType Leaf) { return $d }
    }
    return $null
}

# The strings of a string-array key ("authors") in a manifest's [package]
# table. Cargo normalizes the manifests it publishes, so a small pattern is
# enough; anything it can't read gives an empty list.
function Get-ManifestList([string]$dir, [string]$key) {
    $m = (Read-TextFile (Join-Path $dir "Cargo.toml")).Replace("`r`n", "`n")
    $found = @()
    if ($m -cmatch '(?ms)^\[package\]\n(.*?)(?=^\[|\z)') {
        $pkg = $Matches[1]
        if ($pkg -cmatch ('(?ms)^' + [regex]::Escape($key) + '\s*=\s*\[(.*?)\]\s*$')) {
            foreach ($s in [regex]::Matches($Matches[1], '"((?:[^"\\]|\\.)*)"')) {
                $found += $s.Groups[1].Value.Replace('\"', '"').Replace('\\', '\')
            }
        }
    }
    return , $found
}

# First difference between two LF texts, as a message; "" if they are equal.
function Get-FirstDifference([string]$have, [string]$want) {
    if ([string]::Equals($have, $want, [StringComparison]::Ordinal)) { return "" }
    $a = $have.Split([char]10)
    $b = $want.Split([char]10)
    for ($i = 0; $i -lt [Math]::Max($a.Length, $b.Length); $i++) {
        $n = $i + 1
        if ($i -ge $a.Length) { return "line ${n}: the file ends, expected '$(Limit-Line $b[$i])'" }
        if ($i -ge $b.Length) { return "line ${n}: the file goes on with '$(Limit-Line $a[$i])', expected its end" }
        if (-not [string]::Equals($a[$i], $b[$i], [StringComparison]::Ordinal)) {
            return "line ${n}: the file has '$(Limit-Line $a[$i])', expected '$(Limit-Line $b[$i])'"
        }
    }
    return "the texts differ"
}
function Limit-Line([string]$s) {
    if ($s.Length -gt 100) { return $s.Substring(0, 100) + "..." }
    return $s
}

# --- The dependency graph ---------------------------------------------------------
$cargoHome = $env:CARGO_HOME
if (-not $cargoHome) { $cargoHome = Join-Path $env:USERPROFILE ".cargo" }
$cargo = Join-Path $cargoHome "bin\cargo.exe"
if (-not (Test-Path -LiteralPath $cargo -PathType Leaf)) {
    $onPath = Get-Command cargo.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $onPath) { Fail 40 "cargo not found at $cargo or on PATH -- install Rust via rustup" }
    $cargo = $onPath.Path
}
$registrySrc = Join-Path $cargoHome "registry\src"

# One line per package: "{p}|{l}|{r}" = "name vX.Y.Z[ (proc-macro)][ (source)][ (*)]",
# license expression, repository. The first line is the app itself.
# --color never: CI sets CARGO_TERM_COLOR=always, and cargo then wraps the
# "(*)" repeat marker in ANSI escape codes even when its output is captured.
$lines = Invoke-Cargo tree --color never --locked --target $target --edges "normal,build" --prefix none --format "{p}|{l}|{r}"
if ($rc -ne 0) { Fail 40 "cargo tree failed (exit $rc; is Cargo.lock out of date with Cargo.toml?): $errText" }
$lines = @($lines | Where-Object { $_.Trim() })
if ($lines.Count -lt 1 -or $lines[0] -cnotmatch '^\S+ v\S+ \(.+\)\|[^|]*\|') {
    Fail 41 "unexpected first line from cargo tree (expected the app itself, with its path)"
}

$crates = New-Object 'System.Collections.Generic.Dictionary[string,object]'
for ($i = 1; $i -lt $lines.Count; $i++) {
    $l = $lines[$i]
    # Captured console output: anything outside ASCII could decode differently
    # from one machine to the next.
    if ($l -cmatch '[^\x20-\x7E]') { Fail 41 "cargo tree printed a line with characters outside ASCII: $l" }
    $f = $l -csplit '\|', 3
    if ($f.Count -ne 3) { Fail 41 "unexpected line from cargo tree: $l" }
    $m = $rxPackage.Match($f[0])
    if (-not $m.Success) {
        Fail 41 "'$($f[0])' is not a plain crates.io crate (a path, git or other-registry dependency?) - decide how its notices are included, then extend scripts\notices.ps1"
    }
    $name = $m.Groups[1].Value
    $version = $m.Groups[2].Value
    $key = Get-SortKey $name $version
    if (-not $crates.ContainsKey($key)) {
        $crates.Add($key, @{ name = $name; version = $version; license = $f[1].Trim(); repository = $f[2].Trim(); dir = $null })
    }
}
if ($crates.Count -eq 0) { Fail 41 "cargo tree lists no dependencies" }
$keys = New-Object 'System.Collections.Generic.List[string]'
foreach ($k in $crates.Keys) { $keys.Add($k) }
$keys.Sort($ordinal)

# Locate every crate's unpacked package; download what is missing, once.
$fetched = $false
foreach ($k in $keys) {
    $c = $crates[$k]
    $c.dir = Find-CrateDir $c.name $c.version
    if (-not $c.dir -and -not $fetched) {
        $fetched = $true
        Write-Output "notices: $($c.name) $($c.version) is not in the cargo registry -- running cargo fetch --locked"
        $null = Invoke-Cargo fetch --color never --locked --target $target
        if ($rc -ne 0) { Fail 40 "cargo fetch failed (exit $rc): $errText" }
        $c.dir = Find-CrateDir $c.name $c.version
    }
    if (-not $c.dir) { Fail 42 "the source of $($c.name) $($c.version) is not under $registrySrc\*\$($c.name)-$($c.version)" }
}

# --- Components -------------------------------------------------------------------
# Each component: a title and its indented lines. The Rust entry comes first,
# so the two standard texts are numbers 1 and 2 and its Unicode text is 3.
$components = New-Object 'System.Collections.Generic.List[object]'

$body = New-Object 'System.Collections.Generic.List[string]'
foreach ($l in $rustLines) { $body.Add($l) }
$label = "  Texts:     "
foreach ($n in @("Apache-2.0", "MIT")) {
    $number = Add-Text (Format-Text $standardText[$n]) "${rustTitle}: the standard $n text"
    $body.Add("$label$n [$number]")
    $label = "             "
}
$number = Add-Text (Format-Text $rustUnicodeText) "${rustTitle}: the Unicode-3.0 text, for core's Unicode tables"
foreach ($l in $rustNote) { $body.Add($l.Replace("{UNICODE}", "$number")) }
$components.Add(@{ title = $rustTitle; lines = $body })

$withoutFiles = New-Object 'System.Collections.Generic.List[string]'
foreach ($k in $keys) {
    $c = $crates[$k]
    $id = "$($c.name) $($c.version)"
    $body = New-Object 'System.Collections.Generic.List[string]'
    $body.Add("  License:   $(if ($c.license) { $c.license } else { 'not declared as an expression; see its files' })")
    $body.Add("  Source:    $(if ($c.repository) { $c.repository } else { "https://crates.io/crates/$($c.name)" })")

    # Every file of the package, at any depth: a crate that has copied code
    # in keeps that code's license next to it (tracing-core: src/spin/LICENSE).
    # Paths are relative to the package, with forward slashes. A Rust source
    # file is never a notice file, whatever its name (src/license.rs). A
    # folder named for license texts must hold nothing else this script
    # would leave out: LICENSES/MIT.txt has to stop it, not be skipped.
    $root = $c.dir.TrimEnd([char]92)
    $paths = $null
    try { $paths = [IO.Directory]::GetFiles($root, "*", [IO.SearchOption]::AllDirectories) } catch { $paths = $null }
    if ($null -eq $paths) { Fail 42 "cannot list the files of $id under $registrySrc" }
    $found = New-Object 'System.Collections.Generic.List[string]'
    foreach ($p in $paths) {
        $rel = $p.Substring($root.Length + 1).Replace([char]92, [char]47)
        $parts = $rel.Split([char]47)
        $leaf = $parts[$parts.Length - 1]
        if ($leaf.EndsWith(".rs", [StringComparison]::OrdinalIgnoreCase)) { continue }
        if ($rxNoticeFile.IsMatch($leaf)) { $found.Add($rel); continue }
        for ($j = 0; $j -lt $parts.Length - 1; $j++) {
            if ($rxNoticeDir.IsMatch($parts[$j])) {
                Fail 43 "$id has a folder named for license texts ($($parts[$j])) with a file this script does not take for one ($rel) - decide whether it belongs in the notices, then extend scripts\notices.ps1"
            }
        }
    }
    [string[]]$files = $found.ToArray()
    [Array]::Sort($files, $ordinal)

    if ($files.Length -gt 0) {
        $label = "  Files:     "
        foreach ($f in $files) {
            $raw = Read-TextFile (Join-Path $root $f.Replace([char]47, [char]92))
            if ($raw.IndexOf([char]0) -ge 0) { Fail 43 "${id}: $f is not a text file" }
            $text = Format-Text $raw
            if (-not $text) { Fail 43 "${id}: $f is empty" }
            $number = Add-Text $text "${id}: $f"
            $body.Add("$label$f [$number]")
            $label = "             "
        }
    } else {
        # No license file in the package: say so, give the authors as declared,
        # and attach the standard text of each license the expression names.
        $withoutFiles.Add($id)
        $authors = Get-ManifestList $c.dir "authors"
        $body.Add("  Authors:   $(if ($authors.Count -gt 0) { $authors -join ', ' } else { 'not declared' })")
        $body.Add("  Files:     none - the package contains no license or notice file and")
        $body.Add("             states no copyright holder. Standard texts of the licenses")
        $body.Add("             it declares:")
        $names = @($c.license -csplit '[^A-Za-z0-9.+-]+' | Where-Object { $_ -and @("OR", "AND", "WITH") -cnotcontains $_ })
        if ($names.Count -eq 0) { Fail 43 "$id ships no license file and declares no license" }
        $seen = New-Object 'System.Collections.Generic.Dictionary[string,bool]'
        foreach ($n in $names) {
            if ($seen.ContainsKey($n)) { continue }
            $seen.Add($n, $true)
            # Exact spelling only (ordinal): the ids as $standardText has them.
            if (-not $standardText.ContainsKey($n)) {
                Fail 43 "$id ships no license file, and scripts\notices.ps1 has no standard text for '$n' (license: $($c.license)) - add one to `$standardText"
            }
            $number = Add-Text (Format-Text $standardText[$n]) "${id}: the standard $n text (no license file in the package)"
            $body.Add("             $n [$number]")
        }
    }
    $components.Add(@{ title = $id; lines = $body })
}

# --- The document -----------------------------------------------------------------
$rule = "=" * 78
$out = New-Object 'System.Collections.Generic.List[string]'
$out.Add("THIRD-PARTY NOTICES")
$out.Add("")
$out.Add("WinThemeSwitcher (win-theme-switcher.exe) is released under the MIT license;")
$out.Add("see the LICENSE file. The exe is built from, and contains code of, the")
$out.Add("third-party components listed here: the Rust standard library and $($crates.Count) crates")
$out.Add("from crates.io - the dependency graph of the $target build")
$out.Add("as locked in Cargo.lock, including crates that are only used while building")
$out.Add("(procedural macros, build scripts).")
$out.Add("")
$out.Add("Part 1 lists each component with its license as its authors declare it, and")
$out.Add("the license and notice files its package contains, in any of its folders.")
$out.Add("Part 2 holds the texts of those files. A text that several components ship")
$out.Add("in the same form (line endings and trailing spaces aside) is printed once;")
$out.Add("the numbers in square brackets refer to the texts in part 2.")
$out.Add("")
$out.Add("This file is generated by scripts\notices.ps1 - do not edit it by hand.")
$out.Add("")
$out.Add($rule)
$out.Add("PART 1 - COMPONENTS")
$out.Add($rule)
foreach ($c in $components) {
    $out.Add("")
    $out.Add($c.title)
    foreach ($l in $c.lines) { $out.Add($l) }
}
$out.Add("")
$out.Add($rule)
$out.Add("PART 2 - LICENSE AND NOTICE TEXTS")
$out.Add($rule)
for ($i = 0; $i -lt $textBodies.Count; $i++) {
    $out.Add("")
    $out.Add($rule)
    $out.Add("[$($i + 1)]")
    foreach ($u in $textUsers[$i]) { $out.Add("    $u") }
    $out.Add($rule)
    $out.Add("")
    $out.Add($textBodies[$i])
}
$generated = [string]::Join("`n", $out.ToArray()) + "`n"
$summary = "$($crates.Count) crates + the Rust standard library, $($textBodies.Count) texts"

# --- Check or write ---------------------------------------------------------------
$current = $null
if (Test-Path -LiteralPath $Path -PathType Leaf) {
    try {
        # Not ReadAllText: a BOM must show up as a difference, not be skipped.
        $current = $utf8NoBom.GetString([IO.File]::ReadAllBytes($Path)).Replace("`r`n", "`n")
    } catch {
        Fail $(if ($Check) { 44 } else { 45 }) "cannot read ${Path}: $_"
    }
}
$diff = if ($null -eq $current) { "the file does not exist" } else { Get-FirstDifference $current $generated }

if ($Check) {
    if ($diff) {
        Fail 44 "$Path is not what Cargo.lock and the crates' license files give - $diff. Run scripts\notices.ps1 and commit the file (with Cargo.lock)."
    }
    Write-Output "notices ok: $Path is current ($summary)"
    Quit 0
}
if (-not $diff) {
    Write-Output "notices: $Path is already up to date ($summary)"
} else {
    try {
        [IO.File]::WriteAllBytes($Path, $utf8NoBom.GetBytes($generated))
    } catch {
        Fail 45 "cannot write ${Path}: $_"
    }
    Write-Output "notices: wrote $Path ($summary)"
}
if ($withoutFiles.Count -gt 0) {
    Write-Output "notices: no license file in the package, standard texts used: $($withoutFiles -join ', ')"
}
Quit 0
