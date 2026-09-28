# Publish a signed GitHub release for the commit at HEAD. The ONLY thing in
# this repo that writes releases: there is deliberately no release workflow
# (the signing key exists only on the maintainer's machine, and a workflow
# that writes releases could replace signed assets with unsigned CI builds).
#
# Usage (from anywhere):
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\release.ps1 -Version 0.5.0 -DryRun
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\release.ps1 -Version 0.5.0            # builds, stops at a verified DRAFT
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\release.ps1 -Version 0.5.0 -Publish   # publishes that draft as-is
#
# -Publish never builds: it requires the verified draft of a previous plain
# run for this exact commit, re-verifies it, and publishes it.
#
# Optional: -NotesFile <path> (markdown header placed above GitHub's
# generated notes when the draft is created; default: a standard "signed
# release" header), -NotesStartTag vX.Y.Z (previous tag for the generated
# notes; default: the newest v* tag before this commit).
#
# Order: everything that can fail runs BEFORE anything public is written.
# A draft release is private (only collaborators see it), and the git tag is
# created and pushed only in the publish step - so a failure before publish
# never leaves a public tag behind that a fix-up commit could not move past.
#   0 preflight  gh authenticated; clean tree; Cargo.toml version; the
#                commit's workflows are exactly ci.yml and it can neither run
#                on a tag nor write; an existing tag must point at HEAD (and
#                be on origin/main), else HEAD must be origin/main
#   1 CI gate    ci.yml passed on this exact commit (full SHA)
#   2 state      never modify a PUBLISHED release (re-verify it read-only
#                instead); at most one release per tag; a draft whose assets
#                still verify and that targets HEAD is reused, not rebuilt
#   3 build      scripts\build.ps1 -SkipCopy; stage exe, .cer and a zip
#                {exe, README.md, LICENSE, .cer}; verify them locally
#   4 draft      create the draft (not a prerelease) if absent
#   5 upload     the 3 assets (--clobber; only ever onto a draft), verify
#                them (exact set, state, server digests == local hashes,
#                download, zip contents, zip's exe == bare exe, signatures),
#                and only THEN point a stale draft at HEAD
#   6 publish    only with -Publish: tag + push, publish as a normal release,
#                mark latest, confirm /releases/latest serves it
# -DryRun runs 0-3 for real (all read-only checks, plus the local build and
# local verification) and prints what 4-6 would do.
#
# Exit codes (disjoint from build.ps1's 1-8):
#   20 preflight failed               25 creating/uploading the draft failed
#   21 gh missing/unauthenticated     26 verification of the release failed
#   22 CI gate failed                 27 publishing / post-publish check failed
#   23 release state forbids going on 28 tag creation / push failed
#   24 build.ps1 / local assets bad   29 the commit's workflows fail the guard
param(
    [Parameter(Mandatory = $true)][string]$Version,
    [switch]$Publish,
    [switch]$DryRun,
    [string]$NotesFile = "",
    [string]$NotesStartTag = ""
)

$ErrorActionPreference = "Continue"
# gh writes UTF-8; decode captured native output as UTF-8 (no console when
# run non-interactively - then there is nothing to set).
try { [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false) } catch { }
$repo = Split-Path -Parent $PSScriptRoot
$ownerRepo = "atefalshehri/WinThemeSwitcher"
$thumbprint = "40E0D1EB58DAC255EB37E9D64FF34448E3D33D12"
$exeName = "win-theme-switcher.exe"
$cerName = "WinThemeSwitcher-publisher.cer"
# Relative to the caller's location, not the repo (resolved before we move).
if ($NotesFile) { $NotesFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($NotesFile) }

Push-Location -LiteralPath $repo
function Quit([int]$code) {
    Pop-Location
    exit $code
}
function Fail([int]$code, [string]$msg) {
    [Console]::Error.WriteLine("release: FAILED (exit $code): $msg")
    Quit $code
}
function Step([string]$msg) { Write-Output "== $msg" }

# Run a native command. Returns its stdout lines as strings (an empty array
# when there are none); sets $script:rc to the exit code (seeded, so a launch
# failure can't read as 0) and $script:errText to its stderr.
# NOTE: PowerShell 5.1 strips embedded double quotes from native-command
# arguments, so no argument passed through here may contain a '"'. JSON is
# never parsed here: gh's --jq emits one tab-separated line per item.
function Run {
    $global:LASTEXITCODE = 1
    $rest = @()
    if ($args.Count -gt 1) { $rest = $args[1..($args.Count - 1)] }
    $all = & $args[0] @rest 2>&1
    $script:rc = $LASTEXITCODE
    $script:errText = (@($all | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } |
                ForEach-Object { "$_" }) -join "`n")
    return , @($all | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] } |
            ForEach-Object { "$_" })
}
function Out1([string[]]$lines) { return (($lines -join "`n").Trim()) }

function Test-SignedExe([string]$path) {
    $s = Get-AuthenticodeSignature -LiteralPath $path
    return ($s.Status -eq "Valid" -and $s.TimeStamperCertificate -and
        $s.SignerCertificate.Thumbprint -eq $thumbprint)
}
function Sha256([string]$path) { return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLower() }
function New-TempDir([string]$what) {
    $d = Join-Path ([IO.Path]::GetTempPath()) ("wts-$what-$tag-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
    New-Item -ItemType Directory -Force -Path $d | Out-Null
    return $d
}
# The tree must still be exactly the commit everything was checked against.
function Assert-Head([int]$code, [string]$when) {
    $h = Out1 (Run git rev-parse HEAD)
    if ($h -ne $sha) { Fail $code "HEAD moved $when ($sha -> $h)" }
    $d = Run git status --porcelain
    if ($rc -ne 0 -or (($d -join "").Trim())) { Fail $code "working tree changed $when`:`n$($d -join "`n")" }
}
# Checks the unpacked zip dir: exact entries, its exe identical to $exePath,
# both signed. Returns "" or a reason.
function Test-ZipContents([string]$zipPath, [string]$exePath) {
    $unz = New-TempDir "unzip"
    try {
        Expand-Archive -LiteralPath $zipPath -DestinationPath $unz -Force -ErrorAction Stop
    } catch {
        return "cannot expand ${zipPath}: $_"
    }
    $entries = @(Get-ChildItem -LiteralPath $unz -Recurse -File | ForEach-Object { $_.FullName.Substring($unz.Length + 1) } | Sort-Object)
    $want = @($exeName, "README.md", "LICENSE", $cerName | Sort-Object)
    if (($entries -join "|") -ne ($want -join "|")) { return "zip contains [$($entries -join ', ')], expected [$($want -join ', ')]" }
    $inner = Join-Path $unz $exeName
    if ((Sha256 $inner) -ne (Sha256 $exePath)) { return "the exe inside the zip is not the same file as the bare exe" }
    foreach ($p in @($exePath, $inner)) {
        if (-not (Test-SignedExe $p)) { return "$p is not Valid + timestamped + signed by $thumbprint" }
    }
    return ""
}
# Verifies the release's assets as GitHub serves them. $localHash (name ->
# sha256) is compared too when given. Returns "" or a reason.
function Test-RemoteAssets($localHash) {
    $lines = Run gh release view $tag --repo $ownerRepo --json assets --jq ".assets[] | [.name,.state,.digest] | @tsv"
    if ($rc -ne 0) { return "gh release view (assets) failed: $errText" }
    $digest = @{}
    foreach ($l in $lines) {
        $f = $l -split "`t"
        if ($f.Count -ne 3) { return "unexpected asset line [$l]" }
        if ($f[1] -ne "uploaded") { return "asset $($f[0]) is in state '$($f[1])'" }
        if ($f[2] -notmatch '^sha256:[0-9a-f]{64}$') { return "asset $($f[0]) has no sha256 digest ('$($f[2])')" }
        $digest[$f[0]] = $f[2].Substring(7)
    }
    $names = @($digest.Keys | Sort-Object)
    if (($names -join "|") -ne ($expectedAssets -join "|")) { return "asset set is [$($names -join ', ')], expected [$($expectedAssets -join ', ')]" }
    if ($localHash) {
        foreach ($n in $expectedAssets) {
            if ($digest[$n] -ne $localHash[$n]) { return "server digest of $n differs from the local build" }
        }
    }
    $dl = New-TempDir "download"
    $null = Run gh release download $tag --repo $ownerRepo --dir $dl
    if ($rc -ne 0) { return "gh release download failed: $errText" }
    foreach ($n in $expectedAssets) {
        $p = Join-Path $dl $n
        if (-not (Test-Path -LiteralPath $p)) { return "download is missing $n" }
        if ((Sha256 $p) -ne $digest[$n]) { return "downloaded $n does not match its server digest" }
    }
    $why = Test-ZipContents (Join-Path $dl $zipName) (Join-Path $dl $exeName)
    if ($why) { return "downloaded assets: $why" }
    return ""
}
# The release notes header: -NotesFile, or the standard one.
function Get-HeaderText {
    if ($NotesFile) { return [IO.File]::ReadAllText((Resolve-Path -LiteralPath $NotesFile).Path) }
    return "## WinThemeSwitcher $tag`n`n" +
    "The exe - standalone and inside the zip - is Authenticode-signed with the project's publisher certificate " +
    "(``CN=WinThemeSwitcher Self-Signed``, thumbprint ``$thumbprint``) and carries an RFC 3161 timestamp. " +
    "Install and verification steps: [README](https://github.com/$ownerRepo/blob/$tag/README.md#install).`n"
}
# The notes must be intact; with $header (a draft created by THIS run) they
# must also contain its first line. A reused draft's header came from the run
# that created it, so only its integrity is checked.
function Test-Body([string]$header) {
    $body = (Run gh release view $tag --repo $ownerRepo --json body --jq .body) -join "`n"
    if ($rc -ne 0) { return "gh release view (body) failed: $errText" }
    if (-not $body.Trim()) { return "the release notes are empty" }
    if ($body.Contains([string][char]0xFFFD) -or $body.Contains([string][char]0)) { return "the release notes contain replacement/NUL characters (encoding problem)" }
    if ($header) {
        $first = @($header -split "`n" | Where-Object { $_.Trim() })[0].Trim()
        if (-not $body.Contains($first)) { return "the release notes don't contain the header's first line ('$first')" }
    }
    return ""
}
# Helpers take the exit code of the phase they are called from.
function Get-ReleaseState([int]$code = 23) {
    # Returns $null (no release) or @{ draft; prerelease; target; url }.
    $l = Run gh release view $tag --repo $ownerRepo --json "isDraft,isPrerelease,targetCommitish,url" --jq "[.isDraft,.isPrerelease,.targetCommitish,.url] | @tsv"
    if ($rc -ne 0) {
        if ($errText.Trim() -eq "release not found") { return $null }
        Fail $code "gh release view failed: $errText"
    }
    $f = (Out1 $l) -split "`t"
    if ($f.Count -ne 4) { Fail $code "unexpected gh release view output [$(Out1 $l)]" }
    return @{ draft = ($f[0] -eq "true"); prerelease = ($f[1] -eq "true"); target = $f[2]; url = $f[3] }
}
function Get-ReleaseCount([int]$code = 23) {
    $l = Run gh api "repos/$ownerRepo/releases?per_page=100" --paginate --jq ".[].tag_name"
    if ($rc -ne 0) { Fail $code "listing releases failed: $errText" }
    return @($l | Where-Object { $_.Trim() -eq $tag }).Count
}
function Get-RemoteTagCommit([int]$code = 20) {
    $l = Run git ls-remote --tags origin "refs/tags/$tag" "refs/tags/$tag^{}"
    if ($rc -ne 0) { Fail $code "git ls-remote failed: $errText" }
    $plain = ""; $peeled = ""
    foreach ($x in $l) {
        $p = $x -split "\s+"
        if ($p.Count -lt 2) { continue }
        if ($p[1] -eq "refs/tags/$tag^{}") { $peeled = $p[0] } elseif ($p[1] -eq "refs/tags/$tag") { $plain = $p[0] }
    }
    if ($peeled) { return $peeled }
    return $plain
}

# --- 0 preflight --------------------------------------------------------------
Step "preflight"
if ($Version -notmatch '^\d+\.\d+\.\d+$') { Fail 20 "-Version must look like 1.2.3 (got '$Version')" }
$tag = "v$Version"
$zipName = "win-theme-switcher-$tag-windows-x64.zip"
$expectedAssets = @($zipName, $exeName, $cerName | Sort-Object)
if ($NotesFile -and -not (Test-Path -LiteralPath $NotesFile -PathType Leaf)) { Fail 20 "-NotesFile $NotesFile not found" }

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { Fail 21 "GitHub CLI (gh) not found" }
$null = Run gh auth status
if ($rc -ne 0) { Fail 21 "gh is not authenticated (run: gh auth login -h github.com -w)" }

$null = Run git fetch origin main --tags --force
if ($rc -ne 0) { Fail 20 "git fetch failed: $errText" }
$sha = Out1 (Run git rev-parse HEAD)
if ($sha -notmatch '^[0-9a-f]{40}$') { Fail 20 "cannot resolve HEAD" }
$dirty = Run git status --porcelain
if ($rc -ne 0 -or (($dirty -join "").Trim())) { Fail 20 "working tree is not clean:`n$($dirty -join "`n")" }

$cargoToml = [IO.File]::ReadAllText((Join-Path $repo "Cargo.toml"))
if ($cargoToml -notmatch '(?ms)^\[package\].*?^version\s*=\s*"([^"]+)"') { Fail 20 "cannot read the package version from Cargo.toml" }
if ($Matches[1] -ne $Version) { Fail 20 "Cargo.toml version is $($Matches[1]), not $Version" }

# A tag push (and publishing) runs the workflows of the tagged tree: none may
# be able to publish (the old release.yml replaced the signed assets). An
# allowlist, not a pattern hunt: the tree must hold exactly ci.yml, and ci.yml
# (comments stripped) must be branch-filtered, with no tag filter, no
# release/create/workflow_run trigger and no write permission. A new
# workflow means reviewing it and extending this list.
$wf = @(Run git -c core.quotePath=false ls-tree -r --name-only $sha -- .github/workflows)
if ($rc -ne 0) { Fail 29 "git ls-tree failed: $errText" }
if (($wf -join "|") -ne ".github/workflows/ci.yml") {
    Fail 29 "the workflows at $sha are [$($wf -join ', ')], expected exactly .github/workflows/ci.yml - review any new workflow, then extend the allowlist in release.ps1"
}
$ci = Run git show "${sha}:.github/workflows/ci.yml"
if ($rc -ne 0 -or $ci.Count -eq 0) { Fail 29 "cannot read ci.yml at ${sha}: $errText" }
$ciText = ($ci | ForEach-Object { $_ -replace '#.*$', '' }) -join "`n"
if ($ciText -match '(?m)\btags(-ignore)?\s*:' -or $ciText -match '(?m)^\s*(release|create|workflow_run)\s*:' -or
    $ciText -match '\bwrite(-all)?\b' -or $ciText -notmatch '(?m)^\s*branches\s*:') {
    Fail 29 "ci.yml at $sha could run on a tag or write (tag filter, release/create/workflow_run trigger, write permission, or no branches filter)"
}

# Tag rules. Absent: HEAD must be origin/main. Present (local or remote):
# it must point at HEAD, and HEAD must be on origin/main.
$localTag = Out1 (Run git rev-parse -q --verify "refs/tags/$tag^{commit}")
$remoteTag = Get-RemoteTagCommit
$originMain = Out1 (Run git rev-parse origin/main)
if ($localTag -or $remoteTag) {
    if ($localTag -and $localTag -ne $sha) { Fail 20 "local tag $tag points at $localTag, not HEAD $sha" }
    if ($remoteTag -and $remoteTag -ne $sha) { Fail 20 "tag $tag on origin points at $remoteTag, not HEAD $sha" }
    $null = Run git merge-base --is-ancestor $sha origin/main
    if ($rc -ne 0) { Fail 20 "HEAD $sha (the tagged commit) is not on origin/main" }
} elseif ($sha -ne $originMain) {
    Fail 20 "HEAD ($sha) is not origin/main ($originMain) - push first, or check out main"
}

if (-not $NotesStartTag) {
    $NotesStartTag = Out1 (Run git describe --tags --abbrev=0 --match "v[0-9]*" "$sha^")
    if ($rc -ne 0 -or -not $NotesStartTag) { Fail 20 "cannot find the previous tag (pass -NotesStartTag)" }
}
Write-Output "   $tag at $sha (previous: $NotesStartTag)$(if ($DryRun) { ' [DRY RUN]' })"

# --- 1 CI gate ------------------------------------------------------------------
Step "CI gate (ci.yml on $($sha.Substring(0, 7)))"
# A short or unknown SHA makes `gh run list --commit` return nothing with
# exit 0: no run is a failure, never a pass. Right after a push the run can
# take a few seconds to appear.
$runs = @()
for ($i = 0; $i -lt 7; $i++) {
    if ($i -gt 0) { Start-Sleep -Seconds 10 }
    $runs = Run gh run list --repo $ownerRepo --commit $sha --workflow ci.yml --event push --limit 20 --json "databaseId,status,conclusion,headSha,event,createdAt" --jq ".[] | [.createdAt,.databaseId,.status,.conclusion,.headSha,.event] | @tsv"
    if ($rc -ne 0) { Fail 22 "gh run list failed: $errText" }
    if ($runs.Count -gt 0) { break }
}
if ($runs.Count -eq 0) { Fail 22 "no ci.yml push run found for $sha (was it pushed to main?)" }
$newest = @($runs | Sort-Object -Descending)[0] -split "`t"
if ($newest.Count -ne 6 -or $newest[4] -ne $sha -or $newest[5] -ne "push") { Fail 22 "unexpected run record [$($newest -join ' ')]" }
$runId = $newest[1]
$status = $newest[2]; $conclusion = $newest[3]
if ($status -ne "completed") {
    Write-Output "   waiting for CI run $runId..."
    $null = Run gh run watch $runId --repo $ownerRepo --exit-status --interval 20
    $v = Run gh run view $runId --repo $ownerRepo --json "status,conclusion,headSha" --jq "[.status,.conclusion,.headSha] | @tsv"
    if ($rc -ne 0) { Fail 22 "gh run view failed: $errText" }
    $f = (Out1 $v) -split "`t"
    if ($f.Count -ne 3 -or $f[2] -ne $sha) { Fail 22 "unexpected run view [$(Out1 $v)]" }
    $status = $f[0]; $conclusion = $f[1]
}
if ($status -ne "completed" -or $conclusion -ne "success") { Fail 22 "CI run $runId for $sha is $status/$conclusion" }
Write-Output "   CI run $runId passed"

# --- 2 release state ------------------------------------------------------------
Step "release state"
$count = Get-ReleaseCount
if ($count -gt 1) { Fail 23 "$count releases use $tag - delete the extra drafts on GitHub first" }
$state = Get-ReleaseState
$reuse = $false
if ($state -and -not $state.draft) {
    # Published: never modified. Re-verify what is being served, read-only.
    $why = Test-RemoteAssets $null
    if ($why) { Fail 23 "$tag is PUBLISHED ($($state.url)) and its assets FAIL verification: $why - ship a new patch version" }
    $latest = Out1 (Run gh api "repos/$ownerRepo/releases/latest" --jq .tag_name)
    if ($state.prerelease) { Fail 23 "$tag is published and verified, but marked prerelease - fix by hand: gh release edit $tag --prerelease=false" }
    if ($latest -ne $tag) {
        # Only a problem if nothing newer is the latest release.
        $newer = $latest -match '^v(\d+\.\d+\.\d+)$' -and ([version]$Matches[1] -gt [version]$Version)
        if (-not $newer) { Fail 23 "$tag is published and verified, but /releases/latest is '$latest' - fix by hand: gh release edit $tag --latest" }
        Write-Output "== $tag is already published and verified (latest is the newer $latest): $($state.url)"
        Quit 0
    }
    Write-Output "== $tag is already published and verified: $($state.url)"
    Quit 0
}
if ($state) {
    if ($state.target -ne $sha) {
        Write-Output "   existing draft targets '$($state.target)', not $sha - it will be rebuilt"
    } else {
        $why = Test-RemoteAssets $null
        if ($why) {
            Write-Output "   existing draft does not verify ($why) - rebuilding"
        } else {
            $reuse = $true
            Write-Output "   existing draft targets HEAD and verifies - reusing it as-is"
        }
    }
} else {
    Write-Output "   no release yet"
}
# -Publish publishes the draft a plain run built and verified - never a
# build it makes itself, so what goes public is what was reviewed.
if ($Publish -and -not $reuse -and -not $DryRun) {
    Fail 23 "-Publish needs a verified draft of $tag for $sha, and there is none - run once without -Publish first"
}

if (-not $reuse) {
    # --- 3 build + stage + local verify -----------------------------------------
    Step "build + sign"
    # The exe is staged from the default target dir; a redirected target dir
    # would leave a stale (validly signed) exe there.
    foreach ($v in @("CARGO_TARGET_DIR", "CARGO_BUILD_TARGET_DIR", "CARGO_BUILD_TARGET")) {
        if ([Environment]::GetEnvironmentVariable($v)) { Fail 24 "$v is set - unset it (the release is staged from target\release)" }
    }
    Assert-Head 24 "before the build"
    $buildStart = Get-Date
    $global:LASTEXITCODE = 1
    & (Join-Path $PSScriptRoot "build.ps1") -SkipCopy
    if ($LASTEXITCODE -ne 0) { Fail 24 "build.ps1 -SkipCopy failed (exit $LASTEXITCODE)" }
    Assert-Head 24 "during the build"
    $builtExe = Join-Path $repo "target\release\$exeName"
    if (-not (Test-SignedExe $builtExe)) { Fail 24 "the built exe is not validly signed + timestamped by $thumbprint" }
    # build.ps1 signs the exe it built, so it must have been written just now.
    if ((Get-Item -LiteralPath $builtExe).LastWriteTime -lt $buildStart.AddSeconds(-2)) { Fail 24 "$builtExe was not written by this build" }

    $stage = New-TempDir "stage"
    $zipDir = Join-Path $stage "zip"
    try {
        New-Item -ItemType Directory -Force -Path $zipDir -ErrorAction Stop | Out-Null
        Copy-Item -LiteralPath $builtExe -Destination (Join-Path $stage $exeName) -ErrorAction Stop
        Copy-Item -LiteralPath (Join-Path $repo $cerName) -Destination (Join-Path $stage $cerName) -ErrorAction Stop
        Copy-Item -LiteralPath $builtExe -Destination (Join-Path $zipDir $exeName) -ErrorAction Stop
        foreach ($f in @("README.md", "LICENSE", $cerName)) {
            Copy-Item -LiteralPath (Join-Path $repo $f) -Destination (Join-Path $zipDir $f) -ErrorAction Stop
        }
        Compress-Archive -Path (Join-Path $zipDir "*") -DestinationPath (Join-Path $stage $zipName) -ErrorAction Stop
    } catch {
        Fail 24 "staging failed: $_"
    }
    $assets = @($expectedAssets | ForEach-Object { Join-Path $stage $_ })
    $localHash = @{}
    foreach ($a in $assets) { $localHash[(Split-Path $a -Leaf)] = Sha256 $a }
    $why = Test-ZipContents (Join-Path $stage $zipName) (Join-Path $stage $exeName)
    if ($why) { Fail 24 "staged assets: $why" }
    if ($localHash[$exeName] -ne (Sha256 $builtExe)) { Fail 24 "the staged exe differs from the build" }
    Assert-Head 24 "while staging"
    Write-Output "   staged and verified locally in $stage"
}

if ($DryRun) {
    if ($reuse) {
        Write-Output "== [dry run] would reuse the verified draft (no build, no upload)"
    } elseif ($Publish) {
        Write-Output "== [dry run] -Publish would REFUSE: there is no verified draft for $sha yet (run without -Publish first)"
        Quit 0
    } else {
        Write-Output "== [dry run] would $(if ($state) { 'replace the existing draft''s assets, verify them, then point it at HEAD' } else { "create a draft: gh release create $tag --draft --target $sha --generate-notes --notes-start-tag $NotesStartTag" })"
        Write-Output "== [dry run] would upload $($expectedAssets -join ', ') (--clobber, draft only) and verify them remotely"
    }
    Write-Output "== [dry run] would $(if ($Publish) { "tag $sha as $tag, push the tag, and PUBLISH as latest (not prerelease)" } else { 'stop at the verified draft' })"
    Quit 0
}

$createdHeader = ""
if (-not $reuse) {
    # --- 4 draft ----------------------------------------------------------------
    Step "draft release"
    if (-not $state) {
        $createdHeader = Get-HeaderText
        $header = Join-Path $stage "notes-header.md"
        [IO.File]::WriteAllText($header, $createdHeader, (New-Object Text.UTF8Encoding($false)))
        # No --verify-tag: the tag is created only at publish. A draft's tag
        # is not created until the draft is published.
        $out = Run gh release create $tag --repo $ownerRepo --draft --target $sha --title $tag --notes-file $header --generate-notes --notes-start-tag $NotesStartTag
        if ($rc -ne 0) { Fail 25 "gh release create failed: $errText" }
        if ((Get-ReleaseCount 25) -ne 1) { Fail 25 "after creating the draft, $tag does not have exactly one release" }
        Write-Output "   draft created"
    }

    # --- 5 upload + verify, then retarget ----------------------------------------
    Step "upload"
    # --clobber deletes each same-named asset before re-uploading: acceptable
    # only on a draft, so re-check right before. A stale draft keeps its old
    # target until the new assets are uploaded AND verified - otherwise an
    # interrupted run would leave "targets HEAD" on the old binaries, which a
    # later run would take for a verified draft of HEAD.
    $s = Get-ReleaseState 25
    if (-not $s -or -not $s.draft) { Fail 25 "$tag is no longer a draft - refusing to upload" }
    $null = Run gh release upload $tag $assets[0] $assets[1] $assets[2] --repo $ownerRepo --clobber
    if ($rc -ne 0) { Fail 25 "gh release upload failed: $errText" }

    Step "verify"
    $why = Test-RemoteAssets $localHash
    if ($why) { Fail 26 $why }
    Write-Output "   assets verified (state, digests, download, zip contents, signatures)"
    if ($s.target -ne $sha) {
        $null = Run gh release edit $tag --repo $ownerRepo --target $sha
        if ($rc -ne 0) { Fail 25 "pointing the draft at $sha failed: $errText" }
        $s = Get-ReleaseState 25
        if (-not $s -or -not $s.draft -or $s.target -ne $sha) { Fail 25 "the draft does not target $sha after retargeting" }
        Write-Output "   draft now targets $sha"
    }
}
$why = Test-Body $createdHeader
if ($why) { Fail 26 $why }

# --- 6 publish ------------------------------------------------------------------
if (-not $Publish) {
    Write-Output "== verified DRAFT ready: $((Get-ReleaseState 26).url)"
    Write-Output "   re-run with -Publish to publish exactly this draft (it is re-verified, not rebuilt)"
    Quit 0
}
Step "publish"
if (-not $localTag) {
    $null = Run git tag -a $tag -m "WinThemeSwitcher $tag" $sha
    if ($rc -ne 0) { Fail 28 "git tag failed: $errText" }
}
if (-not $remoteTag) {
    $null = Run git push origin "refs/tags/$tag"
    if ($rc -ne 0) { Fail 28 "git push of $tag failed: $errText" }
}
if ((Get-RemoteTagCommit 28) -ne $sha) { Fail 28 "origin does not show $tag at $sha after the push" }
Write-Output "   tag $tag is on origin at $sha"

$s = Get-ReleaseState 27
if (-not $s -or -not $s.draft -or $s.target -ne $sha) { Fail 27 "$tag is no longer a draft targeting $sha - not publishing" }
$null = Run gh release edit $tag --repo $ownerRepo --draft=false --prerelease=false --latest
if ($rc -ne 0) { Fail 27 "gh release edit failed: $errText" }
$s = Get-ReleaseState 27
if (-not $s -or $s.draft -or $s.prerelease) { Fail 27 "after publishing: draft=$($s.draft) prerelease=$($s.prerelease)" }
if ((Get-RemoteTagCommit 27) -ne $sha) { Fail 27 "after publishing, $tag no longer points at $sha" }
$latest = ""
for ($i = 0; $i -lt 6 -and $latest -ne $tag; $i++) {
    if ($i -gt 0) { Start-Sleep -Seconds 5 }
    $latest = Out1 (Run gh api "repos/$ownerRepo/releases/latest" --jq .tag_name)
}
if ($latest -ne $tag) { Fail 27 "/releases/latest serves '$latest', not $tag" }
Write-Output "== published: $($s.url) (latest)"
Quit 0
