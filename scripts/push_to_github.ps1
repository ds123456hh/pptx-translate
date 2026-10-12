#Requires -Version 5.1
<#
    Push this repository to GitHub.

    Usage (a token is the only way to do this unattended):

        .\scripts\push_to_github.ps1 -Repo "你的用户名/pptx-translate" -Token "ghp_xxx"

    The token is used for this command only: it is passed to git through an
    ephemeral environment variable, is never written to disk, and is not stored
    in the git config. Create one at https://github.com/settings/tokens with the
    `repo` scope (classic) or Contents: Read and write (fine-grained).

    The repository is created for you if it does not exist yet.
#>
param(
    [Parameter(Mandatory = $true)][string]$Repo,      # owner/name
    [Parameter(Mandatory = $true)][string]$Token,
    [string]$Description = "Format-preserving PowerPoint translation: translate .pptx text without changing a single pixel of layout.",
    [string]$RepoDir = 'D:\AI-tools',
    [switch]$Private
)

$ErrorActionPreference = 'Continue'

function Say($m) { Write-Output ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) }

if ($Repo -notmatch '^[^/]+/[^/]+$') {
    Say "Repo must look like owner/name (got '$Repo')"
    exit 2
}

Set-Location $RepoDir
Say "repository : $RepoDir"
Say "target     : https://github.com/$Repo"

# --- local state -------------------------------------------------------------
$branch = (git branch --show-current).Trim()
Say "branch     : $branch"
$dirty = git status --porcelain
if ($dirty) {
    Say "WARNING: uncommitted changes present; they will NOT be pushed:"
    $dirty | Select-Object -First 10 | ForEach-Object { Say "   $_" }
}
$commits = (git log --oneline | Measure-Object).Count
Say "commits    : $commits"
if ($commits -eq 0) { Say 'nothing to push'; exit 1 }

# --- refuse to publish secrets or huge artefacts -----------------------------
Say "scanning the committed tree for credentials"
$scan = git grep -n -I -E 'sk-[A-Za-z0-9]{16,}|ghp_[A-Za-z0-9]{20,}|hf_[A-Za-z0-9]{20,}' HEAD -- . 2>$null
if ($LASTEXITCODE -eq 0 -and $scan) {
    Say "REFUSING TO PUSH - possible credential in the tree:"
    $scan | Select-Object -First 10 | ForEach-Object { Say "   $_" }
    exit 3
}
Say "  clean"

$largest = git ls-files | ForEach-Object {
    if (Test-Path -LiteralPath $_) { [pscustomobject]@{ MB = (Get-Item -LiteralPath $_).Length / 1MB; Path = $_ } }
} | Sort-Object MB -Descending | Select-Object -First 3
foreach ($f in $largest) { Say ("  largest tracked file: {0:N2} MB  {1}" -f $f.MB, $f.Path) }

# --- authenticate git for this run only --------------------------------------
$env:GIT_TERMINAL_PROMPT = '0'
$askPass = Join-Path $env:TEMP ("gh-askpass-{0}.cmd" -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
# git calls this for the password; the token comes from the environment, never argv
@"
@echo off
echo %GH_PUSH_TOKEN%
"@ | Set-Content -LiteralPath $askPass -Encoding ASCII
$env:GH_PUSH_TOKEN = $Token
$oldAskPass = $env:GIT_ASKPASS
$env:GIT_ASKPASS = $askPass
$remoteUrl = "https://github.com/$Repo.git"

try {
    Say "creating the repository if needed (via the GitHub API)"
    $headers = @{
        Authorization          = "Bearer $Token"
        Accept                 = 'application/vnd.github+json'
        'User-Agent'           = 'pptx-translate'
        'X-GitHub-Api-Version' = '2022-11-28'
    }
    $owner = $Repo.Split('/')[0]
    $name = $Repo.Split('/')[1]
    try {
        Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo" -Headers $headers -Method Get -TimeoutSec 30 | Out-Null
        Say "  repository already exists"
    } catch {
        $status = $_.Exception.Response.StatusCode.value__
        if ($status -ne 404) { Say "  lookup failed (HTTP $status): $($_.Exception.Message)"; exit 4 }
        Say "  creating $Repo"
        $body = @{
            name        = $name
            description = $Description
            private     = [bool]$Private
            auto_init   = $false
        } | ConvertTo-Json
        try {
            Invoke-RestMethod -Uri "https://api.github.com/user/repos" -Headers $headers -Method Post `
                -Body $body -ContentType 'application/json' -TimeoutSec 60 | Out-Null
            Say "  created"
        } catch {
            Say "  create failed: $($_.Exception.Message)"
            # a fine-grained token cannot create repos; the user must make it first
            Say "  -> create an empty repository named '$name' in the browser, then re-run"
            exit 5
        }
    }

    Say "configuring the remote"
    $existing = git remote 2>$null
    if ($existing -contains 'origin') { git remote set-url origin $remoteUrl } else { git remote add origin $remoteUrl }
    Say "  origin -> $remoteUrl"

    Say "pushing $branch"
    git -c credential.helper= -c core.askPass="$askPass" push -u origin $branch 2>&1 |
        ForEach-Object { Say "  git: $_" }
    $code = $LASTEXITCODE
    if ($code -ne 0) {
        Say "push failed (exit $code)"
        Say "  - 401/403 usually means the token lacks the 'repo' scope or is expired"
        Say "  - 404 can mean the token cannot see the repository"
        exit $code
    }
    Say ""
    Say "DONE -> https://github.com/$Repo"
} finally {
    $env:GIT_ASKPASS = $oldAskPass
    Remove-Item Env:\GH_PUSH_TOKEN -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $askPass -Force -ErrorAction SilentlyContinue
    Say "token discarded from the environment"
}
