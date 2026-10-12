#Requires -Version 5.1
<#
    Push this repository to GitHub using an ephemeral token.

    The token never touches disk, the git config, or the command line of any
    process: git reads it through a temporary askpass helper whose content is a
    single `echo %GH_PUSH_TOKEN%`, and the variable is cleared in `finally`.
#>
param(
    [Parameter(Mandatory = $true)][string]$Repo,     # owner/name
    [Parameter(Mandatory = $true)][string]$Token,
    [string]$RepoDir = 'D:\AI-tools'
)
$ErrorActionPreference = 'Continue'
function Say($m) { Write-Output ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) }

Set-Location $RepoDir
$branch = (git branch --show-current).Trim()

Say "=== pushing $RepoDir ($branch) -> github.com/$Repo ==="

# refuse to publish a credential
$scan = git grep -n -I -E 'sk-[A-Za-z0-9]{16,}|ghp_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,}|hf_[A-Za-z0-9]{20,}' HEAD -- . 2>$null
if ($LASTEXITCODE -eq 0 -and $scan) {
    Say 'REFUSING: a credential pattern is in the committed tree:'
    $scan | Select-Object -First 5 | ForEach-Object { Say "   $_" }
    exit 3
}
Say 'pre-flight: no credential pattern in the committed tree'

$askPass = Join-Path $env:TEMP ("gh-askpass-{0}.cmd" -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
"@echo off`r`necho %GH_PUSH_TOKEN%" | Set-Content -LiteralPath $askPass -Encoding ASCII
$env:GH_PUSH_TOKEN = $Token
$env:GIT_ASKPASS = $askPass
$env:GIT_TERMINAL_PROMPT = '0'
$oldAskPass = $env:GIT_ASKPASS

try {
    $url = "https://github.com/$Repo.git"
    $remotes = git remote 2>$null
    if ($remotes -contains 'origin') { git remote set-url origin $url } else { git remote add origin $url }
    Say "remote origin -> $url"

    Say "git push -u origin $branch"
    git -c credential.helper= -c core.askPass="$askPass" push -u origin $branch 2>&1 |
        ForEach-Object { Say "  git: $_" }
    $code = $LASTEXITCODE

    if ($code -ne 0) {
        Say ""
        Say "push failed (exit $code). Common causes:"
        Say "  403 / 'Write access to repository not granted'"
        Say "      -> the fine-grained token does not cover this new repository."
        Say "         Open the token settings and add 'ds123456hh/pptx-translate'"
        Say "         under 'Only select repositories', or switch it to 'All repositories'."
        Say "  403 / 'Contents: write' missing -> the token needs Contents: Read and write."
        exit $code
    }

    Say ""
    Say "DONE -> https://github.com/$Repo"
    git remote -v | ForEach-Object { Say "  $_" }
    Say ""
    Say "remote refs:"
    git ls-remote --heads origin 2>&1 | ForEach-Object { Say "  $_" }
} finally {
    $env:GIT_ASKPASS = $oldAskPass
    Remove-Item Env:\GH_PUSH_TOKEN -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $askPass -Force -ErrorAction SilentlyContinue
    Say "token cleared from the environment"
}
