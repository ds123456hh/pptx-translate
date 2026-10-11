# Requires -Version 5.1
<#
    Step 2a: download the Hy-MT2 1.8B Q4_K_M GGUF into D:\AI-tools\Models.

    Completion is decided by comparing the local byte count against the remote
    Content-Length, not by a "big enough" threshold: a half-downloaded file is
    already hundreds of MB, so a size heuristic would call it done and silently
    register a corrupt model.

    Order of attempts (mirror fallback, as requested):
      1. hf-mirror.com   (fast in China)
      2. huggingface.co  (direct)
    Resumable: a partial download continues instead of restarting.
#>
param(
    [string]$Repo = 'tencent/Hy-MT2-1.8B-GGUF',
    [string]$File = 'Hy-MT2-1.8B-Q4_K_M.gguf',
    [string]$Dest = 'D:\AI-tools\Models'
)
$ErrorActionPreference = 'Continue'
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$target = Join-Path $Dest $File
$log = 'D:\AI-tools\out\ollama_download.log'

function Say($m) {
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m
    Write-Output $line
    Add-Content -LiteralPath $log -Value $line -Encoding UTF8
}

function Get-RemoteSize([string]$Url) {
    $raw = & curl.exe -sIL --max-time 60 $Url 2>&1
    $sizes = @()
    foreach ($line in $raw) {
        if ($line -match '^\s*[Cc]ontent-[Ll]ength:\s*(\d+)') { $sizes += [int64]$Matches[1] }
    }
    if ($sizes.Count -eq 0) { return 0 }
    return ($sizes | Measure-Object -Maximum).Maximum
}

Say "=== $Repo/$File -> $target ==="

$urls = @(
    "https://hf-mirror.com/$Repo/resolve/main/$File",
    "https://huggingface.co/$Repo/resolve/main/$File"
)

foreach ($url in $urls) {
    $remote = Get-RemoteSize $url
    if ($remote -le 0) { Say "  no Content-Length from $url; skipping"; continue }
    Say "remote size: $remote bytes ($([math]::Round($remote/1MB,1)) MB) via $url"

    if (Test-Path -LiteralPath $target) {
        $local = (Get-Item -LiteralPath $target).Length
        if ($local -eq $remote) {
            Say "  already complete ($([math]::Round($local/1MB,1)) MB) — nothing to do"
            Say 'RESULT: downloaded'
            exit 0
        }
        if ($local -gt $remote) {
            Say "  local is larger than remote ($local > $remote); restarting from scratch"
            Remove-Item -LiteralPath $target -Force
        } else {
            Say "  resuming: have $([math]::Round($local/1MB,1)) MB of $([math]::Round($remote/1MB,1)) MB"
        }
    }

    & curl.exe -L --fail --retry 12 --retry-delay 5 --retry-all-errors `
        --connect-timeout 30 --continue-at - --output $target $url 2>&1 |
        Select-Object -Last 2 | ForEach-Object { Say "  curl: $_" }

    if (Test-Path -LiteralPath $target) {
        $local = (Get-Item -LiteralPath $target).Length
        if ($local -eq $remote) {
            Say "OK: $local bytes ($([math]::Round($local/1MB,1)) MB) — verified against remote"
            Say 'RESULT: downloaded'
            exit 0
        }
        Say "  incomplete: $local / $remote bytes — trying next mirror"
    }
}

Say 'RESULT: download-failed'
exit 1
