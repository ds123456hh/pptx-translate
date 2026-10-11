# Requires -Version 5.1
<#
    Step 2a: download the Hy-MT2 1.8B Q4_K_M GGUF into D:\AI-tools\Models.

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

Say "=== downloading $Repo/$File -> $target ==="

if ((Test-Path -LiteralPath $target) -and (Get-Item -LiteralPath $target).Length -gt 100MB) {
    Say "already present: $([math]::Round((Get-Item -LiteralPath $target).Length/1MB)) MB — skipping"
    exit 0
}

$urls = @(
    "https://hf-mirror.com/$Repo/resolve/main/$File",
    "https://huggingface.co/$Repo/resolve/main/$File",
    "https://hf-mirror.com/$Repo/resolve/main/$File?download=true"
)

foreach ($url in $urls) {
    Say "trying $url"
    try {
        # curl.exe gives resume + retry, and ships with Windows 10+
        & curl.exe -L --fail --retry 12 --retry-delay 5 --retry-all-errors `
            --connect-timeout 30 --continue-at - --output $target $url 2>&1 |
            ForEach-Object { Say "  curl: $_" }
        if ((Test-Path -LiteralPath $target) -and (Get-Item -LiteralPath $target).Length -gt 100MB) {
            Say "OK: $([math]::Round((Get-Item -LiteralPath $target).Length/1MB)) MB"
            Say 'RESULT: downloaded'
            exit 0
        }
        Say "too small after attempt; trying next mirror"
    } catch {
        Say "attempt failed: $($_.Exception.Message.Split([char]10)[0])"
    }
}

Say 'RESULT: download-failed'
exit 1
