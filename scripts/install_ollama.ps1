# Requires -Version 5.1
<#
    Step 1 of the local deployment: install Ollama (silent, no GUI, no reboot).

    Tries winget first, then falls back to the official installer downloaded from
    ollama.com. Idempotent: exits immediately if ollama is already on PATH.
#>
$ErrorActionPreference = 'Continue'
$log = 'D:\AI-tools\out\ollama_install.log'
New-Item -ItemType Directory -Force -Path 'D:\AI-tools\out' | Out-Null

function Say($msg) {
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $msg
    Write-Output $line
    Add-Content -LiteralPath $log -Value $line -Encoding UTF8
}

Say "=== Ollama install ==="

$existing = Get-Command ollama -ErrorAction SilentlyContinue
if ($existing) {
    Say "ollama already installed at $($existing.Source)"
    & ollama --version 2>&1 | ForEach-Object { Say "  $_" }
    Say "RESULT: already-present"
    exit 0
}

# --- try winget -------------------------------------------------------------
$winget = Get-Command winget -ErrorAction SilentlyContinue
if ($winget) {
    Say "trying: winget install --id Ollama.Ollama -e --silent"
    & winget install --id Ollama.Ollama -e --silent `
        --accept-source-agreements --accept-package-agreements --disable-interactivity 2>&1 |
        ForEach-Object { Say "  winget: $_" }
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
    if (Get-Command ollama -ErrorAction SilentlyContinue) {
        Say "RESULT: winget-installed"
        exit 0
    }
    Say "winget did not produce ollama on PATH; falling back to the official installer"
} else {
    Say "winget not available"
}

# --- fallback: official installer, silent -----------------------------------
$installer = Join-Path $env:TEMP 'OllamaSetup.exe'
$urls = @(
    'https://ollama.com/download/OllamaSetup.exe',
    'https://github.com/ollama/ollama/releases/latest/download/OllamaSetup.exe'
)
$downloaded = $false
foreach ($url in $urls) {
    try {
        Say "downloading $url"
        Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing -TimeoutSec 900
        if ((Get-Item $installer).Length -gt 10MB) { $downloaded = $true; Say "downloaded $([math]::Round((Get-Item $installer).Length/1MB)) MB"; break }
        Say "file too small, trying next mirror"
    } catch {
        Say "download failed: $($_.Exception.Message.Split([char]10)[0])"
    }
}
if (-not $downloaded) { Say 'RESULT: download-failed'; exit 1 }

Say "running installer silently (Inno Setup switches)"
$proc = Start-Process -FilePath $installer -ArgumentList '/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES', '/SP-' -PassThru -Wait
Say "installer exit code: $($proc.ExitCode)"

$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
            [Environment]::GetEnvironmentVariable('Path', 'User')
$candidates = @(
    (Join-Path $env:LOCALAPPDATA 'Programs\Ollama\ollama.exe'),
    'C:\Program Files\Ollama\ollama.exe'
)
foreach ($c in $candidates) {
    if (Test-Path -LiteralPath $c) {
        Say "found binary at $c"
        & $c --version 2>&1 | ForEach-Object { Say "  $_" }
        Say "RESULT: installed-fallback"
        exit 0
    }
}
Say 'RESULT: binary-not-found'
exit 1
