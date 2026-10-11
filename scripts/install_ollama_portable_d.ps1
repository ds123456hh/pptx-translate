# Requires -Version 5.1
<#
    Install Ollama on D: from the official portable zip, bypassing the installer.

    The GUI/Inno installer refused to honour /DIR (exit code 1, nothing written
    to D:), and it also writes into %LOCALAPPDATA% on C:. The portable archive
    contains exactly the same payload (ollama.exe + lib\) and extracts wherever
    we tell it, so nothing can land on C:.

    Steps: download (resumable) -> extract to D:\Ollama -> PATH -> restart ->
    verify. Model storage is pointed at D:\Ollama\Models via OLLAMA_MODELS.
#>
param(
    [string]$InstallDir = 'D:\Ollama',
    [string]$ModelsDir  = 'D:\Ollama\Models',
    [string]$Version    = '0.40.3'
)
$ErrorActionPreference = 'Continue'
$log = 'D:\AI-tools\out\ollama_migrate.log'
New-Item -ItemType Directory -Force -Path 'D:\AI-tools\out' | Out-Null

function Say($m) {
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m
    Write-Output $line
    Add-Content -LiteralPath $log -Value $line -Encoding UTF8
}

Say "=== Ollama portable install onto D: (v$Version) ==="

# --- download ---------------------------------------------------------------
$zip = "D:\AI-tools\out\ollama-windows-amd64-$Version.zip"
$url = "https://github.com/ollama/ollama/releases/download/v$Version/ollama-windows-amd64.zip"

if ((Test-Path -LiteralPath $zip) -and (Get-Item -LiteralPath $zip).Length -gt 1GB) {
    Say "archive already complete: $([math]::Round((Get-Item -LiteralPath $zip).Length/1GB,2)) GB"
} else {
    Say "downloading $url"
    & curl.exe -L --fail --retry 12 --retry-delay 5 --retry-all-errors `
        --connect-timeout 30 --continue-at - --output $zip $url 2>&1 |
        Select-Object -Last 3 | ForEach-Object { Say "  curl: $_" }
}
if (-not (Test-Path -LiteralPath $zip)) { Say 'RESULT: download-failed'; exit 1 }
$zipSize = (Get-Item -LiteralPath $zip).Length
Say "archive: $([math]::Round($zipSize/1GB,2)) GB"
if ($zipSize -lt 1GB) { Say 'RESULT: archive-too-small (download incomplete)'; exit 1 }

# --- stop anything running --------------------------------------------------
Say "stopping Ollama processes"
Get-Process -Name 'ollama','ollama app' -ErrorAction SilentlyContinue | ForEach-Object {
    Say "  stopping PID $($_.Id) $($_.ProcessName)"
    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
}
Start-Sleep -Seconds 3

# --- extract ----------------------------------------------------------------
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
New-Item -ItemType Directory -Force -Path $ModelsDir | Out-Null
Say "extracting to $InstallDir"
try {
    Expand-Archive -LiteralPath $zip -DestinationPath $InstallDir -Force -ErrorAction Stop
    Say "  expand-archive done"
} catch {
    Say "  Expand-Archive failed ($($_.Exception.Message.Split([char]10)[0])); trying tar"
    & tar.exe -xf $zip -C $InstallDir 2>&1 | Select-Object -Last 3 | ForEach-Object { Say "  tar: $_" }
}

$exe = Join-Path $InstallDir 'ollama.exe'
if (-not (Test-Path -LiteralPath $exe)) {
    Say "ollama.exe missing at $exe; contents:"
    Get-ChildItem $InstallDir -ErrorAction SilentlyContinue | Select-Object -First 10 Name |
        ForEach-Object { Say "    $($_.Name)" }
    Say 'RESULT: extract-failed'
    exit 1
}
Say "installed binary: $exe"
& $exe --version 2>&1 | ForEach-Object { Say "  $_" }

# --- environment ------------------------------------------------------------
[Environment]::SetEnvironmentVariable('OLLAMA_MODELS', $ModelsDir, 'User')
$env:OLLAMA_MODELS = $ModelsDir
Say "OLLAMA_MODELS=$ModelsDir"

$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
$machinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
$stale = Join-Path $env:LOCALAPPDATA 'Programs\Ollama'
$keep = @()
foreach ($entry in ($userPath -split ';')) {
    if ([string]::IsNullOrWhiteSpace($entry)) { continue }
    if ($entry.TrimEnd('\') -ieq $stale.TrimEnd('\')) { Say "  dropping stale PATH entry: $entry"; continue }
    if ($entry.TrimEnd('\') -ieq $InstallDir.TrimEnd('\')) { continue }
    $keep += $entry
}
$keep = @($InstallDir) + $keep
[Environment]::SetEnvironmentVariable('Path', ($keep -join ';'), 'User')
$env:Path = ($keep -join ';') + ';' + $machinePath
Say "  PATH updated"

# --- start + verify ---------------------------------------------------------
Say "starting server"
Start-Process -FilePath $exe -ArgumentList 'serve' -WindowStyle Hidden -ErrorAction SilentlyContinue
$ok = $false
for ($i = 1; $i -le 15; $i++) {
    Start-Sleep -Seconds 3
    try {
        $r = Invoke-WebRequest -Uri 'http://localhost:11434/api/version' -TimeoutSec 5 -UseBasicParsing
        Say "  server OK: $($r.Content)"
        $ok = $true; break
    } catch { }
}
if (-not $ok) { Say "  WARNING: 11434 did not answer yet (helpful if the tray app is starting)" }

# --- report -----------------------------------------------------------------
Say ""
Say "=== disk ==="
Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -in 'C','D' } | ForEach-Object {
    Say ("  {0}: free {1} GB" -f $_.Name, [math]::Round($_.Free/1GB,1))
}
Say "=== footprint ==="
foreach ($p in @($InstallDir, $ModelsDir, $stale)) {
    if (Test-Path -LiteralPath $p) {
        $s = (Get-ChildItem -LiteralPath $p -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
        Say ("  {0,-26} {1,8} MB" -f $p, [math]::Round($s/1MB,1))
    } else { Say ("  {0,-26} (none on C:)" -f $p) }
}
Say 'RESULT: installed-on-d'
exit 0
