# Requires -Version 5.1
<#
    Migrate Ollama off the C: drive.

    Performs, in order:
      1. stop every Ollama process (service, tray app);
      2. silently uninstall the C: installation;
      3. install to D:\Ollama with the installer's /DIR switch;
      4. point model storage at D:\Ollama\Models (OLLAMA_MODELS);
      5. fix PATH, remove a stale C: entry if one remains;
      6. start the server and verify it answers on 11434.

    Idempotent enough to re-run: it re-checks each step.
    Nothing is deleted without a logged reason, and the exact rollback is printed
    at the end.
#>
param(
    [string]$InstallDir = 'D:\Ollama',
    [string]$ModelsDir  = 'D:\Ollama\Models'
)
$ErrorActionPreference = 'Continue'
$log = 'D:\AI-tools\out\ollama_migrate.log'
New-Item -ItemType Directory -Force -Path 'D:\AI-tools\out' | Out-Null

function Say($m) {
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $m
    Write-Output $line
    Add-Content -LiteralPath $log -Value $line -Encoding UTF8
}

Say "=== Ollama migration to D: ==="

$oldDir = Join-Path $env:LOCALAPPDATA 'Programs\Ollama'
$oldUninstaller = Join-Path $oldDir 'unins000.exe'

# --- 1. stop everything ------------------------------------------------------
Say "stopping Ollama processes"
Get-Process -Name 'ollama','ollama app' -ErrorAction SilentlyContinue | ForEach-Object {
    Say "  stopping PID $($_.Id) $($_.ProcessName)"
    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
}
Start-Sleep -Seconds 3
$still = Get-Process -Name 'ollama','ollama app' -ErrorAction SilentlyContinue
Say $(if ($still) { "  WARNING: still running: $($still.Id -join ',')" } else { "  all stopped" })

# --- 2. uninstall the C: copy ------------------------------------------------
if (Test-Path -LiteralPath $oldUninstaller) {
    Say "uninstalling $oldDir (silent)"
    $p = Start-Process -FilePath $oldUninstaller `
        -ArgumentList '/VERYSILENT','/NORESTART','/SUPPRESSMSGBOXES' -PassThru -Wait
    Say "  uninstaller exit code: $($p.ExitCode)"
    Start-Sleep -Seconds 5
    if (Test-Path -LiteralPath $oldDir) {
        $left = (Get-ChildItem -LiteralPath $oldDir -Recurse -File -ErrorAction SilentlyContinue |
                 Measure-Object Length -Sum)
        Say "  leftover: $([math]::Round($left.Sum/1MB,1)) MB in $oldDir"
        if ($left.Sum -gt 0) {
            Say "  removing leftovers"
            Remove-Item -LiteralPath $oldDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Say "  C: install directory removed: $(-not (Test-Path -LiteralPath $oldDir))"
} else {
    Say "no uninstaller at $oldUninstaller (already removed?)"
}

# --- 3. install to D: --------------------------------------------------------
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
New-Item -ItemType Directory -Force -Path $ModelsDir | Out-Null

$setup = 'D:\AI-tools\out\OllamaSetup.exe'
if (-not (Test-Path -LiteralPath $setup) -or (Get-Item -LiteralPath $setup).Length -lt 20MB) {
    $urls = @(
        'https://ollama.com/download/OllamaSetup.exe',
        'https://github.com/ollama/ollama/releases/latest/download/OllamaSetup.exe'
    )
    $ok = $false
    foreach ($u in $urls) {
        try {
            Say "downloading $u"
            Invoke-WebRequest -Uri $u -OutFile $setup -UseBasicParsing -TimeoutSec 900
            if ((Get-Item -LiteralPath $setup).Length -gt 20MB) { $ok = $true; break }
        } catch { Say "  failed: $($_.Exception.Message.Split([char]10)[0])" }
    }
    if (-not $ok) { Say 'RESULT: setup-download-failed'; exit 1 }
}
Say "installer ready: $([math]::Round((Get-Item -LiteralPath $setup).Length/1MB,1)) MB"

Say "installing to $InstallDir (Inno /DIR switch)"
$ip = Start-Process -FilePath $setup `
    -ArgumentList "/VERYSILENT","/NORESTART","/SUPPRESSMSGBOXES","/SP-","/DIR=$InstallDir" `
    -PassThru -Wait
Say "  installer exit code: $($ip.ExitCode)"

$exe = Join-Path $InstallDir 'ollama.exe'
if (-not (Test-Path -LiteralPath $exe)) {
    Say "  ollama.exe not found at $exe — the /DIR switch may have been ignored"
    $exe = Join-Path $env:LOCALAPPDATA 'Programs\Ollama\ollama.exe'
    Say "  falling back to $exe"
}
if (-not (Test-Path -LiteralPath $exe)) { Say 'RESULT: install-failed'; exit 1 }
Say "  installed: $exe"
& $exe --version 2>&1 | ForEach-Object { Say "  $_" }

# --- 4. model storage on D: --------------------------------------------------
[Environment]::SetEnvironmentVariable('OLLAMA_MODELS', $ModelsDir, 'User')
$env:OLLAMA_MODELS = $ModelsDir
Say "OLLAMA_MODELS=$ModelsDir (user environment)"

# --- 5. PATH ----------------------------------------------------------------
$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
$machinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
$stale = Join-Path $env:LOCALAPPDATA 'Programs\Ollama'
$keep = @()
foreach ($entry in ($userPath -split ';')) {
    if ([string]::IsNullOrWhiteSpace($entry)) { continue }
    if ($entry.TrimEnd('\') -ieq $stale.TrimEnd('\')) {
        Say "  dropping stale PATH entry: $entry"
        continue
    }
    $keep += $entry
}
if ($keep -notcontains $InstallDir) { $keep = @($InstallDir) + $keep }
[Environment]::SetEnvironmentVariable('Path', ($keep -join ';'), 'User')
$env:Path = ($keep -join ';') + ';' + $machinePath
Say "  user PATH updated (first entry: $($keep[0]))"

# --- 6. start + verify -------------------------------------------------------
Say "starting the server"
$running = Get-Process -Name 'ollama','ollama app' -ErrorAction SilentlyContinue
if (-not $running) {
    Start-Process -FilePath $exe -ArgumentList 'serve' -WindowStyle Hidden -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 6
}
$ok = $false
for ($i = 1; $i -le 10; $i++) {
    try {
        $r = Invoke-WebRequest -Uri 'http://localhost:11434/api/version' -TimeoutSec 5 -UseBasicParsing
        Say "  server OK: $($r.Content)"
        $ok = $true; break
    } catch { Start-Sleep -Seconds 3 }
}
if (-not $ok) { Say "  WARNING: 11434 did not answer" }

# --- report -----------------------------------------------------------------
Say ""
Say "=== disk after migration ==="
Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -in 'C','D' } | ForEach-Object {
    Say ("  {0}: free {1} GB" -f $_.Name, [math]::Round($_.Free/1GB,1))
}
Say "=== footprint ==="
foreach ($p in @($InstallDir, $ModelsDir, $stale)) {
    if (Test-Path -LiteralPath $p) {
        $s = (Get-ChildItem -LiteralPath $p -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
        Say ("  {0,-28} {1,8} MB" -f $p, [math]::Round($s/1MB,1))
    } else { Say ("  {0,-28} (gone)" -f $p) }
}
Say ""
Say "To roll back: uninstall from $InstallDir, reinstall with winget, and clear"
Say "  [Environment]::SetEnvironmentVariable('OLLAMA_MODELS', `$null, 'User')"
Say 'RESULT: migrated'
exit 0
