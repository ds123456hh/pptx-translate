#Requires -Version 5.1
<#
    手动启动 Ollama（本地翻译模型服务）。

    用法：
        scripts\start_ollama.cmd                                # 双击或敲这个（推荐）
        powershell -File scripts\start_ollama.ps1               # 后台启动，脚本立即返回
        powershell -File scripts\start_ollama.ps1 -Foreground   # 前台运行，Ctrl+C 停止
        powershell -File scripts\start_ollama.ps1 -Status       # 只看状态，什么都不改

    为什么要有这个脚本，而不是直接敲 ollama serve：
      ollama serve 没有可用参数，模型目录完全由环境变量 OLLAMA_MODELS 决定。
      如果启动时这个变量没生效，Ollama 会去 C:\Users\<你>\.ollama\models 找，
      那里是空的，于是所有请求都报 "model not found" —— 尽管你的权重好好地
      躺在 D:\Ollama\Models。脚本负责把变量设对再启动。

    实现要点：用 WMI 启动进程，让它完全脱离当前会话。不要用
    Start-Process -RedirectStandardOutput —— PowerShell 5.1 会一直等那个输出
    管道关闭，而服务永不退出，结果就是启动命令永远卡住不返回。
#>
param(
    [string]$OllamaExe = 'D:\Ollama\ollama.exe',
    [string]$ModelsDir = 'D:\Ollama\Models',
    [string]$LogFile   = 'D:\Ollama\ollama-server.log',
    [switch]$Foreground,
    [switch]$Status
)

$ErrorActionPreference = 'Continue'

function Test-Server {
    try { return (Invoke-WebRequest 'http://localhost:11434/api/version' -TimeoutSec 4 -UseBasicParsing).Content }
    catch { return $null }
}

function Show-State {
    $version = Test-Server
    if ($version) {
        Write-Host "服务状态 : 运行中  $version" -ForegroundColor Green
        Get-Process ollama -ErrorAction SilentlyContinue |
            Select-Object Id, @{n = 'MemMB'; e = { [math]::Round($_.WorkingSet64 / 1MB) } }, StartTime |
            Format-Table -AutoSize | Out-String | Write-Host
    } else {
        Write-Host '服务状态 : 未运行' -ForegroundColor Yellow
    }
    Write-Host "程序     : $OllamaExe  $(if (Test-Path $OllamaExe) {'(存在)'} else {'(缺失!)'})"
    Write-Host "模型目录 : $ModelsDir"
    $blobs = Join-Path $ModelsDir 'blobs'
    if (Test-Path $blobs) {
        $sum = (Get-ChildItem $blobs -Recurse -File -ErrorAction SilentlyContinue |
                Measure-Object Length -Sum).Sum
        Write-Host "           $([math]::Round($sum/1MB,1)) MB"
    }
    $manifests = Join-Path $ModelsDir 'manifests-v2'
    if (Test-Path $manifests) {
        $names = Get-ChildItem $manifests -Recurse -File -ErrorAction SilentlyContinue |
                 ForEach-Object { "$($_.Directory.Name):$($_.Name)" }
        Write-Host "已注册模型: $($names -join ', ')"
    }
    Write-Host "环境变量 : OLLAMA_MODELS(User) = $([Environment]::GetEnvironmentVariable('OLLAMA_MODELS','User'))"
}

if ($Status) { Show-State; exit 0 }

if (Test-Server) {
    Write-Host '服务已经在运行，不需要重复启动。' -ForegroundColor Green
    Show-State
    exit 0
}

if (-not (Test-Path $OllamaExe)) {
    Write-Host "找不到 $OllamaExe" -ForegroundColor Red
    Write-Host '重装：powershell -File D:\AI-tools\scripts\install_ollama_portable_d.ps1'
    exit 1
}

# 关键一步：让服务进程继承正确的模型目录
$env:OLLAMA_MODELS = $ModelsDir

if ($Foreground) {
    Write-Host '前台启动（Ctrl+C 停止）'
    Write-Host "  OLLAMA_MODELS = $env:OLLAMA_MODELS"
    & $OllamaExe serve
    exit $LASTEXITCODE
}

Write-Host '正在启动 Ollama ...'
Write-Host "  OLLAMA_MODELS = $env:OLLAMA_MODELS"

# WMI 启动 = 真正脱离本会话，父进程不会被拖住
$cmdline = "`"$OllamaExe`" serve"
try {
    $startup = ([wmiclass]'Win32_ProcessStartup').CreateInstance()
    $startup.ShowWindow = 0            # 隐藏窗口
    $proc = ([wmiclass]'Win32_Process').Create($cmdline, $null, $startup)
    if ($proc.ReturnValue -ne 0) { throw "WMI Create 返回 $($proc.ReturnValue)" }
    Write-Host "  已启动 (PID $($proc.ProcessId))"
} catch {
    Write-Host "  WMI 启动失败：$($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host '  改用普通方式启动'
    Start-Process -FilePath $OllamaExe -ArgumentList 'serve' -WindowStyle Hidden
}

# 等端口就绪，这样你敲完就知道能不能用了
$ready = $null
foreach ($i in 1..25) {
    Start-Sleep -Milliseconds 800
    $ready = Test-Server
    if ($ready) { break }
}

if (-not $ready) {
    Write-Host '20 秒内没有就绪。看日志：' -ForegroundColor Yellow
    foreach ($f in @("$LogFile.err", $LogFile)) {
        if (Test-Path $f) {
            Write-Host "--- $f"
            Get-Content $f -Tail 8 -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "    $_" }
        }
    }
    exit 1
}

Write-Host ''
Write-Host "服务已就绪: $ready" -ForegroundColor Green
& $OllamaExe list 2>&1 | ForEach-Object { Write-Host "  $_" }
Write-Host ''
Write-Host '现在可以翻译了：'
Write-Host '  cd D:\AI-tools'
Write-Host '  python -m pptx_translate "你的课件.pptx" --provider ollama --ollama-model hy-mt-local'
exit 0
