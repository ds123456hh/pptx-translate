@echo off
rem Double-click this file to start the Ollama local translation server.
rem The window closes by itself; the server keeps running in the background.
rem
rem Why this wrapper exists: ollama serve reads its model directory from the
rem OLLAMA_MODELS environment variable, and the PowerShell script sets it before
rem launching. Running "ollama serve" by hand from a shell that lacks that
rem variable makes Ollama look in C:\Users\<you>\.ollama\models, find nothing,
rem and report "model not found" even though the weights are on D:.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0start_ollama.ps1"
if errorlevel 1 (
    echo.
    echo Startup failed. Copy the output above and send it to the maintainer.
    pause
)
