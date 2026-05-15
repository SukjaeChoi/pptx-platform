@echo off
cd /d "%~dp0"
set PPTX_SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create([System.IO.File]::ReadAllText('%~dp0_env_check.ps1', [System.Text.Encoding]::UTF8)))"
pause
