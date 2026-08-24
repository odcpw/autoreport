@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%sync-autobericht.ps1"

if not exist "%PS1%" (
  echo Missing "%PS1%".
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%
