@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist "安装中心.exe" (
  start "" "安装中心.exe"
) else (
  start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "安装中心.ps1"
)
exit /b 0
