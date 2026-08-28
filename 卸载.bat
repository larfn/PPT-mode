@chcp 65001 >nul
@echo off
setlocal EnableExtensions

set "INSTALL_DIR=%LOCALAPPDATA%\PPT-AI-Addin"
set "MANIFEST_PATH=%INSTALL_DIR%\manifest.xml"

echo ============================================
echo    模板助手 - 卸载
echo ============================================
echo.

echo [1/4] 停止本地服务 ...
taskkill /F /IM ppt-ai-addin.exe >nul 2>nul

echo [2/4] 删除登录自启动 ...
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "PPT-AI-Addin" /f >nul 2>nul

echo [3/4] 从 PowerPoint 注销加载项 ...
set "ADDIN_ID=2a7f4c3e-8b1d-4a90-9e6f-3b5a1c9d8e2f"
if exist "%MANIFEST_PATH%" (
  for /f "delims=" %%I in ('powershell -NoProfile -Command "[xml]$m=Get-Content -Raw -Encoding UTF8 -LiteralPath '%MANIFEST_PATH%'; Write-Output $m.OfficeApp.Id"') do set "ADDIN_ID=%%I"
)
reg delete "HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer" /v "%ADDIN_ID%" /f >nul 2>nul

echo [4/4] 清理安装目录 ...
if /i "%~1"=="/y" goto :delete_dir
choice /c YN /m "是否删除安装目录 %INSTALL_DIR%？"
if errorlevel 2 goto :done

:delete_dir
set "CHECK_DIR="
for %%D in ("%INSTALL_DIR%") do set "CHECK_DIR=%%~nxD"
if /i "%CHECK_DIR%"=="PPT-AI-Addin" (
  if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"
)

:done
echo.
echo 卸载完成。
pause
exit /b 0
