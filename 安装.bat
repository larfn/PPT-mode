@chcp 65001 >nul
@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

echo ============================================
echo    模板助手 - 安装（Build → Package → Install → Verify）
echo ============================================
echo.
echo   用法：安装.bat              快速安装（版本号+构建+安装+自检，默认跳过测试）
echo        安装.bat --with-tests   额外运行 npm test（部分电脑测试较慢）
echo        安装.bat --skip-build  跳过构建，直接安装现有构建产物（无 Node 也可用）
echo        安装.bat --no-pause     全程不暂停（静默安装）
echo        安装.bat --skip-e2e    跳过黄金路径 E2E 自检（默认每次安装都会跑，约 10 秒）
echo.

rem ---------- 解析参数 ----------
set "SKIP_BUILD="
set "NO_PAUSE="
set "WITH_TESTS="
set "SKIP_E2E="
for %%A in (%*) do (
  if /i "%%~A"=="--skip-build" set "SKIP_BUILD=1"
  if /i "%%~A"=="--with-tests" set "WITH_TESTS=1"
  if /i "%%~A"=="--no-pause" set "NO_PAUSE=1"
  if /i "%%~A"=="--skip-e2e" set "SKIP_E2E=1"
)

rem ---------- 0. 构建发布产物 ----------
where node >nul 2>nul
if not errorlevel 1 goto :node_present
if exist "release.json" goto :node_missing_use_artifacts
echo [错误] 未检测到 Node.js，且未找到 release.json 发布清单。
echo       请在开发机（安装 Node.js）运行一次完整 安装.bat 生成构建产物，
echo       或将完整产物（dist-exe、addin\dist、manifest.xml、release.json）拷到本机后用 安装.bat --skip-build。
if not defined NO_PAUSE pause
exit /b 1

:node_missing_use_artifacts
echo [提示] 未检测到 Node.js，使用现有构建产物（--skip-build 模式）。
goto :check_artifacts

:node_present
if not defined SKIP_BUILD goto :prepare_dependencies
echo [0/6] 跳过构建（--skip-build），使用现有构建产物。
goto :check_artifacts

:prepare_dependencies
rem GitHub 首次克隆不会包含 node_modules；缺依赖时自动按 lockfile 安装。
if not exist "server\node_modules\express\package.json" (
  echo [准备] 安装后端依赖（npm ci）...
  call npm ci --prefix server
  if errorlevel 1 goto :dependency_fail
)
if not exist "addin\node_modules\vite\package.json" (
  echo [准备] 安装前端依赖（npm ci）...
  call npm ci --prefix addin
  if errorlevel 1 goto :dependency_fail
)

rem 默认跳过 npm test（本机较慢），但黄金路径 E2E 每次安装都跑（约 10 秒，见 release.js）
set "RELEASE_ARGS="
if not defined WITH_TESTS set "RELEASE_ARGS=--skip-tests"
if defined SKIP_E2E set "RELEASE_ARGS=%RELEASE_ARGS% --skip-e2e"
echo [0/6] 运行发布流水线（统一版本号 + 前端构建 + exe 打包 + release 清单）...
if defined WITH_TESTS echo       本次包含 npm test（可能较慢，请耐心等待）
if defined SKIP_E2E echo       已跳过黄金路径 E2E（--skip-e2e）
echo       若失败会在这里停下，不会安装旧产物。
node scripts\release.js %RELEASE_ARGS%
if not errorlevel 1 goto :check_artifacts
echo.
echo [错误] 发布流水线失败（测试未通过或构建出错），未执行安装。
if not defined NO_PAUSE pause
exit /b 1

:check_artifacts
rem ---------- 1. 校验发布产物 ----------
if not exist "release.json" goto :missing_release
if not exist "dist-exe\ppt-ai-addin.exe" goto :missing_exe
if not exist "addin\dist\index.html" goto :missing_dist
if not exist "manifest.xml" goto :missing_manifest

set "INSTALL_DIR=%LOCALAPPDATA%\PPT-AI-Addin"
set "EXE_PATH=%INSTALL_DIR%\ppt-ai-addin.exe"
set "MANIFEST_PATH=%INSTALL_DIR%\manifest.xml"
set "VBS_PATH=%INSTALL_DIR%\start-hidden.vbs"
set "RELEASE_PATH=%INSTALL_DIR%\release.json"

rem ---------- 2. 停止旧服务并复制文件 ----------
echo [1/6] 停止旧服务并复制文件到 %INSTALL_DIR% ...
taskkill /F /T /IM ppt-ai-addin.exe >nul 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\stop-service.ps1" -ProcessName "ppt-ai-addin" -RuntimeFile "%APPDATA%\ppt-ai-addin\runtime.json" -TimeoutSeconds 15
if errorlevel 1 goto :fail

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
copy /y "dist-exe\ppt-ai-addin.exe" "%EXE_PATH%" >nul
if errorlevel 1 goto :fail
copy /y "manifest.xml" "%MANIFEST_PATH%" >nul
if errorlevel 1 goto :fail
copy /y "release.json" "%RELEASE_PATH%" >nul
if errorlevel 1 goto :fail
if not exist "%INSTALL_DIR%\dist" mkdir "%INSTALL_DIR%\dist"
xcopy /y /e /q "addin\dist\*" "%INSTALL_DIR%\dist\" >nul
if errorlevel 1 goto :fail

echo [2/6] 生成后台无窗口启动脚本 ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$exe='%EXE_PATH%'; $vbs='%VBS_PATH%'; $q=[char]34; $line='CreateObject(' + $q + 'Wscript.Shell' + $q + ').Run ' + $q + $q + $q + $exe + $q + $q + $q + ', 0, False'; Set-Content -LiteralPath $vbs -Value $line -Encoding Default"
if not exist "%VBS_PATH%" goto :fail

echo [3/6] 注册到 PowerPoint ...
set "ADDIN_ID="
for /f "delims=" %%I in ('powershell -NoProfile -Command "[xml]$m=Get-Content -Raw -Encoding UTF8 -LiteralPath '%MANIFEST_PATH%'; Write-Output $m.OfficeApp.Id"') do set "ADDIN_ID=%%I"
if "%ADDIN_ID%"=="" goto :fail
reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer" /v "%ADDIN_ID%" /t REG_SZ /d "%MANIFEST_PATH%" /f
if errorlevel 1 goto :fail

echo [4/6] 设置登录自启动 ...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "PPT-AI-Addin" /t REG_SZ /d "wscript.exe \"%VBS_PATH%\"" /f
if errorlevel 1 goto :fail

echo [5/6] 启动本地服务（等待 runtime.json 记录实际端口与 Token，最多约 40 秒）...
del /q "%APPDATA%\ppt-ai-addin\runtime.json" >nul 2>nul
wscript.exe "%VBS_PATH%"
set /a TRY=0
:wait_runtime
set /a TRY+=1
if !TRY! gtr 40 goto :verify
powershell -NoProfile -Command "if (Test-Path -LiteralPath \"$env:APPDATA\ppt-ai-addin\runtime.json\") { exit 0 } else { exit 1 }" >nul 2>nul
if not errorlevel 1 goto :verify
ping -n 2 127.0.0.1 >nul
goto :wait_runtime

:verify
echo [6/6] 安装自检（运行版本 vs 发布清单）...
powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\verify-install.ps1"
set "VERIFY_CODE=%ERRORLEVEL%"

rem 读取发布版本号用于成功提示
set "RELEASE_VERSION="
for /f "delims=" %%V in ('powershell -NoProfile -Command "(Get-Content -Raw -LiteralPath '%RELEASE_PATH%' | ConvertFrom-Json).version"') do set "RELEASE_VERSION=%%V"

if not "%VERIFY_CODE%"=="0" goto :verify_warn

:success
echo.
echo ============================================
echo  安装成功！
echo   - 版本：%RELEASE_VERSION%
echo   - 已注册到 PowerPoint（当前用户）
echo   - 已设置登录自启动
echo   - 本地服务已在后台运行
echo.
echo  请完全退出并重新打开 PowerPoint，
echo  在功能区「开始」选项卡点击「模板助手」。
echo ============================================
echo.
if not defined NO_PAUSE pause
exit /b 0

:verify_warn
echo.
echo [注意] 安装已完成，但自检有未通过项（见上方 [FAIL] 行）：
echo   - 「运行版本=发布版本」失败 = 服务里跑的还是旧 exe，请重新运行本脚本；
echo   - 其余失败请按 [FAIL] 提示排查。
goto :success

:missing_release
echo.
echo [错误] 未找到 release.json 发布清单，说明尚未执行过发布构建。
echo 请在本目录（安装 Node.js）直接运行 安装.bat 自动构建，或使用完整安装包。
if not defined NO_PAUSE pause
exit /b 1

:dependency_fail
echo.
echo [错误] npm 依赖安装失败。请检查网络、npm 配置和 package-lock.json 后重试。
if not defined NO_PAUSE pause
exit /b 1

:missing_exe
echo.
echo [错误] 未找到 dist-exe\ppt-ai-addin.exe。
echo 请直接运行本安装脚本自动构建（需 Node.js），或确认构建产物完整。
if not defined NO_PAUSE pause
exit /b 1

:missing_dist
echo.
echo [错误] 未找到 addin\dist\index.html。
echo 请直接运行本安装脚本自动构建（需 Node.js），或确认构建产物完整。
if not defined NO_PAUSE pause
exit /b 1

:missing_manifest
echo.
echo [错误] 未找到 manifest.xml，请确认项目文件完整后重试。
if not defined NO_PAUSE pause
exit /b 1

:fail
echo.
echo [错误] 安装过程中出现问题，请根据上面的提示检查。
if not defined NO_PAUSE pause
exit /b 1
