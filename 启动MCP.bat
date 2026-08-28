@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   模板助手 - MCP 服务器（供 ChatGPT 等 AI 使用）
echo ============================================
echo.
echo   本窗口保持运行期间，ChatGPT 才可访问模板。
echo   关闭本窗口即停止服务。
echo.
where node >nul 2>nul
if errorlevel 1 (
  echo [错误] 未找到 Node.js，请先安装 Node.js 24+。
  pause
  exit /b 1
)
echo 正在启动 MCP 服务器（stdio）...
echo.
node "%~dp0server\mcp\index.js"
pause
