@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/3] 检查依赖...
where node >nul 2>nul || (echo 未检测到 Node.js，请先安装: https://nodejs.org/ & pause & exit /b 1)
echo [2/3] 安装依赖（首次运行会稍慢）...
if not exist server\node_modules (cd server && call npm install && cd ..)
if not exist addin\node_modules (cd addin && call npm install && cd ..)
echo [3/3] 启动服务...
start "PPT-AI 后端" cmd /k "cd /d %~dp0server && node src/index.js"
start "PPT-AI 前端" cmd /k "cd /d %~dp0addin && npm run dev"
echo 已启动。请按 安装说明.md 完成 sideload。
start notepad 安装说明.md
pause