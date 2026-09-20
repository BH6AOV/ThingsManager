@echo off
rem === ThingsManager one-click start (Windows) ===
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Node.js not found in PATH. Please install Node.js 22+ first.
  pause
  exit /b 1
)
if not exist node_modules (
  echo Installing dependencies, please wait...
  call npm install
)
start "" http://127.0.0.1:3200
echo Starting ThingsManager at http://127.0.0.1:3200
echo Close this window to stop the server.
node server.js
pause
