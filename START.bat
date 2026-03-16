@echo off
title VMix Title Controller
color 0A

echo.
echo  ============================================
echo     VMix Title Controller - Starting
echo  ============================================
echo.

cd /d "%~dp0"

if not exist "node_modules" (
    echo  Installing dependencies - please wait...
    echo.
    call npm install
    if %errorlevel% neq 0 (
        echo.
        echo  [ERROR] npm install failed.
        echo  Make sure Node.js is installed: https://nodejs.org
        echo.
        pause
        exit /b 1
    )
    echo.
)

echo  ============================================
echo   Open in browser: http://localhost:3000
echo   Remote operators: http://YOUR-IP:3000
echo  ============================================
echo.
echo  Keep this window open while running.
echo  Press Ctrl+C to stop the server.
echo.

node server.js

echo.
echo  Server stopped.
pause
