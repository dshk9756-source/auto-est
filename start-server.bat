@echo off
title Estimate Server
cd /d "%~dp0"
echo Open browser: http://localhost:5173
echo Close this window to stop the server.
echo.
node server.js
pause
