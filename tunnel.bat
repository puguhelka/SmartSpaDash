@echo off
title SmartSpaDash Tunnel
cd /d "C:\Users\USER\Documents\smartspa-local"
echo ===============================
echo   SmartSpaDash Public Tunnel
echo ===============================
echo.
echo Pastikan server lokal jalan dulu...
start "SmartSpaDash" /MIN node server.js
timeout /t 3 /nobreak >nul
echo.
echo Membuka tunnel...
echo.
npx localtunnel --port 3000
pause
