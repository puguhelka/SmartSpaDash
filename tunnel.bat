@echo off
title SmartSpaDash — Cloudflare Tunnel
cd /d "C:\Users\USER\Documents\smartspa-local"

echo.
echo   ============================================
echo     SmartSpaDash — PUBLIC TUNNEL (Cloudflare)
echo   ============================================
echo.
echo   Server lokal: http://localhost:3000
echo.
echo   [1/2] Memastikan server berjalan...
start "SmartSpaDash Server" /MIN node server.js
timeout /t 2 /nobreak >nul
echo   [2/2] Membuka Cloudflare Tunnel...
echo.

:loop
echo   Menghubungkan...
echo.
"%USERPROFILE%\cloudflared.exe" tunnel --url http://localhost:3000 2>&1 | findstr /C:"trycloudflare.com"
echo.
echo   ============================================
echo   Tunnel PUTUS. Restart otomatis dalam 3 detik...
echo   (JANGAN tutup window ini!)
echo   ============================================
timeout /t 3 /nobreak >nul
goto loop
