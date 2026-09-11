@echo off
title Suite Logistica - NO CIERRES esta ventana mientras la uses
cd /d "C:\Users\accuv\suite_logistica"
start "" cmd /c "timeout /t 3 /nobreak >nul & start http://127.0.0.1:5000"
py app.py
pause
