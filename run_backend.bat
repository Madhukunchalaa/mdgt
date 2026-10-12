@echo off
title MDGT Backend (Daphne)
echo ===================================================
echo Starting MDGT Django Backend with Daphne on port 8000
echo ===================================================
cd /d "%~dp0meil_backend"
call .\venv\Scripts\activate.bat
daphne -b 127.0.0.1 -p 8000 core.asgi:application
pause
