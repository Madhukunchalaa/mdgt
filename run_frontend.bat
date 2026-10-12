@echo off
title MDGT Frontend (Next.js)
echo ===================================================
echo Starting MDGT Next.js Frontend on port 3000
echo ===================================================
cd /d "%~dp0meil_frontend"
npm run dev
pause
