@echo off
echo Starting MDGT Backend and Frontend...
start "MDGT Backend" cmd /k "call "%~dp0run_backend.bat""
start "MDGT Frontend" cmd /k "call "%~dp0run_frontend.bat""
echo Both servers started in separate windows!
echo Backend:  http://127.0.0.1:8000
echo Frontend: http://localhost:3000
