@echo off
REM Windows launcher for EDB Exporter
cd /d "%~dp0"
python main.py %*
pause
