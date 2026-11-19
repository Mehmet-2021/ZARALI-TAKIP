@echo off
cd /d "%~dp0"
py -3 motor.py >> guncelleme.log 2>&1
if errorlevel 1 exit /b 1
