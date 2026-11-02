@echo off
cd /d %~dp0
python checkin_app.py
if errorlevel 1 pause
