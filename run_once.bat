@echo off
REM ---------------------------------------------------------------------
REM  BSSV Monitor - single pass. Point Windows Task Scheduler at this file.
REM  Task Scheduler > Action > Start a program:
REM      Program : C:\BSSV\bssv-monitor\run_once.bat
REM      Start in: C:\BSSV\bssv-monitor
REM ---------------------------------------------------------------------
setlocal
cd /d "%~dp0"
python bssv_monitor.py --once
set RC=%ERRORLEVEL%
if not "%RC%"=="0" echo [BSSV Monitor] exited with code %RC%
exit /b %RC%
