@echo off
REM ---------------------------------------------------------------------
REM  BSSV Monitor - resident loop. Repeats every run_interval_seconds from
REM  config.ini until the window is closed or the service is stopped.
REM  Leave this running in a console, or wrap it with NSSM to make it a
REM  Windows service:  nssm install BSSVMonitor "C:\BSSV\bssv-monitor\run_loop.bat"
REM ---------------------------------------------------------------------
setlocal
cd /d "%~dp0"
title BSSV Monitor
python bssv_monitor.py --loop
