@echo off
rem serve-local.bat -- runs the local server for the WO Volume Table extension in a visible window.
rem Leave the window open while the dashboard is in use.  Close it (or Ctrl+C) to stop.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0serve-local.ps1"
