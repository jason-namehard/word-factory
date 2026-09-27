@echo off
rem ============================================================
rem  word 工厂 GUI launcher (double-click to start)
rem  Starts the local server on 127.0.0.1:8765 and opens the browser.
rem  Close this black window (or press the "退出" button on the page) to stop.
rem ============================================================
cd /d "%~dp0"
"D:\Hermes\hermes-agent\venv\Scripts\python.exe" -m wordfactory.cli gui
echo.
echo word factory has stopped. Press any key to close this window.
pause >nul
