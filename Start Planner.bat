@echo off
REM YD Ascend - Sales & Dispatch Planner. Opens http://localhost:8600 in your browser. Close this window to stop it.
cd /d "%~dp0"
start "" http://localhost:8600
python -m uvicorn server:app --app-dir planner --host 127.0.0.1 --port 8600
pause
