@echo off
echo Starting WorkflowGenie...

:: Start backend in a new window that auto-restarts
start "WorkflowGenie Backend" cmd /k "cd /d "%~dp0WorkFlow Genie\backend" && :loop && python -m uvicorn main:app --host 0.0.0.0 --port 8000 && goto loop"

:: Wait for backend
timeout /t 20 /nobreak

:: Start frontend in a new window
start "WorkflowGenie Frontend" cmd /k "cd /d "%~dp0WorkFlow Genie\frontend" && npm run dev"

echo Both servers starting...
echo Frontend: http://localhost:3000
pause
