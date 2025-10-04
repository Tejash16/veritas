@echo off
echo ========================================
echo  Starting Veritas AI Auditor (Development)
echo ========================================

:: Start backend server
echo Starting backend server on port 8000...
start "Veritas Backend" cmd /k "cd backend && call venv\Scripts\activate.bat && python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8008"

:: Wait a moment for backend to start
timeout /t 3 /nobreak >nul

:: Start frontend development server
echo Starting frontend server on port 3000...
start "Veritas Frontend" cmd /k "cd frontend && npm start"

echo.
echo ========================================
echo  Development servers started!
echo ========================================
echo.
echo Backend API: http://localhost:8000
echo Frontend: http://localhost:3000
echo API Docs: http://localhost:8000/docs
echo Metrics: http://localhost:8001 (when available)
echo.
echo Press any key to stop all servers...
pause >nul

:: Stop all servers
taskkill /f /im "python.exe" 2>nul
taskkill /f /im "node.exe" 2>nul

echo All servers stopped.