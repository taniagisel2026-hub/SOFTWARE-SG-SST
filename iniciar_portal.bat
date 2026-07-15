@echo off
cd /d "%~dp0"
echo Iniciando el portal SG-SST...
echo (No cierres esta ventana mientras uses el portal)
echo.
start "" /min cmd /c "timeout /t 3 /nobreak >nul & start "" http://localhost:8000/portal/"
uvicorn api:app --port 8000
pause
