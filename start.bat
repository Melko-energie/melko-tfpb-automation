@echo off
echo Demarrage de Melko Energie - TFPB Automation...
echo.

if not exist .venv (
    echo Le venv n'existe pas. Lancez d'abord setup.bat
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
python -m backend.main
pause
