@echo off
echo ============================================
echo   Melko Energie - Installation automatique
echo ============================================
echo.

:: Verifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo Telechargez Python 3.10+ sur https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Afficher la version
for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYVER=%%i
echo [OK] %PYVER% detecte

:: Creer le venv avec fallback
echo.
echo [1/3] Creation de l'environnement virtuel...
if exist .venv (
    echo      .venv existe deja, on le reutilise.
) else (
    python -m venv .venv 2>nul
    if errorlevel 1 (
        echo      venv standard a echoue, tentative avec --without-pip...
        python -m venv .venv --without-pip 2>nul
        if errorlevel 1 (
            echo      Deuxieme tentative avec virtualenv...
            pip install virtualenv 2>nul
            python -m virtualenv .venv 2>nul
            if errorlevel 1 (
                echo [ERREUR] Impossible de creer le venv.
                echo Essayez manuellement : python -m venv .venv --without-pip
                pause
                exit /b 1
            )
        )
    )
    echo      [OK] .venv cree
)

:: Activer le venv
echo.
echo [2/3] Activation du venv...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [ERREUR] Impossible d'activer le venv.
    pause
    exit /b 1
)
echo      [OK] venv active

:: Assurer que pip est installe dans le venv
python -m ensurepip --upgrade >nul 2>&1

:: Installer les dependances
echo.
echo [3/3] Installation des dependances...
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERREUR] Echec de l'installation des dependances.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Installation terminee avec succes !
echo ============================================
echo.
echo Pour lancer l'application :
echo   .venv\Scripts\activate
echo   python -m backend.main
echo.
echo Puis ouvrir http://localhost:8000
echo.
pause
