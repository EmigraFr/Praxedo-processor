@echo off
REM ===================================================
REM  Praxedo Processor - Lanceur Web (Windows)
REM  Double-cliquez pour ouvrir l'app dans le navigateur
REM ===================================================

cd /d "%~dp0"

REM Vérifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo Telechargez-le sur https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Praxedo Processor - Version Web
echo ========================================
echo.
echo [1/2] Installation des dependances...
python -m pip install -q -r requirements_web.txt

echo.
echo [2/2] Demarrage de l'application...
echo.
echo Le navigateur va s'ouvrir automatiquement.
echo Fermez cette fenetre pour arreter l'app.
echo.

python run_web.py

pause
