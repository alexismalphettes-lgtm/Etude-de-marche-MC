@echo off
title Installation - Scraper Marine Center
echo.
echo ============================================
echo   Installation du scraper Marine Center
echo ============================================
echo.

:: Vérifie que Python est installé
python --version >nul 2>&1
if errorlevel 1 (
    echo ERREUR : Python n'est pas installe.
    echo.
    echo Telechargez Python sur : https://www.python.org/downloads/
    echo IMPORTANT : cochez "Add Python to PATH" pendant l'installation !
    echo.
    pause
    exit /b 1
)

echo Python detecte. Installation des dependances...
echo.

pip install playwright beautifulsoup4 openpyxl lxml

echo.
echo Telechargement du navigateur Chromium (environ 150 Mo)...
echo (Cela peut prendre quelques minutes selon votre connexion)
echo.

playwright install chromium

echo.
echo ============================================
echo   Installation terminee avec succes !
echo   Vous pouvez fermer cette fenetre et
echo   double-cliquer sur LANCER.bat
echo ============================================
echo.
pause
