@echo off
title Scraper Alliance Nautique 66 - Marine Center
echo.
echo ============================================
echo   Scraper Alliance Nautique 66
echo   Marine Center - Veille concurrentielle
echo ============================================
echo.
echo Scraping en cours, veuillez patienter...
echo.

python scraper.py

echo.
echo Ouverture du dossier des rapports...
if exist "output\" (
    explorer output
) else (
    echo Le dossier output n'existe pas encore.
)

echo.
echo ============================================
echo   Termine ! Le fichier Excel est dans
echo   le dossier "output" qui vient de s'ouvrir
echo ============================================
echo.
pause
