@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   CONCILIACAO BANCARIA - Metodo Adv Digital
echo ============================================
echo.
echo Preparando componentes (demora so na primeira vez)...
py -m pip install --quiet --disable-pip-version-check pandas openpyxl 2>nul || python -m pip install --quiet --disable-pip-version-check pandas openpyxl
echo.
py index.py . 2>nul || python index.py .
echo.
pause
