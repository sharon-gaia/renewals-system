@echo off
chcp 65001 >nul
echo.
echo ========================================
echo    מערכת שירות לקוחות - מתחילה...
echo ========================================
echo.

pip install -r requirements.txt -q

echo.
echo פותח דפדפן...
start http://localhost:5000

python app.py
pause
