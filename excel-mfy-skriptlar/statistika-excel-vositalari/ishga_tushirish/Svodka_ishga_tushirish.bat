@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   Tumanlar kesimida xatolar svodkasi
echo ============================================
echo.
where python >nul 2>nul
if errorlevel 1 (
  echo [!] Python o'rnatilmagan.
  echo     https://www.python.org/downloads/ dan o'rnating.
  echo     O'rnatishda "Add Python to PATH" katagini belgilang.
  echo.
  pause
  exit /b
)
echo Kerakli kutubxona tekshirilmoqda...
python -m pip install --quiet --disable-pip-version-check openpyxl
echo.
python "%~dp0xatolar_svodka.py"
