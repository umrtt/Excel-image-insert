@echo off
REM Build script for Excel Image Insert

echo.
echo ===================================
echo Excel Image Insert - Build Script
echo ===================================
echo.

REM Check if virtual environment is activated
if not defined VIRTUAL_ENV (
    echo Virtual environment not activated!
    echo Please activate the virtual environment first:
    echo .venv\Scripts\activate
    pause
    exit /b 1
)

echo Installing/updating dependencies...
pip install -r requirements.txt

echo.
echo Building executable...
echo This may take a few minutes...
echo.

pyinstaller build.spec

echo.
echo ===================================
echo Build Complete!
echo ===================================
echo.
echo Your executable is located in:
echo   dist\Excel Image Insert\Excel Image Insert.exe
echo.
pause
