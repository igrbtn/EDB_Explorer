@echo off
echo ==========================================
echo  Exchange EDB Exporter - Windows Install
echo ==========================================
echo.

:: Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python 3.10+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo [OK] Python found:
python --version
echo.

:: Check for pip
pip --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pip is not installed
    echo Installing pip...
    python -m ensurepip --upgrade
)

echo [OK] pip found
echo.

:: Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip
echo.

:: Install dependencies
echo Installing dependencies...
echo.

echo [1/5] Installing PyQt6...
pip install PyQt6>=6.4.0

echo [2/5] Installing libesedb-python...
pip install libesedb-python>=20240420
if errorlevel 1 (
    echo [WARNING] libesedb-python failed. Trying binary-only install...
    pip install --only-binary :all: libesedb-python
)

echo [3/5] Installing dissect.esedb...
pip install dissect.esedb>=3.0

echo [4/5] Installing python-dateutil and chardet...
pip install python-dateutil>=2.8.0 chardet>=5.0.0

echo [5/5] Installing pywin32...
pip install pywin32>=305

echo.
echo ==========================================
echo  Installation complete!
echo ==========================================
echo.
echo To run the application:
echo   python gui_viewer_v2.py
echo.
pause
