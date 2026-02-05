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
echo.
echo NOTE: libesedb-python requires Microsoft Visual C++ Build Tools to compile.
echo.
pip install libesedb-python
if errorlevel 1 (
    echo.
    echo ==========================================
    echo  libesedb-python installation FAILED
    echo ==========================================
    echo.
    echo This package is REQUIRED and needs Visual C++ Build Tools.
    echo.
    echo Please follow these steps:
    echo.
    echo 1. Download Visual C++ Build Tools from:
    echo    https://visualstudio.microsoft.com/visual-cpp-build-tools/
    echo.
    echo 2. Run the installer
    echo.
    echo 3. Select "Desktop development with C++" workload
    echo.
    echo 4. Click Install and wait for completion
    echo.
    echo 5. RESTART your computer
    echo.
    echo 6. Run this script again
    echo.
    echo ==========================================
    pause
    exit /b 1
)

echo [3/5] Installing dissect.esedb...
pip install dissect.esedb>=3.0

echo [4/5] Installing python-dateutil and chardet...
pip install python-dateutil>=2.8.0 chardet>=5.0.0

echo [5/5] Installing pywin32...
pip install pywin32>=305

echo.
echo ==========================================
echo  Verifying installation...
echo ==========================================
echo.
python -c "from PyQt6.QtWidgets import QApplication; print('[OK] PyQt6')" 2>nul || echo [FAIL] PyQt6
python -c "import pyesedb; print('[OK] libesedb-python')" 2>nul || echo [WARN] libesedb-python - not installed
python -c "from dissect.esedb import EseDB; print('[OK] dissect.esedb')" 2>nul || echo [FAIL] dissect.esedb
python -c "import dateutil; print('[OK] python-dateutil')" 2>nul || echo [FAIL] python-dateutil

echo.
echo ==========================================
echo  Installation complete!
echo ==========================================
echo.
echo To run the application:
echo   python gui_viewer_v2.py
echo.
echo If you see [WARN] for libesedb-python, install Visual C++ Build Tools
echo and run this script again for full functionality.
echo.
pause
