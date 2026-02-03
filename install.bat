@echo off
REM Prerequisites installer for EDB Exporter (Windows)

echo ========================================
echo   EDB Exporter - Prerequisites Installer
echo ========================================
echo.

REM Check Python
echo [1/4] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.8 or higher.
    echo   Download from: https://www.python.org/downloads/
    echo   Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYVER=%%i
echo   Found Python %PYVER%

REM Check pip
echo.
echo [2/4] Checking pip...
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo   Installing pip...
    python -m ensurepip --upgrade
)
echo   pip is available

REM System dependencies note
echo.
echo [3/4] System dependencies...
echo   Windows detected
echo   Checking for Microsoft Outlook (for PST export)...
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" >nul 2>&1
if errorlevel 1 (
    echo   Note: Microsoft Outlook not detected.
    echo   PST export will not be available.
) else (
    echo   Microsoft Office detected - PST export should work.
)

REM Install Python packages
echo.
echo [4/4] Installing Python packages...
cd /d "%~dp0"

python -m pip install --upgrade pip

echo.
echo Installing PyQt6...
python -m pip install PyQt6>=6.4.0

echo.
echo Installing pyesedb (ESE database reader)...
python -m pip install pyesedb
if errorlevel 1 (
    echo.
    echo WARNING: pyesedb installation may have failed.
    echo Try: pip install libesedb-python
)

echo.
echo Installing pywin32 (for Outlook COM)...
python -m pip install pywin32>=305

echo.
echo Installing other dependencies...
python -m pip install python-dateutil>=2.8.0 chardet>=5.0.0

REM Verify installation
echo.
echo ========================================
echo   Verifying installation...
echo ========================================

python -c "import sys; print(f'Python: {sys.version}')"
python -c "from PyQt6.QtWidgets import QApplication; print('PyQt6: OK')" 2>nul || echo PyQt6: FAILED
python -c "import pyesedb; print('pyesedb: OK')" 2>nul || echo pyesedb: FAILED
python -c "import win32com.client; print('pywin32: OK')" 2>nul || echo pywin32: FAILED
python -c "import dateutil; print('python-dateutil: OK')" 2>nul || echo python-dateutil: FAILED
python -c "import chardet; print('chardet: OK')" 2>nul || echo chardet: FAILED

echo.
echo ========================================
echo   Installation complete!
echo ========================================
echo.
echo To run the application:
echo   run.bat
echo   or: python main.py
echo.
pause
