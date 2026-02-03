#!/bin/bash
# Prerequisites installer for EDB Exporter (macOS/Linux)

set -e

echo "========================================"
echo "  EDB Exporter - Prerequisites Installer"
echo "========================================"
echo ""

# Check Python version
echo "[1/4] Checking Python..."
if command -v python3 &> /dev/null; then
    PYTHON=python3
    PIP=pip3
elif command -v python &> /dev/null; then
    PYTHON=python
    PIP=pip
else
    echo "ERROR: Python not found. Please install Python 3.8 or higher."
    echo "  macOS: brew install python3"
    echo "  Ubuntu: sudo apt install python3 python3-pip"
    exit 1
fi

PYVER=$($PYTHON --version 2>&1 | cut -d' ' -f2 | cut -d'.' -f1,2)
echo "  Found Python $PYVER"

# Check pip
echo ""
echo "[2/4] Checking pip..."
if ! $PYTHON -m pip --version &> /dev/null; then
    echo "  Installing pip..."
    $PYTHON -m ensurepip --upgrade
fi
echo "  pip is available"

# Install system dependencies for pyesedb (macOS)
echo ""
echo "[3/4] Checking system dependencies..."
if [[ "$OSTYPE" == "darwin"* ]]; then
    echo "  macOS detected"
    if command -v brew &> /dev/null; then
        echo "  Homebrew found, checking for libesedb..."
        if ! brew list libesedb &> /dev/null 2>&1; then
            echo "  Note: libesedb not in Homebrew. pyesedb will be installed via pip."
        fi
    else
        echo "  Note: Homebrew not found. Consider installing: /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
    fi
elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
    echo "  Linux detected"
    if command -v apt-get &> /dev/null; then
        echo "  Debian/Ubuntu detected"
        echo "  You may need: sudo apt-get install build-essential python3-dev"
    fi
fi

# Install Python packages
echo ""
echo "[4/4] Installing Python packages..."
cd "$(dirname "$0")"

$PYTHON -m pip install --upgrade pip

echo ""
echo "Installing PyQt6..."
$PYTHON -m pip install PyQt6>=6.4.0

echo ""
echo "Installing ESE database reader..."
$PYTHON -m pip install libesedb-python || {
    echo "  libesedb-python failed, trying dissect.esedb..."
    $PYTHON -m pip install dissect.esedb || {
        echo ""
        echo "WARNING: ESE library installation failed."
        echo "Try manually: pip install libesedb-python"
    }
}

echo ""
echo "Installing other dependencies..."
$PYTHON -m pip install python-dateutil>=2.8.0 chardet>=5.0.0

# Verify installation
echo ""
echo "========================================"
echo "  Verifying installation..."
echo "========================================"

$PYTHON -c "
import sys
print(f'Python: {sys.version}')

try:
    from PyQt6.QtWidgets import QApplication
    print('PyQt6: OK')
except ImportError as e:
    print(f'PyQt6: FAILED - {e}')

try:
    import pyesedb
    print('pyesedb: OK')
except ImportError as e:
    print(f'pyesedb: FAILED - {e}')
    print('  Note: EDB reading requires pyesedb. Install manually if needed.')

try:
    import dateutil
    print('python-dateutil: OK')
except ImportError as e:
    print(f'python-dateutil: FAILED - {e}')

try:
    import chardet
    print('chardet: OK')
except ImportError as e:
    print(f'chardet: FAILED - {e}')
"

echo ""
echo "========================================"
echo "  Installation complete!"
echo "========================================"
echo ""
echo "To run the application:"
echo "  ./run.sh"
echo "  or: python3 main.py"
echo ""
echo "Note for macOS:"
echo "  - PST export is NOT available (requires Windows + Outlook)"
echo "  - EML and MBOX export work fully"
echo ""
