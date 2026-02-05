#!/bin/bash
echo "=========================================="
echo " Exchange EDB Exporter - macOS Install"
echo "=========================================="
echo

# Check for Homebrew
if ! command -v brew &> /dev/null; then
    echo "[INFO] Homebrew not found. Installing..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
fi
echo "[OK] Homebrew found"

# Check for Python
if ! command -v python3 &> /dev/null; then
    echo "[INFO] Python3 not found. Installing via Homebrew..."
    brew install python@3.11
fi
echo "[OK] Python3 found: $(python3 --version)"
echo

# Upgrade pip
echo "Upgrading pip..."
python3 -m pip install --upgrade pip
echo

# Install dependencies
echo "Installing dependencies..."
echo

echo "[1/4] Installing PyQt6..."
pip3 install PyQt6>=6.4.0

echo "[2/4] Installing libesedb-python..."
pip3 install libesedb-python>=20240420
if [ $? -ne 0 ]; then
    echo "[WARNING] libesedb-python failed. Installing build dependencies..."
    brew install libffi
    pip3 install libesedb-python>=20240420
fi

echo "[3/4] Installing dissect.esedb..."
pip3 install dissect.esedb>=3.0

echo "[4/4] Installing python-dateutil and chardet..."
pip3 install python-dateutil>=2.8.0 chardet>=5.0.0

echo
echo "=========================================="
echo " Installation complete!"
echo "=========================================="
echo
echo "To run the application:"
echo "  python3 gui_viewer_v2.py"
echo
