#!/bin/bash
echo "=========================================="
echo " Exchange EDB Exporter - Ubuntu Install"
echo "=========================================="
echo

# Check for sudo
if [ "$EUID" -ne 0 ]; then
    SUDO="sudo"
else
    SUDO=""
fi

# Update package list
echo "Updating package list..."
$SUDO apt-get update
echo

# Install Python and pip
echo "[1/6] Installing Python3 and pip..."
$SUDO apt-get install -y python3 python3-pip python3-venv
echo "[OK] Python3 found: $(python3 --version)"
echo

# Install system dependencies for PyQt6
echo "[2/6] Installing PyQt6 system dependencies..."
$SUDO apt-get install -y \
    libxcb-xinerama0 \
    libxcb-cursor0 \
    libxkbcommon-x11-0 \
    libgl1-mesa-glx \
    libegl1-mesa \
    libxcb-icccm4 \
    libxcb-image0 \
    libxcb-keysyms1 \
    libxcb-randr0 \
    libxcb-render-util0 \
    libxcb-shape0
echo

# Install build dependencies for libesedb
echo "[3/6] Installing build dependencies..."
$SUDO apt-get install -y \
    build-essential \
    python3-dev \
    libffi-dev
echo

# Upgrade pip
echo "Upgrading pip..."
python3 -m pip install --upgrade pip
echo

# Install Python dependencies
echo "[4/6] Installing PyQt6..."
pip3 install PyQt6>=6.4.0

echo "[5/6] Installing libesedb-python and dissect.esedb..."
pip3 install libesedb-python>=20240420
pip3 install dissect.esedb>=3.0

echo "[6/6] Installing python-dateutil and chardet..."
pip3 install python-dateutil>=2.8.0 chardet>=5.0.0

echo
echo "=========================================="
echo " Installation complete!"
echo "=========================================="
echo
echo "To run the application:"
echo "  python3 gui_viewer_v2.py"
echo
