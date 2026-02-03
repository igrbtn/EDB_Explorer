#!/bin/bash
# Unix launcher for EDB Exporter
cd "$(dirname "$0")"
python3 main.py "$@"
