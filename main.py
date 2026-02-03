#!/usr/bin/env python3
"""
Exchange EDB Email Exporter

Main entry point for the application.
"""

import sys
import logging
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)


def main():
    """Main entry point."""
    try:
        from PyQt6.QtWidgets import QApplication
        from PyQt6.QtCore import Qt
    except ImportError:
        print("Error: PyQt6 is required. Install with: pip install PyQt6")
        sys.exit(1)

    # Create application
    app = QApplication(sys.argv)
    app.setApplicationName("EDB Exporter")
    app.setOrganizationName("EDBExporter")

    # Apply dark theme style (optional)
    app.setStyle("Fusion")

    # Import and create main window
    from src.gui.main_window import MainWindow

    window = MainWindow()
    window.show()

    # If file was passed as argument, open it
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if Path(file_path).exists():
            window.file_path_edit.setText(file_path)
            window._edb_path = file_path
            window._update_ui_state()

    # Run event loop
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
