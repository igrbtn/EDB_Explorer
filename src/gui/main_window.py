"""Main application window for EDB Exporter."""

import logging
import platform
from pathlib import Path
from typing import Optional

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QMessageBox,
    QTreeWidget, QTreeWidgetItem, QSplitter, QTextEdit,
    QProgressBar, QStatusBar, QMenuBar, QMenu, QToolBar,
    QGroupBox, QComboBox, QCheckBox, QSpinBox, QLineEdit,
    QApplication
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QAction, QIcon

from ..core.edb_reader import EDBReader
from ..core.mailbox_parser import MailboxParser
from ..core.message import Message, Folder
from ..exporters.base import ExportOptions, ExportProgress
from ..exporters.eml_exporter import EMLExporter
from ..exporters.mbox_exporter import MBOXExporter
from ..exporters.pst_exporter import PSTExporter, OutlookNotAvailableError
from ..recovery.repair import EDBRecovery

logger = logging.getLogger(__name__)


class ExportWorker(QThread):
    """Background worker for export operations."""

    progress_updated = pyqtSignal(ExportProgress)
    export_finished = pyqtSignal(ExportProgress)
    error_occurred = pyqtSignal(str)

    def __init__(self, exporter, messages, total_count):
        super().__init__()
        self.exporter = exporter
        self.messages = messages
        self.total_count = total_count
        self._should_stop = False

    def run(self):
        try:
            self.exporter.set_progress_callback(self._on_progress)
            result = self.exporter.export(self.messages, self.total_count)
            self.export_finished.emit(result)
        except Exception as e:
            logger.error(f"Export error: {e}")
            self.error_occurred.emit(str(e))

    def _on_progress(self, progress: ExportProgress):
        self.progress_updated.emit(progress)

    def request_stop(self):
        self._should_stop = True
        if self.exporter:
            self.exporter.cancel()


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Exchange EDB Email Exporter")
        self.setMinimumSize(1000, 700)

        self._edb_path: Optional[str] = None
        self._edb_reader: Optional[EDBReader] = None
        self._parser: Optional[MailboxParser] = None
        self._folders: list[Folder] = []
        self._messages: list[Message] = []
        self._export_worker: Optional[ExportWorker] = None

        self._setup_ui()
        self._create_menus()
        self._create_toolbar()
        self._update_ui_state()

    def _setup_ui(self):
        """Set up the main UI layout."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)

        # File selection area
        file_group = QGroupBox("EDB File")
        file_layout = QHBoxLayout(file_group)

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.file_path_edit.setPlaceholderText("Select an Exchange EDB file...")
        file_layout.addWidget(self.file_path_edit)

        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.clicked.connect(self._on_browse_clicked)
        file_layout.addWidget(self.browse_btn)

        self.load_btn = QPushButton("Load")
        self.load_btn.clicked.connect(self._on_load_clicked)
        self.load_btn.setEnabled(False)
        file_layout.addWidget(self.load_btn)

        main_layout.addWidget(file_group)

        # Main content area with splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left panel - folder tree
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        left_layout.addWidget(QLabel("Folders:"))
        self.folder_tree = QTreeWidget()
        self.folder_tree.setHeaderHidden(True)
        self.folder_tree.itemSelectionChanged.connect(self._on_folder_selected)
        left_layout.addWidget(self.folder_tree)

        splitter.addWidget(left_panel)

        # Right panel - message list and preview
        right_splitter = QSplitter(Qt.Orientation.Vertical)

        # Message list
        message_panel = QWidget()
        message_layout = QVBoxLayout(message_panel)
        message_layout.setContentsMargins(0, 0, 0, 0)

        message_layout.addWidget(QLabel("Messages:"))
        self.message_tree = QTreeWidget()
        self.message_tree.setHeaderLabels(["Subject", "From", "Date", "Size"])
        self.message_tree.setColumnWidth(0, 300)
        self.message_tree.setColumnWidth(1, 200)
        self.message_tree.setColumnWidth(2, 150)
        self.message_tree.itemSelectionChanged.connect(self._on_message_selected)
        message_layout.addWidget(self.message_tree)

        right_splitter.addWidget(message_panel)

        # Preview panel
        preview_panel = QWidget()
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(0, 0, 0, 0)

        preview_layout.addWidget(QLabel("Preview:"))
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        preview_layout.addWidget(self.preview_text)

        right_splitter.addWidget(preview_panel)
        right_splitter.setSizes([400, 200])

        splitter.addWidget(right_splitter)
        splitter.setSizes([250, 750])

        main_layout.addWidget(splitter)

        # Export options
        export_group = QGroupBox("Export Options")
        export_layout = QHBoxLayout(export_group)

        export_layout.addWidget(QLabel("Format:"))
        self.format_combo = QComboBox()
        self.format_combo.addItem("EML (Individual files)", "eml")
        self.format_combo.addItem("MBOX (Unix mailbox)", "mbox")
        if platform.system() == 'Windows':
            self.format_combo.addItem("PST (Outlook)", "pst")
        export_layout.addWidget(self.format_combo)

        self.include_attachments_cb = QCheckBox("Include attachments")
        self.include_attachments_cb.setChecked(True)
        export_layout.addWidget(self.include_attachments_cb)

        self.preserve_folders_cb = QCheckBox("Preserve folder structure")
        self.preserve_folders_cb.setChecked(True)
        export_layout.addWidget(self.preserve_folders_cb)

        export_layout.addStretch()

        self.export_btn = QPushButton("Export...")
        self.export_btn.clicked.connect(self._on_export_clicked)
        self.export_btn.setEnabled(False)
        export_layout.addWidget(self.export_btn)

        main_layout.addWidget(export_group)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def _create_menus(self):
        """Create application menus."""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("&File")

        open_action = QAction("&Open EDB...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self._on_browse_clicked)
        file_menu.addAction(open_action)

        file_menu.addSeparator()

        export_action = QAction("&Export...", self)
        export_action.setShortcut("Ctrl+E")
        export_action.triggered.connect(self._on_export_clicked)
        file_menu.addAction(export_action)

        file_menu.addSeparator()

        exit_action = QAction("E&xit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Tools menu
        tools_menu = menubar.addMenu("&Tools")

        analyze_action = QAction("&Analyze EDB...", self)
        analyze_action.triggered.connect(self._on_analyze_clicked)
        tools_menu.addAction(analyze_action)

        recover_action = QAction("&Recover Corrupted...", self)
        recover_action.triggered.connect(self._on_recover_clicked)
        tools_menu.addAction(recover_action)

        # Help menu
        help_menu = menubar.addMenu("&Help")

        about_action = QAction("&About", self)
        about_action.triggered.connect(self._on_about_clicked)
        help_menu.addAction(about_action)

    def _create_toolbar(self):
        """Create application toolbar."""
        toolbar = QToolBar()
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        open_action = QAction("Open", self)
        open_action.triggered.connect(self._on_browse_clicked)
        toolbar.addAction(open_action)

        toolbar.addSeparator()

        export_action = QAction("Export", self)
        export_action.triggered.connect(self._on_export_clicked)
        toolbar.addAction(export_action)

    def _update_ui_state(self):
        """Update UI elements based on current state."""
        has_file = self._edb_path is not None
        has_data = len(self._messages) > 0

        self.load_btn.setEnabled(has_file and self._edb_reader is None)
        self.export_btn.setEnabled(has_data)

    def _on_browse_clicked(self):
        """Handle browse button click."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Exchange Database",
            "",
            "Exchange Database (*.edb);;All Files (*.*)"
        )

        if file_path:
            self._edb_path = file_path
            self.file_path_edit.setText(file_path)
            self._update_ui_state()

    def _on_load_clicked(self):
        """Handle load button click."""
        if not self._edb_path:
            return

        self.status_bar.showMessage("Loading database...")
        QApplication.processEvents()

        try:
            # Close previous reader if any
            if self._edb_reader:
                self._edb_reader.close()

            self._edb_reader = EDBReader(self._edb_path)
            info = self._edb_reader.get_info()

            if not info.is_valid:
                QMessageBox.warning(
                    self,
                    "Invalid File",
                    f"The file is not a valid EDB database:\n{info.error_message}"
                )
                return

            self._edb_reader.open()

            # Parse mailbox data
            self._parser = MailboxParser(self._edb_reader)
            self._folders = self._parser.parse_folders()

            # Populate folder tree
            self._populate_folder_tree()

            # Load messages
            self._messages = list(self._parser.iter_messages())
            self._populate_message_list(self._messages)

            self.status_bar.showMessage(
                f"Loaded {len(self._messages)} messages from {len(self._folders)} folders"
            )

        except Exception as e:
            logger.error(f"Error loading EDB: {e}")
            QMessageBox.critical(
                self,
                "Load Error",
                f"Failed to load database:\n{str(e)}"
            )
            self.status_bar.showMessage("Load failed")

        self._update_ui_state()

    def _populate_folder_tree(self):
        """Populate the folder tree widget."""
        self.folder_tree.clear()

        # Create root item
        root = QTreeWidgetItem(self.folder_tree)
        root.setText(0, "All Folders")
        root.setData(0, Qt.ItemDataRole.UserRole, None)

        # Build folder hierarchy
        folder_items: dict[str, QTreeWidgetItem] = {}

        for folder in self._folders:
            if folder.parent_id and folder.parent_id in folder_items:
                parent = folder_items[folder.parent_id]
            else:
                parent = root

            item = QTreeWidgetItem(parent)
            item.setText(0, folder.name)
            item.setData(0, Qt.ItemDataRole.UserRole, folder.folder_id)
            folder_items[folder.folder_id] = item

        self.folder_tree.expandAll()

    def _populate_message_list(self, messages: list[Message]):
        """Populate the message list widget."""
        self.message_tree.clear()

        for msg in messages:
            item = QTreeWidgetItem(self.message_tree)
            item.setText(0, msg.subject or "(No Subject)")
            item.setText(1, msg.sender or "")
            item.setText(2, msg.display_date.strftime("%Y-%m-%d %H:%M") if msg.display_date else "")
            item.setText(3, f"{len(msg.body_text)} chars" if msg.body_text else "")
            item.setData(0, Qt.ItemDataRole.UserRole, msg)

    def _on_folder_selected(self):
        """Handle folder selection change."""
        items = self.folder_tree.selectedItems()
        if not items:
            return

        folder_id = items[0].data(0, Qt.ItemDataRole.UserRole)

        if folder_id is None:
            # Show all messages
            filtered = self._messages
        else:
            # Filter by folder
            filtered = [m for m in self._messages if m.folder_id == folder_id]

        self._populate_message_list(filtered)

    def _on_message_selected(self):
        """Handle message selection change."""
        items = self.message_tree.selectedItems()
        if not items:
            self.preview_text.clear()
            return

        message: Message = items[0].data(0, Qt.ItemDataRole.UserRole)
        if not message:
            return

        # Build preview text
        preview = f"Subject: {message.subject or '(No Subject)'}\n"
        preview += f"From: {message.sender_name} <{message.sender}>\n" if message.sender_name else f"From: {message.sender}\n"
        preview += f"To: {', '.join(message.recipients_to)}\n" if message.recipients_to else ""
        preview += f"Date: {message.display_date}\n" if message.display_date else ""

        if message.attachments:
            preview += f"Attachments: {len(message.attachments)}\n"

        preview += "\n" + "-" * 40 + "\n\n"
        preview += message.body_text or message.body_html or "(No content)"

        self.preview_text.setPlainText(preview)

    def _on_export_clicked(self):
        """Handle export button click."""
        if not self._messages:
            QMessageBox.information(self, "No Data", "No messages to export.")
            return

        format_code = self.format_combo.currentData()

        # Get output path
        if format_code == "eml":
            output_path = QFileDialog.getExistingDirectory(
                self, "Select Export Folder"
            )
        elif format_code == "mbox":
            output_path, _ = QFileDialog.getSaveFileName(
                self, "Save MBOX File", "", "MBOX Files (*.mbox)"
            )
        else:  # pst
            output_path, _ = QFileDialog.getSaveFileName(
                self, "Save PST File", "", "PST Files (*.pst)"
            )

        if not output_path:
            return

        # Create export options
        options = ExportOptions(
            output_path=output_path,
            include_attachments=self.include_attachments_cb.isChecked(),
            folder_structure=self.preserve_folders_cb.isChecked()
        )

        # Create exporter
        try:
            if format_code == "eml":
                exporter = EMLExporter(options)
            elif format_code == "mbox":
                exporter = MBOXExporter(options)
            else:
                exporter = PSTExporter(options)
        except OutlookNotAvailableError as e:
            QMessageBox.warning(
                self, "Outlook Required",
                f"PST export requires Microsoft Outlook:\n{str(e)}"
            )
            return

        # Start export in background
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(self._messages))
        self.progress_bar.setValue(0)
        self.export_btn.setEnabled(False)

        self._export_worker = ExportWorker(
            exporter,
            iter(self._messages),
            len(self._messages)
        )
        self._export_worker.progress_updated.connect(self._on_export_progress)
        self._export_worker.export_finished.connect(self._on_export_finished)
        self._export_worker.error_occurred.connect(self._on_export_error)
        self._export_worker.start()

    def _on_export_progress(self, progress: ExportProgress):
        """Handle export progress update."""
        self.progress_bar.setValue(progress.exported_messages + progress.failed_messages)
        self.status_bar.showMessage(
            f"Exporting: {progress.exported_messages}/{progress.total_messages}"
        )

    def _on_export_finished(self, progress: ExportProgress):
        """Handle export completion."""
        self.progress_bar.setVisible(False)
        self.export_btn.setEnabled(True)

        if progress.error:
            QMessageBox.warning(
                self, "Export Error",
                f"Export completed with errors:\n{progress.error}"
            )
        else:
            QMessageBox.information(
                self, "Export Complete",
                f"Exported {progress.exported_messages} messages.\n"
                f"Failed: {progress.failed_messages}"
            )

        self.status_bar.showMessage("Export complete")

    def _on_export_error(self, error: str):
        """Handle export error."""
        self.progress_bar.setVisible(False)
        self.export_btn.setEnabled(True)

        QMessageBox.critical(self, "Export Failed", f"Export failed:\n{error}")
        self.status_bar.showMessage("Export failed")

    def _on_analyze_clicked(self):
        """Handle analyze action."""
        if not self._edb_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Select EDB to Analyze", "",
                "Exchange Database (*.edb);;All Files (*.*)"
            )
            if not file_path:
                return
        else:
            file_path = self._edb_path

        recovery = EDBRecovery(file_path)
        info = recovery.analyze()

        msg = f"File: {file_path}\n\n"
        if info.is_corrupted:
            msg += f"Status: CORRUPTED\n"
            msg += f"Type: {info.corruption_type}\n"
            msg += f"Severity: {info.severity}\n"
            msg += f"Recoverable: {'Yes' if info.recoverable else 'No'}\n"
            msg += f"\nDetails: {info.details}"
        else:
            msg += "Status: OK (No corruption detected)"

        QMessageBox.information(self, "EDB Analysis", msg)

    def _on_recover_clicked(self):
        """Handle recover action."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Corrupted EDB", "",
            "Exchange Database (*.edb);;All Files (*.*)"
        )

        if not file_path:
            return

        recovery = EDBRecovery(file_path)
        info = recovery.analyze()

        if not info.is_corrupted:
            QMessageBox.information(
                self, "Not Corrupted",
                "This file does not appear to be corrupted.\n"
                "Use the normal Load function instead."
            )
            return

        if not info.recoverable:
            QMessageBox.warning(
                self, "Not Recoverable",
                f"This file cannot be recovered:\n{info.details}"
            )
            return

        # Attempt recovery
        self.status_bar.showMessage("Attempting recovery...")
        QApplication.processEvents()

        try:
            self._messages = list(recovery.attempt_recovery())
            self._folders = []  # Folders may not be recoverable

            if self._messages:
                self.folder_tree.clear()
                self._populate_message_list(self._messages)

                QMessageBox.information(
                    self, "Recovery Complete",
                    f"Recovered {len(self._messages)} messages."
                )
                self.status_bar.showMessage(f"Recovered {len(self._messages)} messages")
            else:
                QMessageBox.warning(
                    self, "No Data Recovered",
                    "Could not recover any messages from this file."
                )
                self.status_bar.showMessage("Recovery failed - no data found")

        except Exception as e:
            logger.error(f"Recovery error: {e}")
            QMessageBox.critical(
                self, "Recovery Error",
                f"Recovery failed:\n{str(e)}"
            )
            self.status_bar.showMessage("Recovery failed")

        self._update_ui_state()

    def _on_about_clicked(self):
        """Show about dialog."""
        QMessageBox.about(
            self,
            "About EDB Exporter",
            "Exchange EDB Email Exporter\n\n"
            "Version 1.0.0\n\n"
            "Export emails from Microsoft Exchange EDB database files\n"
            "to EML, MBOX, or PST formats.\n\n"
            "Supports recovery from corrupted databases."
        )

    def closeEvent(self, event):
        """Handle window close event."""
        # Cancel any running export
        if self._export_worker and self._export_worker.isRunning():
            self._export_worker.request_stop()
            self._export_worker.wait(5000)

        # Close database
        if self._edb_reader:
            self._edb_reader.close()

        event.accept()
