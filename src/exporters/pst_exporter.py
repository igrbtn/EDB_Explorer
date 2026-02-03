"""PST format exporter using Outlook COM automation."""

import logging
import os
import platform
from pathlib import Path
from typing import Optional

from ..core.message import Message
from .base import BaseExporter, ExportOptions

logger = logging.getLogger(__name__)


class OutlookNotAvailableError(Exception):
    """Raised when Outlook is not available on the system."""
    pass


class PSTExporter(BaseExporter):
    """
    Exports messages to PST format using Outlook COM automation.

    Requires Microsoft Outlook to be installed on Windows.
    Uses pywin32 to interact with Outlook's MAPI interface.
    """

    format_name = "PST"
    file_extension = ".pst"

    def __init__(self, options: ExportOptions):
        super().__init__(options)
        self._outlook = None
        self._namespace = None
        self._pst_store = None
        self._folders: dict[str, any] = {}  # Cache of folder objects

    @staticmethod
    def is_available() -> bool:
        """Check if Outlook COM is available on this system."""
        if platform.system() != 'Windows':
            return False

        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            outlook.GetNamespace("MAPI")
            return True
        except Exception:
            return False

    def _prepare_output(self, output_path: Path):
        """Initialize Outlook and create PST file."""
        if platform.system() != 'Windows':
            raise OutlookNotAvailableError(
                "PST export requires Windows with Microsoft Outlook installed"
            )

        try:
            import win32com.client
            from win32com.client import constants
        except ImportError:
            raise OutlookNotAvailableError(
                "pywin32 package required. Install with: pip install pywin32"
            )

        try:
            # Connect to Outlook
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")

            # Create or open PST file
            pst_path = str(output_path)
            if not pst_path.lower().endswith('.pst'):
                pst_path = str(output_path.with_suffix('.pst'))

            # Remove existing if overwrite enabled
            if Path(pst_path).exists():
                if self.options.overwrite_existing:
                    # Must remove from Outlook first if open
                    self._remove_pst_if_open(pst_path)
                    os.remove(pst_path)
                else:
                    # Try to use existing
                    self._open_existing_pst(pst_path)
                    return

            # Create new PST
            # AddStoreEx creates the PST file
            self._namespace.AddStoreEx(pst_path, 1)  # 1 = Unicode PST

            # Find the newly added store
            for store in self._namespace.Stores:
                if store.FilePath and store.FilePath.lower() == pst_path.lower():
                    self._pst_store = store
                    break

            if not self._pst_store:
                raise RuntimeError("Failed to create PST file")

            logger.info(f"Created PST file: {pst_path}")

        except Exception as e:
            logger.error(f"Failed to initialize Outlook: {e}")
            raise OutlookNotAvailableError(f"Failed to initialize Outlook: {e}")

    def _open_existing_pst(self, pst_path: str):
        """Open an existing PST file."""
        # Check if already open
        for store in self._namespace.Stores:
            if store.FilePath and store.FilePath.lower() == pst_path.lower():
                self._pst_store = store
                return

        # Open it
        self._namespace.AddStore(pst_path)

        for store in self._namespace.Stores:
            if store.FilePath and store.FilePath.lower() == pst_path.lower():
                self._pst_store = store
                return

        raise RuntimeError(f"Failed to open PST: {pst_path}")

    def _remove_pst_if_open(self, pst_path: str):
        """Remove PST from Outlook if it's open."""
        for store in self._namespace.Stores:
            if store.FilePath and store.FilePath.lower() == pst_path.lower():
                try:
                    self._namespace.RemoveStore(store.GetRootFolder())
                except Exception as e:
                    logger.warning(f"Could not remove PST from Outlook: {e}")
                break

    def _finalize_output(self, output_path: Path):
        """Close PST file properly."""
        if self._pst_store:
            try:
                # Get the path before we close
                pst_path = self._pst_store.FilePath

                # Remove from Outlook profile (but keep the file)
                root_folder = self._pst_store.GetRootFolder()
                self._namespace.RemoveStore(root_folder)

                logger.info(f"Finalized PST: {pst_path}")
            except Exception as e:
                logger.warning(f"Error finalizing PST: {e}")

        self._folders.clear()
        self._pst_store = None
        self._namespace = None
        self._outlook = None

    def _get_or_create_folder(self, folder_path: str):
        """Get or create folder hierarchy in PST."""
        if not folder_path:
            return self._pst_store.GetRootFolder()

        if folder_path in self._folders:
            return self._folders[folder_path]

        # Navigate/create folder hierarchy
        parts = folder_path.split('/')
        current_folder = self._pst_store.GetRootFolder()

        for part in parts:
            if not part:
                continue

            # Try to find existing subfolder
            found = None
            for subfolder in current_folder.Folders:
                if subfolder.Name.lower() == part.lower():
                    found = subfolder
                    break

            if found:
                current_folder = found
            else:
                # Create new folder
                try:
                    current_folder = current_folder.Folders.Add(part)
                except Exception as e:
                    logger.warning(f"Could not create folder {part}: {e}")
                    # Use parent folder
                    break

        self._folders[folder_path] = current_folder
        return current_folder

    def _export_message(self, message: Message, output_path: Path):
        """Export a single message to PST via Outlook."""
        try:
            # Get target folder
            if self.options.folder_structure:
                target_folder = self._get_or_create_folder(message.folder_path)
            else:
                target_folder = self._pst_store.GetRootFolder().Folders("Inbox")
                if not target_folder:
                    target_folder = self._pst_store.GetRootFolder().Folders.Add("Inbox")

            # Create mail item
            # 0 = olMailItem
            mail_item = self._outlook.CreateItem(0)

            # Set properties
            mail_item.Subject = message.subject or "(No Subject)"

            if message.body_html:
                mail_item.HTMLBody = message.body_html
            else:
                mail_item.Body = message.body_text or ""

            # Set sender (limited in Outlook automation)
            if message.sender:
                try:
                    mail_item.SenderEmailAddress = message.sender
                except Exception:
                    pass  # Some properties may be read-only

            # Set recipients
            for recipient in message.recipients_to:
                mail_item.Recipients.Add(recipient)

            for recipient in message.recipients_cc:
                recip = mail_item.Recipients.Add(recipient)
                recip.Type = 2  # CC

            # Set dates
            if message.date_received:
                try:
                    mail_item.ReceivedTime = message.date_received
                except Exception:
                    pass

            if message.date_sent:
                try:
                    mail_item.SentOn = message.date_sent
                except Exception:
                    pass

            # Add attachments
            if self.options.include_attachments:
                for attachment in message.attachments:
                    try:
                        # Need to save attachment to temp file first
                        import tempfile
                        with tempfile.NamedTemporaryFile(
                            delete=False,
                            suffix=f"_{attachment.filename}"
                        ) as tmp:
                            tmp.write(attachment.data)
                            tmp_path = tmp.name

                        mail_item.Attachments.Add(tmp_path)
                        os.unlink(tmp_path)
                    except Exception as e:
                        logger.warning(f"Failed to add attachment: {e}")

            # Save to PST folder
            mail_item.Save()
            mail_item.Move(target_folder)

            logger.debug(f"Exported to PST: {message.subject}")

        except Exception as e:
            logger.error(f"Failed to export message to PST: {e}")
            raise
