"""MBOX format exporter - Unix mailbox format."""

import logging
import re
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, formatdate
from pathlib import Path
from typing import Optional
import fcntl
import os

from ..core.message import Message
from .base import BaseExporter, ExportOptions

logger = logging.getLogger(__name__)


class MBOXExporter(BaseExporter):
    """
    Exports messages to MBOX format (Unix mailbox).

    MBOX stores multiple messages in a single file, separated by 'From ' lines.
    Compatible with Thunderbird, Evolution, and other Unix mail clients.
    """

    format_name = "MBOX"
    file_extension = ".mbox"

    def __init__(self, options: ExportOptions):
        super().__init__(options)
        self._mbox_files: dict[str, any] = {}  # folder_path -> file handle

    def _prepare_output(self, output_path: Path):
        """Prepare output - create directory or single mbox file."""
        if self.options.folder_structure:
            # Create directory for multiple mbox files
            output_path.mkdir(parents=True, exist_ok=True)
        else:
            # Single mbox file
            output_path.parent.mkdir(parents=True, exist_ok=True)
            if output_path.exists() and not self.options.overwrite_existing:
                # Append to existing
                pass
            elif output_path.exists():
                # Clear existing
                output_path.unlink()

    def _finalize_output(self, output_path: Path):
        """Close all open mbox files."""
        for path, file_handle in self._mbox_files.items():
            try:
                file_handle.close()
            except Exception as e:
                logger.warning(f"Error closing {path}: {e}")
        self._mbox_files.clear()

    def _get_mbox_file(self, output_path: Path, folder_path: str):
        """Get or create mbox file for the given folder."""
        if self.options.folder_structure and folder_path:
            # Separate mbox file per folder
            sanitized = self._sanitize_path(folder_path)
            mbox_path = output_path / f"{sanitized}.mbox"
            mbox_path.parent.mkdir(parents=True, exist_ok=True)
            key = str(mbox_path)
        else:
            # Single mbox file
            if output_path.is_dir():
                mbox_path = output_path / "export.mbox"
            else:
                mbox_path = output_path
            key = str(mbox_path)

        if key not in self._mbox_files:
            mode = 'a' if mbox_path.exists() else 'w'
            self._mbox_files[key] = open(mbox_path, mode, encoding='utf-8')

        return self._mbox_files[key]

    def _export_message(self, message: Message, output_path: Path):
        """Export a single message to MBOX format."""
        # Build the email message (similar to EML)
        if message.body_html and message.body_text:
            msg = MIMEMultipart('alternative')
            msg.attach(MIMEText(message.body_text, 'plain', 'utf-8'))
            msg.attach(MIMEText(message.body_html, 'html', 'utf-8'))
        elif message.body_html:
            msg = MIMEMultipart('alternative')
            msg.attach(MIMEText(message.body_html, 'html', 'utf-8'))
        else:
            msg = MIMEMultipart()
            msg.attach(MIMEText(message.body_text or '', 'plain', 'utf-8'))

        # Set headers
        msg['Subject'] = message.subject or "(No Subject)"

        if message.sender_name and message.sender:
            msg['From'] = formataddr((message.sender_name, message.sender))
        elif message.sender:
            msg['From'] = message.sender
        else:
            msg['From'] = "unknown@unknown"

        if message.recipients_to:
            msg['To'] = ", ".join(message.recipients_to)

        if message.recipients_cc:
            msg['Cc'] = ", ".join(message.recipients_cc)

        if message.date_sent:
            msg['Date'] = formatdate(message.date_sent.timestamp(), localtime=True)
        elif message.date_received:
            msg['Date'] = formatdate(message.date_received.timestamp(), localtime=True)

        if message.internet_message_id:
            msg['Message-ID'] = message.internet_message_id
        else:
            msg['Message-ID'] = f"<{message.message_id}@edb-export>"

        # Add attachments
        if self.options.include_attachments:
            for attachment in message.attachments:
                try:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.data)
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=attachment.filename
                    )
                    msg.attach(part)
                except Exception as e:
                    logger.warning(f"Failed to attach {attachment.filename}: {e}")

        # Get the appropriate mbox file
        mbox_file = self._get_mbox_file(output_path, message.folder_path)

        # Write MBOX format
        # "From " line with sender and date
        sender_addr = self._extract_email(message.sender) or "unknown@unknown"
        date_str = ""
        if message.display_date:
            date_str = message.display_date.strftime("%a %b %d %H:%M:%S %Y")
        else:
            from datetime import datetime
            date_str = datetime.now().strftime("%a %b %d %H:%M:%S %Y")

        from_line = f"From {sender_addr} {date_str}\n"

        # Escape any "From " lines in the body
        msg_str = msg.as_string()
        msg_str = self._escape_from_lines(msg_str)

        # Write to mbox
        mbox_file.write(from_line)
        mbox_file.write(msg_str)
        mbox_file.write("\n")  # Blank line between messages

        logger.debug(f"Added message to MBOX: {message.subject}")

    def _extract_email(self, addr: str) -> str:
        """Extract email address from a potentially formatted address."""
        if not addr:
            return ""
        # Match email in angle brackets or bare email
        match = re.search(r'<([^>]+)>', addr)
        if match:
            return match.group(1)
        # Check if it looks like an email
        if '@' in addr:
            return addr.strip()
        return addr

    def _escape_from_lines(self, content: str) -> str:
        """Escape 'From ' lines in message body (MBOX format requirement)."""
        lines = content.split('\n')
        escaped = []
        for line in lines:
            if line.startswith('From '):
                escaped.append('>' + line)
            else:
                escaped.append(line)
        return '\n'.join(escaped)

    def _sanitize_path(self, path: str) -> str:
        """Sanitize folder path for filesystem."""
        invalid_chars = r'[<>:"|?*\x00-\x1f]'
        sanitized = re.sub(invalid_chars, '_', path)
        sanitized = sanitized.replace('/', '_').replace('\\', '_')
        return sanitized.strip('. ') or "inbox"
