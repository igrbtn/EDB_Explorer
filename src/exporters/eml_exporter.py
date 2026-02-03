"""EML format exporter - RFC 2822 standard email format."""

import logging
import re
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, formatdate
from pathlib import Path

from ..core.message import Message
from .base import BaseExporter, ExportOptions

logger = logging.getLogger(__name__)


class EMLExporter(BaseExporter):
    """
    Exports messages to EML format (RFC 2822).

    EML files are standard email format readable by most email clients.
    Each message is saved as a separate .eml file.
    """

    format_name = "EML"
    file_extension = ".eml"

    def __init__(self, options: ExportOptions):
        super().__init__(options)

    def _export_message(self, message: Message, output_path: Path):
        """Export a single message to EML format."""
        # Build the email message
        if message.body_html and message.body_text:
            # Multipart alternative for both text and HTML
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

        if message.in_reply_to:
            msg['In-Reply-To'] = message.in_reply_to

        if message.references:
            msg['References'] = " ".join(message.references)

        # Add any original headers
        for header, value in message.headers.items():
            if header.lower() not in ['subject', 'from', 'to', 'cc', 'date', 'message-id']:
                msg[header] = value

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

                    if attachment.content_type:
                        maintype, subtype = attachment.content_type.split('/', 1) \
                            if '/' in attachment.content_type else ('application', 'octet-stream')
                        part.set_type(attachment.content_type)

                    if attachment.content_id:
                        part.add_header('Content-ID', f'<{attachment.content_id}>')

                    msg.attach(part)
                except Exception as e:
                    logger.warning(f"Failed to attach {attachment.filename}: {e}")

        # Determine output file path
        if output_path.is_dir():
            # Create subfolder for folder structure
            if self.options.folder_structure and message.folder_path:
                folder_dir = output_path / self._sanitize_path(message.folder_path)
                folder_dir.mkdir(parents=True, exist_ok=True)
            else:
                folder_dir = output_path

            # Generate filename
            filename = self._generate_filename(message)
            file_path = folder_dir / filename

            # Handle duplicates
            counter = 1
            base_path = file_path.with_suffix('')
            while file_path.exists() and not self.options.overwrite_existing:
                file_path = base_path.parent / f"{base_path.name}_{counter}.eml"
                counter += 1
        else:
            file_path = output_path

        # Write the file
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(msg.as_string())

        logger.debug(f"Exported: {file_path}")

    def _generate_filename(self, message: Message) -> str:
        """Generate a filename for the message."""
        # Use date and subject for filename
        date_str = ""
        if message.display_date:
            date_str = message.display_date.strftime("%Y%m%d_%H%M%S")

        subject = self._sanitize_filename(message.subject or "no_subject")
        subject = subject[:50]  # Truncate long subjects

        if date_str:
            return f"{date_str}_{subject}.eml"
        else:
            return f"{message.message_id[:20]}_{subject}.eml"

    def _sanitize_filename(self, name: str) -> str:
        """Remove invalid characters from filename."""
        # Remove or replace invalid characters
        invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
        sanitized = re.sub(invalid_chars, '_', name)
        sanitized = sanitized.strip('. ')
        return sanitized or "unnamed"

    def _sanitize_path(self, path: str) -> str:
        """Sanitize folder path."""
        parts = path.split('/')
        return '/'.join(self._sanitize_filename(p) for p in parts)
