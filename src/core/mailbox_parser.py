"""Parser for Exchange mailbox data from ESE tables."""

import logging
import struct
from datetime import datetime, timezone
from typing import Iterator, Optional
from .message import Message, Attachment, Folder, MessagePriority
from .edb_reader import EDBReader

logger = logging.getLogger(__name__)


class MailboxParser:
    """
    Parses Exchange mailbox data from ESE database tables.

    Exchange stores data in multiple tables with complex relationships.
    This parser attempts to extract email messages and related data.
    """

    # Exchange property tags (common ones)
    PR_SUBJECT = 0x0037
    PR_BODY = 0x1000
    PR_HTML = 0x1013
    PR_SENDER_EMAIL = 0x0C1F
    PR_SENDER_NAME = 0x0C1A
    PR_RECEIVED_TIME = 0x0E06
    PR_SENT_TIME = 0x0039
    PR_MESSAGE_CLASS = 0x001A
    PR_DISPLAY_TO = 0x0E04
    PR_DISPLAY_CC = 0x0E03
    PR_DISPLAY_BCC = 0x0E02
    PR_ATTACH_FILENAME = 0x3704
    PR_ATTACH_DATA = 0x3701
    PR_ATTACH_MIME_TYPE = 0x370E

    def __init__(self, edb_reader: EDBReader):
        """
        Initialize mailbox parser.

        Args:
            edb_reader: Open EDB reader instance
        """
        self.reader = edb_reader
        self._folders: dict[str, Folder] = {}
        self._message_count = 0

    def parse_folders(self) -> list[Folder]:
        """
        Parse folder structure from the database.

        Returns:
            List of Folder objects
        """
        folders = []

        # Try common Exchange folder table names
        folder_tables = ["Folder", "FolderTable", "Folders", "tblFolder"]

        for table_name in folder_tables:
            if table_name in self.reader.get_table_names():
                try:
                    for record in self.reader.iter_table_records(table_name):
                        folder = self._parse_folder_record(record)
                        if folder:
                            folders.append(folder)
                            self._folders[folder.folder_id] = folder
                except Exception as e:
                    logger.warning(f"Error parsing folder table {table_name}: {e}")
                break

        logger.info(f"Parsed {len(folders)} folders")
        return folders

    def _parse_folder_record(self, record: dict) -> Optional[Folder]:
        """Parse a single folder record."""
        try:
            # Try to find folder ID and name from common column names
            folder_id = None
            name = "Unknown"
            parent_id = None

            for key in ["FolderId", "folder_id", "id", "FolderID"]:
                if key in record and record[key]:
                    folder_id = str(record[key])
                    break

            if not folder_id:
                return None

            for key in ["DisplayName", "display_name", "Name", "name", "FolderName"]:
                if key in record and record[key]:
                    name = self._decode_value(record[key])
                    break

            for key in ["ParentFolderId", "parent_folder_id", "ParentId", "parent_id"]:
                if key in record and record[key]:
                    parent_id = str(record[key])
                    break

            return Folder(
                folder_id=folder_id,
                name=name,
                parent_id=parent_id
            )

        except Exception as e:
            logger.debug(f"Error parsing folder record: {e}")
            return None

    def iter_messages(self, folder_id: Optional[str] = None) -> Iterator[Message]:
        """
        Iterate over messages in the database.

        Args:
            folder_id: Optional folder ID to filter by

        Yields:
            Message objects
        """
        # Try common Exchange message table names
        message_tables = ["Message", "MessageTable", "Messages", "tblMessage", "Msg"]

        for table_name in message_tables:
            if table_name in self.reader.get_table_names():
                logger.info(f"Reading messages from table: {table_name}")

                try:
                    for record in self.reader.iter_table_records(table_name):
                        message = self._parse_message_record(record)
                        if message:
                            if folder_id is None or message.folder_id == folder_id:
                                self._message_count += 1
                                yield message
                except Exception as e:
                    logger.error(f"Error reading message table: {e}")
                break

    def _parse_message_record(self, record: dict) -> Optional[Message]:
        """Parse a single message record into a Message object."""
        try:
            # Extract message ID
            msg_id = None
            for key in ["MessageId", "message_id", "id", "MsgId", "p0E1F001F"]:
                if key in record and record[key]:
                    msg_id = str(record[key])
                    break

            if not msg_id:
                # Generate ID from record position
                msg_id = f"msg_{self._message_count}"

            message = Message(message_id=msg_id)

            # Parse subject
            for key in ["Subject", "subject", "p0037001F", "PR_SUBJECT"]:
                if key in record and record[key]:
                    message.subject = self._decode_value(record[key])
                    break

            # Parse sender
            for key in ["SenderEmailAddress", "sender_email", "p0C1F001F", "From"]:
                if key in record and record[key]:
                    message.sender = self._decode_value(record[key])
                    break

            for key in ["SenderName", "sender_name", "p0C1A001F", "FromName"]:
                if key in record and record[key]:
                    message.sender_name = self._decode_value(record[key])
                    break

            # Parse recipients
            for key in ["DisplayTo", "display_to", "p0E04001F", "To"]:
                if key in record and record[key]:
                    to_str = self._decode_value(record[key])
                    message.recipients_to = self._parse_recipient_list(to_str)
                    break

            for key in ["DisplayCc", "display_cc", "p0E03001F", "Cc"]:
                if key in record and record[key]:
                    cc_str = self._decode_value(record[key])
                    message.recipients_cc = self._parse_recipient_list(cc_str)
                    break

            # Parse dates
            for key in ["MessageDeliveryTime", "received_time", "p0E060040", "ReceivedTime"]:
                if key in record and record[key]:
                    message.date_received = self._parse_datetime(record[key])
                    break

            for key in ["ClientSubmitTime", "sent_time", "p00390040", "SentTime"]:
                if key in record and record[key]:
                    message.date_sent = self._parse_datetime(record[key])
                    break

            # Parse body
            for key in ["Body", "body", "p1000001F", "BodyText"]:
                if key in record and record[key]:
                    message.body_text = self._decode_value(record[key])
                    break

            for key in ["Html", "html", "p1013001F", "BodyHtml", "HtmlBody"]:
                if key in record and record[key]:
                    message.body_html = self._decode_value(record[key])
                    break

            # Parse folder
            for key in ["FolderId", "folder_id", "ParentFolderId"]:
                if key in record and record[key]:
                    message.folder_id = str(record[key])
                    if message.folder_id in self._folders:
                        message.folder_path = self._build_folder_path(message.folder_id)
                    break

            return message

        except Exception as e:
            logger.warning(f"Error parsing message record: {e}")
            # Return partial message for recovery
            if msg_id:
                msg = Message(message_id=msg_id, is_recovered=True)
                msg.recovery_errors.append(str(e))
                return msg
            return None

    def _decode_value(self, value) -> str:
        """Decode a value to string, handling various encodings."""
        if value is None:
            return ""
        if isinstance(value, str):
            return value
        if isinstance(value, bytes):
            # Try UTF-16LE first (common in Exchange)
            try:
                return value.decode('utf-16-le').rstrip('\x00')
            except UnicodeDecodeError:
                pass
            # Try UTF-8
            try:
                return value.decode('utf-8').rstrip('\x00')
            except UnicodeDecodeError:
                pass
            # Fall back to latin-1
            return value.decode('latin-1', errors='replace').rstrip('\x00')
        return str(value)

    def _parse_recipient_list(self, value: str) -> list[str]:
        """Parse a recipient string into list of addresses."""
        if not value:
            return []
        # Split on common delimiters
        for delim in [';', ',', '\n']:
            if delim in value:
                return [r.strip() for r in value.split(delim) if r.strip()]
        return [value.strip()] if value.strip() else []

    def _parse_datetime(self, value) -> Optional[datetime]:
        """Parse a datetime value from various formats."""
        if value is None:
            return None

        # If bytes, try to parse as Windows FILETIME
        if isinstance(value, bytes) and len(value) == 8:
            try:
                filetime = struct.unpack('<Q', value)[0]
                # Convert Windows FILETIME to Unix timestamp
                # FILETIME is 100-nanosecond intervals since Jan 1, 1601
                unix_time = (filetime - 116444736000000000) / 10000000
                return datetime.fromtimestamp(unix_time, tz=timezone.utc)
            except (struct.error, ValueError, OSError):
                pass

        # If already datetime
        if isinstance(value, datetime):
            return value

        # Try string parsing
        if isinstance(value, str):
            from dateutil import parser
            try:
                return parser.parse(value)
            except ValueError:
                pass

        return None

    def _build_folder_path(self, folder_id: str) -> str:
        """Build full folder path from folder ID."""
        path_parts = []
        current_id = folder_id

        while current_id and current_id in self._folders:
            folder = self._folders[current_id]
            path_parts.insert(0, folder.name)
            current_id = folder.parent_id

        return "/".join(path_parts) if path_parts else ""

    def get_message_count_estimate(self) -> int:
        """Get estimated message count."""
        for table_name in ["Message", "MessageTable", "Messages", "tblMessage"]:
            count = self.reader.get_record_count(table_name)
            if count > 0:
                return count
        return 0

    @property
    def message_count(self) -> int:
        """Get number of messages parsed so far."""
        return self._message_count
