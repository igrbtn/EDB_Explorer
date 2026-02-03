"""Exchange 2013+ database parser for partitioned mailbox tables."""

import logging
import struct
import re
from datetime import datetime, timezone
from typing import Iterator, Optional, Dict, List, Tuple
from dataclasses import dataclass

from .message import Message, Attachment, Folder

logger = logging.getLogger(__name__)


def detect_encryption(db) -> Tuple[bool, str]:
    """
    Detect if database content is encrypted.

    Returns:
        (is_encrypted, description)
    """
    # Check MessageClass in first message - if garbled, likely encrypted
    for i in range(db.get_number_of_tables()):
        try:
            table = db.get_table(i)
            if table and table.name.startswith("Message_"):
                for j in range(table.get_number_of_columns()):
                    col = table.get_column(j)
                    if col and col.name == "MessageClass":
                        record = table.get_record(0)
                        if record:
                            val = record.get_value_data(j)
                            if val:
                                text = val.decode('utf-16-le', errors='ignore').rstrip('\x00')
                                # Normal MessageClass starts with "IPM."
                                if text.startswith("IPM."):
                                    return False, "Content is not encrypted"
                                else:
                                    return True, "Database content appears to be encrypted (Exchange 2016+ encryption)"
                break
        except:
            pass

    return False, "Could not determine encryption status"


@dataclass
class MailboxInfo:
    """Information about a mailbox in the database."""
    mailbox_number: int
    mailbox_guid: str
    message_count: int
    display_name: str = ""
    email_address: str = ""


class Exchange2013Parser:
    """
    Parser for Exchange 2013+ EDB databases.

    Exchange 2013+ uses partitioned tables where each mailbox has its own
    set of tables (Message_100, Message_101, Folder_100, etc.)
    """

    def __init__(self, db):
        """
        Initialize parser with an open pyesedb database.

        Args:
            db: Open pyesedb.file() object
        """
        self.db = db
        self._tables: Dict[str, any] = {}
        self._mailboxes: List[MailboxInfo] = []
        self._table_map: Dict[str, int] = {}  # table_name -> table_index

        self._cache_tables()

    def _cache_tables(self):
        """Cache table references and build table map."""
        num_tables = self.db.get_number_of_tables()

        for i in range(num_tables):
            try:
                table = self.db.get_table(i)
                if table:
                    self._tables[table.name] = table
                    self._table_map[table.name] = i
            except Exception as e:
                logger.warning(f"Error caching table {i}: {e}")

        logger.info(f"Cached {len(self._tables)} tables")

    def get_mailboxes(self) -> List[MailboxInfo]:
        """Get list of mailboxes in the database."""
        if self._mailboxes:
            return self._mailboxes

        mailbox_table = self._tables.get("Mailbox")
        if not mailbox_table:
            logger.warning("Mailbox table not found")
            return []

        col_map = self._get_column_map(mailbox_table)

        for i in range(mailbox_table.get_number_of_records()):
            try:
                record = mailbox_table.get_record(i)
                if not record:
                    continue

                mailbox_num = self._get_int_value(record, col_map.get('MailboxNumber', -1))
                msg_count = self._get_int_value(record, col_map.get('MessageCount', -1))
                guid = self._get_bytes_value(record, col_map.get('MailboxGuid', -1))

                if mailbox_num is not None:
                    self._mailboxes.append(MailboxInfo(
                        mailbox_number=mailbox_num,
                        mailbox_guid=guid.hex() if guid else "",
                        message_count=msg_count or 0
                    ))
            except Exception as e:
                logger.warning(f"Error reading mailbox {i}: {e}")

        logger.info(f"Found {len(self._mailboxes)} mailboxes")
        return self._mailboxes

    def get_message_tables(self) -> List[str]:
        """Get list of all Message_* tables."""
        return [name for name in self._tables.keys()
                if name.startswith("Message_") and name[8:].isdigit()]

    def iter_messages(self, mailbox_number: Optional[int] = None) -> Iterator[Message]:
        """
        Iterate over messages in the database.

        Args:
            mailbox_number: Optional mailbox number to filter by

        Yields:
            Message objects
        """
        if mailbox_number is not None:
            table_names = [f"Message_{mailbox_number}"]
        else:
            table_names = self.get_message_tables()

        for table_name in table_names:
            table = self._tables.get(table_name)
            if not table:
                continue

            try:
                yield from self._iter_table_messages(table, table_name)
            except Exception as e:
                logger.error(f"Error reading {table_name}: {e}")

    def _iter_table_messages(self, table, table_name: str) -> Iterator[Message]:
        """Iterate messages from a single table."""
        col_map = self._get_column_map(table)
        num_records = table.get_number_of_records()

        logger.info(f"Reading {num_records} messages from {table_name}")

        for i in range(num_records):
            try:
                record = table.get_record(i)
                if not record:
                    continue

                message = self._parse_message_record(record, col_map, table_name, i)
                if message:
                    yield message

            except Exception as e:
                logger.debug(f"Error parsing record {i} in {table_name}: {e}")

    def _parse_message_record(self, record, col_map: dict, table_name: str, index: int) -> Optional[Message]:
        """Parse a message record into a Message object."""
        try:
            msg_id = f"{table_name}_{index}"

            message = Message(message_id=msg_id)

            # Get timestamps
            date_received = self._get_filetime_value(record, col_map.get('DateReceived', -1))
            date_sent = self._get_filetime_value(record, col_map.get('DateSent', -1))
            message.date_received = date_received
            message.date_sent = date_sent

            # Get flags
            message.is_read = self._get_bool_value(record, col_map.get('IsRead', -1))
            message.has_attachments = self._get_bool_value(record, col_map.get('HasAttachments', -1))

            # Get size
            size_bytes = self._get_bytes_value(record, col_map.get('Size', -1))

            # Parse PropertyBlob for MAPI properties
            prop_blob = self._get_bytes_value(record, col_map.get('PropertyBlob', -1))
            if prop_blob:
                self._parse_property_blob(message, prop_blob)

            # Parse LargePropertyValueBlob for body content
            large_blob = self._get_bytes_value(record, col_map.get('LargePropertyValueBlob', -1))
            if large_blob:
                self._parse_large_property_blob(message, large_blob)

            # Try DisplayTo
            display_to = self._get_string_value(record, col_map.get('DisplayTo', -1))
            if display_to:
                message.recipients_to = [r.strip() for r in display_to.split(';') if r.strip()]

            # Get MessageClass
            msg_class = self._get_string_value(record, col_map.get('MessageClass', -1))
            if msg_class:
                message.headers['X-Message-Class'] = msg_class

            # Try to get subject from SubjectPrefix
            subject_prefix = self._get_string_value(record, col_map.get('SubjectPrefix', -1))
            if subject_prefix and not message.subject:
                message.subject = subject_prefix

            return message

        except Exception as e:
            logger.debug(f"Error parsing message: {e}")
            return None

    def _parse_property_blob(self, message: Message, blob: bytes):
        """Parse MAPI property blob to extract message properties."""
        if not blob or len(blob) < 4:
            return

        # Exchange PropertyBlob is a compressed/encoded format
        # Try to extract readable strings

        # Look for email addresses
        email_pattern = re.compile(rb'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
        emails = email_pattern.findall(blob)

        if emails and not message.sender:
            # First email is usually sender
            message.sender = emails[0].decode('ascii', errors='ignore')

        # Try to find subject-like text (UTF-16LE)
        try:
            text = blob.decode('utf-16-le', errors='ignore')
            # Remove non-printable characters
            readable = ''.join(c if c.isprintable() or c == ' ' else '' for c in text)
            words = readable.split()

            # Look for potential subject
            if not message.subject and words:
                # Find first meaningful text
                for i, word in enumerate(words):
                    if len(word) > 3 and word.isascii():
                        # Take a few words as subject
                        potential_subject = ' '.join(words[i:i+10])
                        if len(potential_subject) > 5:
                            message.subject = potential_subject[:100]
                            break
        except:
            pass

    def _parse_large_property_blob(self, message: Message, blob: bytes):
        """Parse large property blob for body content."""
        if not blob or len(blob) < 10:
            return

        # Try different encodings for body
        for encoding in ['utf-8', 'utf-16-le', 'cp1252']:
            try:
                text = blob.decode(encoding, errors='ignore')
                # Clean up
                readable = ''.join(c if c.isprintable() or c in '\n\r\t ' else '' for c in text)
                if len(readable) > 20:
                    if not message.body_text:
                        message.body_text = readable
                    break
            except:
                continue

    def _get_column_map(self, table) -> dict:
        """Get mapping of column names to indices."""
        col_map = {}
        for j in range(table.get_number_of_columns()):
            try:
                col = table.get_column(j)
                if col:
                    col_map[col.name] = j
            except:
                pass
        return col_map

    def _get_bytes_value(self, record, col_idx: int) -> Optional[bytes]:
        """Get raw bytes value from record."""
        if col_idx < 0:
            return None
        try:
            return record.get_value_data(col_idx)
        except:
            return None

    def _get_string_value(self, record, col_idx: int) -> str:
        """Get string value from record (tries UTF-16LE then UTF-8)."""
        val = self._get_bytes_value(record, col_idx)
        if not val:
            return ""

        for encoding in ['utf-16-le', 'utf-8', 'cp1252']:
            try:
                return val.decode(encoding).rstrip('\x00')
            except:
                continue
        return ""

    def _get_int_value(self, record, col_idx: int) -> Optional[int]:
        """Get integer value from record."""
        val = self._get_bytes_value(record, col_idx)
        if not val:
            return None

        try:
            if len(val) == 4:
                return struct.unpack('<I', val)[0]
            elif len(val) == 8:
                return struct.unpack('<Q', val)[0]
            elif len(val) == 2:
                return struct.unpack('<H', val)[0]
        except:
            pass
        return None

    def _get_bool_value(self, record, col_idx: int) -> bool:
        """Get boolean value from record."""
        val = self._get_bytes_value(record, col_idx)
        if not val:
            return False
        return val != b'\x00' and val != b'\x00\x00'

    def _get_filetime_value(self, record, col_idx: int) -> Optional[datetime]:
        """Get datetime from Windows FILETIME."""
        val = self._get_bytes_value(record, col_idx)
        if not val or len(val) != 8:
            return None

        try:
            filetime = struct.unpack('<Q', val)[0]
            if filetime == 0:
                return None
            # Convert FILETIME to Unix timestamp
            unix_time = (filetime - 116444736000000000) / 10000000
            return datetime.fromtimestamp(unix_time, tz=timezone.utc)
        except:
            return None

    def get_total_message_count(self) -> int:
        """Get total message count across all mailboxes."""
        total = 0
        for table_name in self.get_message_tables():
            table = self._tables.get(table_name)
            if table:
                try:
                    total += table.get_number_of_records()
                except:
                    pass
        return total
