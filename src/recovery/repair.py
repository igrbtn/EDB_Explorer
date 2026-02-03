"""Recovery module for corrupted or damaged EDB files."""

import logging
import struct
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator, Optional

from ..core.message import Message
from ..core.edb_reader import EDBReader, EDBInfo

logger = logging.getLogger(__name__)


@dataclass
class RecoveryResult:
    """Result of a recovery operation."""
    total_messages_found: int = 0
    recovered_messages: int = 0
    failed_recoveries: int = 0
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class CorruptionInfo:
    """Information about detected corruption."""
    is_corrupted: bool = False
    corruption_type: str = ""
    severity: str = "unknown"  # minor, moderate, severe
    recoverable: bool = True
    details: str = ""


class EDBRecovery:
    """
    Recovery handler for corrupted EDB files.

    Attempts to extract data from damaged or improperly closed Exchange databases.
    """

    # ESE page sizes
    PAGE_SIZES = [4096, 8192, 16384, 32768]

    # ESE file states
    STATE_CLEAN = 1
    STATE_DIRTY = 2
    STATE_BACKUP = 3

    def __init__(self, file_path: str):
        """
        Initialize recovery handler.

        Args:
            file_path: Path to the EDB file
        """
        self.file_path = Path(file_path)
        self._page_size: Optional[int] = None
        self._corruption_info: Optional[CorruptionInfo] = None

    def analyze(self) -> CorruptionInfo:
        """
        Analyze the EDB file for corruption.

        Returns:
            CorruptionInfo with details about any detected issues
        """
        if not self.file_path.exists():
            return CorruptionInfo(
                is_corrupted=True,
                corruption_type="file_not_found",
                severity="severe",
                recoverable=False,
                details="File does not exist"
            )

        corruption = CorruptionInfo()

        try:
            with open(self.file_path, 'rb') as f:
                # Read header
                header = f.read(4096)

                if len(header) < 4096:
                    corruption.is_corrupted = True
                    corruption.corruption_type = "truncated_header"
                    corruption.severity = "severe"
                    corruption.recoverable = False
                    corruption.details = "File header is truncated"
                    return corruption

                # Check signature
                signature = header[:4]
                if signature != b'\xef\xcd\xab\x89':
                    corruption.is_corrupted = True
                    corruption.corruption_type = "invalid_signature"
                    corruption.severity = "severe"
                    corruption.recoverable = False
                    corruption.details = f"Invalid ESE signature: {signature.hex()}"
                    return corruption

                # Check database state (offset varies by version)
                # Try common offsets
                for offset in [52, 56, 60]:
                    try:
                        state = struct.unpack('<I', header[offset:offset+4])[0]
                        if state == self.STATE_DIRTY:
                            corruption.is_corrupted = True
                            corruption.corruption_type = "dirty_shutdown"
                            corruption.severity = "minor"
                            corruption.recoverable = True
                            corruption.details = "Database was not cleanly shut down"
                            break
                    except struct.error:
                        continue

                # Check page size
                for offset in [236, 240, 244]:
                    try:
                        page_size = struct.unpack('<I', header[offset:offset+4])[0]
                        if page_size in self.PAGE_SIZES:
                            self._page_size = page_size
                            break
                    except struct.error:
                        continue

                if not self._page_size:
                    corruption.warnings.append("Could not determine page size, using default")
                    self._page_size = 8192  # Default Exchange page size

                # Basic consistency check - verify file size is multiple of page size
                file_size = self.file_path.stat().st_size
                if file_size % self._page_size != 0:
                    if not corruption.is_corrupted:
                        corruption.is_corrupted = True
                        corruption.corruption_type = "size_mismatch"
                        corruption.severity = "moderate"
                        corruption.recoverable = True
                        corruption.details = f"File size not aligned to page size ({self._page_size})"

        except Exception as e:
            corruption.is_corrupted = True
            corruption.corruption_type = "read_error"
            corruption.severity = "severe"
            corruption.recoverable = False
            corruption.details = f"Error reading file: {e}"

        self._corruption_info = corruption
        return corruption

    def attempt_recovery(self) -> Iterator[Message]:
        """
        Attempt to recover messages from a corrupted database.

        Yields:
            Recovered Message objects
        """
        if not self._corruption_info:
            self.analyze()

        if not self._corruption_info.recoverable:
            logger.error(f"Database is not recoverable: {self._corruption_info.details}")
            return

        logger.info(f"Attempting recovery of {self.file_path}")

        # Try normal reading first with error handling
        try:
            yield from self._recover_via_normal_read()
        except Exception as e:
            logger.warning(f"Normal read failed: {e}, trying raw scan")

        # If normal read fails or returns nothing, try raw page scanning
        # This is a last resort for severely corrupted files
        try:
            yield from self._recover_via_raw_scan()
        except Exception as e:
            logger.error(f"Raw scan failed: {e}")

    def _recover_via_normal_read(self) -> Iterator[Message]:
        """Try to read using normal ESE library with error tolerance."""
        try:
            import pyesedb
        except ImportError:
            logger.error("pyesedb not available for recovery")
            return

        try:
            db = pyesedb.file()
            db.open(str(self.file_path))

            # Try to iterate tables and recover what we can
            num_tables = db.get_number_of_tables()

            for i in range(num_tables):
                try:
                    table = db.get_table(i)
                    if not table:
                        continue

                    # Look for message-like tables
                    table_name = table.name.lower()
                    if 'message' in table_name or 'msg' in table_name:
                        yield from self._recover_from_table(table)

                except Exception as e:
                    logger.warning(f"Error reading table {i}: {e}")
                    continue

            db.close()

        except Exception as e:
            logger.error(f"Error in normal recovery: {e}")

    def _recover_from_table(self, table) -> Iterator[Message]:
        """Recover messages from a single table."""
        try:
            num_records = table.get_number_of_records()
            num_columns = table.get_number_of_columns()

            # Get column names
            columns = []
            for i in range(num_columns):
                try:
                    col = table.get_column(i)
                    columns.append(col.name if col else f"col_{i}")
                except Exception:
                    columns.append(f"col_{i}")

            # Iterate records with error handling
            for i in range(num_records):
                try:
                    record = table.get_record(i)
                    if not record:
                        continue

                    message = self._record_to_message(record, columns, i)
                    if message:
                        message.is_recovered = True
                        yield message

                except Exception as e:
                    logger.debug(f"Error reading record {i}: {e}")
                    continue

        except Exception as e:
            logger.warning(f"Error recovering from table: {e}")

    def _record_to_message(self, record, columns: list[str], index: int) -> Optional[Message]:
        """Convert a database record to a Message object."""
        try:
            data = {}
            for j, col_name in enumerate(columns):
                try:
                    value = record.get_value_data(j)
                    data[col_name] = value
                except Exception:
                    pass

            # Build message from available data
            message = Message(message_id=f"recovered_{index}")

            # Try to extract subject
            for key in ['Subject', 'subject', 'p0037001F']:
                if key in data and data[key]:
                    message.subject = self._decode_safe(data[key])
                    break

            # Try to extract body
            for key in ['Body', 'body', 'p1000001F']:
                if key in data and data[key]:
                    message.body_text = self._decode_safe(data[key])
                    break

            # Try to extract sender
            for key in ['SenderEmailAddress', 'From', 'sender']:
                if key in data and data[key]:
                    message.sender = self._decode_safe(data[key])
                    break

            # Only return if we got at least some content
            if message.subject or message.body_text or message.sender:
                return message

            return None

        except Exception as e:
            logger.debug(f"Error converting record to message: {e}")
            return None

    def _recover_via_raw_scan(self) -> Iterator[Message]:
        """
        Last resort: scan raw file for email-like patterns.

        This is very slow and may produce false positives.
        """
        logger.info("Starting raw byte scan for email patterns")

        if not self._page_size:
            self._page_size = 8192

        try:
            with open(self.file_path, 'rb') as f:
                page_num = 0
                message_count = 0

                while True:
                    page = f.read(self._page_size)
                    if not page:
                        break

                    # Look for email-like patterns in the page
                    messages = self._scan_page_for_emails(page, page_num)
                    for msg in messages:
                        msg.is_recovered = True
                        msg.recovery_errors.append("Recovered via raw scan")
                        message_count += 1
                        yield msg

                    page_num += 1

                    # Log progress periodically
                    if page_num % 1000 == 0:
                        logger.debug(f"Scanned {page_num} pages, found {message_count} potential messages")

        except Exception as e:
            logger.error(f"Error in raw scan: {e}")

    def _scan_page_for_emails(self, page: bytes, page_num: int) -> list[Message]:
        """Scan a single page for email-like content."""
        messages = []

        # Look for common email patterns
        # This is heuristic and may produce false positives

        # Look for Subject: or From: patterns
        patterns = [
            b'Subject:',
            b'From:',
            b'To:',
            b'Date:',
            b'Message-ID:',
        ]

        for pattern in patterns:
            if pattern in page:
                # Found potential email content
                try:
                    # Extract surrounding context
                    idx = page.find(pattern)
                    # Get a chunk of text around the pattern
                    start = max(0, idx - 100)
                    end = min(len(page), idx + 500)
                    chunk = page[start:end]

                    # Try to decode and parse
                    text = self._decode_safe(chunk)
                    if text:
                        # Create a minimal message
                        msg = Message(message_id=f"raw_{page_num}_{idx}")
                        msg.body_text = text
                        messages.append(msg)
                        break  # One message per page max

                except Exception:
                    pass

        return messages

    def _decode_safe(self, data) -> str:
        """Safely decode data to string."""
        if data is None:
            return ""
        if isinstance(data, str):
            return data
        if isinstance(data, bytes):
            for encoding in ['utf-16-le', 'utf-8', 'latin-1']:
                try:
                    return data.decode(encoding).rstrip('\x00')
                except UnicodeDecodeError:
                    continue
            return ""
        return str(data)
