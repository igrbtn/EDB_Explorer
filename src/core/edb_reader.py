"""ESE database reader for Exchange EDB files."""

import logging
from pathlib import Path
from typing import Iterator, Optional
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class EDBInfo:
    """Information about an EDB file."""
    file_path: str
    file_size: int
    is_valid: bool
    is_dirty: bool = False
    version: str = ""
    table_count: int = 0
    error_message: str = ""


class EDBReader:
    """
    Reader for Microsoft Exchange EDB (ESE) database files.

    Uses pyesedb library to read the Extensible Storage Engine format.
    """

    # ESE file signature
    ESE_SIGNATURE = b'\xef\xcd\xab\x89'

    # Common Exchange table names
    EXCHANGE_TABLES = {
        "Message": "Contains email messages",
        "Attachment": "Contains message attachments",
        "Folder": "Contains folder structure",
        "Mailbox": "Contains mailbox information",
        "GlobalsTable": "Global database settings",
    }

    def __init__(self, file_path: str):
        """
        Initialize EDB reader.

        Args:
            file_path: Path to the EDB file
        """
        self.file_path = Path(file_path)
        self._db = None
        self._is_open = False
        self._tables: dict[str, any] = {}

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def get_info(self) -> EDBInfo:
        """
        Get information about the EDB file without fully opening it.

        Returns:
            EDBInfo with file details
        """
        if not self.file_path.exists():
            return EDBInfo(
                file_path=str(self.file_path),
                file_size=0,
                is_valid=False,
                error_message="File not found"
            )

        file_size = self.file_path.stat().st_size

        # Check file signature
        try:
            with open(self.file_path, 'rb') as f:
                signature = f.read(4)
                is_valid = signature == self.ESE_SIGNATURE

                if not is_valid:
                    return EDBInfo(
                        file_path=str(self.file_path),
                        file_size=file_size,
                        is_valid=False,
                        error_message="Invalid ESE file signature"
                    )

        except Exception as e:
            return EDBInfo(
                file_path=str(self.file_path),
                file_size=file_size,
                is_valid=False,
                error_message=str(e)
            )

        return EDBInfo(
            file_path=str(self.file_path),
            file_size=file_size,
            is_valid=True
        )

    def open(self) -> bool:
        """
        Open the EDB database for reading.

        Returns:
            True if successful, False otherwise
        """
        if self._is_open:
            return True

        try:
            import pyesedb

            self._db = pyesedb.file()
            self._db.open(str(self.file_path))
            self._is_open = True

            # Cache table references
            self._cache_tables()

            logger.info(f"Opened EDB file: {self.file_path}")
            return True

        except ImportError:
            logger.error("pyesedb library not installed. Run: pip install pyesedb")
            raise ImportError("pyesedb library required. Install with: pip install pyesedb")

        except Exception as e:
            logger.error(f"Failed to open EDB file: {e}")
            self._is_open = False
            raise

    def close(self):
        """Close the database connection."""
        if self._db and self._is_open:
            try:
                self._db.close()
            except Exception as e:
                logger.warning(f"Error closing database: {e}")
            finally:
                self._db = None
                self._is_open = False
                self._tables.clear()

    def _cache_tables(self):
        """Cache references to database tables."""
        if not self._is_open:
            return

        try:
            num_tables = self._db.get_number_of_tables()
            for i in range(num_tables):
                table = self._db.get_table(i)
                if table:
                    self._tables[table.name] = table

            logger.debug(f"Cached {len(self._tables)} tables")

        except Exception as e:
            logger.error(f"Error caching tables: {e}")

    def get_table_names(self) -> list[str]:
        """Get list of all table names in the database."""
        if not self._is_open:
            raise RuntimeError("Database not open")
        return list(self._tables.keys())

    def get_table(self, name: str):
        """
        Get a specific table by name.

        Args:
            name: Table name

        Returns:
            Table object or None if not found
        """
        return self._tables.get(name)

    def iter_table_records(self, table_name: str) -> Iterator[dict]:
        """
        Iterate over records in a table.

        Args:
            table_name: Name of the table

        Yields:
            Dictionary with column names and values for each record
        """
        table = self.get_table(table_name)
        if not table:
            logger.warning(f"Table not found: {table_name}")
            return

        try:
            num_columns = table.get_number_of_columns()
            columns = []
            for i in range(num_columns):
                col = table.get_column(i)
                columns.append(col.name if col else f"column_{i}")

            num_records = table.get_number_of_records()

            for i in range(num_records):
                record = table.get_record(i)
                if record:
                    row = {}
                    for j, col_name in enumerate(columns):
                        try:
                            value = record.get_value_data(j)
                            row[col_name] = value
                        except Exception:
                            row[col_name] = None
                    yield row

        except Exception as e:
            logger.error(f"Error iterating table {table_name}: {e}")
            raise

    def get_record_count(self, table_name: str) -> int:
        """Get the number of records in a table."""
        table = self.get_table(table_name)
        if not table:
            return 0
        try:
            return table.get_number_of_records()
        except Exception:
            return 0

    @property
    def is_open(self) -> bool:
        """Check if database is open."""
        return self._is_open

    @property
    def table_count(self) -> int:
        """Get number of tables."""
        return len(self._tables)
