"""Base exporter class for email export formats."""

import logging
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Callable, Optional, Iterator
from dataclasses import dataclass

from ..core.message import Message

logger = logging.getLogger(__name__)


@dataclass
class ExportProgress:
    """Progress information for export operation."""
    total_messages: int
    exported_messages: int
    failed_messages: int
    current_message: Optional[str] = None
    is_complete: bool = False
    error: Optional[str] = None

    @property
    def progress_percent(self) -> float:
        if self.total_messages == 0:
            return 0.0
        return (self.exported_messages + self.failed_messages) / self.total_messages * 100


@dataclass
class ExportOptions:
    """Options for export operation."""
    output_path: str
    include_attachments: bool = True
    folder_structure: bool = True  # Preserve folder hierarchy
    overwrite_existing: bool = False
    max_messages: int = 0  # 0 = unlimited


class BaseExporter(ABC):
    """
    Abstract base class for email exporters.

    Subclasses implement specific format writers (EML, MBOX, PST).
    """

    format_name: str = "Unknown"
    file_extension: str = ""

    def __init__(self, options: ExportOptions):
        """
        Initialize exporter.

        Args:
            options: Export configuration options
        """
        self.options = options
        self._progress = ExportProgress(0, 0, 0)
        self._cancel_requested = False
        self._progress_callback: Optional[Callable[[ExportProgress], None]] = None

    def set_progress_callback(self, callback: Callable[[ExportProgress], None]):
        """Set callback function for progress updates."""
        self._progress_callback = callback

    def cancel(self):
        """Request cancellation of export."""
        self._cancel_requested = True
        logger.info("Export cancellation requested")

    def _update_progress(self, message: Optional[str] = None, success: bool = True):
        """Update progress and notify callback."""
        if success:
            self._progress.exported_messages += 1
        else:
            self._progress.failed_messages += 1

        self._progress.current_message = message

        if self._progress_callback:
            self._progress_callback(self._progress)

    def export(self, messages: Iterator[Message], total_count: int = 0) -> ExportProgress:
        """
        Export messages to the configured format.

        Args:
            messages: Iterator of Message objects
            total_count: Total number of messages (for progress reporting)

        Returns:
            ExportProgress with final status
        """
        self._progress = ExportProgress(total_count, 0, 0)
        self._cancel_requested = False

        output_path = Path(self.options.output_path)

        try:
            self._prepare_output(output_path)

            for message in messages:
                if self._cancel_requested:
                    logger.info("Export cancelled by user")
                    break

                if self.options.max_messages > 0 and \
                   self._progress.exported_messages >= self.options.max_messages:
                    logger.info(f"Reached max message limit: {self.options.max_messages}")
                    break

                try:
                    self._export_message(message, output_path)
                    self._update_progress(message.subject, success=True)
                except Exception as e:
                    logger.warning(f"Failed to export message {message.message_id}: {e}")
                    self._update_progress(message.subject, success=False)

            self._finalize_output(output_path)

        except Exception as e:
            logger.error(f"Export failed: {e}")
            self._progress.error = str(e)

        self._progress.is_complete = True
        return self._progress

    def _prepare_output(self, output_path: Path):
        """
        Prepare output location.

        Override in subclasses if needed.
        """
        if output_path.is_dir():
            output_path.mkdir(parents=True, exist_ok=True)
        else:
            output_path.parent.mkdir(parents=True, exist_ok=True)

    def _finalize_output(self, output_path: Path):
        """
        Finalize output after export.

        Override in subclasses if needed.
        """
        pass

    @abstractmethod
    def _export_message(self, message: Message, output_path: Path):
        """
        Export a single message.

        Args:
            message: Message to export
            output_path: Output directory or file path
        """
        pass

    @classmethod
    def get_format_info(cls) -> dict:
        """Get information about this export format."""
        return {
            "name": cls.format_name,
            "extension": cls.file_extension,
            "description": cls.__doc__ or ""
        }
