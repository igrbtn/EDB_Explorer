"""Email message data models."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional
from enum import Enum


class MessagePriority(Enum):
    LOW = 0
    NORMAL = 1
    HIGH = 2


@dataclass
class Attachment:
    """Email attachment data."""
    filename: str
    content_type: str
    data: bytes
    size: int = 0
    content_id: Optional[str] = None

    def __post_init__(self):
        if self.size == 0:
            self.size = len(self.data)


@dataclass
class Folder:
    """Mailbox folder."""
    folder_id: str
    name: str
    parent_id: Optional[str] = None
    message_count: int = 0
    unread_count: int = 0

    @property
    def is_root(self) -> bool:
        return self.parent_id is None


@dataclass
class Message:
    """Email message data structure."""
    message_id: str
    subject: str = ""
    sender: str = ""
    sender_name: str = ""
    recipients_to: list[str] = field(default_factory=list)
    recipients_cc: list[str] = field(default_factory=list)
    recipients_bcc: list[str] = field(default_factory=list)
    date_sent: Optional[datetime] = None
    date_received: Optional[datetime] = None
    body_text: str = ""
    body_html: str = ""
    headers: dict[str, str] = field(default_factory=dict)
    attachments: list[Attachment] = field(default_factory=list)
    folder_id: Optional[str] = None
    folder_path: str = ""
    priority: MessagePriority = MessagePriority.NORMAL
    is_read: bool = False
    has_attachments: bool = False
    internet_message_id: str = ""
    in_reply_to: str = ""
    references: list[str] = field(default_factory=list)

    # Recovery metadata
    is_recovered: bool = False
    recovery_errors: list[str] = field(default_factory=list)

    def __post_init__(self):
        self.has_attachments = len(self.attachments) > 0

    @property
    def recipients_all(self) -> list[str]:
        """Get all recipients combined."""
        return self.recipients_to + self.recipients_cc + self.recipients_bcc

    @property
    def display_date(self) -> Optional[datetime]:
        """Get the best available date for display."""
        return self.date_received or self.date_sent

    def to_dict(self) -> dict:
        """Convert message to dictionary for serialization."""
        return {
            "message_id": self.message_id,
            "subject": self.subject,
            "sender": self.sender,
            "sender_name": self.sender_name,
            "recipients_to": self.recipients_to,
            "recipients_cc": self.recipients_cc,
            "recipients_bcc": self.recipients_bcc,
            "date_sent": self.date_sent.isoformat() if self.date_sent else None,
            "date_received": self.date_received.isoformat() if self.date_received else None,
            "body_text": self.body_text,
            "body_html": self.body_html,
            "folder_path": self.folder_path,
            "has_attachments": self.has_attachments,
            "attachment_count": len(self.attachments),
        }
