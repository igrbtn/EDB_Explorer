"""Core components for EDB reading and parsing."""

from .message import Message, Attachment, Folder
from .edb_reader import EDBReader
from .mailbox_parser import MailboxParser

__all__ = ["Message", "Attachment", "Folder", "EDBReader", "MailboxParser"]
