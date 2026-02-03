#!/usr/bin/env python3
"""
Exchange EDB Email Exporter - Console Mode

Usage:
    python cli.py <edb_file> [options]

Options:
    --format eml|mbox       Export format (default: eml)
    --output <path>         Output directory/file
    --mailbox <number>      Export specific mailbox only
    --list-mailboxes        List mailboxes and exit
    --list-tables           List all tables and exit
    --info                  Show database info and exit
    --limit <n>             Limit number of messages
    --verbose               Verbose output
"""

import sys
import os
import argparse
import logging
from pathlib import Path
from datetime import datetime

# Add src to path
sys.path.insert(0, str(Path(__file__).parent))

from src.core.exchange_parser import Exchange2013Parser, MailboxInfo, detect_encryption
from src.core.message import Message
from src.exporters.base import ExportOptions
from src.exporters.eml_exporter import EMLExporter
from src.exporters.mbox_exporter import MBOXExporter


def setup_logging(verbose: bool):
    """Configure logging."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(levelname)s: %(message)s'
    )


def open_database(edb_path: str):
    """Open EDB database."""
    try:
        import pyesedb
    except ImportError:
        print("ERROR: pyesedb not installed. Run: pip install libesedb-python")
        sys.exit(1)

    if not os.path.exists(edb_path):
        print(f"ERROR: File not found: {edb_path}")
        sys.exit(1)

    print(f"Opening: {edb_path}")

    db = pyesedb.file()
    db.open(edb_path)

    return db


def show_info(db, parser: Exchange2013Parser):
    """Show database information."""
    print("\n=== DATABASE INFO ===")
    print(f"Tables: {len(parser._tables)}")

    # Check encryption
    is_encrypted, enc_msg = detect_encryption(db)
    if is_encrypted:
        print(f"\n*** WARNING: {enc_msg} ***")
        print("    Subject and body content cannot be extracted from encrypted databases.")
        print("    Only metadata (dates, flags, email addresses) can be exported.")

    # Get message counts
    msg_tables = parser.get_message_tables()
    print(f"\nMessage tables: {len(msg_tables)}")
    print(f"Total messages: {parser.get_total_message_count()}")

    # Show mailboxes
    mailboxes = parser.get_mailboxes()
    print(f"\nMailboxes: {len(mailboxes)}")

    for mb in mailboxes[:10]:
        print(f"  Mailbox {mb.mailbox_number}: {mb.message_count} messages")

    if len(mailboxes) > 10:
        print(f"  ... and {len(mailboxes) - 10} more")

    if is_encrypted:
        print("\n=== EXTRACTABLE DATA ===")
        print("  - Received/Sent dates")
        print("  - Read status, attachment flags")
        print("  - Email addresses (from PropertyBlob)")
        print("  - Message size")
        print("\n=== NOT EXTRACTABLE (ENCRYPTED) ===")
        print("  - Subject")
        print("  - Body content")
        print("  - Attachment content")


def list_mailboxes(parser: Exchange2013Parser):
    """List all mailboxes."""
    print("\n=== MAILBOXES ===")

    mailboxes = parser.get_mailboxes()

    if not mailboxes:
        print("No mailboxes found")
        return

    print(f"{'#':<5} {'Messages':<10} {'GUID':<40}")
    print("-" * 60)

    for mb in mailboxes:
        print(f"{mb.mailbox_number:<5} {mb.message_count:<10} {mb.mailbox_guid[:36]}")

    # Also show which Message_* tables exist
    print(f"\nMessage tables: {', '.join(parser.get_message_tables()[:10])}...")


def list_tables(parser: Exchange2013Parser):
    """List all tables with record counts."""
    print("\n=== TABLES ===")
    print(f"{'Table Name':<40} {'Records':<10}")
    print("-" * 55)

    for name, table in sorted(parser._tables.items()):
        try:
            count = table.get_number_of_records()
            if count > 0:
                print(f"{name:<40} {count:<10}")
        except:
            pass


def export_messages(parser: Exchange2013Parser, args):
    """Export messages to specified format."""
    output_path = args.output or f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    if args.format == 'mbox':
        if not output_path.endswith('.mbox'):
            output_path += '.mbox'

    print(f"\nExporting to: {output_path}")
    print(f"Format: {args.format.upper()}")

    options = ExportOptions(
        output_path=output_path,
        include_attachments=True,
        folder_structure=True,
        overwrite_existing=True,
        max_messages=args.limit if args.limit else 0
    )

    if args.format == 'eml':
        exporter = EMLExporter(options)
        Path(output_path).mkdir(parents=True, exist_ok=True)
    else:
        exporter = MBOXExporter(options)

    # Get messages
    mailbox_num = args.mailbox if args.mailbox is not None else None

    total = 0
    exported = 0
    failed = 0

    print("\nExporting messages...")

    for message in parser.iter_messages(mailbox_num):
        total += 1

        try:
            if args.format == 'eml':
                exporter._export_message(message, Path(output_path))
            else:
                exporter._export_message(message, Path(output_path))
            exported += 1

            if exported % 100 == 0:
                print(f"  Exported {exported} messages...")

        except Exception as e:
            failed += 1
            if args.verbose:
                print(f"  Failed: {message.message_id} - {e}")

        if args.limit and exported >= args.limit:
            print(f"  Reached limit of {args.limit} messages")
            break

    # Finalize
    exporter._finalize_output(Path(output_path))

    print(f"\n=== EXPORT COMPLETE ===")
    print(f"Total processed: {total}")
    print(f"Exported: {exported}")
    print(f"Failed: {failed}")
    print(f"Output: {output_path}")


def scan_messages(parser: Exchange2013Parser, args):
    """Scan and display message summaries."""
    print("\n=== SCANNING MESSAGES ===")

    mailbox_num = args.mailbox if args.mailbox is not None else None
    limit = args.limit or 20

    count = 0
    for message in parser.iter_messages(mailbox_num):
        if count >= limit:
            break

        print(f"\n--- Message {count + 1} ---")
        print(f"ID: {message.message_id}")
        print(f"Subject: {message.subject or '(none)'}")
        print(f"From: {message.sender or '(none)'}")
        print(f"To: {', '.join(message.recipients_to) if message.recipients_to else '(none)'}")
        print(f"Date: {message.date_received or message.date_sent or '(none)'}")
        print(f"Read: {message.is_read}, Attachments: {message.has_attachments}")

        if message.body_text:
            preview = message.body_text[:200].replace('\n', ' ')
            print(f"Body: {preview}...")

        count += 1

    print(f"\nShowed {count} messages (use --limit to see more)")


def main():
    parser = argparse.ArgumentParser(
        description="Exchange EDB Email Exporter - Console Mode",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python cli.py database.edb --info
    python cli.py database.edb --list-mailboxes
    python cli.py database.edb --format eml --output ./exported
    python cli.py database.edb --format mbox --output emails.mbox --limit 100
    python cli.py database.edb --mailbox 101 --format eml --output ./mailbox_101
        """
    )

    parser.add_argument('edb_file', help='Path to EDB database file')
    parser.add_argument('--format', choices=['eml', 'mbox'], default='eml',
                        help='Export format (default: eml)')
    parser.add_argument('--output', '-o', help='Output path')
    parser.add_argument('--mailbox', '-m', type=int, help='Mailbox number to export')
    parser.add_argument('--list-mailboxes', action='store_true',
                        help='List mailboxes and exit')
    parser.add_argument('--list-tables', action='store_true',
                        help='List all tables and exit')
    parser.add_argument('--info', action='store_true',
                        help='Show database info and exit')
    parser.add_argument('--scan', action='store_true',
                        help='Scan and show message summaries')
    parser.add_argument('--limit', '-n', type=int,
                        help='Limit number of messages')
    parser.add_argument('--verbose', '-v', action='store_true',
                        help='Verbose output')

    args = parser.parse_args()

    setup_logging(args.verbose)

    # Open database
    db = open_database(args.edb_file)
    exchange_parser = Exchange2013Parser(db)

    try:
        if args.info:
            show_info(db, exchange_parser)
        elif args.list_mailboxes:
            list_mailboxes(exchange_parser)
        elif args.list_tables:
            list_tables(exchange_parser)
        elif args.scan:
            scan_messages(exchange_parser, args)
        elif args.output or args.format:
            export_messages(exchange_parser, args)
        else:
            # Default: show info
            show_info(db, exchange_parser)
            print("\nUse --help for export options")

    finally:
        db.close()

    return 0


if __name__ == '__main__':
    sys.exit(main())
