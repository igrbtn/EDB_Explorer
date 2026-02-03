#!/usr/bin/env python3
"""
Comprehensive mailbox analysis - find all objects and understand folder mapping
"""

import sys
import struct
import re
from collections import defaultdict
from datetime import datetime, timezone

# Import folder mapping
try:
    from folder_mapping import get_folder_name, FOLDER_NUM_TO_NAME, FOLDER_ID_TO_PATH
    HAS_MAPPING = True
except ImportError:
    HAS_MAPPING = False
    FOLDER_NUM_TO_NAME = {}


def get_column_map(table):
    col_map = {}
    for j in range(table.get_number_of_columns()):
        col = table.get_column(j)
        if col:
            col_map[col.name] = j
    return col_map


def get_bytes(record, col_idx):
    if col_idx < 0:
        return None
    try:
        return record.get_value_data(col_idx)
    except:
        return None


def get_int(data):
    if not data:
        return None
    if len(data) == 1:
        return data[0]
    elif len(data) == 2:
        return struct.unpack('<H', data)[0]
    elif len(data) == 4:
        return struct.unpack('<I', data)[0]
    elif len(data) == 8:
        return struct.unpack('<Q', data)[0]
    return None


def get_datetime(data):
    if not data or len(data) != 8:
        return None
    try:
        ft = struct.unpack('<Q', data)[0]
        if ft == 0:
            return None
        unix = (ft - 116444736000000000) / 10000000
        return datetime.fromtimestamp(unix, tz=timezone.utc)
    except:
        return None


def extract_subject(blob):
    if not blob or len(blob) < 10:
        return None
    for i in range(len(blob) - 5):
        if blob[i] == 0x4d:  # M marker
            length = blob[i+1]
            if 2 <= length <= 100 and i + 2 + length <= len(blob):
                potential = blob[i+2:i+2+length]
                if all(32 <= b < 127 for b in potential):
                    text = potential.decode('ascii')
                    skip = ['admin', 'exchange', 'recipient', 'fydib', 'pdlt', 'ipm.']
                    if not any(x in text.lower() for x in skip):
                        return text
    return None


def get_folder_display_name(folder_id_hex):
    """Get folder name from folder ID."""
    if not folder_id_hex:
        return "Unknown"

    # Try folder mapping module
    if HAS_MAPPING:
        name = get_folder_name(folder_id_hex)
        if name:
            return name

    # Extract folder number from last 20 chars
    short_id = folder_id_hex[-20:] if len(folder_id_hex) >= 20 else folder_id_hex
    if len(short_id) >= 12:
        folder_num = short_id[8:12]
        if folder_num in FOLDER_NUM_TO_NAME:
            return FOLDER_NUM_TO_NAME[folder_num]

    return f"Folder_{short_id[-8:]}"


def analyze_mailbox(db_path, mailbox_num):
    import pyesedb

    print("=" * 80)
    print(f"MAILBOX ANALYSIS: {db_path}")
    print(f"Mailbox Number: {mailbox_num}")
    print("=" * 80)

    db = pyesedb.file()
    db.open(db_path)

    tables = {}
    for i in range(db.get_number_of_tables()):
        t = db.get_table(i)
        if t:
            tables[t.name] = t

    # =========================================================================
    # ANALYZE MESSAGE TABLE
    # =========================================================================
    msg_table = tables.get(f"Message_{mailbox_num}")
    if not msg_table:
        print(f"ERROR: Message_{mailbox_num} not found")
        return

    col_map = get_column_map(msg_table)
    total_records = msg_table.get_number_of_records()

    print(f"\n{'='*80}")
    print(f"MESSAGE TABLE: Message_{mailbox_num}")
    print(f"{'='*80}")
    print(f"Total Records: {total_records}")
    print(f"Columns: {len(col_map)}")
    print()

    # Categorize all messages
    messages_by_folder = defaultdict(list)
    hidden_messages = []
    visible_messages = []
    messages_with_attachments = []

    print("Scanning all records...")
    print("-" * 80)

    for i in range(total_records):
        record = msg_table.get_record(i)
        if not record:
            continue

        # Get key fields
        folder_id = get_bytes(record, col_map.get('FolderId', -1))
        is_hidden = get_bytes(record, col_map.get('IsHidden', -1))
        is_read = get_bytes(record, col_map.get('IsRead', -1))
        has_attach = get_bytes(record, col_map.get('HasAttachments', -1))
        msg_class = get_bytes(record, col_map.get('MessageClass', -1))
        date_received = get_datetime(get_bytes(record, col_map.get('DateReceived', -1)))
        date_sent = get_datetime(get_bytes(record, col_map.get('DateSent', -1)))
        size = get_int(get_bytes(record, col_map.get('Size', -1)))
        prop_blob = get_bytes(record, col_map.get('PropertyBlob', -1))

        # Parse values
        folder_hex = folder_id.hex() if folder_id else ""
        folder_name = get_folder_display_name(folder_hex)
        is_hidden_val = bool(is_hidden and is_hidden != b'\x00')
        is_read_val = bool(is_read and is_read != b'\x00')
        has_attach_val = bool(has_attach and has_attach != b'\x00')

        # Get message class
        msg_class_str = ""
        if msg_class:
            try:
                msg_class_str = msg_class.decode('utf-8').rstrip('\x00')
            except:
                try:
                    msg_class_str = msg_class.decode('utf-16-le').rstrip('\x00')
                except:
                    msg_class_str = msg_class[:20].hex()

        # Get subject
        subject = extract_subject(prop_blob) if prop_blob else ""

        msg_info = {
            'record': i,
            'folder_id': folder_hex,
            'folder_name': folder_name,
            'is_hidden': is_hidden_val,
            'is_read': is_read_val,
            'has_attach': has_attach_val,
            'msg_class': msg_class_str,
            'date_received': date_received,
            'date_sent': date_sent,
            'size': size or 0,
            'subject': subject,
        }

        messages_by_folder[folder_name].append(msg_info)

        if is_hidden_val:
            hidden_messages.append(msg_info)
        else:
            visible_messages.append(msg_info)

        if has_attach_val:
            messages_with_attachments.append(msg_info)

    # =========================================================================
    # SUMMARY BY FOLDER
    # =========================================================================
    print(f"\n{'='*80}")
    print("MESSAGES BY FOLDER")
    print(f"{'='*80}")

    for folder_name in sorted(messages_by_folder.keys()):
        msgs = messages_by_folder[folder_name]
        visible = [m for m in msgs if not m['is_hidden']]
        hidden = [m for m in msgs if m['is_hidden']]

        print(f"\n{folder_name}:")
        print(f"  Total: {len(msgs)} | Visible: {len(visible)} | Hidden: {len(hidden)}")

        if visible:
            print(f"  Visible messages:")
            for m in visible[:10]:  # Show first 10
                date_str = m['date_received'].strftime('%Y-%m-%d %H:%M') if m['date_received'] else 'N/A'
                attach_str = " [ATTACH]" if m['has_attach'] else ""
                subj = m['subject'][:40] if m['subject'] else m['msg_class'][:40]
                print(f"    #{m['record']:4d}: {date_str} | {subj}{attach_str}")
            if len(visible) > 10:
                print(f"    ... and {len(visible) - 10} more")

    # =========================================================================
    # INBOX ANALYSIS
    # =========================================================================
    print(f"\n{'='*80}")
    print("INBOX DETAILED ANALYSIS")
    print(f"{'='*80}")

    inbox_msgs = messages_by_folder.get('Inbox', [])
    print(f"Total Inbox records: {len(inbox_msgs)}")
    print(f"Visible: {len([m for m in inbox_msgs if not m['is_hidden']])}")
    print(f"Hidden: {len([m for m in inbox_msgs if m['is_hidden']])}")

    print("\nAll Inbox records:")
    for m in inbox_msgs:
        hidden_str = "[HIDDEN]" if m['is_hidden'] else "[VISIBLE]"
        date_str = m['date_received'].strftime('%Y-%m-%d %H:%M') if m['date_received'] else 'N/A'
        attach_str = " [ATTACH]" if m['has_attach'] else ""
        subj = m['subject'][:50] if m['subject'] else f"({m['msg_class'][:30]})"
        print(f"  #{m['record']:4d} {hidden_str}: {date_str} | {subj}{attach_str}")

    # =========================================================================
    # CHECK FOR MESSAGES IN UNEXPECTED PLACES
    # =========================================================================
    print(f"\n{'='*80}")
    print("MESSAGE CLASS DISTRIBUTION")
    print(f"{'='*80}")

    class_counts = defaultdict(int)
    for folder_msgs in messages_by_folder.values():
        for m in folder_msgs:
            class_counts[m['msg_class'] or 'Unknown'] += 1

    for cls, count in sorted(class_counts.items(), key=lambda x: -x[1]):
        print(f"  {cls}: {count}")

    # =========================================================================
    # HIDDEN ITEMS ANALYSIS
    # =========================================================================
    print(f"\n{'='*80}")
    print("HIDDEN ITEMS ANALYSIS")
    print(f"{'='*80}")
    print(f"Total hidden items: {len(hidden_messages)}")

    hidden_by_folder = defaultdict(list)
    for m in hidden_messages:
        hidden_by_folder[m['folder_name']].append(m)

    for folder, msgs in sorted(hidden_by_folder.items()):
        print(f"\n  {folder}: {len(msgs)} hidden items")
        for m in msgs[:5]:
            cls = m['msg_class'][:40] if m['msg_class'] else 'Unknown'
            print(f"    #{m['record']}: {cls}")

    # =========================================================================
    # FOLDER ID MAPPING CHECK
    # =========================================================================
    print(f"\n{'='*80}")
    print("FOLDER ID MAPPING")
    print(f"{'='*80}")

    unique_folder_ids = set()
    for folder_msgs in messages_by_folder.values():
        for m in folder_msgs:
            if m['folder_id']:
                unique_folder_ids.add(m['folder_id'])

    print(f"Unique folder IDs found: {len(unique_folder_ids)}")
    print()

    for fid in sorted(unique_folder_ids):
        short_id = fid[-20:]
        folder_num = short_id[8:12] if len(short_id) >= 12 else "????"
        name = get_folder_display_name(fid)
        count = sum(1 for m in visible_messages if m['folder_id'] == fid)
        hidden_count = sum(1 for m in hidden_messages if m['folder_id'] == fid)
        print(f"  {folder_num} -> {name}: {count} visible, {hidden_count} hidden")

    # =========================================================================
    # ATTACHMENTS CHECK
    # =========================================================================
    print(f"\n{'='*80}")
    print("ATTACHMENTS ANALYSIS")
    print(f"{'='*80}")

    attach_table = tables.get(f"Attachment_{mailbox_num}")
    if attach_table:
        attach_col_map = get_column_map(attach_table)
        print(f"Attachment table records: {attach_table.get_number_of_records()}")
        print(f"Messages with HasAttachments=True: {len(messages_with_attachments)}")

        print("\nMessages with attachments:")
        for m in messages_with_attachments:
            print(f"  #{m['record']} in {m['folder_name']}: {m['subject'] or m['msg_class']}")

    db.close()

    print(f"\n{'='*80}")
    print("ANALYSIS COMPLETE")
    print(f"{'='*80}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python analyze_mailbox.py <edb_file> [mailbox_num]")
        sys.exit(1)

    db_path = sys.argv[1]
    mailbox_num = int(sys.argv[2]) if len(sys.argv) > 2 else 103

    analyze_mailbox(db_path, mailbox_num)
