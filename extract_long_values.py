#!/usr/bin/env python3
"""
Extract Long Value data from ESE database.
Uses pyesedb's get_value_data_as_long_value method for LongBinary columns.
"""

import sys
import struct
import os

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


def extract_filename(prop_blob):
    """Extract filename from PropertyBlob."""
    if not prop_blob:
        return "unknown"

    # Look for filename extensions
    for ext in [b'.xml', b'.txt', b'.doc', b'.pdf', b'.jpg', b'.png', b'.gif', b'.xlsx', b'.pptx']:
        ext_lower = ext.lower()
        blob_lower = prop_blob.lower()
        if ext_lower in blob_lower:
            idx = blob_lower.find(ext_lower)
            start = idx
            while start > 0 and 0x20 <= prop_blob[start-1] < 0x7f:
                start -= 1
            return prop_blob[start:idx+len(ext)].decode('ascii', errors='ignore')
    return "unknown"


def extract_long_values(db_path, mailbox_num, output_dir):
    import pyesedb

    print(f"Opening: {db_path}")
    db = pyesedb.file()
    db.open(db_path)

    tables = {}
    for i in range(db.get_number_of_tables()):
        t = db.get_table(i)
        if t:
            tables[t.name] = t

    attach_table = tables.get(f"Attachment_{mailbox_num}")
    if not attach_table:
        print(f"Attachment_{mailbox_num} not found!")
        return

    col_map = get_column_map(attach_table)
    content_idx = col_map.get('Content', -1)
    size_idx = col_map.get('Size', -1)
    prop_idx = col_map.get('PropertyBlob', -1)
    name_idx = col_map.get('Name', -1)

    os.makedirs(output_dir, exist_ok=True)

    print(f"\n{'='*70}")
    print("EXTRACTING LONG VALUE ATTACHMENTS")
    print(f"{'='*70}")

    for i in range(attach_table.get_number_of_records()):
        record = attach_table.get_record(i)
        if not record:
            continue

        content_raw = get_bytes(record, content_idx)
        prop_blob = get_bytes(record, prop_idx)
        size_data = get_bytes(record, size_idx)

        size_val = struct.unpack('<Q', size_data)[0] if size_data and len(size_data) == 8 else 0
        filename = extract_filename(prop_blob)

        # Check if this is a 4-byte reference (Long Value pointer)
        if content_raw and len(content_raw) == 4:
            ref_val = struct.unpack('<I', content_raw)[0]

            print(f"\nAttachment {i}: {filename}")
            print(f"  Reference: {ref_val} (0x{ref_val:08x})")
            print(f"  Expected size: {size_val} bytes")
            print(f"  is_long_value: {record.is_long_value(content_idx)}")

            # Try to get long value data
            try:
                # This is the key - get_value_data_as_long_value
                lv_data = record.get_value_data_as_long_value(content_idx)
                if lv_data:
                    print(f"  get_value_data_as_long_value type: {type(lv_data)}")

                    # pyesedb long value object has methods
                    if hasattr(lv_data, 'get_data'):
                        actual_data = lv_data.get_data()
                        print(f"  LV get_data: {len(actual_data) if actual_data else 0} bytes")

                        if actual_data and len(actual_data) > 0:
                            # Save the data
                            output_file = os.path.join(output_dir, f"attachment_{i}_{filename}")
                            with open(output_file, 'wb') as f:
                                f.write(actual_data)
                            print(f"  SAVED: {output_file} ({len(actual_data)} bytes)")

                            # Preview
                            if len(actual_data) <= 200:
                                print(f"  Content: {actual_data[:100]}")
                            else:
                                print(f"  Preview: {actual_data[:100]}...")

                    elif hasattr(lv_data, 'data'):
                        print(f"  LV data attr: {len(lv_data.data) if lv_data.data else 0} bytes")

                    else:
                        print(f"  LV methods: {[m for m in dir(lv_data) if not m.startswith('_')]}")

                        # Try to read chunks if available
                        if hasattr(lv_data, 'get_number_of_segments'):
                            num_seg = lv_data.get_number_of_segments()
                            print(f"  LV segments: {num_seg}")

                        if hasattr(lv_data, 'read'):
                            chunk = lv_data.read()
                            print(f"  LV read: {len(chunk) if chunk else 0} bytes")

                else:
                    print(f"  get_value_data_as_long_value returned None")

            except Exception as e:
                print(f"  Error getting long value: {e}")
                import traceback
                traceback.print_exc()

            # Also check SeparatedProperty columns
            for sep_num in range(1, 11):
                sep_col = f"SeparatedProperty{sep_num:02d}"
                if sep_col in col_map:
                    sep_data = get_bytes(record, col_map[sep_col])
                    if sep_data and len(sep_data) > 0:
                        print(f"  {sep_col}: {len(sep_data)} bytes")
                        if len(sep_data) == 4:
                            sep_ref = struct.unpack('<I', sep_data)[0]
                            print(f"    -> Reference: {sep_ref}")

                            # Try long value for separated property
                            try:
                                sep_lv = record.get_value_data_as_long_value(col_map[sep_col])
                                if sep_lv and hasattr(sep_lv, 'get_data'):
                                    sep_actual = sep_lv.get_data()
                                    print(f"    -> LV data: {len(sep_actual) if sep_actual else 0} bytes")
                            except:
                                pass
                        elif len(sep_data) > 100:
                            print(f"    -> Actual data, preview: {sep_data[:50]}...")

        elif content_raw and len(content_raw) > 4:
            # Inline content
            print(f"\nAttachment {i}: {filename}")
            print(f"  Inline content: {len(content_raw)} bytes")

    # Also check messages for body content
    print(f"\n{'='*70}")
    print("CHECKING MESSAGE BODY CONTENT")
    print(f"{'='*70}")

    msg_table = tables.get(f"Message_{mailbox_num}")
    if msg_table:
        msg_col_map = get_column_map(msg_table)
        body_cols = ['NativeBody', 'OffPagePropertyBlob', 'LargePropertyValueBlob']

        for msg_i in range(min(5, msg_table.get_number_of_records())):
            rec = msg_table.get_record(msg_i)
            if not rec:
                continue
            print(f"\nMessage {msg_i}:")

            for col_name in body_cols:
                if col_name not in msg_col_map:
                    continue
                col_idx = msg_col_map[col_name]
                raw_data = get_bytes(rec, col_idx)

                if raw_data and len(raw_data) == 4:
                    ref = struct.unpack('<I', raw_data)[0]
                    print(f"  {col_name}: Reference {ref}")

                    try:
                        lv = rec.get_value_data_as_long_value(col_idx)
                        if lv and hasattr(lv, 'get_data'):
                            actual = lv.get_data()
                            print(f"    -> LV data: {len(actual) if actual else 0} bytes")
                    except:
                        pass
                elif raw_data and len(raw_data) > 10:
                    print(f"  {col_name}: {len(raw_data)} bytes inline")

    db.close()
    print(f"\n{'='*70}")
    print("EXTRACTION COMPLETE")
    print(f"{'='*70}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_long_values.py <edb_file> [mailbox_num] [output_dir]")
        sys.exit(1)

    db_path = sys.argv[1]
    mailbox_num = int(sys.argv[2]) if len(sys.argv) > 2 else 103
    output_dir = sys.argv[3] if len(sys.argv) > 3 else "./extracted_attachments"

    extract_long_values(db_path, mailbox_num, output_dir)
