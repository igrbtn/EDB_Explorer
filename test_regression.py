#!/usr/bin/env python3
"""Test regression fix for AAAA BBBB CCCC extraction."""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import pyesedb
from lzxpress import get_body_preview, extract_body_from_property_blob, _looks_like_repeat_pattern

DB_PATH = '/Users/igorbatin/Documents/VaibPro/MDB_exporter/NewDB/NewDB_ABCD_New/NewDB.edb'

db = pyesedb.file()
db.open(DB_PATH)

tables = {}
for i in range(db.get_number_of_tables()):
    t = db.get_table(i)
    if t:
        tables[t.name] = t

msg_table = tables['Message_101']

col_map = {}
for j in range(msg_table.get_number_of_columns()):
    col = msg_table.get_column(j)
    if col:
        col_map[col.name] = j

# Test cases
test_cases = [
    (293, "Lorem ipsum", ['Lorem', 'ipsum', 'dolor', 'sit', 'amet', 'consectetur',
                          'adipiscing', 'elit', 'sed', 'do', 'eiusmod', 'tempor',
                          'incididunt', 'ut', 'labore']),
    (302, "AAAA BBBB CCCC", ['AAAA', 'BBBB', 'CCCC']),
    (308, "1111 2222 3333 4444", ['1111', '2222', '3333', '4444']),
]

print("=" * 70)
print("REGRESSION TEST - BODY EXTRACTION")
print("=" * 70)

for rec_idx, desc, expected_words in test_cases:
    record = msg_table.get_record(rec_idx)

    print(f"\nRecord {rec_idx}: {desc}")
    print("-" * 50)

    # Get PropertyBlob
    property_blob = None
    pb_idx = col_map.get('PropertyBlob', -1)
    if pb_idx >= 0:
        property_blob = record.get_value_data(pb_idx)

    # Get NativeBody
    native_data = None
    native_idx = col_map.get('NativeBody', -1)
    if native_idx >= 0 and record.is_long_value(native_idx):
        lv = record.get_value_data_as_long_value(native_idx)
        if lv:
            native_data = lv.get_data()

    # Test extract_body_from_property_blob directly
    if property_blob:
        pb_text = extract_body_from_property_blob(property_blob)
        is_repeat = _looks_like_repeat_pattern(pb_text)
        print(f"  PropertyBlob direct: '{pb_text[:60]}...' (repeat={is_repeat})")

    # Test combined extraction
    text = get_body_preview(native_data, 500, property_blob)
    print(f"  Combined result: '{text[:80]}...'")

    # Check word matches
    found = [w for w in expected_words if w in text]
    missing = [w for w in expected_words if w not in text]

    status = "PASS" if len(found) == len(expected_words) else "PARTIAL" if found else "FAIL"
    print(f"  Words: {len(found)}/{len(expected_words)} [{status}]")
    if missing:
        print(f"  Missing: {missing}")

db.close()
