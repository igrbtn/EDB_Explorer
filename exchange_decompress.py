#!/usr/bin/env python3
"""
Exchange NativeBody decompression.
Exchange uses a variant of LZ compression for HTML body content.
"""

import struct


def decompress_exchange_html(data):
    """
    Decompress Exchange NativeBody HTML content.

    Format:
    - Header: 7 bytes
      - Byte 0: 0x18 or 0x19 (compression marker)
      - Bytes 1-2: Uncompressed size (little-endian)
      - Bytes 3-6: Flags/reserved
    - Compressed data follows

    Compression uses LZ-style back-references:
    - Literal bytes are copied as-is (printable ASCII)
    - Control sequences indicate back-references
    """
    if not data or len(data) < 7:
        return data

    # Check header
    if data[0] not in [0x18, 0x19]:
        # Not compressed, return as-is
        return data

    # Parse header
    uncompressed_size = struct.unpack('<H', data[1:3])[0]

    # Skip header
    compressed = data[7:]

    # Decompress using sliding window
    output = bytearray()
    window = bytearray(4096)
    window_pos = 0
    i = 0

    while i < len(compressed) and len(output) < uncompressed_size:
        b = compressed[i]

        # Check for control byte patterns
        if b == 0x00 and i + 1 < len(compressed):
            next_b = compressed[i + 1]
            if next_b == 0x00:
                # Double null = literal null
                output.append(0x00)
                window[window_pos % 4096] = 0x00
                window_pos += 1
                i += 2
            elif next_b < 0x10:
                # Short back-reference
                # Skip control sequence
                i += 2
            else:
                # Might be literal, copy next byte
                i += 1
        elif 0x01 <= b <= 0x1f:
            # Control byte - indicates back-reference
            if i + 1 < len(compressed):
                length = (b & 0x0f) + 3  # Length is encoded
                offset_byte = compressed[i + 1]

                if offset_byte < len(output):
                    # Copy from output buffer
                    start = len(output) - offset_byte - 1
                    for _ in range(min(length, uncompressed_size - len(output))):
                        if start >= 0 and start < len(output):
                            c = output[start]
                            output.append(c)
                            window[window_pos % 4096] = c
                            window_pos += 1
                            start += 1
                i += 2
            else:
                i += 1
        elif 0x80 <= b <= 0xff:
            # High bit set - might be back-reference or literal
            # Check if it's part of UTF-8 or control
            if i + 1 < len(compressed):
                next_b = compressed[i + 1]
                # Try to interpret as back-reference
                length = ((b & 0x70) >> 4) + 3
                offset = ((b & 0x0f) << 8) | next_b

                if offset < len(output) and length > 0:
                    start = len(output) - offset - 1
                    for _ in range(min(length, uncompressed_size - len(output))):
                        if 0 <= start < len(output):
                            c = output[start]
                            output.append(c)
                            window[window_pos % 4096] = c
                            window_pos += 1
                            start += 1
                    i += 2
                else:
                    i += 1
            else:
                i += 1
        else:
            # Literal byte
            output.append(b)
            window[window_pos % 4096] = b
            window_pos += 1
            i += 1

    return bytes(output)


def extract_html_text(html_bytes):
    """Extract visible text from HTML bytes."""
    try:
        html = html_bytes.decode('utf-8', errors='ignore')
    except:
        html = html_bytes.decode('latin-1', errors='ignore')

    # Remove script and style content
    import re
    html = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<!--.*?-->', '', html, flags=re.DOTALL)

    # Extract text between tags
    text_parts = []
    in_tag = False
    current = []

    for c in html:
        if c == '<':
            if current:
                text = ''.join(current).strip()
                if text:
                    text_parts.append(text)
                current = []
            in_tag = True
        elif c == '>':
            in_tag = False
        elif not in_tag:
            current.append(c)

    if current:
        text = ''.join(current).strip()
        if text:
            text_parts.append(text)

    return '\n'.join(text_parts)


def decompress_and_extract(data):
    """Decompress Exchange body and extract text."""
    if not data:
        return ""

    # Try decompression
    decompressed = decompress_exchange_html(data)

    # If decompression didn't produce valid HTML, try raw decode
    if not decompressed or b'<html' not in decompressed.lower():
        decompressed = data[7:] if len(data) > 7 and data[0] in [0x18, 0x19] else data

    # Extract text
    return extract_html_text(decompressed)


if __name__ == '__main__':
    # Test with message 314
    import pyesedb

    db = pyesedb.file()
    db.open('/Users/igorbatin/Documents/VaibPro/MDB_exporter/Mailbox Database 0058949847 3/Mailbox Database 0058949847.edb')

    tables = {}
    for i in range(db.get_number_of_tables()):
        t = db.get_table(i)
        if t:
            tables[t.name] = t

    msg_table = tables['Message_103']
    col_map = {}
    for j in range(msg_table.get_number_of_columns()):
        col = msg_table.get_column(j)
        if col:
            col_map[col.name] = j

    rec = msg_table.get_record(314)
    native_idx = col_map.get('NativeBody', -1)

    if rec.is_long_value(native_idx):
        lv = rec.get_value_data_as_long_value(native_idx)
        if lv:
            data = lv.get_data()
            print(f"Original data: {len(data)} bytes")
            print(f"Header: {data[:7].hex()}")

            decompressed = decompress_exchange_html(data)
            print(f"\nDecompressed: {len(decompressed)} bytes")
            print(decompressed[:500])

            text = extract_html_text(decompressed)
            print(f"\nExtracted text:")
            print(text)

    db.close()
