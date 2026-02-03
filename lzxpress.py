#!/usr/bin/env python3
"""
LZXPRESS Plain LZ77 Decompression

Based on Microsoft MS-XCA specification and MagnetForensics/rust-lzxpress.

The algorithm processes compressed data using:
- 32-bit flag words where each bit indicates literal (0) or match (1)
- Literal bytes are copied directly
- Match references encode offset and length in 16-bit values
- Extended length encoding for matches >= 7 bytes
"""

import struct


def decompress_lzxpress(data: bytes, max_output_size: int = 0) -> bytes:
    """
    Decompress LZXPRESS Plain LZ77 compressed data.

    Args:
        data: Compressed input bytes
        max_output_size: Maximum output size (0 = no limit)

    Returns:
        Decompressed bytes
    """
    if not data:
        return b''

    output = bytearray()
    in_idx = 0
    data_len = len(data)

    while in_idx < data_len:
        # Read 32-bit flag word (little-endian)
        if in_idx + 4 > data_len:
            break

        flags = struct.unpack('<I', data[in_idx:in_idx + 4])[0]
        in_idx += 4

        # Process 32 bits
        for bit_pos in range(32):
            if in_idx >= data_len:
                break

            if max_output_size > 0 and len(output) >= max_output_size:
                return bytes(output[:max_output_size])

            # Check flag bit (LSB first)
            flag_bit = (flags >> bit_pos) & 1

            if flag_bit == 0:
                # Literal byte - copy directly
                output.append(data[in_idx])
                in_idx += 1
            else:
                # Match reference - read 16-bit metadata
                if in_idx + 2 > data_len:
                    break

                match_data = struct.unpack('<H', data[in_idx:in_idx + 2])[0]
                in_idx += 2

                # Decode offset and length
                # Offset = (match_data / 8) + 1 = (match_data >> 3) + 1
                # Length base = match_data % 8 = match_data & 7
                offset = (match_data >> 3) + 1
                length = match_data & 7

                # Extended length encoding
                if length == 7:
                    # Read additional length byte
                    if in_idx >= data_len:
                        break
                    length += data[in_idx]
                    in_idx += 1

                    if length == 7 + 255:
                        # Read 16-bit length
                        if in_idx + 2 > data_len:
                            break
                        length = struct.unpack('<H', data[in_idx:in_idx + 2])[0]
                        in_idx += 2

                        if length == 0:
                            # Read 32-bit length
                            if in_idx + 4 > data_len:
                                break
                            length = struct.unpack('<I', data[in_idx:in_idx + 4])[0]
                            in_idx += 4

                # Add 3 to get actual length (minimum match is 3)
                length += 3

                # Validate offset
                if offset > len(output):
                    # Invalid offset - skip or return what we have
                    continue

                # Copy match from output buffer
                # Note: source and destination may overlap
                match_start = len(output) - offset
                for i in range(length):
                    if max_output_size > 0 and len(output) >= max_output_size:
                        break
                    # Read from current position (may have just been written)
                    output.append(output[match_start + i])

    return bytes(output)


def decompress_exchange_body(data: bytes) -> bytes:
    """
    Decompress Exchange NativeBody which uses LZXPRESS with a custom header.

    Exchange header format (7 bytes):
    - Byte 0: 0x18 or 0x19 (compression type marker)
    - Bytes 1-2: Uncompressed size (little-endian 16-bit)
    - Bytes 3-6: Flags/reserved

    Args:
        data: Raw NativeBody data from Exchange

    Returns:
        Decompressed HTML content
    """
    if not data or len(data) < 7:
        return data

    # Check for Exchange compression header
    if data[0] not in [0x18, 0x19, 0x9a]:
        # Not compressed or different format
        return data

    # Parse header
    uncompressed_size = struct.unpack('<H', data[1:3])[0]

    # Skip 7-byte header
    compressed_data = data[7:]

    # Try LZXPRESS decompression
    try:
        result = decompress_lzxpress(compressed_data, uncompressed_size)
        if result and len(result) > 0:
            return result
    except Exception as e:
        pass

    # Fallback: return data without header
    return compressed_data


def extract_text_from_html(html_bytes: bytes) -> str:
    """
    Extract visible text from HTML bytes.

    Args:
        html_bytes: HTML content (possibly with artifacts)

    Returns:
        Extracted text content
    """
    import re

    try:
        html = html_bytes.decode('utf-8', errors='ignore')
    except:
        html = html_bytes.decode('latin-1', errors='ignore')

    # Remove script and style
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
                if text and len(text) >= 2:
                    text_parts.append(text)
                current = []
            in_tag = True
        elif c == '>':
            in_tag = False
        elif not in_tag:
            if c.isprintable() or c in '\r\n\t':
                current.append(c)

    if current:
        text = ''.join(current).strip()
        if text:
            text_parts.append(text)

    return '\n'.join(text_parts)


if __name__ == '__main__':
    # Test with Exchange data
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

    # Test messages
    for rec_idx in [308, 309, 314]:
        rec = msg_table.get_record(rec_idx)
        native_idx = col_map.get('NativeBody', -1)

        if rec.is_long_value(native_idx):
            lv = rec.get_value_data_as_long_value(native_idx)
            if lv:
                raw_data = lv.get_data()
                print(f"\n{'='*60}")
                print(f"Message {rec_idx}")
                print(f"{'='*60}")
                print(f"Raw data: {len(raw_data)} bytes")
                print(f"Header: {raw_data[:7].hex()}")

                decompressed = decompress_exchange_body(raw_data)
                print(f"Decompressed: {len(decompressed)} bytes")
                print(f"Content preview: {decompressed[:200]}")

                text = extract_text_from_html(decompressed)
                print(f"Extracted text: {text[:200]}")

    db.close()
