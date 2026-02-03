#!/usr/bin/env python3
"""
Exchange Folder ID Mapping
Generated from Exchange Get-MailboxFolderStatistics
Mailbox: administrator@lab.sith.uz
"""

# Folder ID Short (last 20 chars) to folder path
# Format: CA00000000XXXX00000Y where XXXX is folder number, Y is type
FOLDER_ID_TO_PATH = {
    # Standard folders (system)
    '00000000010800000100': '/Top of Information Store',  # 0108 = Root/IPM Subtree
    '00000000010900000100': '/Sent Items',                # 0109
    '00000000010a00000100': '/Deleted Items',             # 010A
    '00000000010b00000100': '/Outbox',                    # 010B
    '00000000010c00000100': '/Inbox',                     # 010C
    '00000000010d00000200': '/Calendar',                  # 010D
    '00000000010e00000300': '/Contacts',                  # 010E
    '00000000010f00000100': '/Drafts',                    # 010F
    '00000000011000000600': '/Journal',                   # 0110
    '00000000011100000500': '/Notes',                     # 0111
    '00000000011200000400': '/Tasks',                     # 0112
    '00000000011300000100': '/Recoverable Items',         # 0113
    '00000000011500000100': '/Deletions',                 # 0115
    '00000000011600000100': '/Versions',                  # 0116
    '00000000011700000100': '/Purges',                    # 0117
    '00000000011800000100': '/Calendar Logging',          # 0118

    # User-created and special folders
    '000024da113c00000300': '/ExternalContacts',
    '000024da114100000100': '/Junk Email',
    '000024da114200000100': '/Conversation Action Settings',
    '000024da114300000300': '/Contacts/Recipient Cache',
    '000024da114400000300': '/Contacts/{06967759-274D-40B2-A3EB-D7F9E73727D7}',
    '000024da114500000300': '/Contacts/{A9E2BC46-B3A0-4243-B315-60D991004455}',
    '000024da114800000300': '/Contacts/GAL Contacts',
    '000024da115d00000100': '/Files',
    '000024da116500000100': '/Yammer Root',
    '000024da116600000100': '/Yammer Root/Inbound',
    '000024da116700000100': '/Yammer Root/Outbound',
    '000024da116800000100': '/Yammer Root/Feeds',
    '000024da116e00000200': '/Calendar/Birthdays',
    '000024da116f00000300': '/Contacts/PeopleCentricConversation Buddies',
    '000024da117000000300': '/Contacts/Organizational Contacts',
    '000024da117100000300': '/Contacts/Companies',

    # User-created folders
    '00017cf45ba800000100': '/Test Folder',
    '00017cf45ba900000100': '/Test Folder/Test Subfolder',
}

# Map folder number (4 hex chars) to name for quick lookup
FOLDER_NUM_TO_NAME = {
    '0100': 'Root',
    '0101': 'Top of Information Store',
    '0106': 'Deleted Items',
    '0108': 'IPM Subtree',
    '0109': 'Sent Items',
    '010a': 'Outbox',
    '010b': 'Outbox',
    '010c': 'Inbox',
    '010d': 'Calendar',
    '010e': 'Contacts',
    '010f': 'Drafts',
    '0110': 'Journal',
    '0111': 'Notes',
    '0112': 'Tasks',
    '0113': 'Recoverable Items',
    '0114': 'System/Hidden',
    '0115': 'Deletions',
    '0116': 'Versions',
    '0117': 'Purges',
    '0118': 'Calendar Logging',
    '1141': 'Junk Email',
    '1142': 'Conversation Actions',
    '1143': 'Recipient Cache',
    '5ba8': 'Test Folder',
    '5ba9': 'Test Subfolder',
}

# Folder types
FOLDER_TYPES = {
    '0108': 'Root',
    '0109': 'SentItems',
    '010a': 'DeletedItems',
    '010b': 'Outbox',
    '010c': 'Inbox',
    '010d': 'Calendar',
    '010e': 'Contacts',
    '010f': 'Drafts',
    '0110': 'Journal',
    '0111': 'Notes',
    '0112': 'Tasks',
    '0113': 'RecoverableItemsRoot',
    '0115': 'RecoverableItemsDeletions',
    '0116': 'RecoverableItemsVersions',
    '0117': 'RecoverableItemsPurges',
    '0118': 'CalendarLogging',
}

# Item counts from Exchange
FOLDER_ITEM_COUNTS = {
    '/Inbox': 11,
    '/Sent Items': 3,
    '/Calendar': 3,
    '/Contacts/Recipient Cache': 2,
    '/Test Folder': 1,
    '/Test Folder/Test Subfolder': 1,
}

# Special folder number mapping (from SpecialFolderNumber column in EDB)
# Note: These may differ from the folder ID numbers
SPECIAL_FOLDER_MAP = {
    1: 'Root',
    2: 'Spooler Queue',
    3: 'Shortcuts',
    4: 'Finder',
    5: 'Views',
    6: 'Common Views',
    7: 'Schedule',
    8: 'Junk Email',
    9: 'IPM Subtree',
    10: 'Inbox',
    11: 'Outbox',
    12: 'Sent Items',
    13: 'Deleted Items',
    14: 'Contacts',
    15: 'Calendar',
    16: 'Drafts',
    17: 'Journal',
    18: 'Notes',
    19: 'Tasks',
    20: 'Recoverable Items',
    21: 'Deletions',
}


def get_folder_name(folder_id_hex, special_folder_num=None):
    """
    Get folder name from folder ID hex string or special folder number.

    Args:
        folder_id_hex: Full hex string of FolderId from database
        special_folder_num: SpecialFolderNumber value (if available)

    Returns:
        Folder name string
    """
    # Try special folder number first
    if special_folder_num is not None and special_folder_num in SPECIAL_FOLDER_MAP:
        return SPECIAL_FOLDER_MAP[special_folder_num]

    if not folder_id_hex:
        return None

    folder_id_hex = folder_id_hex.lower()

    # Try exact match on last 20 chars
    short_id = folder_id_hex[-20:] if len(folder_id_hex) >= 20 else folder_id_hex
    if short_id in FOLDER_ID_TO_PATH:
        path = FOLDER_ID_TO_PATH[short_id]
        # Return just the folder name, not full path
        return path.rstrip('/').split('/')[-1]

    # Try matching folder number (chars 8-12 of the short ID)
    # Short ID format: PPPPPPPPXXXXTTTTTTTT where P=prefix, X=folder num, T=type
    if len(short_id) >= 12:
        folder_num = short_id[8:12]
        if folder_num in FOLDER_NUM_TO_NAME:
            return FOLDER_NUM_TO_NAME[folder_num]

    return None


def get_folder_path(folder_id_hex):
    """Get full folder path from folder ID."""
    if not folder_id_hex:
        return None

    folder_id_hex = folder_id_hex.lower()
    short_id = folder_id_hex[-20:] if len(folder_id_hex) >= 20 else folder_id_hex

    return FOLDER_ID_TO_PATH.get(short_id)


def get_folder_type(folder_id_hex):
    """Get folder type from folder ID."""
    if not folder_id_hex:
        return None

    folder_id_hex = folder_id_hex.lower()
    short_id = folder_id_hex[-20:] if len(folder_id_hex) >= 20 else folder_id_hex

    if len(short_id) >= 12:
        folder_num = short_id[8:12]
        return FOLDER_TYPES.get(folder_num)

    return None


# Print mapping for debugging
if __name__ == '__main__':
    print("=" * 70)
    print("Exchange Folder ID Mapping")
    print("=" * 70)
    print()

    print("Folder ID to Path mapping:")
    print("-" * 70)
    for fid, path in sorted(FOLDER_ID_TO_PATH.items(), key=lambda x: x[1]):
        items = FOLDER_ITEM_COUNTS.get(path, 0)
        item_str = f" [{items} items]" if items > 0 else ""
        # Extract folder number
        folder_num = fid[8:12] if len(fid) >= 12 else "????"
        print(f"  {folder_num}: {path}{item_str}")
        print(f"         ID: {fid}")

    print()
    print("Folder Number to Name (quick lookup):")
    print("-" * 70)
    for num, name in sorted(FOLDER_NUM_TO_NAME.items()):
        print(f"  {num}: {name}")

    print()
    print("Special Folder Numbers:")
    print("-" * 70)
    for num, name in sorted(SPECIAL_FOLDER_MAP.items()):
        print(f"  {num:2d}: {name}")
