#!/usr/bin/env python3
"""
vba_project_builder.py — Create vbaProject.bin from VBA source code.

Pure Python, stdlib only. Supports Cyrillic (cp1251).
Implements MS-OVBA compression + MS-CFB v3 compound file format.

Usage as module:
    from vba_project_builder import build_vba_project
    modules = [
        ("Module1", "standard", "Sub Test()\\nMsgBox \"Hi\"\\nEnd Sub"),
        ("ЭтаКнига", "document", ""),
    ]
    build_vba_project(modules, "vbaProject.bin", code_page=1251)

Usage as CLI:
    python3 vba_project_builder.py output.bin Module1.bas [ThisWorkbook.cls ...]
"""

import math
import random
import struct
import os
import sys
import uuid

# ===========================================================================
# Constants
# ===========================================================================

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF = 0x1000  # 4096
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
FREESECT = 0xFFFFFFFF
NOSTREAM = 0xFFFFFFFF

# OLE header signature
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

# Default VBA version (VBA7)
VBA_MAJOR_VERSION = 0x6BC68EBC
VBA_MINOR_VERSION = 0


# ===========================================================================
# MS-OVBA Compression (Section 2.4.1)
# ===========================================================================

def ovba_compress(data: bytes) -> bytes:
    """Compress data using MS-OVBA compression algorithm."""
    result = bytearray([0x01])  # signature byte
    offset = 0
    while offset < len(data):
        chunk_end = min(offset + 4096, len(data))
        chunk = data[offset:chunk_end]
        result.extend(_compress_chunk(chunk))
        offset = chunk_end
    return bytes(result)


def _compress_chunk(chunk: bytes) -> bytes:
    """Compress a single chunk (up to 4096 bytes)."""
    compressed = bytearray()
    src = 0
    while src < len(chunk):
        flag_byte = 0
        tokens = bytearray()
        for i in range(8):
            if src >= len(chunk):
                break
            offset, length = _find_match(chunk, src)
            if length >= 3:
                token = _pack_copy_token(offset, length, src)
                tokens.extend(struct.pack("<H", token))
                flag_byte |= 1 << i
                src += length
            else:
                tokens.append(chunk[src])
                src += 1
        compressed.append(flag_byte)
        compressed.extend(tokens)

    if len(compressed) <= 4095:
        header = (len(compressed) - 1) | 0xB000
        return struct.pack("<H", header) + bytes(compressed)
    else:
        padded = chunk + b"\x00" * (4096 - len(chunk))
        header = 4095 | 0x3000
        return struct.pack("<H", header) + padded


def _find_match(chunk: bytes, pos: int) -> tuple:
    """Find longest backward match for copy token."""
    if pos == 0:
        return (0, 0)
    bit_count = max(4, (pos - 1).bit_length()) if pos > 1 else 4
    max_length = (0xFFFF >> bit_count) + 3
    best_off, best_len = 0, 0
    search_start = max(0, pos - (1 << bit_count))
    for cand in range(pos - 1, search_start - 1, -1):
        ln = 0
        while pos + ln < len(chunk) and ln < max_length and chunk[cand + ln] == chunk[pos + ln]:
            ln += 1
        if ln > best_len:
            best_len = ln
            best_off = pos - cand
            if ln >= max_length:
                break
    return (best_off, best_len) if best_len >= 3 else (0, 0)


def _pack_copy_token(offset: int, length: int, pos: int) -> int:
    """Pack offset and length into 16-bit CopyToken."""
    bit_count = max(4, (pos - 1).bit_length()) if pos > 1 else 4
    return (((offset - 1) << (16 - bit_count)) | (length - 3)) & 0xFFFF


# ===========================================================================
# dir stream builder (Section 2.3.4.2)
# ===========================================================================

def _record(rec_id: int, data: bytes) -> bytes:
    """Build a dir stream record: Id(u16) + Size(u32) + Data."""
    return struct.pack("<HI", rec_id, len(data)) + data


def _build_dir_stream(modules: list, code_page: int) -> bytes:
    """Build the dir stream (uncompressed) for the VBA project.

    modules: list of (name: str, mod_type: str, source: str)
        mod_type: "standard" | "document"
    """
    cp_name = f"cp{code_page}"
    buf = bytearray()

    # --- PROJECTINFORMATION ---
    # SysKind (Win64 = 3, matches modern Excel)
    buf += _record(0x0001, struct.pack("<I", 3))
    # Note: CompatVersion omitted — valid file doesn't include it
    # Lcid
    buf += _record(0x0002, struct.pack("<I", 0x0409))  # en-US (standard)
    # LcidInvoke
    buf += _record(0x0014, struct.pack("<I", 0x0409))
    # CodePage
    buf += _record(0x0003, struct.pack("<H", code_page))
    # Name
    name_bytes = "VBAProject".encode(cp_name)
    buf += _record(0x0004, name_bytes)
    # DocString + DocStringUnicode (empty)
    buf += _record(0x0005, b"")
    buf += _record(0x0040, b"")
    # HelpFile1 + HelpFile2 (empty)
    buf += _record(0x0006, b"")
    buf += _record(0x003D, b"")
    # HelpContext
    buf += _record(0x0007, struct.pack("<I", 0))
    # LibFlags
    buf += _record(0x0008, struct.pack("<I", 0))
    # Version
    buf += struct.pack("<HI", 0x0009, 4)  # Id + Reserved(=4)
    buf += struct.pack("<IH", VBA_MAJOR_VERSION, VBA_MINOR_VERSION)
    # Constants + ConstantsUnicode (empty)
    buf += _record(0x000C, b"")
    buf += _record(0x003C, b"")

    # --- PROJECTREFERENCES ---
    # Reference: stdole
    _add_reference(buf, "stdole",
                   "{00020430-0000-0000-C000-000000000046}",
                   "2.0", "C:\\Windows\\System32\\stdole2.tlb",
                   "OLE Automation", cp_name)
    # Reference: Office
    _add_reference(buf, "Office",
                   "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}",
                   "2.0",
                   "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL",
                   "Microsoft Office 16.0 Object Library", cp_name)

    # --- PROJECTMODULES ---
    buf += _record(0x000F, struct.pack("<H", len(modules)))
    buf += _record(0x0013, struct.pack("<H", 0xFFFF))  # cookie

    for name, mod_type, _source in modules:
        name_cp = name.encode(cp_name)
        name_u16 = name.encode("utf-16-le")
        # ModuleName
        buf += _record(0x0019, name_cp)
        # ModuleNameUnicode
        buf += _record(0x0047, name_u16)
        # ModuleStreamName + Unicode
        buf += _record(0x001A, name_cp)
        buf += _record(0x0032, name_u16)
        # ModuleDocString + Unicode (empty)
        buf += _record(0x001C, b"")
        buf += _record(0x0048, b"")
        # ModuleOffset (0 = no performance cache / p-code)
        buf += _record(0x0031, struct.pack("<I", 0))
        # ModuleHelpContext
        buf += _record(0x001E, struct.pack("<I", 0))
        # ModuleCookie
        buf += _record(0x002C, struct.pack("<H", 0xFFFF))
        # ModuleType
        if mod_type == "document":
            buf += struct.pack("<HI", 0x0022, 0)  # doc/cls/designer
        else:
            buf += struct.pack("<HI", 0x0021, 0)  # procedural
        # ModuleEnd
        buf += struct.pack("<HI", 0x002B, 0)

    # Terminator
    buf += struct.pack("<HI", 0x0010, 0)

    return bytes(buf)


def _add_reference(buf: bytearray, name: str, clsid: str, version: str,
                   path: str, description: str, cp_name: str) -> None:
    """Add a REFERENCE record to the dir stream buffer."""
    # ReferenceName
    name_cp = name.encode(cp_name)
    buf += _record(0x0016, name_cp)
    # ReferenceNameUnicode
    buf += _record(0x003E, name.encode("utf-16-le"))
    # ReferenceRegistered
    libid = f"*\\G{clsid}#{version}#0#{path}#{description}"
    libid_bytes = libid.encode(cp_name)
    # Body: LibidSize(u32) + Libid + Reserved1(u32=0) + Reserved2(u16=0)
    ref_body = struct.pack("<I", len(libid_bytes)) + libid_bytes
    ref_body += struct.pack("<IH", 0, 0)
    buf += _record(0x000D, ref_body)


# ===========================================================================
# MS-OVBA Data Encryption (Section 2.4.3)
# ===========================================================================

def _ovba_encrypt(clsid: str, data: bytes) -> bytes:
    """Encrypt data per MS-OVBA 2.4.3 Data Encryption.

    clsid: project GUID string "{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"
    data: plaintext bytes to encrypt
    Returns: encrypted bytes (seed + version_enc + proj_key_enc + ignored + length + data)
    """
    length = len(data)
    seed = random.randint(0, 255)

    version = 2
    version_enc = version ^ seed

    proj_key = 0
    for i in range(38):
        proj_key += ord(clsid[i])
    proj_key = proj_key & 255
    proj_key_enc = proj_key ^ seed

    unencrypted_byte_1 = proj_key
    encrypted_byte_1 = proj_key_enc
    encrypted_byte_2 = version_enc

    ignored_length = (seed & 6) // 2
    ignored_enc = b""
    for _ in range(ignored_length):
        temp_value = random.randint(0, 255)
        s = (unencrypted_byte_1 + encrypted_byte_2) & 255
        byte_enc = temp_value ^ s
        ignored_enc += byte_enc.to_bytes(1, "little")
        encrypted_byte_2 = encrypted_byte_1
        encrypted_byte_1 = byte_enc
        unencrypted_byte_1 = temp_value

    data_length_enc = b""
    length_bytes = length.to_bytes(4, "little")
    for i in range(4):
        byte = length_bytes[i]
        byte_enc = byte ^ ((unencrypted_byte_1 + encrypted_byte_2) & 255)
        data_length_enc += byte_enc.to_bytes(1, "little")
        encrypted_byte_2 = encrypted_byte_1
        encrypted_byte_1 = byte_enc
        unencrypted_byte_1 = byte

    data_enc = b""
    for i in range(len(data)):
        data_byte = data[i]
        s = (unencrypted_byte_1 + encrypted_byte_2) & 255
        byte_enc = data_byte ^ s
        data_enc += byte_enc.to_bytes(1, "little")
        encrypted_byte_2 = encrypted_byte_1
        encrypted_byte_1 = byte_enc
        unencrypted_byte_1 = data_byte

    output = struct.pack(
        "<BBB", seed, version_enc, proj_key_enc
    ) + ignored_enc + data_length_enc + data_enc
    return output


# ===========================================================================
# PROJECT stream builder (Section 2.3.1)
# ===========================================================================

def _build_project_stream(modules: list, project_guid: str) -> bytes:
    """Build the PROJECT stream (text, not compressed).

    Includes encrypted CMG/DPB/GC fields per MS-OVBA 2.4.3:
    - CMG: protection state (4 zero bytes = no protection)
    - DPB: password hash (1 zero byte = no password)
    - GC: visibility (0xFF = visible)
    """
    lines = []
    lines.append(f'ID="{project_guid}"')

    for name, mod_type, _source in modules:
        if mod_type == "document":
            lines.append(f"Document={name}/&H00000000")
        else:
            lines.append(f"Module={name}")

    lines.append('Name="VBAProject"')
    lines.append('HelpContextID="0"')
    lines.append('VersionCompatible32="393222000"')

    # Encrypted protection fields (MS-OVBA 2.4.3)
    cmg_bytes = _ovba_encrypt(project_guid, b"\x00\x00\x00\x00")
    dpb_bytes = _ovba_encrypt(project_guid, b"\x00")
    gc_bytes = _ovba_encrypt(project_guid, b"\xFF")
    lines.append(f"CMG=\"{cmg_bytes.hex().upper()}\"")
    lines.append(f"DPB=\"{dpb_bytes.hex().upper()}\"")
    lines.append(f"GC=\"{gc_bytes.hex().upper()}\"")

    lines.append("")
    lines.append("[Host Extender Info]")
    lines.append("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000")
    lines.append("")
    lines.append("[Workspace]")
    for name, _mod_type, _source in modules:
        lines.append(f"{name}=0, 0, 0, 0, C")

    return "\r\n".join(lines).encode("cp1251")


# ===========================================================================
# PROJECTwm stream builder (Section 2.3.3)
# ===========================================================================

def _build_projectwm_stream(modules: list, cp_name: str) -> bytes:
    """Build PROJECTwm stream: module name → unicode name mapping."""
    buf = bytearray()
    for name, _mod_type, _source in modules:
        buf += name.encode(cp_name) + b"\x00"
        buf += name.encode("utf-16-le") + b"\x00\x00"
    return bytes(buf)


# ===========================================================================
# _VBA_PROJECT stream builder (Section 2.3.4.1)
# ===========================================================================

def _build_vba_project_stream() -> bytes:
    """Build _VBA_PROJECT stream with dummy PerformanceCache.

    Excel requires _VBA_PROJECT to be >= 32 bytes even when
    there is no real p-code.  We emit a 5-byte header followed
    by 59 zero bytes (total 64 bytes) as a safe dummy cache.
    Version=0 triggers recompile from source on first open.
    """
    header = struct.pack("<HHB", 0x61CC, 0, 0x00)
    return header + b'\x00' * 59  # 64 bytes total


# ===========================================================================
# VBA Attribute header builder
# ===========================================================================

# CLSIDs for document module VB_Base attribute
_CLSID_WORKBOOK = "0{00020819-0000-0000-C000-000000000046}"
_CLSID_WORKSHEET = "0{00020820-0000-0000-C000-000000000046}"

# Names recognized as ThisWorkbook (EN + RU)
_WORKBOOK_NAMES = {"thisworkbook", "этакнига"}


def _build_attribute_header(name: str, mod_type: str) -> str:
    """Build VBA Attribute header for a module.

    Standard modules get just VB_Name.
    Document modules get full Attribute block with VB_Base CLSID.
    """
    if mod_type == "document":
        clsid = (_CLSID_WORKBOOK if name.lower() in _WORKBOOK_NAMES
                 else _CLSID_WORKSHEET)
        return (
            f'Attribute VB_Name = "{name}"\r\n'
            f'Attribute VB_Base = "{clsid}"\r\n'
            f'Attribute VB_GlobalNameSpace = False\r\n'
            f'Attribute VB_Creatable = False\r\n'
            f'Attribute VB_PredeclaredId = True\r\n'
            f'Attribute VB_Exposed = True\r\n'
            f'Attribute VB_TemplateDerived = False\r\n'
            f'Attribute VB_Customizable = True\r\n'
        )
    else:
        return f'Attribute VB_Name = "{name}"\r\n'


# ===========================================================================
# MS-CFB v3 Compound File Writer
# ===========================================================================

class _DirEntry:
    """A single OLE directory entry (128 bytes)."""

    def __init__(self, name: str, obj_type: int, data: bytes = b""):
        self.name = name
        self.obj_type = obj_type  # 1=storage, 2=stream, 5=root
        self.data = data
        self.color = 1  # black
        self.left_sibling = NOSTREAM
        self.right_sibling = NOSTREAM
        self.child = NOSTREAM
        self.start_sector = ENDOFCHAIN
        self.stream_size = len(data)
        self.clsid = b"\x00" * 16

    def pack(self) -> bytes:
        """Serialize to 128 bytes."""
        name_u16 = self.name.encode("utf-16-le")
        # name field is 64 bytes, null-padded, name_size includes null terminator
        name_buf = name_u16[:62] + b"\x00\x00"
        name_buf = name_buf.ljust(64, b"\x00")
        name_size = len(name_u16) + 2  # +2 for null terminator

        return struct.pack(
            "<64sHBBIII16sIQQIQ",
            name_buf,
            name_size,
            self.obj_type,
            self.color,
            self.left_sibling,
            self.right_sibling,
            self.child,
            self.clsid,
            0,  # state bits
            0,  # creation time
            0,  # modified time
            self.start_sector,
            self.stream_size,
        )


def _cfb_name_key(entry):
    """Sort key per MS-CFB spec: shorter names first, then uppercase comparison."""
    name_u16 = entry.name.encode("utf-16-le")
    upper_name = entry.name.upper()
    return (len(name_u16), upper_name)


def _build_balanced_tree(indices: list, entries: list) -> list:
    """Build a balanced BST with proper red-black coloring.

    indices: entry indices to organize as siblings (must be sorted by CFB name).
    entries: the full entry list (to read names for sorting).
    Returns list of (index, left, right) tuples.
    Also sets entry.color (0=Red, 1=Black).
    """
    if not indices:
        return []

    # Sort indices by CFB name comparison (shorter names first, then uppercase)
    sorted_indices = sorted(indices, key=lambda i: _cfb_name_key(entries[i]))

    # Red-black coloring: nodes at depth >= floor(log2(n+1)) are Red
    n = len(sorted_indices)
    min_nil_depth = int(math.log2(n + 1)) if n > 0 else 0

    def _build(arr, depth):
        if not arr:
            return None
        mid = len(arr) // 2
        node = arr[mid]
        # Assign color: Red if at or beyond min_nil_depth, else Black
        entries[node].color = 0 if depth >= min_nil_depth else 1
        left = _build(arr[:mid], depth + 1)
        right = _build(arr[mid + 1:], depth + 1)
        return (node, left, right)

    def _flatten(tree_node):
        if tree_node is None:
            return []
        idx, left, right = tree_node
        result = [(idx,
                    left[0] if left else NOSTREAM,
                    right[0] if right else NOSTREAM)]
        result.extend(_flatten(left))
        result.extend(_flatten(right))
        return result

    tree = _build(sorted_indices, 0)
    return _flatten(tree)


def _write_cfb(entries: list, output_path: str) -> None:
    """Write a MS-CFB v3 compound file from directory entries.

    entries[0] must be the Root Entry (type 5).
    Streams < 4096 bytes go into the mini-stream (via mini-FAT).
    Streams >= 4096 bytes go into regular sectors (via FAT).
    """
    # ── Phase 1: Classify streams ──
    mini_entries = []
    regular_entries = []
    for entry in entries:
        if entry.obj_type == 2 and entry.stream_size > 0:
            if entry.stream_size < MINI_STREAM_CUTOFF:
                mini_entries.append(entry)
            else:
                regular_entries.append(entry)

    # ── Phase 2: Build mini-stream + mini-FAT ──
    mini_stream = bytearray()
    mini_fat_entries = []

    for entry in mini_entries:
        start_mini = len(mini_stream) // MINI_SECTOR_SIZE
        entry.start_sector = start_mini

        data = entry.data
        padded_len = ((len(data) + MINI_SECTOR_SIZE - 1)
                      // MINI_SECTOR_SIZE) * MINI_SECTOR_SIZE
        mini_stream.extend(data)
        mini_stream.extend(b"\x00" * (padded_len - len(data)))

        num_mini = padded_len // MINI_SECTOR_SIZE
        for i in range(num_mini - 1):
            mini_fat_entries.append(start_mini + i + 1)
        mini_fat_entries.append(ENDOFCHAIN)

    entries[0].stream_size = len(mini_stream)

    # ── Phase 3: Plan sector layout ──
    num_dir_sectors = (len(entries) + 3) // 4

    # Mini-FAT bytes (pad entries to sector boundary with FREESECT)
    if mini_fat_entries:
        while len(mini_fat_entries) % 128 != 0:
            mini_fat_entries.append(FREESECT)
        mini_fat_bytes = b"".join(
            struct.pack("<I", v) for v in mini_fat_entries)
        num_mini_fat_sectors = len(mini_fat_bytes) // SECTOR_SIZE
    else:
        mini_fat_bytes = b""
        num_mini_fat_sectors = 0

    # Mini-stream bytes (pad to sector boundary)
    if mini_stream:
        mini_stream_padded = bytes(mini_stream)
        pad = (SECTOR_SIZE - len(mini_stream_padded) % SECTOR_SIZE) % SECTOR_SIZE
        mini_stream_padded += b"\x00" * pad
        num_mini_stream_sectors = len(mini_stream_padded) // SECTOR_SIZE
    else:
        mini_stream_padded = b""
        num_mini_stream_sectors = 0

    # Regular stream data (each padded to sector boundary)
    regular_stream_data = []
    for entry in regular_entries:
        data = entry.data
        pad = (SECTOR_SIZE - len(data) % SECTOR_SIZE) % SECTOR_SIZE
        regular_stream_data.append(data + b"\x00" * pad)
    num_regular_sectors = sum(
        len(d) // SECTOR_SIZE for d in regular_stream_data)

    # Iterative FAT sizing: need enough FAT sectors to index themselves + content
    content_sectors = (num_dir_sectors + num_mini_fat_sectors
                       + num_mini_stream_sectors + num_regular_sectors)
    num_fat_sectors = 1
    while num_fat_sectors * 128 < num_fat_sectors + content_sectors:
        num_fat_sectors += 1

    total_sectors = num_fat_sectors + content_sectors

    # Assign sector offsets
    dir_start = num_fat_sectors
    mini_fat_start = dir_start + num_dir_sectors
    mini_stream_start = mini_fat_start + num_mini_fat_sectors
    regular_start = mini_stream_start + num_mini_stream_sectors

    # Root Entry → mini-stream container location
    if num_mini_stream_sectors > 0:
        entries[0].start_sector = mini_stream_start
    else:
        entries[0].start_sector = ENDOFCHAIN

    # Regular stream entries → their sector offsets
    offset = regular_start
    for i, entry in enumerate(regular_entries):
        entry.start_sector = offset
        offset += len(regular_stream_data[i]) // SECTOR_SIZE

    # ── Phase 4: Build FAT ──
    fat = [FREESECT] * (num_fat_sectors * 128)

    # FAT sectors → FATSECT
    for i in range(num_fat_sectors):
        fat[i] = FATSECT

    # Directory chain
    for i in range(num_dir_sectors):
        s = dir_start + i
        fat[s] = (dir_start + i + 1
                  if i < num_dir_sectors - 1 else ENDOFCHAIN)

    # Mini-FAT chain
    for i in range(num_mini_fat_sectors):
        s = mini_fat_start + i
        fat[s] = (mini_fat_start + i + 1
                  if i < num_mini_fat_sectors - 1 else ENDOFCHAIN)

    # Mini-stream chain
    for i in range(num_mini_stream_sectors):
        s = mini_stream_start + i
        fat[s] = (mini_stream_start + i + 1
                  if i < num_mini_stream_sectors - 1 else ENDOFCHAIN)

    # Regular stream chains
    offset = regular_start
    for i in range(len(regular_entries)):
        num_s = len(regular_stream_data[i]) // SECTOR_SIZE
        for j in range(num_s):
            s = offset + j
            fat[s] = offset + j + 1 if j < num_s - 1 else ENDOFCHAIN
        offset += num_s

    # ── Phase 5: Header (512 bytes) ──
    difat = [FREESECT] * 109
    for i in range(min(num_fat_sectors, 109)):
        difat[i] = i

    header = bytearray(512)
    struct.pack_into("<8s", header, 0x00, OLE_SIGNATURE)
    struct.pack_into("<16s", header, 0x08, b"\x00" * 16)  # CLSID
    struct.pack_into("<H", header, 0x18, 0x003E)  # minor version
    struct.pack_into("<H", header, 0x1A, 0x0003)  # major version (v3)
    struct.pack_into("<H", header, 0x1C, 0xFFFE)  # byte order (little-endian)
    struct.pack_into("<H", header, 0x1E, 0x0009)  # sector shift (512)
    struct.pack_into("<H", header, 0x20, 0x0006)  # mini sector shift (64)
    struct.pack_into("<6s", header, 0x22, b"\x00" * 6)  # reserved
    struct.pack_into("<I", header, 0x28, 0)  # num dir sectors (0 for v3)
    struct.pack_into("<I", header, 0x2C, num_fat_sectors)
    struct.pack_into("<I", header, 0x30, dir_start)  # first dir sector
    struct.pack_into("<I", header, 0x34, 0)  # transaction sig
    struct.pack_into("<I", header, 0x38, MINI_STREAM_CUTOFF)
    struct.pack_into("<I", header, 0x3C,
                     mini_fat_start if num_mini_fat_sectors > 0
                     else ENDOFCHAIN)
    struct.pack_into("<I", header, 0x40, num_mini_fat_sectors)
    struct.pack_into("<I", header, 0x44, ENDOFCHAIN)  # first DIFAT sector
    struct.pack_into("<I", header, 0x48, 0)  # num DIFAT sectors

    for i, val in enumerate(difat):
        struct.pack_into("<I", header, 0x4C + i * 4, val)

    # ── Phase 6: Write file ──
    with open(output_path, "wb") as f:
        f.write(header)

        # FAT sectors
        fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
        f.write(fat_bytes)

        # Directory sectors
        dir_data = bytearray()
        for entry in entries:
            dir_data += entry.pack()
        # Pad with empty entries to fill complete sectors
        while len(dir_data) % SECTOR_SIZE != 0:
            empty = bytearray(128)
            struct.pack_into("<I", empty, 0x44, NOSTREAM)   # left sibling
            struct.pack_into("<I", empty, 0x48, NOSTREAM)   # right sibling
            struct.pack_into("<I", empty, 0x4C, NOSTREAM)   # child
            dir_data += empty
        f.write(dir_data)

        # Mini-FAT sectors
        if mini_fat_bytes:
            f.write(mini_fat_bytes)

        # Mini-stream sectors
        if mini_stream_padded:
            f.write(mini_stream_padded)

        # Regular stream sectors
        for data in regular_stream_data:
            f.write(data)

    # Verify file size matches expected layout
    file_size = os.path.getsize(output_path)
    expected = (1 + total_sectors) * SECTOR_SIZE
    assert file_size == expected, (
        f"CFB size mismatch: {file_size} != {expected}")


# ===========================================================================
# VBA Project Assembly
# ===========================================================================

def build_vba_project(
    modules: list,
    output_path: str,
    code_page: int = 1251,
) -> None:
    """Build a complete vbaProject.bin file.

    Args:
        modules: list of (name, mod_type, source_code) tuples.
            name: module name (supports Cyrillic)
            mod_type: "standard" or "document"
            source_code: VBA source as string
        output_path: path for the output vbaProject.bin
        code_page: Windows code page (1251 for Cyrillic)
    """
    cp_name = f"cp{code_page}"
    project_guid = "{" + str(uuid.uuid4()).upper() + "}"

    # Build streams
    dir_raw = _build_dir_stream(modules, code_page)
    dir_compressed = ovba_compress(dir_raw)

    project_data = _build_project_stream(modules, project_guid)
    projectwm_data = _build_projectwm_stream(modules, cp_name)
    vba_project_data = _build_vba_project_stream()

    # Build module streams (each: compressed source with Attribute headers)
    module_streams = {}
    for name, mod_type, source in modules:
        attr_header = _build_attribute_header(name, mod_type)
        full_source = attr_header + (source if source else "")
        source_bytes = full_source.encode(cp_name)
        compressed = ovba_compress(source_bytes)
        module_streams[name] = compressed

    # --- Assemble OLE directory ---
    # Entry 0: Root Entry
    root = _DirEntry("Root Entry", 5)
    root.color = 0  # Red, as Excel expects

    # Root-level children: PROJECT, PROJECTwm, VBA
    e_project = _DirEntry("PROJECT", 2, project_data)
    e_projectwm = _DirEntry("PROJECTwm", 2, projectwm_data)
    e_vba = _DirEntry("VBA", 1)

    # VBA-level children: _VBA_PROJECT, dir, module streams
    e_vba_proj = _DirEntry("_VBA_PROJECT", 2, vba_project_data)
    e_dir = _DirEntry("dir", 2, dir_compressed)

    module_entries = []
    for name, _mod_type, _source in modules:
        e = _DirEntry(name, 2, module_streams[name])
        module_entries.append(e)

    # All entries in order
    all_entries = [root, e_vba, e_project, e_projectwm,
                   e_vba_proj, e_dir] + module_entries

    # Build sibling trees (sorted by CFB name comparison, red-black colored)
    # Root children: indices 1 (VBA), 2 (PROJECT), 3 (PROJECTwm)
    root_children = [1, 2, 3]
    tree_nodes = _build_balanced_tree(root_children, all_entries)
    for idx, left, right in tree_nodes:
        all_entries[idx].left_sibling = left
        all_entries[idx].right_sibling = right
    root.child = tree_nodes[0][0] if tree_nodes else NOSTREAM

    # VBA children: indices 4 (_VBA_PROJECT), 5 (dir), 6+ (modules)
    vba_children = list(range(4, len(all_entries)))
    tree_nodes = _build_balanced_tree(vba_children, all_entries)
    for idx, left, right in tree_nodes:
        all_entries[idx].left_sibling = left
        all_entries[idx].right_sibling = right
    e_vba.child = tree_nodes[0][0] if tree_nodes else NOSTREAM

    _write_cfb(all_entries, output_path)


# ===========================================================================
# CLI
# ===========================================================================

def main():
    if len(sys.argv) < 3:
        print("Usage: python3 vba_project_builder.py <output.bin> <module1.bas> [module2.cls ...]")
        print()
        print("File extensions determine module type:")
        print("  .bas → standard module")
        print("  .cls → document module (ThisWorkbook, Sheet1, ...)")
        print()
        print("Supports Cyrillic file names and VBA code (cp1251).")
        sys.exit(1)

    output_path = sys.argv[1]
    module_files = sys.argv[2:]

    modules = []
    for filepath in module_files:
        if not os.path.exists(filepath):
            print(f"Error: file not found: {filepath}", file=sys.stderr)
            sys.exit(1)

        basename = os.path.basename(filepath)
        name, ext = os.path.splitext(basename)
        ext = ext.lower()

        with open(filepath, "r", encoding="utf-8") as f:
            source = f.read()
        # Normalize line endings to CRLF (VBA requirement)
        source = source.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")

        if ext == ".cls":
            modules.append((name, "document", source))
        else:
            modules.append((name, "standard", source))

    if not modules:
        print("Error: no modules specified", file=sys.stderr)
        sys.exit(1)

    build_vba_project(modules, output_path)
    print(f"Created: {output_path} ({len(modules)} module(s))")


if __name__ == "__main__":
    main()
