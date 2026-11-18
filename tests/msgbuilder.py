"""Build Outlook .msg files for the tests, without Outlook.

A .msg is an OLE compound file ([MS-CFB]) holding MAPI properties
([MS-OXMSG]). olefile only reads that format, so this is a small writer:
version 3, 512-byte sectors, streams under 4096 bytes in the mini stream.
It exists so that the tests need no binary fixture and no real mailbox --
every message here is made up, with example.com addresses.

    data = build_msg(subject="Hi", body="text", sender=("A", "a@example.com"),
                     to=[("B", "b@example.com")], attachments=[...])
"""
from __future__ import annotations

import datetime as _dt
import struct

ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF
SECTOR, MINI, CUTOFF = 512, 64, 4096

PT_LONG, PT_BOOLEAN, PT_SYSTIME = 0x0003, 0x000B, 0x0040


# ----------------------------------------------------------------- CFB ---

class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data      # kind: 5 root, 1 storage, 2 stream
        self.children: list[_Entry] = []
        self.sid = 0
        self.start, self.size = ENDOFCHAIN, 0


def _tree(storage: dict, entry: _Entry):
    """{name: bytes | dict} -> entries; a dict is a sub-storage."""
    for name, value in storage.items():
        child = _Entry(name, 1, b"") if isinstance(value, dict) else _Entry(name, 2, value)
        entry.children.append(child)
        if isinstance(value, dict):
            _tree(value, child)


def _flatten(entry, out):
    entry.sid = len(out)
    out.append(entry)
    for child in entry.children:
        _flatten(child, out)


def build_cfb(storage: dict) -> bytes:
    root = _Entry("Root Entry", 5)
    _tree(storage, root)
    entries: list[_Entry] = []
    _flatten(root, entries)

    # Mini stream: every stream under the cutoff, in 64-byte mini sectors.
    mini = bytearray()
    minifat: list[int] = []
    big = []
    for e in entries:
        if e.kind != 2:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
        elif e.size < CUTOFF:
            first = len(mini) // MINI
            count = -(-e.size // MINI)
            mini += e.data + b"\0" * (count * MINI - e.size)
            minifat += [first + i + 1 for i in range(count - 1)] + [ENDOFCHAIN]
            e.start = first
        else:
            big.append(e)

    def sectors(n_bytes):
        return -(-n_bytes // SECTOR)

    n_dir = sectors(len(entries) * 128)
    n_minifat = sectors(len(minifat) * 4)
    n_mini = sectors(len(mini))
    n_big = sum(sectors(e.size) for e in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_mini + n_big
        if n_fat * (SECTOR // 4) >= total:
            break
        n_fat += 1
    assert n_fat <= 109, "this writer keeps the FAT inside the header DIFAT"

    fat = [FATSECT] * n_fat
    cursor = n_fat

    def chain(count):
        nonlocal cursor
        first = cursor
        fat.extend([cursor + i + 1 for i in range(count - 1)] + [ENDOFCHAIN] if count else [])
        cursor += count
        return first if count else ENDOFCHAIN

    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat)
    mini_start = chain(n_mini)
    for e in big:
        e.start = chain(sectors(e.size))
    root.start, root.size = (mini_start, len(mini)) if mini else (ENDOFCHAIN, 0)
    fat += [FREESECT] * (n_fat * (SECTOR // 4) - len(fat))

    # Directory: each storage points at its first child; siblings are a
    # right-only chain, all black. olefile walks left, right and child.
    def dir_entry(e: _Entry, right):
        name = e.name.encode("utf-16-le") + b"\0\0"
        child = e.children[0].sid if e.children else NOSTREAM
        return (name.ljust(64, b"\0") + struct.pack("<HBB", len(name), e.kind, 1)
                + struct.pack("<III", NOSTREAM, right, child) + b"\0" * 16 + b"\0" * 4
                + b"\0" * 16 + struct.pack("<IQ", e.start, e.size))
    rights = {}
    for e in entries:
        for a, b in zip(e.children, e.children[1:]):
            rights[a.sid] = b.sid
    directory = b"".join(dir_entry(e, rights.get(e.sid, NOSTREAM)) for e in entries)
    empty = b"\0" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) + b"\0" * 48
    directory += empty * (n_dir * 4 - len(entries))

    header = (bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 16
              + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
              + struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, CUTOFF,
                            minifat_start if n_minifat else ENDOFCHAIN, n_minifat, ENDOFCHAIN, 0))
    header += b"".join(struct.pack("<I", i) for i in range(n_fat)) + struct.pack("<I", FREESECT) * (109 - n_fat)
    assert len(header) == SECTOR

    def pad(blob, count):
        return bytes(blob) + b"\0" * (count * SECTOR - len(blob))
    body = b"".join(struct.pack("<I", v) for v in fat)
    body += pad(directory, n_dir)
    body += pad(b"".join(struct.pack("<I", v) for v in minifat), n_minifat)
    body += pad(mini, n_mini)
    for e in big:
        body += pad(e.data, sectors(e.size))
    return header + body


# ----------------------------------------------------------------- MSG ---

def _filetime(when: _dt.datetime) -> int:
    epoch = _dt.datetime(1601, 1, 1, tzinfo=_dt.timezone.utc)
    return int((when - epoch).total_seconds() * 10_000_000)


def _props(header_size, fixed):
    """__properties_version1.0: header, then 16-byte entries for fixed props."""
    out = b"\0" * header_size
    for pid, (ptype, value) in fixed.items():
        if ptype == PT_SYSTIME:
            raw = struct.pack("<Q", _filetime(value))
        else:
            raw = struct.pack("<i", int(value)) + b"\0" * 4
        out += struct.pack("<II", (pid << 16) | ptype, 0x06) + raw
    return out


def _string(pid, text, unicode=True, codepage="cp1252"):
    if unicode:
        return {f"__substg1.0_{pid:04X}001F": text.encode("utf-16-le")}
    return {f"__substg1.0_{pid:04X}001E": text.encode(codepage)}


def _message_streams(subject="", body=None, html=None, rtf=None, sender=None, to=(), cc=(),
                     sent=None, attachments=(), message_class="IPM.Note", unicode=True,
                     codepage=None, header_size=32):
    s: dict = {}
    enc = f"cp{codepage}" if codepage else "cp1252"
    s.update(_string(0x001A, message_class, unicode, enc))
    s.update(_string(0x0037, subject, unicode, enc))
    if body is not None:
        s.update(_string(0x1000, body, unicode, enc))
    if html is not None:
        s["__substg1.0_10130102"] = html.encode("utf-8")
    if rtf is not None:
        s["__substg1.0_10090102"] = rtf
    if sender:
        s.update(_string(0x0C1A, sender[0], unicode, enc))
        s.update(_string(0x5D01, sender[1], unicode, enc))
    fixed = {}
    if sent:
        fixed[0x0039] = (PT_SYSTIME, sent)
    if codepage:
        fixed[0x3FFD] = (PT_LONG, codepage)
    recipients = [(1, r) for r in to] + [(2, r) for r in cc]
    for i, (kind, (name, address)) in enumerate(recipients):
        r = {}
        r.update(_string(0x3001, name, unicode, enc))
        r.update(_string(0x39FE, address, unicode, enc))
        r["__properties_version1.0"] = _props(8, {0x0C15: (PT_LONG, kind)})
        s[f"__recip_version1.0_#{i:08X}"] = r
    for i, att in enumerate(attachments):
        a = {}
        a.update(_string(0x3707, att["name"], unicode, enc))
        afixed = {0x3705: (PT_LONG, att.get("method", 1))}
        if att.get("mime"):
            a.update(_string(0x370E, att["mime"], unicode, enc))
        if att.get("content_id"):
            a.update(_string(0x3712, att["content_id"], unicode, enc))
        if "data" in att:
            a["__substg1.0_37010102"] = att["data"]
        if "message" in att:      # an attached e-mail, method 5
            afixed[0x3705] = (PT_LONG, 5)
            a["__substg1.0_3701000D"] = _message_streams(**att["message"], header_size=24)
        if att.get("ole"):        # an OLE object, method 6
            afixed[0x3705] = (PT_LONG, 6)
            a["__substg1.0_3701000D"] = {"CONTENTS": b"\x01\x02"}
        a["__properties_version1.0"] = _props(8, afixed)
        s[f"__attach_version1.0_#{i:08X}"] = a
    s["__properties_version1.0"] = _props(header_size, fixed)
    return s


def build_msg(**kwargs) -> bytes:
    return build_cfb(_message_streams(**kwargs))
