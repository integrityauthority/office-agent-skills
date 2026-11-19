"""Read an Outlook .msg file without Outlook.

A .msg is an OLE compound file holding MAPI properties ([MS-OXMSG]). Every
string property is a stream named ``__substg1.0_<id><type>``; the fixed-size
ones (dates, integers) sit together in ``__properties_version1.0``.
Recipients and attachments are sub-storages, and an attached e-mail is a
whole message again inside its attachment.

Only olefile (BSD) is needed. extract-msg would do the same job but is GPL-3,
which this MIT repository does not ship against.

    from msgfile import open_msg
    msg = open_msg(data)          # bytes of the .msg
    msg.summary()                 # subject, from, to, cc, date, body, attachments
    name, payload = msg.attachment(2)   # bytes, or a Msg for an attached e-mail
"""
from __future__ import annotations

import codecs
import datetime as _dt
import io
import re
import struct

PT_LONG = 0x0003
PT_BOOLEAN = 0x000B
PT_OBJECT = 0x000D
PT_STRING8 = 0x001E
PT_UNICODE = 0x001F
PT_SYSTIME = 0x0040
PT_BINARY = 0x0102

# Header length of __properties_version1.0 before the 16-byte entries.
_HEADER_TOP = 32
_HEADER_EMBEDDED = 24
_HEADER_SUB = 8

_ATTACH_EMBEDDED_MSG = 5
_ATTACH_OLE = 6
_RECIPIENT_KIND = {1: "to", 2: "cc", 3: "bcc"}


class MsgError(Exception):
    pass


def is_msg(data: bytes) -> bool:
    """An OLE file that carries MAPI message streams."""
    if not data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return False
    try:
        import olefile
    except ImportError:
        return False
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
    except Exception:
        return False
    try:
        return any(entry and entry[0].startswith("__substg1.0_") for entry in ole.listdir())
    finally:
        ole.close()


def _codec(codepage):
    if not codepage:
        return None
    name = "utf-8" if codepage == 65001 else f"cp{codepage}"
    if codepage in (28591, 28592):
        name = f"iso8859-{codepage - 28590}"
    try:
        codecs.lookup(name)
        return name
    except LookupError:
        return None


class _Node:
    """One MAPI property set: the message, a recipient or an attachment."""

    def __init__(self, ole, prefix, header, codepage=None):
        self.ole = ole
        self.prefix = list(prefix)
        self.header = header
        self.codepage = codepage
        self._fixed = None

    def _path(self, name):
        return "/".join(self.prefix + [name])

    def _has(self, name):
        return self.ole.exists(self._path(name))

    def _read(self, name):
        return self.ole.openstream(self._path(name)).read()

    def string(self, pid):
        base = f"__substg1.0_{pid:04X}"
        if self._has(base + "001F"):
            return self._read(base + "001F").decode("utf-16-le", "replace").rstrip("\0")
        if self._has(base + "001E"):
            raw = self._read(base + "001E").rstrip(b"\0")
            return raw.decode(self.codepage or "cp1252", "replace")
        return None

    def binary(self, pid):
        name = f"__substg1.0_{pid:04X}0102"
        return self._read(name) if self._has(name) else None

    def fixed(self):
        if self._fixed is None:
            self._fixed = {}
            name = "__properties_version1.0"
            if self._has(name):
                data = self._read(name)
                for offset in range(self.header, len(data) - 15, 16):
                    tag = struct.unpack_from("<I", data, offset)[0]
                    self._fixed[tag >> 16] = (tag & 0xFFFF, data[offset + 8:offset + 16])
        return self._fixed

    def long(self, pid):
        value = self.fixed().get(pid)
        if value and value[0] in (PT_LONG, PT_BOOLEAN):
            return struct.unpack_from("<i", value[1])[0]
        return None

    def time(self, pid):
        value = self.fixed().get(pid)
        if value and value[0] == PT_SYSTIME:
            ticks = struct.unpack_from("<Q", value[1])[0]
            if ticks:
                epoch = _dt.datetime(1601, 1, 1, tzinfo=_dt.timezone.utc)
                return epoch + _dt.timedelta(microseconds=ticks // 10)
        return None

    def substorages(self, start):
        depth = len(self.prefix)
        names = set()
        for entry in self.ole.listdir(streams=True, storages=True):
            if len(entry) > depth and entry[:depth] == self.prefix and entry[depth].startswith(start):
                names.add(entry[depth])
        return sorted(names)


def _is_embedded_msg(node):
    method = node.long(0x3705)
    if method is not None:
        return method == _ATTACH_EMBEDDED_MSG
    return node._has("__substg1.0_3701000D")


def _is_inline(node, mime):
    """A picture placed in the body (a signature logo), not a real attachment."""
    if node.long(0x7FFE):
        return True
    flags = node.long(0x3714) or 0
    looks_like_image = (mime or "").lower().startswith("image/")
    return looks_like_image and (bool(flags & 4) or bool(node.string(0x3712)))


def _address(name, address):
    name = (name or "").strip()
    address = (address or "").strip()
    if address.upper().startswith("/O="):
        address = ""  # an Exchange X500 path says nothing to a reader
    if name and address and name.lower() != address.lower():
        return f"{name} <{address}>"
    return name or address


def _header_value(headers, field):
    if not headers:
        return ""
    match = re.search(rf"^{field}:[ \t]*(.*(?:\r?\n[ \t].*)*)", headers, re.I | re.M)
    if not match:
        return ""
    from email.header import decode_header, make_header
    value = re.sub(r"\r?\n[ \t]+", " ", match.group(1)).strip()
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


class Msg:
    def __init__(self, ole, prefix=(), header=_HEADER_TOP, codepage=None):
        self.ole = ole
        self.node = _Node(ole, prefix, header, codepage)
        cp = self.node.long(0x3FFD) or self.node.long(0x3FDE)
        self.node.codepage = _codec(cp) or codepage

    # -- headers ------------------------------------------------------------
    def recipients(self):
        out = {"to": [], "cc": [], "bcc": []}
        for storage in self.node.substorages("__recip_version1.0_#"):
            node = _Node(self.ole, self.node.prefix + [storage], _HEADER_SUB, self.node.codepage)
            kind = _RECIPIENT_KIND.get(node.long(0x0C15) or 1, "to")
            smtp = node.string(0x39FE) or node.string(0x3003)
            out[kind].append(_address(node.string(0x3001), smtp))
        return out

    def sender(self):
        n = self.node
        name = n.string(0x0C1A) or n.string(0x0042)
        address = n.string(0x5D01) or n.string(0x5D02) or n.string(0x0C1F) or n.string(0x0065)
        value = _address(name, address)
        if not address or address.upper().startswith("/O="):
            from_header = _header_value(n.string(0x007D), "From")
            if from_header:
                return from_header
        return value

    def date(self):
        return self.node.time(0x0039) or self.node.time(0x0E06) or self.node.time(0x3007)

    # -- body ---------------------------------------------------------------
    def body(self):
        """(text, source): plain if present, else HTML, else the compressed RTF."""
        plain = self.node.string(0x1000)
        if plain and plain.strip():
            return plain, "plain"
        html = self.node.binary(0x1013)
        if html is not None:
            return html.decode(_html_charset(html) or self.node.codepage or "utf-8", "replace"), "html"
        html_text = self.node.string(0x1013)
        if html_text:
            return html_text, "html"
        rtf = self.node.binary(0x1009)
        if rtf:
            try:
                return rtf_to_text(decompress_rtf(rtf)), "rtf"
            except Exception:
                pass
        return "", ""

    # -- attachments --------------------------------------------------------
    def _attachment_nodes(self):
        nodes = []
        for storage in self.node.substorages("__attach_version1.0_#"):
            nodes.append(_Node(self.ole, self.node.prefix + [storage], _HEADER_SUB, self.node.codepage))
        return nodes

    def attachments(self):
        out = []
        for index, node in enumerate(self._attachment_nodes(), 1):
            method = node.long(0x3705)
            embedded = _is_embedded_msg(node)
            name = node.string(0x3707) or node.string(0x3704) or node.string(0x3001) or ""
            if method == _ATTACH_OLE:
                size = None
                mime = "application/x-ole-object"
            elif embedded:
                inner = Msg(self.ole, node.prefix + ["__substg1.0_3701000D"], _HEADER_EMBEDDED,
                            self.node.codepage)
                name = name or (inner.node.string(0x0037) or "message")
                size = None
                mime = "message/rfc822"
            else:
                data = node.binary(0x3701)
                size = len(data) if data is not None else None
                mime = node.string(0x370E) or ""
            out.append({
                "index": index,
                "name": name or f"attachment-{index}",
                "mime": mime,
                "size": size,
                "embedded_email": embedded,
                "inline": _is_inline(node, mime),
            })
            if method == _ATTACH_OLE:
                out[-1]["ole_object"] = True
        return out

    def attachment(self, index):
        """(name, bytes) for a file, (name, Msg) for an attached e-mail."""
        nodes = self._attachment_nodes()
        if not 1 <= index <= len(nodes):
            raise MsgError(f"there is no attachment {index}; this message has {len(nodes)}")
        node = nodes[index - 1]
        name = node.string(0x3707) or node.string(0x3704) or node.string(0x3001) or f"attachment-{index}"
        if node.long(0x3705) == _ATTACH_OLE:
            raise MsgError(f"attachment {index} ({name}) is an embedded OLE object, such as a "
                           "contact photo or a pasted object; it cannot be extracted as a file")
        if _is_embedded_msg(node):
            return name, Msg(self.ole, node.prefix + ["__substg1.0_3701000D"], _HEADER_EMBEDDED,
                             self.node.codepage)
        data = node.binary(0x3701)
        if data is None:
            raise MsgError(f"attachment {index} ({name}) has no data stream; it may be a link or an OLE object")
        return name, data

    def summary(self):
        recipients = self.recipients()
        headers = self.node.string(0x007D) or ""
        if not any(recipients.values()) and headers:
            recipients["to"] = [v for v in [_header_value(headers, "To")] if v]
            recipients["cc"] = [v for v in [_header_value(headers, "Cc")] if v]
        when = self.date()
        text, source = self.body()
        return {
            "message_class": self.node.string(0x001A) or "",
            "subject": self.node.string(0x0037) or "",
            "from": self.sender(),
            "to": recipients["to"],
            "cc": recipients["cc"],
            "bcc": recipients["bcc"],
            "date": when.astimezone().isoformat(timespec="minutes") if when else "",
            "body": text,
            "body_source": source,
            "attachments": self.attachments(),
        }


def open_msg(data: bytes) -> Msg:
    try:
        import olefile
    except ImportError as exc:
        raise MsgError("reading .msg needs the olefile package (pip install olefile)") from exc
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
    except Exception as exc:
        raise MsgError(f"not an Outlook .msg file ({type(exc).__name__})") from exc
    return Msg(ole)


def _html_charset(raw: bytes):
    match = re.search(rb'charset\s*=\s*["\']?([A-Za-z0-9_-]+)', raw[:4096], re.I)
    if not match:
        return None
    name = match.group(1).decode("ascii", "ignore")
    try:
        codecs.lookup(name)
        return name
    except LookupError:
        return None


# ------------------------------------------------------------ RTF (LZFu) ---

_RTF_PREBUF = (
    b"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss "
    b"\\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier"
    b"{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx"
)


def decompress_rtf(data: bytes) -> bytes:
    """[MS-OXRTFCP] compressed RTF to RTF bytes."""
    if len(data) < 16:
        raise MsgError("compressed RTF is shorter than its header")
    comp_size, raw_size, magic, _crc = struct.unpack_from("<IIII", data, 0)
    body = data[16:16 + comp_size - 12]
    if magic == 0x414C454D:  # "MELA": stored uncompressed
        return body[:raw_size]
    if magic != 0x75465A4C:  # "LZFu"
        raise MsgError("unknown compressed RTF format")
    window = bytearray(4096)
    window[:len(_RTF_PREBUF)] = _RTF_PREBUF
    write = len(_RTF_PREBUF)
    out = bytearray()
    pos = 0
    while pos < len(body):
        control = body[pos]
        pos += 1
        for bit in range(8):
            if pos >= len(body):
                break
            if control & (1 << bit):
                if pos + 1 >= len(body):
                    return bytes(out)
                word = (body[pos] << 8) | body[pos + 1]
                pos += 2
                offset, length = word >> 4, (word & 0xF) + 2
                if offset == write:
                    return bytes(out)
                for i in range(length):
                    byte = window[(offset + i) % 4096]
                    out.append(byte)
                    window[write] = byte
                    write = (write + 1) % 4096
            else:
                byte = body[pos]
                pos += 1
                out.append(byte)
                window[write] = byte
                write = (write + 1) % 4096
    return bytes(out)


_RTF_SKIP_DESTINATIONS = {
    "fonttbl", "colortbl", "stylesheet", "info", "pict", "object", "header", "footer",
    "headerl", "headerr", "footerl", "footerr", "listtable", "listoverridetable",
    "rsidtbl", "generator", "xmlnstbl", "themedata", "colorschememapping", "latentstyles",
    "datastore", "mmathPr",
}


def rtf_to_text(rtf: bytes) -> str:
    """Plain text from RTF, including Outlook's HTML-in-RTF (\\fromhtml1)."""
    text = rtf.decode("latin-1")
    codepage = "cp1252"
    match = re.search(r"\\ansicpg(\d+)", text)
    if match and _codec(int(match.group(1))):
        codepage = _codec(int(match.group(1)))
    out = []
    stack = []
    skip = False
    htmlrtf = False
    uc_skip = 1
    pending_bytes = bytearray()
    i = 0
    n = len(text)

    def flush():
        if pending_bytes:
            out.append(pending_bytes.decode(codepage, "replace"))
            pending_bytes.clear()

    while i < n:
        ch = text[i]
        if ch == "{":
            stack.append((skip, htmlrtf))
            i += 1
            if text.startswith("\\*", i):
                skip = True
            continue
        if ch == "}":
            flush()
            skip, htmlrtf = stack.pop() if stack else (False, False)
            i += 1
            continue
        if ch == "\\":
            i += 1
            if i >= n:
                break
            nxt = text[i]
            if nxt == "'":
                if not skip and not htmlrtf:
                    try:
                        pending_bytes.append(int(text[i + 1:i + 3], 16))
                    except ValueError:
                        pass
                i += 3
                continue
            if nxt in "\\{}":
                flush()
                if not skip and not htmlrtf:
                    out.append(nxt)
                i += 1
                continue
            if nxt == "~":
                flush()
                if not skip and not htmlrtf:
                    out.append("\u00a0")
                i += 1
                continue
            match = re.match(r"([a-zA-Z]+)(-?\d+)? ?", text[i:])
            if not match:
                i += 1
                continue
            word, arg = match.group(1), match.group(2)
            i += match.end()
            flush()
            if word in _RTF_SKIP_DESTINATIONS:
                skip = True
            elif word == "htmlrtf":
                htmlrtf = arg != "0"
            elif word == "uc" and arg:
                uc_skip = int(arg)
            elif skip or htmlrtf:
                pass
            elif word in ("par", "line", "row"):
                out.append("\n")
            elif word in ("tab", "cell"):
                out.append("\t")
            elif word == "u" and arg:
                value = int(arg)
                out.append(chr(value + 65536 if value < 0 else value))
                # The next uc_skip characters are the ANSI fallback; a \'hh
                # escape counts as one of them.
                for _ in range(uc_skip):
                    if text.startswith("\\'", i):
                        i += 4
                    elif i < n and text[i] not in "\\{}":
                        i += 1
            continue
        if ch in "\r\n":
            i += 1
            continue
        if not skip and not htmlrtf:
            flush()
            out.append(ch)
        i += 1
    flush()
    result = "".join(out)
    result = re.sub(r"[ \t]+\n", "\n", result)
    return re.sub(r"\n{3,}", "\n\n", result).strip()
