"""Zip archives: list them, read one member in place, or unpack them safely.

The safety rules, proven in production document workflows: a member
that would write outside the destination is skipped (an absolute
path, a drive letter, a `..`, a name Windows cannot hold); an encrypted
member is skipped; a member whose compression ratio looks like a zip bomb is
skipped; and the archive, each member and the unpacked total are capped.
Nothing in an archive is ever run.

    from archive import list_members, read_member, unpack
    list_members(data)                 # [{index, name, size, ...}]
    read_member(data, "2")             # (name, bytes), by index or by name
    unpack(r"C:\\path\\a.zip")         # next to the zip, or dest=
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

MAX_ZIP_BYTES = 200 * 1024 * 1024
MAX_MEMBERS = 2000
MAX_MEMBER_BYTES = 100 * 1024 * 1024
MAX_TOTAL_BYTES = 500 * 1024 * 1024
BOMB_RATIO = 80                 # uncompressed / compressed, for members over 1 MB
PREVIEW = 40


class ArchiveError(Exception):
    def __init__(self, message, hint=""):
        super().__init__(message)
        self.hint = hint


def _open(data):
    try:
        return zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise ArchiveError(f"not a readable zip ({exc})") from exc


def _refusal(info):
    """Why a member is not read or unpacked, or None."""
    if info.flag_bits & 0x1:
        return "encrypted"
    if info.file_size > MAX_MEMBER_BYTES:
        return "larger than the per-file limit"
    if info.file_size > 1_000_000 and info.compress_size and info.file_size / info.compress_size > BOMB_RATIO:
        return "compression ratio too high"
    return None


def list_members(data):
    """Files in the archive (folders left out), numbered from 1."""
    with _open(data) as archive:
        files = [info for info in archive.infolist() if not info.is_dir()]
        out = []
        for index, info in enumerate(files[:MAX_MEMBERS], 1):
            row = {"index": index, "name": info.filename, "size": info.file_size,
                   "suffix": Path(info.filename).suffix.lower()}
            refusal = _refusal(info)
            if refusal:
                row["unreadable"] = refusal
            out.append(row)
        return out, len(files)


def read_member(data, selector):
    """(name, bytes) of one member, chosen by its 1-based index or its name
    (case does not matter; a bare file name matches a member in a folder)."""
    with _open(data) as archive:
        files = [info for info in archive.infolist() if not info.is_dir()]
        chosen = None
        text = str(selector).strip()
        if text.isdigit():
            number = int(text)
            if not 1 <= number <= len(files):
                raise ArchiveError(f"there is no member {number}; the archive has {len(files)}")
            chosen = files[number - 1]
        else:
            wanted = text.replace("\\", "/").lower()
            exact = [i for i in files if i.filename.lower() == wanted]
            tail = [i for i in files if i.filename.lower().rsplit("/", 1)[-1] == wanted]
            chosen = (exact or tail or [None])[0]
            if chosen is None:
                raise ArchiveError(f"no member named {selector!r}",
                                   "Available: " + ", ".join(f"{n}={i.filename}" for n, i in enumerate(files[:30], 1)))
        refusal = _refusal(chosen)
        if refusal:
            raise ArchiveError(f"{chosen.filename} is not read: {refusal}")
        with archive.open(chosen) as handle:
            return chosen.filename.rsplit("/", 1)[-1], handle.read(MAX_MEMBER_BYTES + 1)[:MAX_MEMBER_BYTES]


def _target(root: Path, name: str):
    """Where a member lands inside root, or None when it would escape it."""
    raw = str(name).replace("\\", "/")
    if raw.startswith("/") or re.match(r"^[A-Za-z]:", raw):
        return None
    parts = []
    for part in raw.split("/"):
        if part in ("", "."):
            continue
        if part == ".." or ":" in part or part.endswith((" ", ".")):
            return None
        parts.append(part)
    if not parts:
        return None
    dest = root.joinpath(*parts).resolve()
    base = root.resolve()
    if dest != base and base not in dest.parents:
        return None
    return dest


def _unique(path: Path) -> Path:
    if not path.exists():
        return path
    for n in range(2, 1000):
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not candidate.exists():
            return candidate
    raise ArchiveError(f"too many files named {path.name}")


def unpack(path, dest=None):
    """Extract a zip next to itself (or into dest). Never overwrites: a name
    that exists gets " (2)". A nested zip is extracted as a file, not opened.
    Returns a summary and the first files, not their contents."""
    source = Path(path)
    if source.stat().st_size > MAX_ZIP_BYTES:
        raise ArchiveError(f"the zip is larger than {MAX_ZIP_BYTES} bytes")
    if dest:
        folder = Path(dest).expanduser()
        folder = (Path.cwd() / folder) if not folder.is_absolute() else folder
    else:
        folder = source.with_suffix("")
        if folder.exists() and not folder.is_dir():
            folder = source.parent / f"{source.stem}-unpacked"
    folder = folder.resolve()
    if folder == source.resolve() or source.resolve() in folder.parents:
        raise ArchiveError("the destination would contain the zip itself")
    folder.mkdir(parents=True, exist_ok=True)

    extracted, skipped, total = [], [], 0
    with zipfile.ZipFile(source) as archive:
        members = [info for info in archive.infolist() if not info.is_dir()]
        if len(members) > MAX_MEMBERS:
            raise ArchiveError(f"the zip has more than {MAX_MEMBERS} files")
        for info in members:
            target = _target(folder, info.filename)
            if target is None:
                skipped.append({"name": info.filename, "reason": "path escapes the destination"})
                continue
            refusal = _refusal(info)
            if refusal:
                skipped.append({"name": info.filename, "reason": refusal})
                continue
            if total + info.file_size > MAX_TOTAL_BYTES:
                skipped.append({"name": info.filename, "reason": "uncompressed total exceeded"})
                break
            target = _unique(target)
            target.parent.mkdir(parents=True, exist_ok=True)
            written = 0
            with archive.open(info) as src, target.open("wb") as out:
                while written < info.file_size:
                    chunk = src.read(min(1024 * 1024, info.file_size - written))
                    if not chunk:
                        break
                    out.write(chunk)
                    written += len(chunk)
            total += written
            extracted.append({"path": str(target.relative_to(folder)), "size": written,
                              "suffix": target.suffix.lower()})
    by_suffix = {}
    for item in extracted:
        by_suffix[item["suffix"] or "(none)"] = by_suffix.get(item["suffix"] or "(none)", 0) + 1
    result = {"archive": str(source), "folder": str(folder), "count": len(extracted),
              "by_suffix": by_suffix, "files": extracted[:PREVIEW],
              "files_omitted": max(0, len(extracted) - PREVIEW)}
    if skipped:
        result["skipped"] = skipped[:PREVIEW]
    return result
