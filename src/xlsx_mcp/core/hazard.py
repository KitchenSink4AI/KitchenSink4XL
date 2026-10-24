"""core/hazard.py: the round-trip hazard scan and per-tool routing decision.

One of the most important modules. A cheap zip central-directory read detects
the fragile parts openpyxl
would drop on save (slicers, timelines, drawings/shapes, ActiveX, threaded
comments, Power Query DataMashup, connections, dynamic-array metadata, pivot
caches, VBA) WITHOUT an object-model load, then routes each mutation to
openpyxl / raw OOXML surgery / COM / HAZARD_REFUSED.

The knowledge table in HAZARD_SPECS encodes what openpyxl preserves vs drops.
The fidelity harness (tests/fixtures + scripts) VALIDATES every
`survives_openpyxl` claim in this table empirically against real fixtures,
and part-injection round-trips validated the additions
(embeddings, persons, richData, namedSheetViews, queryTables, data model).
Any file whose parts openpyxl would drop MUST be flagged here; a lossy file
this scan calls clean is the exact incumbent failure the brand exists to
prevent, so the table errs toward flagging (a false positive costs a needless
raw/COM route; a false negative silently destroys user content).

Two deliberate NON-flags, verified empirically (round-trips):
- xl/externalLinks/ is PRESERVED by openpyxl (modeled; re-serialized but
  semantically intact), so it carries no spec. Only externalBook links were
  tested; ddeLink/oleLink variants are untested and would surface in
  verify-after-write's part-loss diff if they ever dropped.
- xl/printerSettings/ IS dropped by openpyxl, but flagging it would refuse
  mutations on nearly every workbook that ever printed, for a loss that is
  cosmetic (page-setup device settings; the pageSetup element itself
  survives). It stays unflagged by policy and is treated as a routine drop.
  RE-VERIFIED against the 1.1 fidelity benchmark's printsetup probe, which
  listed print setup as a silent degradation: orientation, print area, print
  titles, fit-to-height and all six header and footer slots came back
  identical, and the only casualties were printerSettings1.bin and its
  relationship. The benchmark's separate claim that header/footer slots are
  reshuffled did not reproduce. The policy stands, on a measurement rather
  than on the original reasoning alone.

- LEGACY COMMENTS ARE NOT A LOSS, and the same benchmark row says they are.
  Excel names the parts xl/comments1.xml and xl/drawings/vmlDrawing1.vml;
  openpyxl models the notes and rewrites them as xl/comments/comment1.xml
  and xl/drawings/commentsDrawing1.vml. A part-name diff sees two parts
  leave and calls it a drop. Reading the file instead: both notes, both
  authors, both note-box dimensions came back identical. This is the same
  finding as the refusal of old files that carry a note, seen from the other
  direction, and it is
  why _refine_comment_vml demotes content-verified comment-only anchors
  rather than refusing on them.

IN-PART HAZARDS (the former known limit, now closed). Some fragile content
does not live in a part of its own: x14 conditional formatting (every modern
data bar and icon set), sparkline groups, worksheet slicer lists and the rest
of the worksheet-level extLst extensions sit INSIDE xl/worksheets/sheetN.xml,
which survives the save at full size. The fidelity gate measured what that
cost: an Excel-authored data bar plus icon set (tests/fixtures corpus
x14_condformat.xlsx) added no part at all, the namelist scan reported CLEAN,
the mutation applied without a murmur, and the x14 rules were gone from the
saved file. openpyxl's worksheet reader warns "<name> extension is not
supported and will be removed" for every ext it meets and writes none of them
back, so the whole block is DROP-class, not degrade-class.

The scan therefore reads the worksheet parts and looks at the top-level
extLst itself (EXT_DROP_LABELS / the "in_part_extensions" spec below). That
costs a decompress-and-parse of each worksheet, so scan_path is no longer
namelist-only for workbooks with sheets; the parse is a streaming expat walk
with no tree built, but the read is linear in sheet size and that is the price
of not destroying rules in silence. scan_names WITHOUT a part reader cannot
see in-part content and reports only the name-detectable hazards; core/
package.py keeps openpyxl's own load-time warnings as a second, independent
detector and refuses on those too, so a URI this table has never heard of is
still caught.

The package policy governs the outcome: drop-risk refuses
without allow_loss, degrade-risk warns and proceeds. In-part extensions are
drop-risk, so a mutation that would lose them now raises HAZARD_REFUSED naming
the rules at risk and the ways out. Under allow_loss the loss is still
announced item by item. This replaces the previous warn-and-proceed behavior,
which announced the loss but let it happen.

The refusal used to name a COM route (see refusal_routes below). It does not
any more: the com pack holds no tool that writes a cell, a format, or a row,
so for a refused file-tier mutation that route was a dead end the caller had
to walk before finding out. Preservation is the real fix and it is not in
this release; until it lands the refusal states the exact losses and the two
options that exist.
"""

from __future__ import annotations

import os
import re
import zipfile
from dataclasses import dataclass, field
from typing import Callable, Iterable

from . import xmlsafe as _xmlsafe

# ------------------------------------------- legacy OLE containers (.xls)
#
# A genuine .xls handed to any tool escaped the
# envelope entirely (openpyxl's InvalidFileException is a bare Exception
# subclass), or -- worse -- half-parsed. Modern Excel's .xls files embed a
# zip FRAGMENT (the theme package), so Python's zipfile "successfully" opens
# a BIFF file, finds [Content_Types].xml with no workbook part, and the
# caller of a perfectly healthy 2003 workbook was told the file was corrupt.
# Four different behaviors for one user situation, two out of envelope, and
# the one leaked message recommended xlrd, a Python library an MCP caller
# cannot invoke, instead of the server's own com_convert_format.
#
# The fix is an 8-byte signature sniff at every file entry point. An OLE
# compound file (D0CF11E0A1B11AE1) is never a readable OOXML package, but it
# IS one of two very different user situations, told apart by the container's
# top-level stream names: a legacy BIFF workbook carries a "Workbook" (or
# BIFF5 "Book") stream, while Excel's real encryption wraps an OOXML package
# in streams named "EncryptionInfo"/"EncryptedPackage". Diagnosing an
# encrypted heirloom as "an old .xls" (or vice versa) sends the user down
# the wrong remedy, so the sniff reads the first directory sector and names
# the situation it actually found.

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

OLE_KIND_BIFF = "biff"            # legacy Excel 97-2003 (or older) workbook
OLE_KIND_ENCRYPTED = "encrypted"  # Excel's real password encryption
OLE_KIND_UNKNOWN = "ole"          # an OLE container this sniff cannot place


def _ole_stream_names(fh, header: bytes) -> set[str]:
    """Lower-cased top-level stream names from the FIRST directory sector.
    The root entry and the main streams (Workbook / Book / EncryptionInfo /
    EncryptedPackage / WordDocument) live there in every real file; a parse
    failure returns what was read (the caller stays conservative)."""
    import struct
    names: set[str] = set()
    try:
        sector_shift = struct.unpack_from("<H", header, 30)[0]
        if not 6 <= sector_shift <= 20:
            return names
        sector = 1 << sector_shift
        first_dir = struct.unpack_from("<i", header, 48)[0]
        if first_dir < 0:
            return names
        fh.seek((first_dir + 1) * sector)
        data = fh.read(sector)
    except (OSError, struct.error):
        return names
    for off in range(0, len(data) - 127, 128):
        nlen = int.from_bytes(data[off + 64:off + 66], "little")
        if not 2 <= nlen <= 64:
            continue
        try:
            names.add(
                data[off:off + nlen - 2].decode("utf-16-le").strip().lower())
        except UnicodeDecodeError:
            continue
    return names


def ole_container_kind(path: str) -> str | None:
    """None when the file is not an OLE compound container; otherwise one of
    OLE_KIND_BIFF / OLE_KIND_ENCRYPTED / OLE_KIND_UNKNOWN."""
    try:
        with open(path, "rb") as fh:
            header = fh.read(512)
            if header[:8] != _OLE_MAGIC:
                return None
            if len(header) < 512:
                return OLE_KIND_UNKNOWN
            names = _ole_stream_names(fh, header)
    except OSError:
        return None  # unreadable here: the normal open path diagnoses locks
    if {"workbook", "book"} & names:
        return OLE_KIND_BIFF
    if {"encryptedpackage", "encryptioninfo"} & names:
        return OLE_KIND_ENCRYPTED
    return OLE_KIND_UNKNOWN


def ole_refusal_text(kind: str, path: str) -> str:
    """The one honest message per container kind, shared by every entry
    point so a .xls meets the same words whatever tool it hits first."""
    from pathlib import Path as _P
    name = _P(path).name
    if kind == OLE_KIND_BIFF:
        return (
            f"{name} is a legacy Excel 97-2003 workbook (BIFF format in an "
            "OLE compound container), not an OOXML .xlsx/.xlsm package -- "
            "whatever its extension says. The file-based tools cannot read "
            "BIFF. Remedy: com_convert_format(path=..., output='...xlsx') "
            "has Excel itself write a faithful .xlsx copy; then run this "
            "tool on the converted file.")
    if kind == OLE_KIND_ENCRYPTED:
        return (
            f"{name} is password-protected (Excel's real encryption wraps "
            "the whole package in an OLE container); the file-based tools "
            "cannot decrypt it. The com tools accept password=... "
            "(com_validate_opens_clean verifies it opens; "
            "com_save_with_password with password='' removes the "
            "encryption so the file tier can work on it).")
    return (
        f"{name} is an OLE compound file, not an OOXML package: most likely "
        "a legacy Office file (.xls/.doc/.ppt) or an encrypted workbook "
        "this scan could not classify. If it is a legacy Excel workbook, "
        "com_convert_format(path=..., output='...xlsx') converts it via "
        "Excel; if it is password-protected, the com tools accept "
        "password=... .")


def refuse_ole_container(path: str) -> None:
    """Raise the in-envelope refusal when `path` is an OLE container.
    UNSUPPORTED_CONTENT for BIFF/unknown (UnsupportedStructure),
    WorkbookProtected for the encrypted kind; both carry hint_tools so the
    envelope names the com pack when it is disabled."""
    kind = ole_container_kind(path)
    if kind is None:
        # Not a compound file: a Strict Open XML package is refused here too,
        # since every read and edit passes this gate before it opens a workbook.
        refuse_strict_package(path)
        return
    from .errors import UnsupportedStructure, WorkbookProtected
    if kind == OLE_KIND_ENCRYPTED:
        exc: Exception = WorkbookProtected(ole_refusal_text(kind, path))
        exc.hint_tools = ("com_validate_opens_clean",
                          "com_save_with_password")
    else:
        exc = UnsupportedStructure(ole_refusal_text(kind, path))
        exc.hint_tools = ("com_convert_format",)
    raise exc


# ------------------------------------------------- Strict Open XML packages


#: The refusal for a workbook saved as Strict Open XML. Excel opens it
#: cleanly, but its namespaces are not the Transitional ones this server
#: reads, so a read used to answer "ok" with no sheets and an edit answered
#: that the workbook had no active sheet: an AI given that calls the
#: workbook empty.
STRICT_REFUSAL = (
    "{name} is saved as Strict Open XML, which this server does not read; read "
    "as it is, it would look like a workbook with no sheets. Nothing was "
    "changed.")
STRICT_HINT = (
    "Open the workbook in Excel and save it as an Excel Workbook (.xlsx), "
    "then use the new file.")


def is_strict_package(path: str) -> bool:
    """True when the OOXML package at `path` is a Strict Open XML workbook:
    its root relationship names the workbook with the purl.oclc.org
    relationship type, or its workbook part (the one the root relationship
    names, else xl/workbook.xml) is in the Strict namespace. Any file that is
    not a readable zip answers False (the open path diagnoses the real
    reason). The package gate runs first, so a decompression bomb or a DTD is
    refused (UnsafePackage) from the zip directory before this sniff inflates
    any part, as every other open path does. The two parts are read as XML
    (any encoding, quote style or namespace prefix the XML rules allow), by
    namespace, never as text."""
    from pathlib import Path as _P
    from . import opcxml as _opc
    from . import xmlsafe as _xmlsafe
    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError, ValueError):
        return False
    with zf:
        _xmlsafe.check_package(zf, doc=_P(path).name)
    try:
        with zipfile.ZipFile(path) as zf:
            names = {n.lower(): n for n in zf.namelist()}
            workbooks: list[str] = []
            rels = names.get("_rels/.rels")
            if rels is not None:
                try:
                    found = _opc.relationships(
                        zf.read(rels), doc=_P(path).name, part="_rels/.rels")
                except _xmlsafe.XML_ERRORS:
                    found = []
                for rel in found:
                    if not rel.type.endswith("/officeDocument"):
                        continue
                    if rel.type.startswith(_opc.STRICT_REL_TYPE_ROOT):
                        return True
                    if not rel.external:
                        workbooks.append(_opc.resolve_part("", rel.target))
            workbooks.append("xl/workbook.xml")
            for want in workbooks:
                real = names.get(want.lower())
                if real is None:
                    continue
                with zf.open(real) as fh:
                    head = fh.read(_opc.ROOT_LOOK_BYTES)
                root = _opc.root_tag(head, doc=_P(path).name, part=real)
                if root == (_opc.STRICT_MAIN_NS, "workbook"):
                    return True
    except (zipfile.BadZipFile, OSError, KeyError, ValueError):
        return False
    return False


def refuse_strict_package(path: str) -> None:
    """Raise the named refusal when `path` is a Strict Open XML workbook."""
    if not is_strict_package(path):
        return
    from pathlib import Path as _P
    from .errors import UnsupportedStructure
    exc = UnsupportedStructure(STRICT_REFUSAL.format(name=_P(path).name))
    exc.hint = STRICT_HINT
    raise exc


# ---------------------------------------------- Excel 4.0 (XLM) macro sheets

#: An XLM macro sheet lives at xl/macrosheets/<name>.xml with the content
#: type application/vnd.ms-excel.macrosheet+xml. (Excel 5 dialog sheets, at
#: xl/dialogsheets/, are not XLM macro sheets and are not matched here.)
_MACROSHEET_PREFIX = "xl/macrosheets/"
_MACROSHEET_CT = "application/vnd.ms-excel.macrosheet+xml"


def has_xlm_macro_sheet(path: str) -> bool:
    """True when the OOXML package carries an Excel 4.0 (XLM) macro sheet.
    An XLM macro sheet can hold an auto-run macro (a defined name Auto_Open),
    and Excel's AutomationSecurity cannot disable XLM macros, so the COM
    route never opens such a file. Reads the package only; never opens it in
    Excel. A file that is not a readable OOXML zip answers False (the open
    path diagnoses the real reason)."""
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if name.lower().startswith(_MACROSHEET_PREFIX):
                    return True
            try:
                ct = zf.read("[Content_Types].xml").decode("utf-8", "replace")
            except KeyError:
                return False
            return _MACROSHEET_CT in ct.lower()
    except (zipfile.BadZipFile, OSError, KeyError):
        return False


def xlm_macro_sheet_refusal(path: str):
    """The refusal every COM entry point raises for an XLM macro-sheet
    workbook, with the file-route remedy."""
    from pathlib import Path as _P
    from .errors import UnsupportedStructure
    name = _P(path).name
    exc = UnsupportedStructure(
        f"{name} contains an Excel 4.0 (XLM) macro sheet. The COM route does "
        "not open it: Excel's automation security can disable VBA and "
        "auto-run macros, but not XLM 4.0 macros, so opening the file in "
        "Excel could run a macro the workbook carries. Excel was not "
        "started and nothing was changed. The file-based tools read and "
        "edit this workbook without running anything and keep the macro "
        "sheet intact; remove the macro sheet to use a COM operation on it.")
    exc.hint_tools = ("diagnose_workbook",)
    return exc


def refuse_xlm_macro_sheet(path: str) -> None:
    """Raise the XLM refusal when `path` carries an XLM macro sheet."""
    if has_xlm_macro_sheet(path):
        raise xlm_macro_sheet_refusal(path)


# ------------------------------------------------- password to modify

_FILE_SHARING_RE = re.compile(r"<(?:[A-Za-z_][\w.-]*:)?fileSharing\b([^>]*)>")
_ATTR_RE = re.compile(r"([A-Za-z_][\w.:-]*)\s*=\s*(\"[^\"]*\"|'[^']*')")

#: Byte-order marks, longest first (UTF-32 LE starts with the UTF-16 LE one).
_BOMS = ((b"\x00\x00\xfe\xff", "utf-32-be"), (b"\xff\xfe\x00\x00", "utf-32-le"),
         (b"\xef\xbb\xbf", "utf-8"), (b"\xfe\xff", "utf-16-be"),
         (b"\xff\xfe", "utf-16-le"))
_XML_DECL_ENCODING_RE = re.compile(
    rb"<\?xml[^>]*?encoding\s*=\s*[\"']([A-Za-z0-9._-]+)[\"']")


def decode_xml_part(data: bytes) -> str:
    """An XML part as text, in the encoding it is actually stored in.
    Excel opens a workbook.xml stored as UTF-16, where a byte scan
    for "<fileSharing" finds nothing because every character is followed by
    a zero byte. Order of evidence, as XML 1.0 Appendix F has it: a
    byte-order mark; then the width of the first "<" (UTF-16 or UTF-32
    without a mark); then the encoding the declaration names. Undecodable
    bytes become U+FFFD rather than an error, so a scan never fails."""
    for bom, enc in _BOMS:
        if data.startswith(bom):
            return data[len(bom):].decode(enc, "replace")
    for head, enc in ((b"\x00\x00\x00<", "utf-32-be"), (b"<\x00\x00\x00", "utf-32-le"),
                      (b"\x00<", "utf-16-be"), (b"<\x00", "utf-16-le")):
        if data.startswith(head):
            return data.decode(enc, "replace")
    enc = "utf-8"
    m = _XML_DECL_ENCODING_RE.match(data)
    if m is not None:
        import codecs
        try:
            enc = codecs.lookup(m.group(1).decode("ascii")).name
        except (LookupError, UnicodeDecodeError):
            enc = "utf-8"
    return data.decode(enc, "replace")


def _attrs(text: str) -> dict[str, str]:
    return {k.split(":")[-1]: v[1:-1] for k, v in _ATTR_RE.findall(text)}


def _workbook_part_name(zf: zipfile.ZipFile) -> str:
    """The workbook part the root relationships name (xl/workbook.xml in
    every file Excel writes, xl/workbook.bin in an .xlsb; a few generators
    put it elsewhere)."""
    try:
        rels = decode_xml_part(zf.read("_rels/.rels"))
    except (KeyError, OSError, zipfile.BadZipFile, ValueError,
            NotImplementedError, RuntimeError):
        return "xl/workbook.xml"
    names = set(zf.namelist())
    for m in re.finditer(r"<(?:\w+:)?Relationship\b([^>]*)>", rels):
        attrs = _attrs(m.group(1))
        if attrs.get("Type", "").endswith("/officeDocument"):
            target = attrs.get("Target", "").lstrip("/")
            if target and target in names:
                return target
    return "xl/workbook.xml"


# BIFF12 (.xlsb) records that hold the password to modify (MS-XLSB 2.4).
_BRT_BEGIN_BOOK = 0x83
_BRT_FILE_SHARING = 0x224       # fReadOnlyRec(2) wResPass(2) stUserName
_BRT_FILE_SHARING_ISO = 0x2A4   # spinCount(4) fReadOnlyRec(2) stUserName hash


class _Unreadable(ValueError):
    """The binary workbook part does not frame as BIFF12 records."""


def _biff12_records(data: bytes):
    """(record type, body) for each record of a BIFF12 stream: a 1-2 byte
    type and a 1-4 byte size, 7 bits a byte with the high bit meaning
    "another byte follows"."""
    i, n = 0, len(data)
    while i < n:
        b = data[i]
        i += 1
        rt = b & 0x7F
        if b & 0x80:
            if i >= n or data[i] & 0x80:
                raise _Unreadable("record type runs past two bytes")
            rt |= data[i] << 7
            i += 1
        size = 0
        for k in range(4):
            if i >= n:
                raise _Unreadable("record size cut off")
            b = data[i]
            i += 1
            size |= (b & 0x7F) << (7 * k)
            if not b & 0x80:
                break
        else:
            raise _Unreadable("record size runs past four bytes")
        if i + size > n:
            raise _Unreadable("record runs past the end of the part")
        yield rt, data[i:i + size]
        i += size


def _xl_wide_string(body: bytes, off: int) -> str | None:
    """An XLNullableWideString (4-byte character count, UTF-16LE) at off."""
    if len(body) < off + 4:
        return None
    cch = int.from_bytes(body[off:off + 4], "little")
    if cch == 0xFFFFFFFF or len(body) < off + 4 + 2 * cch:
        return None
    return body[off + 4:off + 4 + 2 * cch].decode("utf-16-le", "replace") or None


def _xlsb_reservation(data: bytes) -> dict | None:
    """write_reservation for a binary workbook part. Excel
    saves a password to modify as BrtFileSharingIso (the ISO hash; Excel
    writes it only when there is a password) followed by BrtFileSharing
    with a zero legacy verifier; an older file carries a non-zero verifier
    in BrtFileSharing alone. Read-only recommended alone is the
    BrtFileSharing record with both password fields empty."""
    reserved = False
    ro_rec = False
    user = None
    try:
        first = True
        for rt, body in _biff12_records(data):
            if first and rt != _BRT_BEGIN_BOOK:
                raise _Unreadable("the part does not start a workbook")
            first = False
            if rt == _BRT_FILE_SHARING_ISO:
                reserved = True
                ro_rec = ro_rec or body[4:6] not in (b"", b"\x00\x00")
                user = user or _xl_wide_string(body, 6)
            elif rt == _BRT_FILE_SHARING:
                if len(body) < 4:
                    raise _Unreadable("a file-sharing record too short to read")
                ro_rec = ro_rec or body[0:2] != b"\x00\x00"
                if body[2:4] != b"\x00\x00":
                    reserved = True
                user = user or _xl_wide_string(body, 4)
    except _Unreadable:
        return _UNREADABLE_RESERVATION
    if not reserved:
        return None
    return {"password_to_modify": True, "reserved_by": user,
            "read_only_recommended": ro_rec}


#: What write_reservation returns when the workbook part cannot be read
#: well enough to rule a password to modify out. Not None, so every
#: writing route refuses it; password_to_modify None says "unknown".
_UNREADABLE_RESERVATION = {"password_to_modify": None, "reserved_by": None,
                           "read_only_recommended": False, "unreadable": True}


def _is_binary_part(name: str, data: bytes) -> bool:
    return name.lower().endswith(".bin") or data[:2] == b"\x83\x01"


def write_reservation(path: str) -> dict | None:
    """The workbook's file-sharing settings when they make Excel stop and
    ask on a read-write open, else None.

    A "password to modify" (Excel: Save As, Tools, General Options) is
    stored in workbook.xml as <fileSharing reservationPassword=...> (legacy
    hash) or algorithmName/hashValue/saltValue/spinCount (modern hash), and
    in an .xlsb's binary workbook.bin as a BrtFileSharingIso or
    BrtFileSharing record. The file is an ordinary zip, not encrypted, so
    the encryption sniff passes it; but Excel opening it read-write shows a
    Password dialog that DisplayAlerts does not suppress, and in this
    server's hidden Excel that dialog appeared on the user's desktop. The
    .xlsb and UTF-16 forms are covered too. The COM routes that write call
    this before Excel starts. "Read-only recommended"
    alone is not returned: the open passes IgnoreReadOnlyRecommended, so
    Excel does not ask about it. A binary workbook part that cannot be read
    returns a result with password_to_modify None (cannot be excluded),
    which the writing routes refuse like a password."""
    try:
        with zipfile.ZipFile(path) as zf:
            part = _workbook_part_name(zf)
            data = zf.read(part)
    except (KeyError, OSError, zipfile.BadZipFile, ValueError,
            NotImplementedError, RuntimeError):
        return None
    if _is_binary_part(part, data):
        return _xlsb_reservation(data)
    m = _FILE_SHARING_RE.search(decode_xml_part(data))
    if m is None:
        return None
    attrs = _attrs(m.group(1))
    if not (attrs.get("reservationPassword") or attrs.get("hashValue")):
        return None
    return {"password_to_modify": True,
            "reserved_by": attrs.get("userName") or None,
            "read_only_recommended": attrs.get("readOnlyRecommended")
            in ("1", "true", "on")}


def write_reserved_refusal(path: str, *, unreadable: bool = False,
                           excel_started: bool = False):
    """The refusal for a workbook with a password to modify, sent to a COM
    route that would open it for editing. unreadable: the settings could
    not be read, so a password cannot be excluded. excel_started: the
    refusal comes from inside the shared open, where Excel is running (it
    reported the password, or the file could not be checked), so the "not
    started" sentence would be false."""
    from pathlib import Path as _P
    from .errors import WorkbookProtected
    name = _P(path).name
    if unreadable:
        lead = (f"This server cannot read {name}'s settings to rule out a "
                "password to modify, which would make Excel stop to ask for "
                "it when it opens the file for editing, and this server's "
                "hidden Excel cannot answer. ")
        remedy = ("Open it in Excel and check File, Save As, Tools, General "
                  "Options: clear Password to modify if there is one (and "
                  "remove the file's encryption if it has any), save, and "
                  "retry. ")
    else:
        lead = ((f"Excel reports that {name} has a password to modify, so it "
                 "would stop to ask for it when it opens the file for "
                 "editing, and this server's hidden Excel cannot answer. ")
                if excel_started else
                (f"{name} has a password to modify, so Excel would stop to "
                 "ask for it when it opens the file for editing, and this "
                 "server's hidden Excel cannot answer. "))
        remedy = ("Remove the password in Excel (File, Save As, Tools, "
                  "General Options: clear Password to modify), save, and "
                  "retry. ")
    exc = WorkbookProtected(
        lead
        + ("Nothing was changed, and Excel did not open the file for "
           "editing. " if excel_started else
           "Nothing was changed and Excel was not started. ")
        + remedy
        + "The file tools edit it as it is, and the read-only Excel tools "
        "(com_validate_opens_clean, com_export_pdf, com_render_sheet, "
        "com_convert_format) open it without asking.")
    exc.hint_tools = ("com_validate_opens_clean",)
    return exc


def write_reserved_notice(path: str, *, kept: bool = True) -> str:
    """The warning a successful FILE-route write carries when the workbook
    has a password to modify. The file tools edit it as it is
    and never touch the password; the answer says both, because "done" with
    no word about the password reads as if there were none. kept=False is the
    sentence for a saved file that no longer carries the password, which the
    file tools are tested never to produce and which is said plainly if it
    ever happens."""
    from pathlib import Path as _P
    name = _P(path).name
    if not kept:
        return (f"{name} had a password to modify, and the saved file no longer "
                "carries it. Tell the user, and restore the file as it was "
                "with xl_workbook manage_backups restore, source 'prev'.")
    return (f"{name} has a password to modify. This edit was made without that "
            "password, and the password is kept. If it exists to control who "
            "may edit the file, tell the user.")


def refuse_if_write_reserved(path: str) -> None:
    """Raise write_reserved_refusal when `path` has a password to modify,
    or when its workbook part cannot be read well enough to rule one out."""
    got = write_reservation(path)
    if got is not None:
        raise write_reserved_refusal(
            path, unreadable=got.get("password_to_modify") is None)


# ------------------------------------------------------------ routing verbs

ROUTE_OPENPYXL = "openpyxl"       # clean workbook, model rewrite is safe
ROUTE_RAW_OOXML = "raw_ooxml"     # hazard present, surgical edit, byte-preserve
ROUTE_COM = "com"                 # hazard present, structural edit, Excel saves
ROUTE_REFUSE = "refuse"           # hazard present, no safe route, no override

# ------------------------------------------------------------ severity

SEV_DROPS = "drops"        # openpyxl silently loses this part on save
SEV_DEGRADES = "degrades"  # openpyxl round-trips it through a lossy model
SEV_CONDITIONAL = "conditional"  # preserved only under a special load flag


@dataclass(frozen=True)
class HazardSpec:
    """One fragile-part class: how to detect it in the zip and what openpyxl
    does to it. `matcher` is a predicate over a single archive member name."""

    key: str
    label: str
    severity: str
    survives_openpyxl: bool
    note: str
    matcher: "callable"


def _prefix(*prefixes: str):
    def m(name: str) -> bool:
        low = name.lower()
        return any(low.startswith(p) for p in prefixes)
    return m


def _exact(*names: str):
    wanted = {n.lower() for n in names}
    def m(name: str) -> bool:
        return name.lower() in wanted
    return m


def _basename(*names: str):
    wanted = {n.lower() for n in names}
    def m(name: str) -> bool:
        return name.rsplit("/", 1)[-1].lower() in wanted
    return m


def _basename_or_prefix(basenames: tuple[str, ...], prefixes: tuple[str, ...]):
    wanted = {n.lower() for n in basenames}
    def m(name: str) -> bool:
        low = name.lower()
        if low.rsplit("/", 1)[-1] in wanted:
            return True
        return any(low.startswith(p) for p in prefixes)
    return m


def _never(_name: str) -> bool:
    """Matcher for a hazard that has no part of its own. The in-part
    extension hazard is detected by READING worksheet content, never by a
    member name, and this matters beyond the scan: verify.part_loss_check
    walks HAZARD_SPECS matchers over the parts that VANISHED, and the
    worksheet part does not vanish (its extLst does). A never-matching
    matcher keeps that check honest instead of teaching it to excuse a lost
    worksheet."""
    return False


def _drawings_matcher(name: str) -> bool:
    """Match drawing parts under xl/drawings/, EXCEPT legacy-comment VML
    anchors named commentsDrawing*.vml (openpyxl's own output naming).
    openpyxl models legacy notes and regenerates their VML anchor on the
    round-trip (empirically confirmed), so a comment anchor is not shape
    loss and must not trip the SEV_DROPS drawings hazard.

    EXCEL'S naming is different: real Excel writes comment anchors as
    vmlDrawing*.vml -- the same basename form controls use -- so a NAME
    cannot tell a sticky note from a checkbox (this docstring used to claim
    the opposite, and every real old file with a comment refused plain edits
    on it). vmlDrawing*.vml therefore stays matched
    here, and _refine_comment_vml then READS each one: an anchor whose every
    shape is a ClientData ObjectType=\"Note\" is demoted from the hazard
    (content-verified comment-only), while anything carrying form controls
    or drawn shapes -- or that cannot be read or parsed -- stays flagged."""
    low = name.lower()
    if not low.startswith("xl/drawings/"):
        return False
    base = low.rsplit("/", 1)[-1]
    if base.startswith("commentsdrawing") and base.endswith(".vml"):
        return False
    return True


#: VML element local names that mean the part carries a DRAWN object of its
#: own (not a comment note): any of these keeps the part flagged.
_VML_DRAWN_LOCALS = frozenset({
    "oval", "rect", "roundrect", "line", "polyline", "arc", "curve",
    "image", "group",
})


def _vml_local(tag) -> str:
    return tag.rsplit("}", 1)[-1].lower() if isinstance(tag, str) else ""


def _vml_is_comment_only(data: bytes) -> bool | None:
    """True when a VML part holds nothing but legacy-comment note anchors:
    at least one shape, every shape carrying exactly one
    x:ClientData ObjectType="Note", and no drawn VML object anywhere.
    False for form controls (Checkbox/Button/Drop/... ClientData), drawn
    shapes, or an empty part; None when the XML cannot be parsed (the
    caller stays conservative and keeps the part flagged)."""
    try:
        root = _xmlsafe.fromstring(data, part="a VML drawing")
    except Exception:  # noqa: BLE001
        return None
    shapes = 0
    for el in root.iter():
        local = _vml_local(el.tag)
        if local in _VML_DRAWN_LOCALS:
            return False
        if local != "shape":
            continue
        shapes += 1
        client_data = [c for c in el.iter()
                       if _vml_local(c.tag) == "clientdata"]
        if len(client_data) != 1 or \
                client_data[0].get("ObjectType") != "Note":
            return False
    return shapes > 0


#: A legacy comments part, either spelling (Excel's xl/comments1.xml or
#: openpyxl's xl/comments/comment1.xml). The comment-anchor demotion requires
#: one: openpyxl regenerates a VML anchor only for notes it actually models,
#: so an ORPHANED Note-anchor part (no comments anywhere) stays flagged.
_COMMENTS_PART_RE = re.compile(
    r"^xl/(comments\d+\.xml|comments/comment\d+\.xml)$")


def _refine_comment_vml(
    found: dict[str, "Hazard"],
    nameset: set[str],
    reader: Callable[[str], bytes] | None,
) -> list[str]:
    """Demote content-verified comment-only vmlDrawing*.vml anchors from the
    SEV_DROPS drawings hazard, returning the demoted part names.

    Excel names comment anchors vmlDrawing*.vml (see _drawings_matcher), so
    the single most common legacy-file scar -- a sticky note on an old business
    file -- refused every plain edit while the note in fact round-trips
    (openpyxl regenerates the anchor under its own commentsDrawing name,
    text, author, and position intact; empirically confirmed on the
    corpus of old files). A namelist alone cannot make this call, because the
    same basename carries form controls; the demotion happens ONLY after
    reading the part and finding nothing but Note anchors. Without a reader
    (scan_names namelist-only) everything stays flagged, the conservative
    direction."""
    hz = found.get("drawings")
    if hz is None or reader is None:
        return []
    if not any(_COMMENTS_PART_RE.match(_normalized(n).lower())
               for n in nameset):
        return []  # no legacy comments to anchor: nothing to demote
    demoted: list[str] = []
    for part in hz.parts:
        base = part.lower().rsplit("/", 1)[-1]
        if not (base.startswith("vmldrawing") and base.endswith(".vml")):
            continue
        try:
            data = reader(part)
        except Exception:  # noqa: BLE001
            continue
        if _vml_is_comment_only(data) is True:
            demoted.append(part)
    if demoted:
        kept = [p for p in hz.parts if p not in demoted]
        if kept:
            hz.parts = kept
        else:
            del found["drawings"]
    return demoted


# ------------------------------------------------- in-part extension blocks

#: Hazard key for worksheet-level extLst content. It has no part of its own,
#: so scan_names detects it by reading the worksheet rather than by matching a
#: member name (see _never and _detect_in_part_extensions).
IN_PART_EXT_KEY = "in_part_extensions"

#: Three more classes that live inside a part which survives, found by the
#: 1.1 fidelity benchmark. Same shape as IN_PART_EXT_KEY: no member name to
#: match on, so _never is the matcher and a content reader finds them.
PROTECTED_RANGES_KEY = "protected_ranges"
SORT_STATE_KEY = "sort_state"
DOC_PROPERTIES_KEY = "doc_properties"
MERGED_HIDDEN_KEY = "merged_hidden"
BROKEN_SHARED_KEY = "broken_shared_formula"

#: The worksheet elements openpyxl neither reads nor writes, as (local name,
#: hazard key). Both are plain CT_Worksheet children, NOT extLst content, so
#: the extLst walk above does not see them. The x14 protected-range extension
#: is a different element that the extLst walk does cover.
_IN_PART_WORKSHEET_ELEMENTS: tuple[tuple[bytes, str], ...] = (
    (b"protectedRange", PROTECTED_RANGES_KEY),
    (b"sortState", SORT_STATE_KEY),
)

#: docProps/app.xml fields a human authored, which openpyxl's writer does not
#: carry across (it emits a fresh ExtendedProperties every save). An EMPTY
#: element is not a loss and does not flag: Excel writes <Company/> on most
#: installs, and flagging that would put a hazard on nearly every workbook.
_APP_PROPERTY_TAGS: tuple[str, ...] = ("Company", "Manager")
_APP_PART = "docprops/app.xml"

#: Worksheet extLst URIs and what each one carries, mirroring
#: openpyxl.xml.constants.EXT_TYPES so these labels line up with the
#: "... extension is not supported and will be removed" warnings the package
#: surfaces from the load. openpyxl drops EVERY worksheet-level ext, listed or
#: not, so a URI missing from this table is still reported (as an
#: unrecognized extension) rather than passed over.
EXT_DROP_LABELS: dict[str, str] = {
    "{78C0D931-6437-407D-A8EE-F0AAD7539E65}":
        "x14 conditional formatting (data bars, icon sets)",
    "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}": "x14 data validation",
    "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}": "sparkline groups",
    "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}": "worksheet slicer list",
    "{3A4CF648-6AED-40F4-86FF-DC5316D8AED3}": "worksheet slicer list",
    "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}": "protected ranges",
    "{01252117-D84E-4E92-8308-4BE1C098FCBB}": "ignored-error markers",
    "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}": "web extensions",
    "{7E03D99C-DC04-49D9-9315-930204A7B6E9}": "timeline references",
}


def _top_level_ext_uris(data: bytes) -> list[str]:
    """URIs of the <ext> elements directly under the worksheet's top-level
    <extLst>, in document order, de-duplicated.

    Depth matters. A cfRule carries its OWN nested extLst holding the x14 rule
    id ({B025F937-...}), and that one is not a dropped extension block, it is
    a pointer into the block. Matching URIs with a regex over the whole part
    would report it as a second hazard and misdescribe what is at risk, so the
    walk only accepts worksheet > extLst > ext.

    Streaming, no tree (core/xmlsafe.walk): a worksheet can be hundreds of
    megabytes and this runs on every open. A malformed part yields whatever
    was read before the error, which is the conservative direction (report,
    do not swallow)."""
    uris: list[str] = []
    stack: list[str] = []

    def start(name: str, attrs: dict) -> None:
        local = name.rsplit("}", 1)[-1]
        stack.append(local)
        if len(stack) == 3 and stack[1] == "extLst" and local == "ext":
            uri = attrs.get("uri")
            if uri and uri not in uris:
                uris.append(uri)

    def end(_name: str) -> None:
        if stack:
            stack.pop()

    try:
        _xmlsafe.walk(data, start=start, end=end, part="a worksheet")
    except Exception:  # noqa: BLE001
        return uris
    return uris


def _is_worksheet_part(name: str) -> bool:
    low = _normalized(name).lower()
    return (low.startswith("xl/worksheets/") and low.endswith(".xml")
            and "/_rels/" not in low)


def _detect_in_part_extensions(
    found: dict[str, "Hazard"], names: Iterable[str],
    part_reader: Callable[[str], bytes] | None,
) -> None:
    """Add the in-part extension hazard when a worksheet carries a top-level
    extLst. Needs part content, so it is a no-op on the namelist-only path."""
    if part_reader is None:
        return
    labels: list[str] = []
    parts: list[str] = []
    for name in names:
        if not _is_worksheet_part(name):
            continue
        try:
            data = part_reader(name)
        except Exception:  # noqa: BLE001
            continue
        # Cheap reject before the parse. It is a UTF-8 byte scan, so a part
        # with a UTF-16 byte-order mark (legal XML, and openpyxl loads one)
        # skips the reject and goes straight to the parse, which reads the
        # declared encoding itself.
        if not data[:2] in (b"\xff\xfe", b"\xfe\xff") and b"extLst" not in data:
            continue
        uris = _top_level_ext_uris(data)
        if not uris:
            continue
        parts.append(name)
        for uri in uris:
            label = EXT_DROP_LABELS.get(
                uri.upper(), f"unrecognized extension {uri}")
            if label not in labels:
                labels.append(label)
    if not parts:
        return
    spec = _SPEC_BY_KEY[IN_PART_EXT_KEY]
    found[IN_PART_EXT_KEY] = Hazard(
        spec.key, f"{spec.label}: {', '.join(labels)}", spec.severity,
        spec.survives_openpyxl, spec.note, parts,
    )


def _detect_in_part_worksheet_elements(
    found: dict[str, "Hazard"], names: Iterable[str],
    part_reader: Callable[[str], bytes] | None,
) -> None:
    """Add a hazard for each worksheet element openpyxl silently discards.

    The extLst walk above answers a different question: it asks what is in
    the extension block. <protectedRanges> and <sortState> are ordinary
    CT_Worksheet children written by Excel in the plain namespace, so they
    are invisible to that walk, and openpyxl's reader has no handler for
    either. The 1.1 fidelity benchmark measured both going out with the
    save on a file that reported clean.

    A byte scan is enough. Both names are distinctive, neither appears in
    an attribute value or a shared string that reaches this part, and a
    false positive costs a needless refusal rather than lost content, which
    is the direction this table errs in by policy.
    """
    if part_reader is None:
        return
    hits: dict[str, list[str]] = {}
    for name in names:
        if not _is_worksheet_part(name):
            continue
        try:
            data = part_reader(name)
        except Exception:  # noqa: BLE001
            continue
        for needle, key in _IN_PART_WORKSHEET_ELEMENTS:
            if b"<" + needle in data or b":" + needle in data:
                hits.setdefault(key, []).append(name)
    for key, parts in hits.items():
        spec = _SPEC_BY_KEY[key]
        found[key] = Hazard(spec.key, spec.label, spec.severity,
                            spec.survives_openpyxl, spec.note, parts)


_CELL_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
#: A merged range up to this many cells is indexed cell by cell; a larger
#: one (a whole-column merge) is tested by bounds.
_MERGE_INDEX_CELLS = 4096


def _col_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _cell_rc(ref: str) -> tuple[int, int] | None:
    m = _CELL_REF_RE.match(ref.strip())
    if m is None:
        return None
    return int(m.group(2)), _col_number(m.group(1))


#: Declared encodings under which the part's bytes spell its element names
#: literally (a byte search for "mergeCell" is then a fair first filter).
_PLAIN_ENCODINGS = (b"utf-8", b"utf8", b"us-ascii", b"ascii", b"iso-8859-1",
                    b"latin1", b"windows-1252")
_DECLARED_ENCODING = re.compile(
    rb"\A(?:\xef\xbb\xbf)?\s*<\?xml\b[^>]*?\bencoding\s*=\s*([\"'])([^\"']*)\1")


def _may_mention(data: bytes, needle: bytes) -> bool:
    """False only when `data` cannot contain the element name `needle`: its
    bytes are an ASCII-compatible encoding that the part declares plainly (or
    does not declare) and `needle` is not among them. A part in UTF-16, or one
    that declares another encoding (UTF-7 spells letters in other bytes), is
    always parsed."""
    from . import xmlsafe as _xmlsafe
    if _xmlsafe.byte_encoding(data) != "utf-8":
        return True
    m = _DECLARED_ENCODING.match(data[:1024])
    if m is not None and m.group(2).strip().lower() not in _PLAIN_ENCODINGS:
        return True
    return needle in data


def _main_tags(*locals_: str) -> tuple[str, ...]:
    """The "{namespace}local" tags of spreadsheet elements, in both the
    namespace Excel writes and the Strict one."""
    from . import opcxml as _opc
    return tuple("{%s}%s" % (ns, name) for name in locals_
                 for ns in _opc.MAIN_NAMESPACES)


def _local(tag: str) -> str:
    return tag.rpartition("}")[2]


def _sheet_merges_and_links(data: bytes) -> tuple[list[str], list[str]]:
    """(mergeCell refs, hyperlink refs) of one worksheet, read as XML by
    namespace (the rows are cleared as they pass, so a large sheet stays
    small in memory)."""
    from . import xmlsafe as _xmlsafe
    merges: list[str] = []
    links: list[str] = []
    for el in _xmlsafe.stream_elements(
            data, _main_tags("row", "mergeCell", "hyperlink"), part="a worksheet"):
        name = _local(el.tag)
        if name == "row":
            _discard(el)
        elif el.get("ref"):
            (merges if name == "mergeCell" else links).append(el.get("ref"))
    return merges, links


def _discard(el) -> None:
    """Drop an element already read, and its earlier siblings."""
    el.clear()
    parent = el.getparent()
    if parent is not None:
        while el.getprevious() is not None:
            del parent[0]


def _merge_tester(merges: list[str]):
    """A function (row, col) -> True for a cell that lies inside a merged
    range and is not its top-left cell, or None when the sheet merges
    nothing. Its attribute `rows` is the set of rows that can hold such a
    cell (None when a merge is too big to list)."""
    small: set[tuple[int, int]] = set()
    big: list[tuple[int, int, int, int]] = []
    for ref in merges:
        a, _sep, b = ref.strip().partition(":")
        top, bottom = _cell_rc(a), _cell_rc(b or a)
        if top is None or bottom is None:
            continue
        (r1, c1), (r2, c2) = top, bottom
        if (r2 - r1 + 1) * (c2 - c1 + 1) <= _MERGE_INDEX_CELLS:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) != (r1, c1):
                        small.add((r, c))
        else:
            big.append((r1, c1, r2, c2))
    if not small and not big:
        return None

    def hidden(r: int, c: int) -> bool:
        if (r, c) in small:
            return True
        return any(r1 <= r <= r2 and c1 <= c <= c2 and (r, c) != (r1, c1)
                   for r1, c1, r2, c2 in big)
    hidden.rows = {r for r, _c in small}
    hidden.big = [(r1, r2) for r1, _c1, r2, _c2 in big]
    return hidden


def _hidden_cell_has_content(data: bytes, hidden) -> bool:
    """True when a cell that `hidden` says lies under a merge (not its
    top-left cell) holds a formula, or a value with text in it. Reads the
    sheet row by row as XML, so the spelling of the markup (quotes,
    prefixes, encoding, character references) does not matter; only the rows
    that can hold a hidden cell are looked into. A cell with no `r` sits
    where the cell order puts it, as Excel and openpyxl read it."""
    from . import xmlsafe as _xmlsafe
    tags = _main_tags("row")
    c_tags = set(_main_tags("c"))
    f_tags = set(_main_tags("f"))
    value_tags = set(_main_tags("v", "is"))
    row_no = 0
    for row in _xmlsafe.stream_elements(data, tags, part="a worksheet"):
        parent = row.getparent()
        if parent is None or _local(parent.tag) != "sheetData":
            continue
        given = (row.get("r") or "").strip()
        row_no = int(given) if given.isdigit() else row_no + 1
        if row_no in hidden.rows or any(a <= row_no <= b for a, b in hidden.big):
            col = 0
            for cell in row:
                if cell.tag not in c_tags:
                    continue
                rc = _cell_rc(cell.get("r") or "")
                if rc is not None:
                    row_no, col = rc
                else:
                    col += 1
                if not hidden(row_no, col):
                    continue
                for child in cell:
                    if child.tag in f_tags:
                        return True
                    if child.tag in value_tags and "".join(child.itertext()).strip():
                        return True
        _discard(row)
    return False


def _sheet_comment_refs(name: str,
                        part_reader: Callable[[str], bytes]) -> list[str]:
    """The cells the worksheet's legacy notes (xl/comments*.xml, reached
    through the sheet's own relationships) are attached to."""
    from . import opcxml as _opc
    from . import xmlsafe as _xmlsafe
    rels_name = _opc.rels_part_for(name)
    try:
        rels = _opc.relationships(part_reader(rels_name), part=rels_name)
    except Exception:  # noqa: BLE001
        return []
    refs: list[str] = []
    for rel in rels:
        if rel.external or not rel.type.endswith("/relationships/comments"):
            continue
        try:
            data = part_reader(_opc.resolve_part(name, rel.target))
            refs.extend(el.get("ref") for el in _xmlsafe.stream_elements(
                data, _main_tags("comment"), part="a comments part")
                if el.get("ref"))
        except Exception:  # noqa: BLE001
            continue
    return refs


def _detect_merged_hidden(
    found: dict[str, "Hazard"], names: Iterable[str],
    part_reader: Callable[[str], bytes] | None,
) -> None:
    """Add a hazard when a value, a note or a hyperlink is stored under a
    merged range in a cell that is not the range's top-left one.

    Excel keeps that content (hidden) and shows it again when the cells are
    unmerged. openpyxl's model holds the top-left cell only: the other cells
    of a range are not read as cells, so their values and notes go with the
    save, and a hyperlink is re-attached to the top-left cell. Only the
    full rewrite does this; a cell-level edit copies the sheet through.

    The sheet, its relationships and its notes are read as XML by namespace
    (core/opcxml, core/xmlsafe), so single or double quotes, a prefixed
    spreadsheet namespace, UTF-16 and character references all read the same.
    A sheet whose bytes cannot spell "mergeCell" is passed over without
    parsing; one in another encoding is always parsed."""
    if part_reader is None:
        return
    parts: list[str] = []
    kinds: list[str] = []
    for name in names:
        if not _is_worksheet_part(name):
            continue
        try:
            data = part_reader(name)
        except Exception:  # noqa: BLE001
            continue
        if not _may_mention(data, b"mergeCell"):
            continue
        hit: set[str] = set()
        try:
            merges, links = _sheet_merges_and_links(data)
            hidden = _merge_tester(merges)
            if hidden is None:
                continue
            if _hidden_cell_has_content(data, hidden):
                hit.add("a value")
        except Exception:  # noqa: BLE001
            continue
        for ref in links:
            rc = _cell_rc(ref.split(":")[0])
            if rc is not None and hidden(*rc):
                hit.add("a hyperlink")
        for ref in _sheet_comment_refs(name, part_reader):
            rc = _cell_rc(ref)
            if rc is not None and hidden(*rc):
                hit.add("a note")
        if hit:
            parts.append(name)
            for k in ("a value", "a note", "a hyperlink"):
                if k in hit and k not in kinds:
                    kinds.append(k)
    if not parts:
        return
    spec = _SPEC_BY_KEY[MERGED_HIDDEN_KEY]
    found[MERGED_HIDDEN_KEY] = Hazard(
        spec.key, f"{spec.label}: {', '.join(kinds)}", spec.severity,
        spec.survives_openpyxl, spec.note, parts,
    )


def _detect_broken_shared(
    found: dict[str, "Hazard"], names: Iterable[str],
    part_reader: Callable[[str], bytes] | None,
) -> None:
    """Add a hazard when a worksheet holds shared-formula members whose first
    cell is missing or comes later (core/sharedformulas.py)."""
    if part_reader is None:
        return
    from . import sharedformulas as _sf
    parts: list[str] = []
    for name in names:
        if not _is_worksheet_part(name):
            continue
        try:
            data = part_reader(name)
            if _sf.broken_members(data):
                parts.append(name)
        except Exception:  # noqa: BLE001
            continue
    if not parts:
        return
    spec = _SPEC_BY_KEY[BROKEN_SHARED_KEY]
    found[BROKEN_SHARED_KEY] = Hazard(
        spec.key, spec.label, spec.severity, spec.survives_openpyxl,
        spec.note, parts,
    )


def _detect_doc_properties(
    found: dict[str, "Hazard"], names: Iterable[str],
    part_reader: Callable[[str], bytes] | None,
) -> None:
    """Add the document-properties hazard when app.xml carries authored text.

    Only non-empty fields count. openpyxl rewrites app.xml unconditionally,
    but rewriting <Company/> into nothing loses nothing, and this file's
    whole discipline is that a hazard means content actually goes missing.
    """
    if part_reader is None:
        return
    part = next((n for n in names if _normalized(n).lower() == _APP_PART),
                None)
    if part is None:
        return
    try:
        data = part_reader(part)
    except Exception:  # noqa: BLE001
        return
    text = data.decode("utf-8", "replace")
    carried = [tag for tag in _APP_PROPERTY_TAGS
               if re.search(rf"<{tag}>\s*\S[^<]*</{tag}>", text)]
    if not carried:
        return
    spec = _SPEC_BY_KEY[DOC_PROPERTIES_KEY]
    found[DOC_PROPERTIES_KEY] = Hazard(
        spec.key, f"{spec.label}: {', '.join(carried)}", spec.severity,
        spec.survives_openpyxl, spec.note, [part],
    )


# The knowledge table. survives_openpyxl is the CLAIM the fidelity harness
# checks. Ordered most-to-least destructive for readable reports.
HAZARD_SPECS: tuple[HazardSpec, ...] = (
    HazardSpec(
        "slicers", "slicers", SEV_DROPS, False,
        "openpyxl has no slicer model; slicers are removed on save.",
        _prefix("xl/slicers/", "xl/slicercaches/"),
    ),
    HazardSpec(
        "timelines", "timelines", SEV_DROPS, False,
        "openpyxl has no timeline model; timelines are removed on save.",
        _prefix("xl/timelines/", "xl/timelinecaches/"),
    ),
    HazardSpec(
        "activex", "ActiveX controls", SEV_DROPS, False,
        "ActiveX controls are dropped; even with keep_vba they break.",
        _prefix("xl/activex/"),
    ),
    HazardSpec(
        "ctrlprops", "form/ActiveX control properties", SEV_DROPS, False,
        "control property parts are dropped on the openpyxl round-trip.",
        _prefix("xl/ctrlprops/"),
    ),
    HazardSpec(
        "drawings", "drawings / shapes", SEV_DROPS, False,
        "openpyxl tutorial states shapes are lost from existing files; only "
        "charts are re-modeled, textboxes and shapes vanish. Legacy-comment "
        "VML anchors are excluded (openpyxl preserves legacy notes).",
        _drawings_matcher,
    ),
    HazardSpec(
        "media", "embedded media / images", SEV_DROPS, False,
        "images embedded in an existing file are not carried through the "
        "openpyxl load+save of that file.",
        _prefix("xl/media/"),
    ),
    HazardSpec(
        "threaded_comments", "threaded comments", SEV_DROPS, False,
        "openpyxl models legacy notes only; threadedComments parts and the "
        "xl/persons/ author registry they reference are lost.",
        _prefix("xl/threadedcomments/", "xl/persons/"),
    ),
    HazardSpec(
        "macro_sheets", "Excel 4.0 macro sheets", SEV_DROPS, False,
        "xl/macrosheets/ holds Excel 4.0 (XLM) macro sheets. openpyxl reads "
        "one as an ordinary worksheet and writes it back as one, so its "
        "macros stop being macros and the macro sheet part is gone. Cell "
        "edits keep it intact; nothing here ever runs it.",
        _prefix("xl/macrosheets/", "xl/dialogsheets/"),
    ),
    HazardSpec(
        "embeddings", "embedded OLE objects", SEV_DROPS, False,
        "xl/embeddings/ holds embedded documents (Word, PDF, other OLE "
        "payloads); openpyxl does not model them and drops them on save.",
        _prefix("xl/embeddings/"),
    ),
    HazardSpec(
        "power_query", "customXml (incl. Power Query DataMashup)", SEV_DROPS,
        False,
        "customXml parts are not modeled by openpyxl and are dropped. Power "
        "Query M lives here (the DataMashup item), but so do add-in stores "
        "and SharePoint property sets; a namelist scan cannot tell them "
        "apart, so every customXml part is flagged.",
        _prefix("customxml/"),
    ),
    HazardSpec(
        "connections", "data connections", SEV_DROPS, False,
        "xl/connections.xml is not modeled by openpyxl and is dropped.",
        _exact("xl/connections.xml"),
    ),
    HazardSpec(
        "query_tables", "legacy query tables", SEV_DROPS, False,
        "xl/queryTables/ (legacy web/database query definitions) is not "
        "modeled by openpyxl and is dropped.",
        _prefix("xl/querytables/"),
    ),
    HazardSpec(
        "data_model", "Power Pivot data model", SEV_DROPS, False,
        "xl/model/ holds the Power Pivot / data-model payload; openpyxl does "
        "not model it and drops it on save.",
        _prefix("xl/model/"),
    ),
    HazardSpec(
        "named_sheet_views", "named sheet views", SEV_DROPS, False,
        "xl/namedSheetViews/ (saved temporary filter/sort views) is not "
        "modeled by openpyxl and is dropped.",
        _prefix("xl/namedsheetviews/"),
    ),
    HazardSpec(
        "rich_metadata", "dynamic-array / rich-value metadata", SEV_DROPS,
        False,
        "xl/metadata.xml carries dynamic-array spill and rich-value metadata, "
        "and xl/richData/ holds the rich-value payloads (stock/geo types, "
        "images-in-cells); openpyxl models neither, so spill ranges and rich "
        "values lose their backing.",
        _basename_or_prefix(("metadata.xml", "richdata.xml"),
                            ("xl/richdata/",)),
    ),
    HazardSpec(
        IN_PART_EXT_KEY, "in-sheet extension blocks", SEV_DROPS, False,
        "worksheet-level extLst blocks (x14 conditional formatting, sparkline "
        "groups, slicer lists, protected ranges) live inside the worksheet "
        "part, which survives the save at full size. openpyxl parses them, "
        "warns that each is unsupported, and writes none of them back, so the "
        "rules are gone from a file that looks untouched. Detected by reading "
        "the worksheet, not by a part name.",
        _never,
    ),
    HazardSpec(
        PROTECTED_RANGES_KEY, "allow-edit ranges", SEV_DROPS, False,
        "a protected sheet's <protectedRanges> block names the ranges that "
        "stay editable, each with its own name and optional password. "
        "openpyxl has no model for it and writes none of it back, so the "
        "sheet returns protected with its exceptions gone and cells that "
        "were editable no longer are. The sheetProtection element itself "
        "survives, which is what makes the loss quiet.",
        _never,
    ),
    HazardSpec(
        SORT_STATE_KEY, "autofilter sort state", SEV_DROPS, False,
        "a worksheet's <sortState> records the sort the user applied under "
        "an autofilter: which column, which direction, over which range. "
        "Excel writes it beside the autoFilter element, openpyxl's reader "
        "looks for it only INSIDE that element, so it is read by nobody and "
        "written by nobody. The filter survives and the sort it was sorted "
        "by does not.",
        _never,
    ),
    HazardSpec(
        MERGED_HIDDEN_KEY, "content hidden under merged cells", SEV_DROPS,
        False,
        "a merged range keeps one visible cell, its top-left one, and Excel "
        "keeps whatever is stored in the other cells of the range (a value, "
        "a note or a hyperlink) out of sight, to show it again if the cells "
        "are unmerged. The file writer holds the top-left cell only: the "
        "others are not read, so their values and notes are lost, and a "
        "hyperlink on one of them is attached to the top-left cell instead.",
        _never,
    ),
    HazardSpec(
        BROKEN_SHARED_KEY, "shared formulas with no usable first cell",
        SEV_DEGRADES, False,
        "a shared formula is stored once, in its first cell, and the cells "
        "that use it carry only its index. Here that first cell is missing "
        "or comes after a cell that uses it, so the formula of those cells "
        "cannot be rebuilt: Excel drops it when it repairs the file and keeps "
        "the stored value, reads here show that value, and a full rewrite "
        "saves it as a plain value.",
        _never,
    ),
    HazardSpec(
        DOC_PROPERTIES_KEY, "document properties", SEV_DEGRADES, False,
        "docProps/app.xml is rewritten from an empty template on every save, "
        "so the authored fields it carries are cleared. The core properties "
        "in docProps/core.xml (title, subject, creator, keywords, category, "
        "description) survive.",
        _never,
    ),
    HazardSpec(
        "pivot", "pivot tables / caches", SEV_DEGRADES, True,
        "openpyxl has read-support for pivots (preserved, light edits only); "
        "attached slicers still die and pivot charts follow chart fidelity.",
        _prefix("xl/pivotcache/", "xl/pivottables/"),
    ),
    HazardSpec(
        "charts", "charts", SEV_DEGRADES, True,
        "charts are re-serialized through openpyxl's chart model; features "
        "the model does not know are degraded or lost.",
        _prefix("xl/charts/"),
    ),
    HazardSpec(
        "vba", "VBA project", SEV_CONDITIONAL, True,
        "vbaProject.bin survives ONLY with load_workbook(keep_vba=True) and an "
        ".xlsm or .xltm save; a default load drops it. Form controls break regardless.",
        _basename("vbaproject.bin"),
    ),
)

_SPEC_BY_KEY = {s.key: s for s in HAZARD_SPECS}


@dataclass
class Hazard:
    key: str
    label: str
    severity: str
    survives_openpyxl: bool
    note: str
    parts: list[str] = field(default_factory=list)


@dataclass
class HazardReport:
    path: str
    parts: list[str]
    hazards: list[Hazard]
    error: str | None = None
    #: uncompressed part sizes from the zip central directory (scan_path only;
    #: empty for scan_names). verify-after-write uses these to catch a fragile
    #: part REPLACED with empty content rather than dropped outright.
    sizes: dict[str, int] = field(default_factory=dict)
    #: vmlDrawing*.vml parts the content classifier verified as comment-only
    #: anchors and demoted from the drawings hazard.
    #: openpyxl regenerates them under its own commentsDrawing naming on
    #: save, so verify-after-write excuses exactly these names vanishing.
    comment_anchor_vml: list[str] = field(default_factory=list)
    #: WHY the scan failed: "corrupt" | "missing" | "locked" | None. A
    #: byte-range-locked (but healthy) file makes zipfile raise BadZipFile
    #: when the lock covers the central directory, and telling the user of
    #: a one-of-a-kind file the file is CORRUPT when an indexer/AV merely holds
    #: a region lock invites a destructive "repair".
    error_kind: str | None = None

    @property
    def clean(self) -> bool:
        """No fragile part present at all: openpyxl round-trip is safe."""
        return not self.hazards and self.error is None

    @property
    def lossy_keys(self) -> list[str]:
        """Hazards openpyxl would DROP outright (not merely degrade or
        conditionally preserve). These are the never-silently-lose set."""
        return [h.key for h in self.hazards if h.severity == SEV_DROPS]

    @property
    def would_lose(self) -> bool:
        return bool(self.lossy_keys)

    def labels(self) -> list[str]:
        return [h.label for h in self.hazards]

    def as_dict(self) -> dict:
        return {
            "path": self.path,
            "clean": self.clean,
            "would_lose": self.would_lose,
            "hazards": [
                {
                    "key": h.key,
                    "label": h.label,
                    "severity": h.severity,
                    "survives_openpyxl": h.survives_openpyxl,
                    "parts": h.parts,
                    "note": h.note,
                }
                for h in self.hazards
            ],
            "error": self.error,
        }


# --------------------------------------------------- chart-vs-shape drawings
#
# The over-refusal this fixes: every part under xl/drawings/ matched the
# SEV_DROPS "drawings" (shape-loss) spec, so a workbook whose ONLY drawing is a
# chart was treated as unrecoverable shape loss and refused. But a chart's
# drawing is just its anchor: openpyxl re-models charts and round-trips both the
# chart part and its anchor drawing byte-for-byte (fidelity table +
# empirically re-confirmed here: chart.xlsx loses NO parts on an openpyxl
# round-trip). Only genuine shapes (textboxes, rectangles, form/ActiveX
# controls) and pictures are dropped.
#
# THE HEURISTIC (cheap, and honest about its limits): a drawing part is a pure
# chart anchor iff BOTH hold:
#   (a) its relationship part (xl/drawings/_rels/drawingN.xml.rels) exists and
#       EVERY relationship in it targets a chart (Type ending "/chart"), and
#   (b) the drawing XML itself contains no inline shape content (sp / pic /
#       grpSp / cxnSp elements). This second check matters: a textbox drawn
#       NEXT TO a chart in the same drawing needs NO relationship of its own,
#       so a rels-only test would call that drawing chart-only and let the
#       textbox be silently dropped.
#   - No rels part at all  -> inline shapes (textbox/rectangle), a real drop
#     (this is exactly shape.xlsx: xl/drawings/drawing1.xml with no rels).
#   - rels all chart AND no inline shape elements -> chart anchor, NOT a
#     shape-loss drop; chart fidelity is already tracked by the separate
#     SEV_DEGRADES "charts" spec.
#   - rels present with any image / oleObject / control target -> a drop
#     (this is image.xlsx: a /image relationship; media survival is Pillow- and
#     authorship-dependent, so it stays conservatively flagged).
#
# LIMITS: the classification needs to read the tiny rels part and the drawing
# part, so it is not decidable from the central-directory namelist ALONE.
# scan_path supplies a reader (two extra small-part reads only when drawings
# are present, still milliseconds); scan_names WITHOUT a reader cannot see
# part content and so stays CONSERVATIVE, keeping every drawing flagged as a
# potential drop (a false positive costs a needless raw/COM route, never
# silent loss). An unparseable drawing part is likewise conservatively
# flagged. A chart drawn with no relationship part (not produced by Excel or
# openpyxl in practice) would be conservatively flagged.

_CHART_REL_SUFFIX = "/chart"
_TYPE_RE = re.compile(r'Type="([^"]+)"')


def _drawing_rels_for(drawing_part: str) -> str:
    head, base = drawing_part.rsplit("/", 1)
    return f"{head}/_rels/{base}.rels"


#: Local element names that mean a drawing carries inline shape content
#: openpyxl would drop: shapes, pictures, shape groups, connectors.
_SHAPE_LOCALS = frozenset({"sp", "pic", "grpSp", "cxnSp"})


def _drawing_has_shape_content(data: bytes) -> bool | None:
    """True when the drawing XML contains inline shape elements (dropped by
    openpyxl even when every relationship targets a chart). None when the part
    cannot be parsed (the caller stays conservative)."""
    try:
        root = _xmlsafe.fromstring(data, part="a drawing")
    except Exception:
        return None
    for el in root.iter():
        tag = el.tag
        local = tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""
        if local in _SHAPE_LOCALS:
            return True
    return False


def _drawing_is_chart_only(
    drawing_part: str, nameset: set[str],
    rels_reader: Callable[[str], bytes] | None,
) -> bool | None:
    """True: pure chart anchor (survives). False: shape/picture/control drawing
    (drops). None: cannot tell without reading part content (namelist-only
    path, or an unreadable/unparseable part)."""
    rels = _drawing_rels_for(drawing_part)
    if rels not in nameset:
        return False  # no rels: inline shapes, a genuine shape-loss drop
    if rels_reader is None:
        return None   # rels exists but its content is not readable here
    try:
        data = rels_reader(rels)
    except Exception:
        return None
    if not data:
        return None
    types = _TYPE_RE.findall(data.decode("utf-8", "replace"))
    if not types:
        return False
    if not all(t.rstrip("/").lower().endswith(_CHART_REL_SUFFIX)
               for t in types):
        return False
    # Every relationship targets a chart; now confirm the drawing XML itself
    # holds no inline shape (a textbox beside the chart needs no rel).
    try:
        drawing_data = rels_reader(drawing_part)
    except Exception:
        return None
    has_shape = _drawing_has_shape_content(drawing_data)
    if has_shape is None:
        return None
    return not has_shape


def _refine_drawings(
    found: dict[str, "Hazard"], nameset: set[str],
    rels_reader: Callable[[str], bytes] | None,
) -> None:
    """Drop pure chart-anchor drawings from the SEV_DROPS drawings hazard.
    Chart fidelity is covered by the separate SEV_DEGRADES charts spec, so a
    chart-only workbook must not be treated as shape loss."""
    hz = found.get("drawings")
    if hz is None:
        return
    drawing_xmls = [
        p for p in hz.parts
        if p.lower().startswith("xl/drawings/")
        and "/_rels/" not in p.lower()
        and p.lower().endswith(".xml")
    ]
    if not drawing_xmls:
        return
    chart_only = {
        d for d in drawing_xmls
        if _drawing_is_chart_only(d, nameset, rels_reader) is True
    }
    exclude = set(chart_only)
    for d in chart_only:
        rels = _drawing_rels_for(d)
        if rels in nameset:
            exclude.add(rels)
    # Subtract ONLY the demoted chart anchors (and their rels). The old
    # code rebuilt the part list from the surviving .xml drawings, which
    # silently dropped every NON-xml member -- an Excel-authored
    # vmlDrawing*.vml sharing the hazard with a chart anchor vanished from
    # the report, was therefore never warned, and its loss at verify time
    # was unexcusable even under allow_loss: the advertised remedy could
    # never succeed on such a file.
    kept = [p for p in hz.parts if p not in exclude]
    if not kept:
        del found["drawings"]  # every drawing is a pure chart anchor
        return
    hz.parts = kept


def _normalized(name: str) -> str:
    r"""A part name reduced to the spelling the matchers expect.

    Zip member names are attacker-controlled text, and a package can carry the
    same part under a spelling the prefix matchers miss: './xl/slicers/x.xml',
    '/xl/slicers/x.xml', 'xl//slicers/x.xml', or backslash separators.
    Those four evaded the hazard scan, so the mutation was
    NOT refused up front (the default-fail inventory check in
    verify-after-write still caught the loss and refused the save, which is
    why nothing was destroyed). Normalizing for the MATCH restores the loud,
    specific up-front refusal; the reported part keeps its real name, because
    the inventory diff compares against the real one."""
    low = name.replace("\\", "/")
    while low.startswith("./"):
        low = low[2:]
    low = low.lstrip("/")
    while "//" in low:
        low = low.replace("//", "/")
    return low


def scan_names(names: Iterable[str], path: str = "<names>", *,
               part_reader: Callable[[str], bytes] | None = None,
               rels_reader: Callable[[str], bytes] | None = None,
               ) -> HazardReport:
    """Classify a pre-listed set of archive member names. Split out so the
    detection logic is testable without a real file on disk.

    `part_reader`, when supplied (scan_path always does), reads a member's
    bytes. It serves two content-dependent checks that a namelist cannot
    answer: whether a drawing is a pure chart anchor rather than shape loss
    (chart-vs-shape fix), and whether a worksheet carries a top-level extLst
    whose contents openpyxl drops. Without it the scan reports only the
    name-detectable hazards; `rels_reader` is the former name of the same
    argument and is still accepted."""
    reader = part_reader or rels_reader
    names = list(names)
    nameset = set(names)
    found: dict[str, Hazard] = {}
    for name in names:
        probe = _normalized(name)
        for spec in HAZARD_SPECS:
            if spec.matcher(name) or (probe != name and spec.matcher(probe)):
                h = found.get(spec.key)
                if h is None:
                    h = Hazard(
                        spec.key, spec.label, spec.severity,
                        spec.survives_openpyxl, spec.note,
                    )
                    found[spec.key] = h
                h.parts.append(name)
    _refine_drawings(found, nameset, reader)
    demoted_vml = _refine_comment_vml(found, nameset, reader)
    _detect_in_part_extensions(found, names, reader)
    _detect_in_part_worksheet_elements(found, names, reader)
    _detect_merged_hidden(found, names, reader)
    _detect_broken_shared(found, names, reader)
    _detect_doc_properties(found, names, reader)
    ordered = [found[s.key] for s in HAZARD_SPECS if s.key in found]
    return HazardReport(path=path, parts=names, hazards=ordered,
                        comment_anchor_vml=demoted_vml)


def scan_path(path: str) -> HazardReport:
    """Open the .xlsx/.xlsm as a zip, classify its members, and return the
    report.

    A package whose XML declares a document type (<!DOCTYPE ...>) is
    refused here, before any part is parsed (core/xmlsafe.py): UnsafePackage
    is raised, not reported, because no caller may read such a package.

    The member list comes from the central directory alone (no decompression),
    and that is still all most hazards need. Two checks read part content: the
    chart-vs-shape refinement reads a drawing and its tiny rels part, and the
    in-part extension detector reads each worksheet to see whether openpyxl
    would drop a top-level extLst. The worksheet read is what makes this scan
    proportional to sheet size rather than the old flat milliseconds; the
    alternative is a data bar that vanishes without a word, so the read
    stays."""
    kind = ole_container_kind(path)
    if kind is not None:
        # An OLE container (legacy .xls, or an encrypted package). Modern
        # Excel's .xls files embed a zip fragment, so without this sniff the
        # zip open below SUCCEEDS on a BIFF file and the scan mistakes the
        # theme fragment for the workbook.
        return HazardReport(path=path, parts=[], hazards=[],
                            error=ole_refusal_text(kind, path),
                            error_kind="ole")
    try:
        with zipfile.ZipFile(path) as zf:
            _xmlsafe.check_package(zf, doc=path)
            # A ZIP directory entry (a name ending in '/') is not an OOXML part: Excel
            # ignores it and a full rewrite does not write it back, so it is left out
            # of the part inventory rather than reported as a lost part.
            infos = [i for i in zf.infolist() if not i.is_dir()]
            names = [i.filename for i in infos]
            sizes = {i.filename: i.file_size for i in infos}
            # Only drawings are cached: the refinement reads a drawing and
            # its rels part and may revisit them. Worksheets are read once
            # and released, so a big workbook is not held in memory whole
            # just to look at its extLst.
            cache: dict[str, bytes] = {}

            def reader(member: str, _zf=zf, _cache=cache) -> bytes:
                if not member.lower().startswith("xl/drawings/"):
                    return _zf.read(member)
                if member not in _cache:
                    _cache[member] = _zf.read(member)
                return _cache[member]

            rep = scan_names(names, path=path, part_reader=reader)
            rep.sizes = sizes
            return rep
    except zipfile.BadZipFile:
        # zipfile answers "not a zip" for a file it could not fully READ
        # too: a byte-range lock over the central directory surfaces as
        # BadZipFile. Probe before diagnosing corruption; the wrong word
        # here sends a panicked user to a destructive "repair".
        if file_read_blocked(path):
            return HazardReport(
                path=path, parts=[], hazards=[], error=_LOCKED_ERROR,
                error_kind="locked")
        return HazardReport(path=path, parts=[], hazards=[],
                            error="not a valid zip / OOXML package",
                            error_kind="corrupt")
    except FileNotFoundError:
        return HazardReport(path=path, parts=[], hazards=[],
                            error="file not found", error_kind="missing")
    except PermissionError:
        # PermissionError means two different things and they need
        # different words. On Windows it is usually a lock: another
        # process holds the file and the OS denies the open. On POSIX it
        # is usually the permission bits, and telling the user of a file
        # they cannot read that "another process holds a lock" sends them
        # hunting for a process that does not exist. Ask the filesystem
        # which one it is rather than guessing from the exception type.
        # os.access answers only the read-only attribute on
        # Windows, so a file whose permissions deny reading was reported as
        # "another process holds a lock"; the read probe tells them apart.
        from .access import read_denied
        if not os.access(path, os.R_OK) or read_denied(path):
            return HazardReport(
                path=path, parts=[], hazards=[], error=_UNREADABLE_ERROR,
                error_kind="permissions")
        return HazardReport(
            path=path, parts=[], hazards=[], error=_LOCKED_ERROR,
            error_kind="locked")
    except OSError as exc:
        return HazardReport(
            path=path, parts=[], hazards=[],
            error=f"cannot read the file ({type(exc).__name__}: {exc})",
            error_kind="os")


_UNREADABLE_ERROR = (
    "this process is not allowed to read the file. The file itself is "
    "intact and nothing is holding it; its permissions deny read access "
    "to the account this server runs as. Grant read permission, or run "
    "the server as an account that has it")

_LOCKED_ERROR = (
    "another process holds a lock on part of this file (an antivirus "
    "scanner, indexer, or app with a byte-range lock); the file itself may "
    "be perfectly intact. Wait for the other process to release it, or "
    "close whatever holds it open, then retry")


def file_read_blocked(path: str) -> bool:
    """True when reading the file end to end fails with an access error,
    i.e. some process holds a byte-range or exclusive lock. Only consulted
    on failure paths, so the full read is paid rarely."""
    try:
        with open(path, "rb") as fh:
            while fh.read(1 << 20):
                pass
        return False
    except PermissionError:
        return True
    except OSError as exc:
        return getattr(exc, "errno", None) == 13


def route(
    report: HazardReport,
    *,
    edit: str = "surgical",
    com_available: bool = False,
    allow_loss: bool = False,
) -> tuple[str, str]:
    """The per-tool routing decision. Returns
    (route_verb, reason).

    edit: 'surgical' (single cells/values/formulas, raw OOXML can preserve) or
    'structural' (row/col shifts, beyond safe raw surgery).
    """
    if report.error is not None:
        return ROUTE_REFUSE, f"cannot scan package: {report.error}"
    if report.clean:
        return ROUTE_OPENPYXL, "no fragile parts; openpyxl round-trip is safe"
    if not report.would_lose:
        # only degrade/conditional hazards (charts, pivots, VBA). openpyxl can
        # proceed but the caller must opt into fidelity loss on charts/pivots.
        if allow_loss:
            return ROUTE_OPENPYXL, (
                "only degradable parts present; proceeding with allow_loss"
            )
        if com_available:
            return ROUTE_COM, (
                "degradable parts present (" + ", ".join(report.labels())
                + "); COM preserves full fidelity"
            )
        return ROUTE_REFUSE, (
            "degradable parts present (" + ", ".join(report.labels())
            + "); pass allow_loss to accept the degradation, or leave the "
              "workbook to Excel. xl_app drives Excel for "
              "recalculation, pivots, export, rendering and conversion, "
              "and writes no cells, so it cannot carry this edit"
        )
    # would_lose: parts openpyxl drops outright.
    if edit == "surgical" and _raw_safe(report):
        return ROUTE_RAW_OOXML, (
            "lossy parts present (" + ", ".join(
                _SPEC_BY_KEY[k].label for k in report.lossy_keys)
            + "); surgical edit routes to byte-preserving raw OOXML"
        )
    if com_available:
        return ROUTE_COM, (
            "lossy parts present (" + ", ".join(
                _SPEC_BY_KEY[k].label for k in report.lossy_keys)
            + "); COM lets Excel save with everything intact"
        )
    if allow_loss:
        return ROUTE_OPENPYXL, (
            "lossy parts present but allow_loss set; the backup is the safety "
            "net and the loss is explicit"
        )
    return ROUTE_REFUSE, (
        "would lose " + ", ".join(
            _SPEC_BY_KEY[k].label for k in report.lossy_keys)
        + "; a surgical raw-OOXML edit preserves them, or pass allow_loss "
        "with a backup to accept the loss. xl_app writes no cells "
        "and is not a route for this"
    )


# ------------------------------------------------- refusal routes (honest)
# The gate used to offer "enable COM (com pack)" as the remedy for any
# refused file-tier mutation. The com pack holds eleven tools -- recalculate,
# com_manage_pivot, com_goal_seek, com_export_pdf, com_render_sheet,
# com_convert_format, com_save_with_password, com_set_sparkline, com_autofit,
# com_validate_opens_clean, com_status -- and NONE of them writes a cell, a
# format, or a row. A caller who enabled the pack as instructed found nothing
# there that could perform the operation that was refused. Naming a remedy
# that does not exist is worse than naming none, in a product whose whole
# claim is that its labels are honest.
#
# So the routes are now the two that are real: accept the loss explicitly,
# with the exact cost stated per class, or do not put this workbook through
# the file tier. Preservation is the fix and it is not in this release.

#: The route strings, kept as data so the refusal text, the .detail payload,
#: and the tests all read the same list.
ROUTE_ALLOW_LOSS = "allow_loss=true: proceed while accepting exactly that loss (backed up first; the loss is permanent in the saved file)"
ROUTE_LEAVE_ALONE = "make this edit in Excel itself, which preserves everything; copy_workbook branches the file byte-for-byte if you want a working copy"


def loss_costs(keys) -> list[str]:
    """"<label>: <what openpyxl does to it>" for each refused hazard class.

    The caller deciding whether to pass allow_loss is deciding to destroy
    specific content, so the refusal states which content and what happens
    to it, rather than a class name the caller has to go look up."""
    out = []
    for k in keys:
        spec = _SPEC_BY_KEY.get(k)
        if spec is None:
            continue
        out.append(f"{spec.label}: {spec.note.rstrip('.')}")
    return out


def refusal_routes() -> list[str]:
    """The routes a refused file-tier mutation actually has."""
    return [ROUTE_ALLOW_LOSS, ROUTE_LEAVE_ALONE]


def _raw_safe(report: HazardReport) -> bool:
    """Raw OOXML surgery preserves untouched parts byte-identical, so it is
    safe for any lossy hazard as long as the edit itself is surgical. The one
    class raw surgery must not touch structurally is the pivot cache index,
    but a surgical CELL edit never touches it, so raw is safe here."""
    return True


__all__ = [
    "HazardReport", "Hazard", "HazardSpec", "HAZARD_SPECS",
    "EXT_DROP_LABELS", "IN_PART_EXT_KEY",
    "PROTECTED_RANGES_KEY", "SORT_STATE_KEY", "DOC_PROPERTIES_KEY",
    "MERGED_HIDDEN_KEY", "BROKEN_SHARED_KEY",
    "loss_costs", "refusal_routes", "ROUTE_ALLOW_LOSS", "ROUTE_LEAVE_ALONE",
    "ole_container_kind", "ole_refusal_text", "refuse_ole_container",
    "is_strict_package", "refuse_strict_package", "STRICT_REFUSAL",
    "OLE_KIND_BIFF", "OLE_KIND_ENCRYPTED", "OLE_KIND_UNKNOWN",
    "scan_path", "scan_names", "route",
    "ROUTE_OPENPYXL", "ROUTE_RAW_OOXML", "ROUTE_COM", "ROUTE_REFUSE",
    "SEV_DROPS", "SEV_DEGRADES", "SEV_CONDITIONAL",
]
