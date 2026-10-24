"""A made-up printer-settings part for the Excel fixtures.

Excel stores the active printer's DEVMODE structure in
xl/printerSettings/printerSettings1.bin. A workbook saved on a real machine
carries that machine's printer model and, in the driver's private data, often
the Windows user name and the computer name too. A fixture must not. This
module builds a DEVMODE from constants alone (a printer that does not exist,
no user, no computer) and rewrites a workbook so that its printer-settings part
is that DEVMODE. Every other member is copied across unchanged.

The product treats the part as an opaque blob, and the preserve tests check
that it comes through an edit byte for byte, so any well-formed DEVMODE serves.
Run this after authoring or re-saving a fixture in Excel:

    python -X utf8 tests/fixtures/build_printer_settings.py            # rewrite
    python -X utf8 tests/fixtures/build_printer_settings.py --check    # exit 1 if a part differs
    python -X utf8 tests/fixtures/build_printer_settings.py FILE ...   # named workbooks only
"""

from __future__ import annotations

import struct
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
PART_PREFIX = "xl/printerSettings/"
WORKBOOK_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam"}

DEVICE_NAME = "Example Laser 100"
FORM_NAME = "Letter"

#: Bytes of driver-private data after the 220-byte DEVMODEW structure.
DRIVER_EXTRA = 1024

#: wingdi.h dmFields flags for the members this module fills in.
DM_FIELDS = {
    "ORIENTATION": 0x1, "PAPERSIZE": 0x2, "PAPERLENGTH": 0x4, "PAPERWIDTH": 0x8,
    "SCALE": 0x10, "COPIES": 0x100, "DEFAULTSOURCE": 0x200, "PRINTQUALITY": 0x400,
    "COLOR": 0x800, "DUPLEX": 0x1000, "YRESOLUTION": 0x2000, "TTOPTION": 0x4000,
    "COLLATE": 0x8000, "FORMNAME": 0x10000,
}


def _wide(text: str, chars: int) -> bytes:
    """`text` as a fixed-width UTF-16LE field of `chars` characters."""
    assert len(text) < chars
    return text.encode("utf-16-le").ljust(chars * 2, b"\x00")


def _private_block() -> bytes:
    """Driver-private data of a driver that does not exist: a tagged header, a
    short capability table and a label, then zeros."""
    block = struct.pack("<4sHHII", b"EXLP", 1, 0, DRIVER_EXTRA, 0b111)
    block += struct.pack("<I3I", 3, 300, 600, 1200)        # supported resolutions, dpi
    block += struct.pack("<I2H", 2, 7, 1)                   # trays: automatic, upper
    block += _wide(DEVICE_NAME, 32)
    return block.ljust(DRIVER_EXTRA, b"\x00")


def build_devmode() -> bytes:
    """The DEVMODEW (220 bytes) followed by DRIVER_EXTRA bytes of private data.
    Landscape Letter, 600 dpi, color, one copy, automatic tray, one-sided: the
    page setup the fixture's sheet already declares."""
    flags = sum(DM_FIELDS.values())
    head = (
        _wide(DEVICE_NAME, 32)                                  # dmDeviceName
        + struct.pack("<HHHHI",
                      0x0401,                                   # dmSpecVersion
                      0x0100,                                   # dmDriverVersion
                      220,                                      # dmSize
                      DRIVER_EXTRA,                             # dmDriverExtra
                      flags)                                    # dmFields
        + struct.pack("<8h",
                      2,                                        # dmOrientation: landscape
                      1,                                        # dmPaperSize: Letter
                      2794, 2159,                               # dmPaperLength, dmPaperWidth (0.1 mm)
                      100,                                      # dmScale
                      1,                                        # dmCopies
                      7,                                        # dmDefaultSource: automatic
                      600)                                      # dmPrintQuality: dpi
        + struct.pack("<5h",
                      2,                                        # dmColor: color
                      1,                                        # dmDuplex: one-sided
                      600,                                      # dmYResolution
                      1,                                        # dmTTOption: bitmap
                      0)                                        # dmCollate: off
        + _wide(FORM_NAME, 32)                                  # dmFormName
        + struct.pack("<H9I", *([0] * 10))                      # log pixels, display members
        + struct.pack("<4I", 0, 0, 0, 0)                        # reserved, panning
    )
    assert len(head) == 220, len(head)
    return head + _private_block()


def fixtures() -> list[Path]:
    return sorted(p for p in HERE.rglob("*")
                  if p.suffix.lower() in WORKBOOK_SUFFIXES and zipfile.is_zipfile(p))


def differing(path: Path, blob: bytes) -> list[str]:
    """Printer-settings members of `path` whose bytes are not `blob`."""
    with zipfile.ZipFile(path) as z:
        return [i.filename for i in z.infolist()
                if i.filename.startswith(PART_PREFIX) and z.read(i.filename) != blob]


def replace_printer_settings(path: Path, blob: bytes) -> bool:
    """Make every printer-settings member of `path` equal `blob`, in place.
    Returns True when the file changed. Member order, names, dates and
    attributes are kept; only the printer-settings bytes differ."""
    with zipfile.ZipFile(path) as z:
        members = [(i, z.read(i.filename)) for i in z.infolist()]
    changed = False
    out = []
    for info, data in members:
        if info.filename.startswith(PART_PREFIX) and data != blob:
            data, changed = blob, True
        out.append((info, data))
    if not changed:
        return False
    tmp = path.with_name(path.name + ".new")
    with zipfile.ZipFile(tmp, "w") as zo:
        for info, data in out:
            zo.writestr(info, data)
    tmp.replace(path)
    return True


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    check = "--check" in args
    named = [Path(a) for a in args if not a.startswith("--")]
    blob = build_devmode()
    bad = 0
    for path in named or fixtures():
        if check:
            for member in differing(path, blob):
                print(f"differs: {path.name}: {member}")
                bad += 1
        elif replace_printer_settings(path, blob):
            print(f"rewrote {path.name}")
    if check:
        print("printer-settings parts all match the made-up DEVMODE" if not bad
              else f"{bad} printer-settings part(s) differ; run without --check to rewrite")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
