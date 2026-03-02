"""
processor.py
------------
Strips sheetProtection and workbookProtection nodes from OOXML workbooks.

OOXML files (.xlsx / .xlsm / .xltx / .xltm) are just ZIP archives containing
XML files.  We iterate over every entry, remove the relevant XML elements via
regex (they are self-closing in 99.9 % of real-world files), and re-zip.

No third-party Excel library is required – only stdlib.
"""

import io
import re
import zipfile
from typing import Dict, Tuple

# ---------------------------------------------------------------------------
# Compiled patterns – covers optional namespace prefix (e.g. <x:sheetProtection>)
# ---------------------------------------------------------------------------

_SP_SELF_CLOSE = re.compile(
    r"<(?:[\w]+:)?sheetProtection\b[^>]*/\s*>",
    re.IGNORECASE,
)
_SP_BLOCK = re.compile(
    r"<(?:[\w]+:)?sheetProtection\b[^>]*>.*?</(?:[\w]+:)?sheetProtection\s*>",
    re.IGNORECASE | re.DOTALL,
)
_WP_SELF_CLOSE = re.compile(
    r"<(?:[\w]+:)?workbookProtection\b[^>]*/\s*>",
    re.IGNORECASE,
)
_WP_BLOCK = re.compile(
    r"<(?:[\w]+:)?workbookProtection\b[^>]*>.*?</(?:[\w]+:)?workbookProtection\s*>",
    re.IGNORECASE | re.DOTALL,
)

# Matches xl/worksheets/sheet1.xml, sheet2.xml, etc.
_SHEET_PATH = re.compile(r"^xl/worksheets/sheet\d+\.xml$", re.IGNORECASE)


def _strip_patterns(xml_bytes: bytes, *patterns: re.Pattern) -> Tuple[bytes, bool]:
    """Apply all patterns to xml_bytes, return (result_bytes, was_changed)."""
    text = xml_bytes.decode("utf-8", errors="replace")
    original = text
    for pat in patterns:
        text = pat.sub("", text)
    return text.encode("utf-8"), text != original


def remove_protection(file_bytes: bytes) -> Tuple[bytes, Dict]:
    """
    Remove sheetProtection and workbookProtection from an OOXML workbook.

    Returns
    -------
    (processed_bytes, stats)
        processed_bytes – the modified workbook as bytes
        stats – dict with keys:
            sheets_unprotected  (int)  number of sheets whose protection was removed
            workbook_unprotected (bool) whether workbook-level protection was removed
    """
    stats: Dict = {"sheets_unprotected": 0, "workbook_unprotected": False}

    src = io.BytesIO(file_bytes)
    dst = io.BytesIO()

    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dst, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == "xl/workbook.xml":
                data, changed = _strip_patterns(data, _WP_SELF_CLOSE, _WP_BLOCK)
                if changed:
                    stats["workbook_unprotected"] = True

            elif _SHEET_PATH.match(item.filename):
                data, changed = _strip_patterns(data, _SP_SELF_CLOSE, _SP_BLOCK)
                if changed:
                    stats["sheets_unprotected"] += 1

            # Preserve the original ZipInfo (keeps timestamps, compression flags, etc.)
            zout.writestr(item, data)

    return dst.getvalue(), stats
