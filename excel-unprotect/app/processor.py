"""
processor.py
------------
Strips sheetProtection and workbookProtection nodes from OOXML workbooks
and returns the display names of every sheet that was unlocked.
"""

import io
import re
import zipfile
from typing import Dict, List, Tuple

# ---------------------------------------------------------------------------
# Compiled patterns
# ---------------------------------------------------------------------------

_SP_SELF_CLOSE = re.compile(r"<(?:[\w]+:)?sheetProtection\b[^>]*/\s*>", re.IGNORECASE)
_SP_BLOCK      = re.compile(r"<(?:[\w]+:)?sheetProtection\b[^>]*>.*?</(?:[\w]+:)?sheetProtection\s*>", re.IGNORECASE | re.DOTALL)
_WP_SELF_CLOSE = re.compile(r"<(?:[\w]+:)?workbookProtection\b[^>]*/\s*>", re.IGNORECASE)
_WP_BLOCK      = re.compile(r"<(?:[\w]+:)?workbookProtection\b[^>]*>.*?</(?:[\w]+:)?workbookProtection\s*>", re.IGNORECASE | re.DOTALL)
_SHEET_PATH    = re.compile(r"^xl/worksheets/sheet\d+\.xml$", re.IGNORECASE)


def _strip_patterns(xml_bytes: bytes, *patterns: re.Pattern) -> Tuple[bytes, bool]:
    text = xml_bytes.decode("utf-8", errors="replace")
    original = text
    for pat in patterns:
        text = pat.sub("", text)
    return text.encode("utf-8"), text != original


def _build_sheet_name_map(zin: zipfile.ZipFile) -> Dict[str, str]:
    """
    Returns {normalised_zip_path: display_name}
    e.g. {"xl/worksheets/sheet2.xml": "Budget"}

    1. Parse xl/_rels/workbook.xml.rels  → rId → normalised path
    2. Parse xl/workbook.xml             → display name → rId
    3. Combine                           → path → display name
    """
    rid_to_path: Dict[str, str] = {}
    try:
        rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="replace")
        for m in re.finditer(
            r'<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bType="[^"]*worksheet[^"]*"[^>]*\bTarget="([^"]+)"',
            rels_xml, re.IGNORECASE
        ):
            rid, target = m.group(1), m.group(2)
            if not target.startswith("xl/"):
                target = "xl/" + target
            rid_to_path[rid] = target
    except Exception:
        pass

    path_to_name: Dict[str, str] = {}
    try:
        wb_xml = zin.read("xl/workbook.xml").decode("utf-8", errors="replace")
        for m in re.finditer(
            r'<(?:[\w]+:)?sheet\b[^>]*\bname="([^"]+)"[^>]*\br:id="([^"]+)"',
            wb_xml, re.IGNORECASE
        ):
            name, rid = m.group(1), m.group(2)
            if rid in rid_to_path:
                path_to_name[rid_to_path[rid]] = name
    except Exception:
        pass

    return path_to_name


def remove_protection(file_bytes: bytes) -> Tuple[bytes, Dict]:
    """
    Strip sheetProtection / workbookProtection from an OOXML workbook.

    Returns (processed_bytes, stats) where stats contains:
        sheets_unprotected    int   – number of sheets unlocked
        workbook_unprotected  bool  – whether workbook protection was removed
        unlocked_sheet_names  list  – display names of unlocked sheets
    """
    stats: Dict = {
        "sheets_unprotected":   0,
        "workbook_unprotected": False,
        "unlocked_sheet_names": [],
    }

    src = io.BytesIO(file_bytes)
    dst = io.BytesIO()

    with zipfile.ZipFile(src, "r") as zin, \
         zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:

        sheet_name_map = _build_sheet_name_map(zin)

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
                    display = sheet_name_map.get(
                        item.filename,
                        item.filename.split("/")[-1].replace(".xml", "")
                    )
                    stats["unlocked_sheet_names"].append(display)

            zout.writestr(item, data)

    return dst.getvalue(), stats
