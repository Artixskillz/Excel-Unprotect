"""
main.py
-------
FastAPI application for Excel Unprotect.

Endpoints
---------
POST   /api/upload              – upload & process an Excel file
GET    /api/files               – list all processed files (newest first)
GET    /api/download/{id}       – download the unlocked file
GET    /api/download/{id}/original – download the original file
DELETE /api/files/{id}          – delete a single entry + its files
DELETE /api/files               – purge all entries + files
GET    /health                  – health check
GET    /                        – SPA (served from /app/static)
"""

import json
import sqlite3
import uuid
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from processor import remove_protection

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

DATA_DIR    = Path("/data")
UPLOADS_DIR = DATA_DIR / "uploads"
DB_PATH     = DATA_DIR / "db.sqlite"

UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

app = FastAPI(title="Excel Unprotect", docs_url=None, redoc_url=None)


# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------


def _get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def _init_db() -> None:
    with _get_db() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS uploads (
                id                   TEXT PRIMARY KEY,
                original_filename    TEXT    NOT NULL,
                upload_time          TEXT    NOT NULL,
                file_size_bytes      INTEGER,
                sheets_unprotected   INTEGER DEFAULT 0,
                workbook_unprotected INTEGER DEFAULT 0,
                sheet_names          TEXT    DEFAULT '[]'
            )
            """
        )
        conn.commit()
        # Migration: add sheet_names column if it doesn't exist (for pre-existing DBs)
        try:
            conn.execute("ALTER TABLE uploads ADD COLUMN sheet_names TEXT DEFAULT '[]'")
            conn.commit()
        except sqlite3.OperationalError:
            pass  # column already exists


_init_db()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _validate_id(file_id: str) -> None:
    if not all(c in "0123456789abcdef-" for c in file_id.lower()):
        raise HTTPException(status_code=400, detail="Invalid file ID.")


def _delete_files_for_row(row) -> None:
    ext = Path(row["original_filename"]).suffix.lower()
    for suffix in ("", "_orig"):
        p = UPLOADS_DIR / f"{row['id']}{suffix}{ext}"
        if p.exists():
            p.unlink()


# ---------------------------------------------------------------------------
# API routes
# ---------------------------------------------------------------------------


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file type '{ext}'. Please upload an Excel file (.xlsx, .xlsm, .xltx, .xltm).",
        )

    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    try:
        processed_bytes, stats = remove_protection(contents)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"Failed to process file: {exc}")

    file_id  = str(uuid.uuid4())
    safe_name = Path(file.filename).name

    # Save both the original and the processed file
    (UPLOADS_DIR / f"{file_id}_orig{ext}").write_bytes(contents)
    (UPLOADS_DIR / f"{file_id}{ext}").write_bytes(processed_bytes)

    sheet_names_json = json.dumps(stats["unlocked_sheet_names"])

    with _get_db() as conn:
        conn.execute(
            "INSERT INTO uploads VALUES (?, ?, ?, ?, ?, ?, ?)",
            (
                file_id,
                safe_name,
                datetime.utcnow().isoformat(timespec="seconds"),
                len(contents),
                stats["sheets_unprotected"],
                1 if stats["workbook_unprotected"] else 0,
                sheet_names_json,
            ),
        )
        conn.commit()

    return {
        "id":                   file_id,
        "filename":             safe_name,
        "file_size_bytes":      len(contents),
        "sheets_unprotected":   stats["sheets_unprotected"],
        "workbook_unprotected": stats["workbook_unprotected"],
        "unlocked_sheet_names": stats["unlocked_sheet_names"],
    }


@app.get("/api/files")
def list_files():
    with _get_db() as conn:
        rows = conn.execute("SELECT * FROM uploads ORDER BY upload_time DESC").fetchall()

    result = []
    for r in rows:
        d = dict(r)
        try:
            d["unlocked_sheet_names"] = json.loads(d.get("sheet_names") or "[]")
        except Exception:
            d["unlocked_sheet_names"] = []
        # Let the frontend know whether the original file is still on disk
        ext = Path(d["original_filename"]).suffix.lower()
        d["has_original"] = (UPLOADS_DIR / f"{d['id']}_orig{ext}").exists()
        result.append(d)
    return result


@app.get("/api/download/{file_id}")
def download_unlocked(file_id: str):
    return _serve_file(file_id, original=False)


@app.get("/api/download/{file_id}/original")
def download_original(file_id: str):
    return _serve_file(file_id, original=True)


def _serve_file(file_id: str, original: bool) -> FileResponse:
    _validate_id(file_id)

    with _get_db() as conn:
        row = conn.execute("SELECT * FROM uploads WHERE id = ?", (file_id,)).fetchone()

    if not row:
        raise HTTPException(status_code=404, detail="Record not found.")

    ext = Path(row["original_filename"]).suffix.lower()

    if original:
        file_path     = UPLOADS_DIR / f"{file_id}_orig{ext}"
        download_name = row["original_filename"]
    else:
        file_path     = UPLOADS_DIR / f"{file_id}{ext}"
        download_name = f"unlocked_{row['original_filename']}"

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found on disk.")

    return FileResponse(
        str(file_path),
        filename=download_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.delete("/api/files/{file_id}")
def delete_file(file_id: str):
    _validate_id(file_id)

    with _get_db() as conn:
        row = conn.execute("SELECT * FROM uploads WHERE id = ?", (file_id,)).fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Record not found.")
        _delete_files_for_row(row)
        conn.execute("DELETE FROM uploads WHERE id = ?", (file_id,))
        conn.commit()

    return {"status": "deleted"}


@app.delete("/api/files")
def purge_all():
    with _get_db() as conn:
        rows = conn.execute("SELECT id, original_filename FROM uploads").fetchall()
        for row in rows:
            _delete_files_for_row(row)
        conn.execute("DELETE FROM uploads")
        conn.commit()

    return {"status": "purged"}


# ---------------------------------------------------------------------------
# Static files – must be mounted LAST (catch-all)
# ---------------------------------------------------------------------------

app.mount("/", StaticFiles(directory="/app/static", html=True), name="static")
