"""
main.py
-------
FastAPI application for Excel Unprotect.

Endpoints
---------
POST /api/upload          – upload & process an Excel file
GET  /api/files           – list all processed files (newest first)
GET  /api/download/{id}   – download a processed file
GET  /                    – SPA (served from /app/static)
"""

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

DATA_DIR = Path("/data")
UPLOADS_DIR = DATA_DIR / "uploads"
DB_PATH = DATA_DIR / "db.sqlite"

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
                workbook_unprotected INTEGER DEFAULT 0
            )
            """
        )
        conn.commit()


_init_db()


# ---------------------------------------------------------------------------
# API routes
# ---------------------------------------------------------------------------


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
        raise HTTPException(
            status_code=422,
            detail=f"Failed to process file: {exc}",
        )

    file_id = str(uuid.uuid4())
    safe_name = Path(file.filename).name  # strip any path components
    stored_path = UPLOADS_DIR / f"{file_id}{ext}"

    stored_path.write_bytes(processed_bytes)

    with _get_db() as conn:
        conn.execute(
            "INSERT INTO uploads VALUES (?, ?, ?, ?, ?, ?)",
            (
                file_id,
                safe_name,
                datetime.utcnow().isoformat(timespec="seconds"),
                len(contents),
                stats["sheets_unprotected"],
                1 if stats["workbook_unprotected"] else 0,
            ),
        )
        conn.commit()

    return {
        "id": file_id,
        "filename": safe_name,
        "file_size_bytes": len(contents),
        "sheets_unprotected": stats["sheets_unprotected"],
        "workbook_unprotected": stats["workbook_unprotected"],
    }


@app.get("/api/files")
def list_files():
    with _get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM uploads ORDER BY upload_time DESC"
        ).fetchall()
    return [dict(r) for r in rows]


@app.get("/api/download/{file_id}")
def download_file(file_id: str):
    # Basic validation to prevent path-traversal
    if not all(c in "0123456789abcdef-" for c in file_id.lower()):
        raise HTTPException(status_code=400, detail="Invalid file ID.")

    with _get_db() as conn:
        row = conn.execute(
            "SELECT * FROM uploads WHERE id = ?", (file_id,)
        ).fetchone()

    if not row:
        raise HTTPException(status_code=404, detail="Record not found.")

    ext = Path(row["original_filename"]).suffix.lower()
    file_path = UPLOADS_DIR / f"{file_id}{ext}"

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found on disk.")

    return FileResponse(
        str(file_path),
        filename=f"unlocked_{row['original_filename']}",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Static files – must be mounted LAST (catch-all)
# ---------------------------------------------------------------------------

app.mount("/", StaticFiles(directory="/app/static", html=True), name="static")
