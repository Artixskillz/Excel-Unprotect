# 📊 Excel Unprotect

[![Docker Pulls](https://img.shields.io/docker/pulls/artixskillz/excel-unprotect?style=flat-square&logo=docker&color=2496ED)](https://hub.docker.com/r/artixskillz/excel-unprotect)
[![Docker Image Size](https://img.shields.io/docker/image-size/artixskillz/excel-unprotect/latest?style=flat-square&logo=docker&color=2496ED)](https://hub.docker.com/r/artixskillz/excel-unprotect)
[![Build & Push](https://img.shields.io/github/actions/workflow/status/Artixskillz/Excel-Unprotect/docker-publish.yml?style=flat-square&logo=github-actions&label=build)](https://github.com/Artixskillz/Excel-Unprotect/actions)
[![License: MIT](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)

A clean, self-hosted web UI for removing **sheet protection** and **workbook protection** from Excel files (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`) — no password required, no Excel installation needed.

> Works entirely through XML manipulation inside the OOXML file format. No third-party Excel libraries required.

---

## ✨ Features

- 🔓 Removes **sheet-level protection** from every sheet in the workbook
- 🔓 Removes **workbook-level protection** (structure lock)
- 📦 **Batch upload** — drop or select multiple files at once with a live per-file progress queue
- 📁 **Upload history** — every processed file is stored and re-downloadable at any time
- ⬇️ Download both the **original** and **unlocked** file from history
- 🔍 **Search & sort history** — live filter by filename; sort by name, date, size, or sheets unlocked
- 🌙 **Dark mode** toggle (remembers your preference)
- 🗑️ Delete individual files or clear all history (with confirmation)
- 🏷️ Shows which sheet names were unlocked
- 💻 **Installable as a desktop app** (PWA) — works in Chrome & Edge, runs in its own window
- 📱 **Mobile responsive** — history converts to a clean card layout on small screens
- 🔒 **Encryption detection** — gives a helpful error message if a file has a file-open password instead of crashing
- 🖥️ Clean, modern web UI with animated drag & drop support
- 🐳 Single Docker container, no dependencies to install
- 💾 Persistent storage via Docker volume — history survives restarts

---

## 🚀 Quick Start

### Option 1 — Clone & run (build from source)

```bash
git clone https://github.com/Artixskillz/Excel-Unprotect.git
cd Excel-Unprotect
docker compose up --build -d
```

Then open **[http://localhost:5757](http://localhost:5757)** in your browser. That's it — no Docker Hub account needed.

---

### Option 2 — Docker Compose (pull pre-built image)

No need to clone the repo. Just save this as `docker-compose.yml` and run `docker compose up -d`:

```yaml
version: "3.9"
services:
  excel-unprotect:
    image: artixskillz/excel-unprotect:latest
    container_name: excel-unprotect
    ports:
      - "5757:8000"
    volumes:
      - excel-data:/data
    restart: unless-stopped
volumes:
  excel-data:
```

---

### Option 3 — Docker run (one-liner)

```bash
docker run -d \
  --name excel-unprotect \
  -p 5757:8000 \
  -v excel-data:/data \
  --restart unless-stopped \
  artixskillz/excel-unprotect:latest
```

Then open **[http://localhost:5757](http://localhost:5757)**.

---

### Option 4 — Portainer Stack

1. In Portainer, go to **Stacks → + Add stack**
2. Paste the `docker-compose.yml` from Option 2 above
3. Click **Deploy the stack**

---

### Option 5 — No Docker, plain Python (Windows & Mac)

**Requires:** Python 3.11+ → [python.org/downloads](https://www.python.org/downloads/)

**Step 1 — Clone the repo and install dependencies**

```bash
git clone https://github.com/Artixskillz/Excel-Unprotect.git
cd Excel-Unprotect
pip install -r excel-unprotect/requirements.txt
```

> 💡 Recommended: use a virtual environment first
> ```bash
> python -m venv venv
> # Windows:
> venv\Scripts\activate
> # Mac / Linux:
> source venv/bin/activate
> ```

**Step 2 — Run it**

```bash
python run.py
```

Then open **[http://localhost:5757](http://localhost:5757)**.

Files are stored in a `data/` folder next to the repo. To stop the server press `Ctrl+C`.

**Optional flags:**
```bash
python run.py --port 8080          # use a different port
python run.py --host 0.0.0.0       # expose to your local network
```

---

## 💻 Install as a Desktop App (PWA)

When you open the app in **Chrome or Edge**, the browser will offer to install it as a standalone desktop app. You can also click the **"Install App"** button that appears in the top-right of the header.

Once installed, Excel Unprotect runs in its own window (no browser tabs or address bar) and appears in your taskbar / Start menu / Dock like any native app.

---

## 📖 How It Works

Excel files (`.xlsx` and friends) are ZIP archives containing XML files. Protection in Excel is stored as XML elements:

| Protection type | XML element | Location |
|---|---|---|
| Sheet protection | `<sheetProtection ... />` | `xl/worksheets/sheet*.xml` |
| Workbook protection | `<workbookProtection ... />` | `xl/workbook.xml` |

This tool unzips the file, strips those XML elements using regex, and re-zips — preserving all content, formatting, and formulas. No password is needed because these protections don't encrypt the file; they only instruct Excel's UI to restrict editing.

> **Note:** This tool removes *editing protection* only. It does **not** remove *file open passwords* (encryption), as those require the actual password to decrypt. If you upload a file with an open password, the app will detect it and show a clear error explaining how to fix it.

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3.11 + FastAPI |
| File processing | stdlib only (`zipfile` + `re`) |
| Database | SQLite (via stdlib `sqlite3`) |
| Frontend | Vanilla HTML / CSS / JS (no frameworks) |
| PWA | Web App Manifest + Service Worker |
| Container | Docker (python:3.11-slim base) |

---

## 📁 Project Structure

```
Excel-Unprotect/
├── run.py                   # Local launcher (no Docker needed)
├── docker-compose.yml       # Root compose — build from source
└── excel-unprotect/
    ├── Dockerfile
    ├── docker-compose.yml   # Compose — pull pre-built image
    ├── requirements.txt
    └── app/
        ├── main.py          # FastAPI app & API routes
        ├── processor.py     # Core XML stripping + encryption detection
        └── static/
            ├── index.html   # Single-page UI
            ├── style.css    # Styles (light/dark themes, responsive)
            ├── app.js       # Client-side logic
            ├── manifest.json# PWA manifest
            ├── sw.js        # Service worker (offline cache)
            └── icon.svg     # PWA app icon
```

---

## 🔄 Updating

**If you cloned the repo (build from source):**
```bash
git pull
docker compose up --build -d
```

**If you're pulling the pre-built image:**
```bash
docker compose pull && docker compose up -d
```

**If running locally with Python:**
```bash
git pull
python run.py
```

---

## 🤝 Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

---

## 📄 License

[MIT](LICENSE) © Artixskillz
