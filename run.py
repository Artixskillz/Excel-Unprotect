#!/usr/bin/env python3
"""
run.py — start Excel Unprotect locally without Docker.

Usage
-----
  python run.py              # default port 5757
  python run.py --port 8080  # custom port

Requirements
------------
  pip install -r excel-unprotect/requirements.txt
"""

import argparse
import os
import subprocess
import sys
from pathlib import Path

ROOT    = Path(__file__).parent
APP_DIR = ROOT / "excel-unprotect" / "app"
DATA_DIR = ROOT / "data"


def main():
    parser = argparse.ArgumentParser(description="Run Excel Unprotect locally")
    parser.add_argument("--port", type=int, default=5757, help="Port to listen on (default: 5757)")
    parser.add_argument("--host", default="127.0.0.1", help="Host to bind to (default: 127.0.0.1)")
    args = parser.parse_args()

    # Validate paths
    if not APP_DIR.exists():
        print("ERROR: Could not find excel-unprotect/app/")
        print("Make sure you are running this script from the repo root.")
        sys.exit(1)

    req_file = ROOT / "excel-unprotect" / "requirements.txt"
    if not req_file.exists():
        print("ERROR: Could not find excel-unprotect/requirements.txt")
        sys.exit(1)

    # Create local data directories
    (DATA_DIR / "uploads").mkdir(parents=True, exist_ok=True)

    # Tell the app where to store its data
    os.environ["DATA_DIR"] = str(DATA_DIR)

    print("=" * 50)
    print("  Excel Unprotect")
    print("=" * 50)
    print(f"  URL:       http://localhost:{args.port}")
    print(f"  Data:      {DATA_DIR}")
    print(f"  Press Ctrl+C to stop")
    print("=" * 50 + "\n")

    # Check uvicorn is available
    try:
        subprocess.run(
            [sys.executable, "-m", "uvicorn", "--version"],
            check=True, capture_output=True
        )
    except subprocess.CalledProcessError:
        print("uvicorn not found. Installing dependencies first...\n")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", str(req_file)],
            check=True
        )

    # Start the server
    subprocess.run([
        sys.executable, "-m", "uvicorn", "main:app",
        "--host", args.host,
        "--port", str(args.port),
        "--reload",
    ], cwd=str(APP_DIR))


if __name__ == "__main__":
    main()
