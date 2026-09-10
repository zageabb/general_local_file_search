#!/usr/bin/env python3
"""
Generic Local File Search GUI
=============================

A Tkinter desktop search tool for large local/OneDrive folders.

Features
--------
- Generic free-text searches rather than one hard-coded supplier search.
- Builds a persistent SQLite content index so later searches are fast.
- Incremental re-indexing: unchanged files are not re-read.
- Searches Office files, PDFs, spreadsheets and normal text/code files.
- Optional local Ollama re-ranking.
- Default Ollama: http://192.168.1.249:11434
- Default model: llama3.2:latest
- Uses LibreOffice as a fallback converter for legacy Office formats.
- Does not modify source files.
- Avoids openpyxl keep_vba=True, preventing the noisy Python 3.13
  ZipFile.__del__ cleanup warning seen during large XLSM scans.

Version: 2.3.0
"""

from __future__ import annotations

import csv
import json
import os
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import httpx
except ImportError:
    httpx = None

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

try:
    import xlrd
except ImportError:
    xlrd = None

try:
    from pyxlsb import open_workbook as open_xlsb
except ImportError:
    open_xlsb = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from pptx import Presentation
except ImportError:
    Presentation = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None


VERSION = "2.2.0"
DEFAULT_OLLAMA_URL = "http://192.168.1.249:11434"
DEFAULT_OLLAMA_MODEL = "llama3.2:latest"

APP_DIR = Path.home() / ".generic_file_search"
DB_PATH = APP_DIR / "file_index.sqlite3"

MAX_INDEX_CHARS_DEFAULT = 120_000
MAX_FILE_MB_DEFAULT = 150

fPAEADSHEET_EXTS = {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods", ".csv", ".tsv"}
OFFICE_EXTS = {".docx", ".doc", ".pptx", ".ppt", ".pdf"} | fPAEADSHEET_EXTS
TEXT_EXTS = {
    ".txt", ".md", ".markdown", ".csv", ".tsv", ".json", ".jsonl",
    ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf", ".log", ".sql",
    ".py", ".js", ".jsx", ".ts", ".tsx", ".css", ".scss", ".html", ".htm",
    ".java", ".cs", ".cpp", ".c", ".h", ".hpp", ".sh", ".bash", ".zsh",
    ".ps1", ".bat", ".cmd", ".toml", ".properties", ".rst",
}
SUPPORTED_EXTS = OFFICE_EXTS | TEXT_EXTS

FILE_TYPE_PAESETS: dict[str, set[str] | None] = {
    "All supported files": None,
    "Spreadsheets": {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods", ".csv", ".tsv"},
    "Excel workbooks": {".xlsx", ".xlsm", ".xls", ".xlsb"},
    "CSV / TSV": {".csv", ".tsv"},
    "Documents": {".docx", ".doc", ".pdf", ".txt", ".md", ".markdown"},
    "Word documents": {".docx", ".doc"},
    "Presentations": {".pptx", ".ppt"},
    "PowerPoint": {".pptx", ".ppt"},
    "PDF": {".pdf"},
    "Text / Markdown": {".txt", ".md", ".markdown"},
    "Code / config": {
        ".json", ".jsonl", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf",
        ".log", ".sql", ".py", ".js", ".jsx", ".ts", ".tsx", ".css", ".scss",
        ".html", ".htm", ".java", ".cs", ".cpp", ".c", ".h", ".hpp", ".sh",
        ".bash", ".zsh", ".ps1", ".bat", ".cmd", ".toml", ".properties", ".rst",
    },
}

SKIP_DIRS = {
    ".git", ".svn", ".hg", "node_modules", "__pycache__", ".venv", "venv",
    ".idea", ".vscode", "$recycle.bin", "system volume information",
}

STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "has",
    "have", "i", "in", "is", "it", "of", "on", "or", "that", "the", "this",
    "to", "was", "were", "what", "which", "with", "about", "find", "file",
    "files", "looking", "search", "show", "me", "my",
}


@dataclass
class SearchResult:
    path: str
    extension: str
    modified: float
    size: int
    local_score: float
    final_score: float
    snippet: str
    status: str
    ollama_score: int | None = None
    ollama_reason: str = ""


class StopRequested(RuntimeError):
    pass


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def human_size(value: int) -> str:
    units = ["B", "KB", "MB", "GB"]
    size = float(value)
    for unit in units:
        if size < 1024 or unit == units[-1]:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{value} B"


def discover_onedrive_roots() -> list[Path]:
    roots: list[Path] = []
    for name in ("OneDriveCommercial", "OneDriveConsumer", "OneDrive"):
        value = os.environ.get(name)
        if value:
            roots.append(Path(value).expanduser())

    home = Path.home()
    roots.extend([home / "OneDrive", home / "OneDrive - Hitachi Energy"])

    cloud = home / "Library" / "CloudStorage"
    if cloud.is_dir():
        try:
            roots.extend(
                p for p in cloud.iterdir()
                if p.is_dir() and p.name.casefold().startswith("onedrive")
            )
        except OSError:
            pass

    seen: set[str] = set()
    result: list[Path] = []
    for path in roots:
        try:
            key = str(path.resolve()).casefold()
        except OSError:
            key = str(path).casefold()
        if path.is_dir() and key not in seen:
            seen.add(key)
            result.append(path)
    return result


def find_soffice() -> str | None:
    candidates = [
        shutil.which("libreoffice"),
        shutil.which("soffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def db_connect() -> sqlite3.Connection:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute(
        """
        CAEATE TABLE IF NOT EXISTS files (
            path TEXT PRIMARY KEY,
            root TEXT NOT NULL,
            extension TEXT NOT NULL,
            size INTEGER NOT NULL,
            mtime_ns INTEGER NOT NULL,
            modified AEAL NOT NULL,
            indexed_at TEXT NOT NULL,
            status TEXT NOT NULL,
            error TEXT NOT NULL DEFAULT '',
            content TEXT NOT NULL DEFAULT ''
        )
        """
    )
    conn.execute("CAEATE INDEX IF NOT EXISTS idx_files_root ON files(root)")
    conn.execute("CAEATE INDEX IF NOT EXISTS idx_files_modified ON files(modified)")
    conn.commit()
    return conn


def close_openpyxl_book(book: Any) -> None:
    """Defensively close all openpyxl archive handles."""
    try:
        book.close()
    except Exception:
        pass

    for attr in ("vba_archive", "_archive"):
        archive = getattr(book, attr, None)
        if archive is not None:
            try:
                archive.close()
            except Exception:
                pass
            try:
                setattr(book, attr, None)
            except Exception:
                pass


def append_text(parts: list[str], value: Any, current: int, limit: int) -> int:
    if value is None or current >= limit:
        return current
    text = re.sub(r"\s+", " ", str(value)).strip()
    if not text:
        return current
    room = limit - current
    text = text[:room]
    parts.append(text)
    return current + len(text) + 1


def extract_xlsx(path: Path, limit: int) -> str:
    if load_workbook is None:
        raise RuntimeError("openpyxl is not installed")

    # Read-only search: never use keep_vba=True. It creates an unnecessary
    # secondary ZIP archive and can cause ZipFile.__del__ warnings on cleanup.
    book = load_workbook(
        path,
        read_only=True,
        data_only=False,
        keep_vba=False,
        keep_links=False,
    )
    parts: list[str] = []
    count = 0
    try:
        for ws in book.worksheets:
            count = append_text(parts, f"[Worksheet: {ws.title}]", count, limit)
            for row in ws.iter_rows():
                values: list[str] = []
                for cell in row:
                    if cell.value is not None:
                        value = re.sub(r"\s+", " ", str(cell.value)).strip()
                        if value:
                            values.append(value)
                if values:
                    count = append_text(parts, " | ".join(values), count, limit)
                if count >= limit:
                    break
            if count >= limit:
                break
    finally:
        close_openpyxl_book(book)
    return "\n".join(parts)


def extract_xls(path: Path, limit: int) -> str:
    if xlrd is None:
        raise RuntimeError("xlrd is not installed")
    book = xlrd.open_workbook(path, on_demand=True)
    parts: list[str] = []
    count = 0
    try:
        for sheet_name in book.sheet_names():
            ws = book.sheet_by_name(sheet_name)
            count = append_text(parts, f"[Worksheet: {sheet_name}]", count, limit)
            for row_no in range(ws.nrows):
                values = [
                    re.sub(r"\s+", " ", str(ws.cell_value(row_no, col))).strip()
                    for col in range(ws.ncols)
                    if ws.cell_value(row_no, col) not in (None, "")
                ]
                if values:
                    count = append_text(parts, " | ".join(values), count, limit)
                if count >= limit:
                    break
            if count >= limit:
                break
    finally:
        try:
            book.release_resources()
        except Exception:
            pass
    return "\n".join(parts)


def extract_xlsb(path: Path, limit: int) -> str:
    if open_xlsb is None:
        raise RuntimeError("pyxlsb is not installed")
    parts: list[str] = []
    count = 0
    with open_xlsb(path) as book:
        for sheet_name in book.sheets:
            count = append_text(parts, f"[Worksheet: {sheet_name}]", count, limit)
            with book.get_sheet(sheet_name) as ws:
                for row in ws.rows():
                    values = [
                        re.sub(r"\s+", " ", str(cell.v)).strip()
                        for cell in row
                        if cell.v not in (None, "")
                    ]
                    if values:
                        count = append_text(parts, " | ".join(values), count, limit)
                    if count >= limit:
                        break
            if count >= limit:
                break
    return "\n".join(parts)


def extract_docx(path: Path, limit: int) -> str:
    if Document is None:
        raise RuntimeError("python-docx is not installed")
    doc = Document(path)
    parts: list[str] = []
    count = 0

    for paragraph in doc.paragraphs:
        count = append_text(parts, paragraph.text, count, limit)
        if count >= limit:
            return "\n".join(parts)

    for table in doc.tables:
        for row in table.rows:
            count = append_text(
                parts,
                " | ".join(cell.text for cell in row.cells),
                count,
                limit,
            )
            if count >= limit:
                return "\n".join(parts)
    return "\n".join(parts)


def extract_pptx(path: Path, limit: int) -> str:
    if Presentation is None:
        raise RuntimeError("python-pptx is not installed")
    deck = Presentation(path)
    parts: list[str] = []
    count = 0
    for number, slide in enumerate(deck.slides, start=1):
        count = append_text(parts, f"[Slide {number}]", count, limit)
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                count = append_text(parts, shape.text, count, limit)
            if count >= limit:
                return "\n".join(parts)
    return "\n".join(parts)


def extract_pdf(path: Path, limit: int) -> str:
    if PdfReader is None:
        raise RuntimeError("pypdf is not installed")
    reader = PdfReader(path)
    parts: list[str] = []
    count = 0
    for page_no, page in enumerate(reader.pages, start=1):
        count = append_text(parts, f"[Page {page_no}]", count, limit)
        count = append_text(parts, page.extract_text() or "", count, limit)
        if count >= limit:
            break
    return "\n".join(parts)


def extract_delimited(path: Path, limit: int) -> str:
    delimiter = "\t" if path.suffix.casefold() == ".tsv" else ","
    parts: list[str] = []
    count = 0
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as handle:
        reader = csv.reader(handle, delimiter=delimiter)
        for row in reader:
            count = append_text(parts, " | ".join(row), count, limit)
            if count >= limit:
                break
    return "\n".join(parts)


def extract_plain_text(path: Path, limit: int) -> str:
    with path.open("r", encoding="utf-8", errors="replace") as handle:
        return handle.read(limit)


def convert_with_libreoffice(path: Path) -> tuple[tempfile.TemporaryDirectory, Path]:
    executable = find_soffice()
    if not executable:
        raise RuntimeError("LibreOffice/soffice is not installed")

    suffix = path.suffix.casefold()
    target = {
        ".doc": "docx",
        ".ppt": "pptx",
        ".ods": "xlsx",
        ".xls": "xlsx",
    }.get(suffix, "xlsx")

    temp = tempfile.TemporaryDirectory(prefix="generic-search-convert-")
    outdir = Path(temp.name)
    result = subprocess.run(
        [executable, "--headless", "--convert-to", target, "--outdir", str(outdir), str(path)],
        capture_output=True,
        text=True,
        timeout=180,
    )
    converted = outdir / f"{path.stem}.{target}"
    if result.returncode != 0 or not converted.exists():
        temp.cleanup()
        message = (result.stderr or result.stdout or "conversion failed").strip()
        raise RuntimeError(f"LibreOffice conversion failed: {message}")
    return temp, converted


def extract_content(path: Path, limit: int) -> tuple[str, str]:
    ext = path.suffix.casefold()

    try:
        if ext in {".xlsx", ".xlsm"}:
            return extract_xlsx(path, limit), "direct"
        if ext == ".xls":
            try:
                return extract_xls(path, limit), "direct"
            except Exception:
                temp, converted = convert_with_libreoffice(path)
                try:
                    return extract_xlsx(converted, limit), "libreoffice-fallback"
                finally:
                    temp.cleanup()
        if ext == ".xlsb":
            return extract_xlsb(path, limit), "direct"
        if ext in {".csv", ".tsv"}:
            return extract_delimited(path, limit), "direct"
        if ext == ".docx":
            return extract_docx(path, limit), "direct"
        if ext == ".pptx":
            return extract_pptx(path, limit), "direct"
        if ext == ".pdf":
            return extract_pdf(path, limit), "direct"
        if ext in {".doc", ".ppt", ".ods"}:
            temp, converted = convert_with_libreoffice(path)
            try:
                if converted.suffix.casefold() == ".docx":
                    return extract_docx(converted, limit), "libreoffice"
                if converted.suffix.casefold() == ".pptx":
                    return extract_pptx(converted, limit), "libreoffice"
                return extract_xlsx(converted, limit), "libreoffice"
            finally:
                temp.cleanup()
        if ext in TEXT_EXTS:
            return extract_plain_text(path, limit), "direct"
        return "", "unsupported"
    except Exception:
        raise


def iter_supported_files(
    root: Path,
    max_file_mb: int,
    allowed_exts: set[str] | None = None,
) -> Iterable[Path]:
    max_bytes = int(max_file_mb * 1024 * 1024)
    extensions = SUPPORTED_EXTS if allowed_exts is None else allowed_exts
    for base, dirs, files in os.walk(root):
        dirs[:] = [
            name for name in dirs
            if name.casefold() not in SKIP_DIRS
            and not name.startswith(".")
            and name.casefold() != ".generic_file_search"
        ]
        base_path = Path(base)
        for name in files:
            if name.startswith("~$") or name.startswith("."):
                continue
            path = base_path / name
            if path.suffix.casefold() not in extensions:
                continue
            try:
                if path.stat().st_size <= max_bytes:
                    yield path
            except OSError:
                continue


def query_terms(query: str) -> list[str]:
    tokens = re.findall(r"[A-Za-z0-9][A-Za-z0-9_.-]*", query.casefold())
    terms = []
    seen = set()
    for token in tokens:
        if len(token) < 2 or token in STOPWORDS:
            continue
        if token not in seen:
            seen.add(token)
            terms.append(token)
    return terms


def find_best_snippet(content: str, terms: list[str], width: int = 500) -> str:
    if not content:
        return ""
    low = content.casefold()
    positions = [low.find(term) for term in terms if low.find(term) >= 0]
    pos = min(positions) if positions else 0
    start = max(0, pos - width // 3)
    end = min(len(content), start + width)
    snippet = re.sub(r"\s+", " ", content[start:end]).strip()
    if start > 0:
        snippet = "…" + snippet
    if end < len(content):
        snippet += "…"
    return snippet


def deterministic_score(query: str, path: str, content: str) -> tuple[float, str]:
    query_low = query.casefold().strip()
    path_low = path.casefold()
    content_low = content.casefold()
    terms = query_terms(query)

    raw = 0.0

    if query_low:
        if query_low in path_low:
            raw += 90
        if query_low in content_low:
            raw += 75

    matched = 0
    for term in terms:
        path_count = path_low.count(term)
        content_count = content_low.count(term)
        if path_count or content_count:
            matched += 1
        raw += min(path_count, 3) * 22
        raw += min(content_count, 8) * 4
        if content_count:
            raw += 7

    if terms:
        coverage = matched / len(terms)
        raw += coverage * 55
        if matched == len(terms):
            raw += 25

    # Stable 0-100 local score.
    score = round(min(100.0, raw / 2.5), 2)
    return score, find_best_snippet(content, terms)


def ollama_status(url: str) -> tuple[bool, list[str], str]:
    if httpx is None:
        return False, [], "httpx is not installed"
    try:
        with httpx.Client(base_url=url.rstrip("/"), timeout=httpx.Timeout(8, connect=3)) as client:
            response = client.get("/api/tags")
            response.raise_for_status()
            models = [
                m.get("name") or m.get("model")
                for m in response.json().get("models", [])
                if m.get("name") or m.get("model")
            ]
        return True, models, ""
    except Exception as exc:
        return False, [], str(exc)


def ollama_rank(url: str, model: str, query: str, result: SearchResult, content: str) -> tuple[int, str]:
    if httpx is None:
        raise RuntimeError("httpx is not installed")

    prompt = f"""You are ranking a local file search result.

USER IS LOOKING FOR:
{query}

FILE:
{result.path}

EXTRACTED FILE CONTENT:
{content[:22000]}

Judge whether this file actually satisfies the search request.
Generic keyword coincidence is not enough; use the meaning and context.

Return JSON only:
{{
  "relevance": 0-100,
  "reason": "one concise explanation using concrete evidence from the file"
}}
"""
    payload = {
        "model": model,
        "stream": False,
        "format": "json",
        "options": {"temperature": 0.1},
        "messages": [
            {
                "role": "system",
                "content": "You are a careful local document-search relevance judge. Return JSON only.",
            },
            {"role": "user", "content": prompt},
        ],
    }

    with httpx.Client(base_url=url.rstrip("/"), timeout=httpx.Timeout(180, connect=5)) as client:
        response = client.post("/api/chat", json=payload)
        response.raise_for_status()
        message = (response.json().get("message") or {}).get("content", "").strip()

    try:
        data = json.loads(message)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", message, flags=re.S)
        if not match:
            raise RuntimeError("Ollama returned invalid JSON")
        data = json.loads(match.group(0))

    try:
        score = max(0, min(100, int(float(data.get("relevance", 0)))))
    except (TypeError, ValueError):
        score = 0
    return score, re.sub(r"\s+", " ", str(data.get("reason", ""))).strip()


def open_path(path: str) -> None:
    if sys.platform == "darwin":
        subprocess.Popen(["open", path])
    elif os.name == "nt":
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", path])


def reveal_path(path: str) -> None:
    if sys.platform == "darwin":
        subprocess.Popen(["open", "-R", path])
    elif os.name == "nt":
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    else:
        subprocess.Popen(["xdg-open", str(Path(path).parent)])


class GenericSearchApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"Generic Local File Search v{VERSION}")
        self.geometry("1260x820")
        self.minsize(1000, 650)

        self.events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.stop_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.result_map: dict[str, SearchResult] = {}

        roots = discover_onedrive_roots()
        default_root = str(roots[0]) if roots else str(Path.home())

        self.root_var = tk.StringVar(value=default_root)
        self.ollama_url_var = tk.StringVar(value=DEFAULT_OLLAMA_URL)
        self.model_var = tk.StringVar(value=DEFAULT_OLLAMA_MODEL)
        self.use_ollama_var = tk.BooleanVar(value=True)
        self.max_results_var = tk.IntVar(value=100)
        self.ollama_top_var = tk.IntVar(value=20)
        self.max_file_mb_var = tk.IntVar(value=MAX_FILE_MB_DEFAULT)
        self.file_type_var = tk.StringVar(value="All supported files")
        self.date_from_var = tk.StringVar()
        self.date_to_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.index_status_var = tk.StringVar(value=f"Index: {DB_PATH}")

        self._build_ui()
        self.after(120, self._poll_events)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        source = ttk.LabelFrame(outer, text="Search location", padding=8)
        source.pack(fill="x")

        ttk.Entry(source, textvariable=self.root_var).pack(side="left", fill="x", expand=True)
        ttk.Button(source, text="Browse…", command=self.choose_root).pack(side="left", padx=(8, 0))
        ttk.Button(source, text="Build / Update Index", command=self.start_index).pack(side="left", padx=(8, 0))

        query_frame = ttk.LabelFrame(outer, text="What are you looking for?", padding=8)
        query_frame.pack(fill="x", pady=(8, 0))
        self.query_text = tk.Text(query_frame, height=4, wrap="word")
        self.query_text.pack(fill="x")
        self.query_text.insert(
            "1.0",
            "approved suppliers by MDF with supplier ratings and approved/qualified status, possibly SPE Dashboard or Europe PGGI Hub Dashboard",
        )

        options = ttk.Frame(outer)
        options.pack(fill="x", pady=(8, 0))

        ollama_box = ttk.LabelFrame(options, text="Local Ollama", padding=8)
        ollama_box.pack(side="left", fill="x", expand=True)

        ttk.Checkbutton(ollama_box, text="Use Ollama to re-rank best matches", variable=self.use_ollama_var).grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(ollama_box, text="URL").grid(row=1, column=0, sticky="w")
        ttk.Entry(ollama_box, textvariable=self.ollama_url_var, width=34).grid(row=1, column=1, sticky="ew", padx=(5, 12))
        ttk.Label(ollama_box, text="Model").grid(row=1, column=2, sticky="w")
        ttk.Entry(ollama_box, textvariable=self.model_var, width=22).grid(row=1, column=3, sticky="ew", padx=(5, 0))
        ollama_box.columnconfigure(1, weight=1)
        ollama_box.columnconfigure(3, weight=1)

        filters = ttk.LabelFrame(options, text="Filters", padding=8)
        filters.pack(side="left", fill="x", padx=(8, 0))

        ttk.Label(filters, text="From year").grid(row=0, column=0)
        ttk.Entry(filters, textvariable=self.date_from_var, width=7).grid(row=0, column=1, padx=(4, 10))
        ttk.Label(filters, text="To year").grid(row=0, column=2)
        ttk.Entry(filters, textvariable=self.date_to_var, width=7).grid(row=0, column=3, padx=(4, 10))
        ttk.Label(filters, text="Results").grid(row=1, column=0)
        ttk.Spinbox(filters, from_=10, to=1000, textvariable=self.max_results_var, width=7).grid(row=1, column=1, padx=(4, 10))
        ttk.Label(filters, text="Ollama top").grid(row=1, column=2)
        ttk.Spinbox(filters, from_=0, to=100, textvariable=self.ollama_top_var, width=7).grid(row=1, column=3, padx=(4, 10))
        ttk.Label(filters, text="Max MB/file").grid(row=2, column=0)
        ttk.Spinbox(filters, from_=1, to=2000, textvariable=self.max_file_mb_var, width=7).grid(row=2, column=1, padx=(4, 10))

        ttk.Label(filters, text="File type (index + search)").grid(row=2, column=2, sticky="e")
        self.file_type_combo = ttk.Combobox(
            filters,
            textvariable=self.file_type_var,
            values=list(FILE_TYPE_PAESETS.keys()),
            state="readonly",
            width=22,
        )
        self.file_type_combo.grid(row=2, column=3, padx=(4, 10), sticky="w")

        action = ttk.Frame(outer)
        action.pack(fill="x", pady=(8, 0))
        self.search_button = ttk.Button(action, text="Search Indexed Files", command=self.start_search)
        self.search_button.pack(side="left")
        self.stop_button = ttk.Button(action, text="Stop", command=self.stop_work, state="disabled")
        self.stop_button.pack(side="left", padx=(8, 0))
        ttk.Button(action, text="Check Ollama", command=self.check_ollama).pack(side="left", padx=(8, 0))
        ttk.Button(action, text="Clear Results", command=self.clear_results).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(outer, mode="determinate")
        self.progress.pack(fill="x", pady=(8, 0))
        ttk.Label(outer, textvariable=self.status_var).pack(anchor="w", pady=(4, 0))
        ttk.Label(outer, textvariable=self.index_status_var).pack(anchor="w")
        ttk.Label(
            outer,
            text="The File type selector applies to both Build / Update Index and Search Indexed Files.",
        ).pack(anchor="w")

        result_box = ttk.LabelFrame(outer, text="Ranked results — double-click to open", padding=5)
        result_box.pack(fill="both", expand=True, pady=(8, 0))

        columns = ("score", "ollama", "modified", "type", "size", "file", "reason")
        self.tree = ttk.Treeview(result_box, columns=columns, show="headings", selectmode="browse")
        headings = {
            "score": "Score",
            "ollama": "Ollama",
            "modified": "Modified",
            "type": "Type",
            "size": "Size",
            "file": "File",
            "reason": "Evidence / reason",
        }
        widths = {
            "score": 70, "ollama": 70, "modified": 100, "type": 65,
            "size": 85, "file": 390, "reason": 460,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w")

        yscroll = ttk.Scrollbar(result_box, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(result_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        result_box.rowconfigure(0, weight=1)
        result_box.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.open_selected)
        self.tree.bind("<Button-3>", self.show_context_menu)

        self.context_menu = tk.Menu(self, tearoff=False)
        self.context_menu.add_command(label="Open file", command=self.open_selected)
        self.context_menu.add_command(label="Show in folder", command=self.reveal_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Copy full path", command=self.copy_selected_path)

    def choose_root(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.root_var.get() or str(Path.home()))
        if selected:
            self.root_var.set(selected)

    def set_busy(self, busy: bool) -> None:
        self.stop_button.configure(state="normal" if busy else "disabled")
        self.search_button.configure(state="disabled" if busy else "normal")
        self.stop_event.clear()

    def stop_work(self) -> None:
        self.stop_event.set()
        self.status_var.set("Stopping after the current file…")

    def check_stop(self) -> None:
        if self.stop_event.is_set():
            raise StopRequested()

    def start_index(self) -> None:
        root = Path(self.root_var.get()).expanduser()
        if not root.is_dir():
            messagebox.showerror("Invalid folder", "Choose an existing folder first.")
            return
        if self.worker and self.worker.is_alive():
            return

        self.set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)
        self.status_var.set("Finding supported files…")
        max_file_mb = int(self.max_file_mb_var.get())
        selected_type = self.file_type_var.get().strip() or "All supported files"
        self.worker = threading.Thread(
            target=self._index_worker,
            args=(root, max_file_mb, selected_type),
            daemon=True,
        )
        self.worker.start()

    def _index_worker(self, root: Path, max_file_mb: int, selected_type: str) -> None:
        conn = db_connect()
        try:
            allowed_exts = FILE_TYPE_PAESETS.get(selected_type)
            files = list(iter_supported_files(root, max_file_mb, allowed_exts))
            total = len(files)
            self.events.put(("progress_mode", ("determinate", max(1, total))))
            self.events.put((
                "status",
                f"Found {total:,} files for '{selected_type}'. Updating index…",
            ))

            if allowed_exts:
                placeholders = ",".join("?" for _ in allowed_exts)
                existing_rows = conn.execute(
                    f"SELECT path, size, mtime_ns FROM files WHERE root = ? "
                    f"AND extension IN ({placeholders})",
                    [str(root), *sorted(allowed_exts)],
                ).fetchall()
            else:
                existing_rows = conn.execute(
                    "SELECT path, size, mtime_ns FROM files WHERE root = ?",
                    (str(root),),
                ).fetchall()
            existing = {row[0]: (row[1], row[2]) for row in existing_rows}
            seen: set[str] = set()

            updated = 0
            unchanged = 0
            failed = 0

            for index, path in enumerate(files, start=1):
                self.check_stop()
                path_text = str(path)
                seen.add(path_text)

                try:
                    stat = path.stat()
                except OSError as exc:
                    failed += 1
                    self.events.put(("progress", (index, f"Cannot stat: {path.name}")))
                    continue

                fingerprint = (stat.st_size, stat.st_mtime_ns)
                if existing.get(path_text) == fingerprint:
                    unchanged += 1
                    self.events.put(("progress", (index, f"Unchanged {index:,}/{total:,}: {path.name}")))
                    continue

                try:
                    content, method = extract_content(path, MAX_INDEX_CHARS_DEFAULT)
                    status = method
                    error = ""
                    updated += 1
                except Exception as exc:
                    content = ""
                    status = "unreadable"
                    error = str(exc)[:1500]
                    failed += 1

                conn.execute(
                    """
                    INSERT INTO files(path, root, extension, size, mtime_ns, modified, indexed_at, status, error, content)
                    VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(path) DO UPDATE SET
                        root=excluded.root,
                        extension=excluded.extension,
                        size=excluded.size,
                        mtime_ns=excluded.mtime_ns,
                        modified=excluded.modified,
                        indexed_at=excluded.indexed_at,
                        status=excluded.status,
                        error=excluded.error,
                        content=excluded.content
                    """,
                    (
                        path_text, str(root), path.suffix.casefold(), stat.st_size,
                        stat.st_mtime_ns, stat.st_mtime, now_iso(), status, error, content,
                    ),
                )
                if index % 25 == 0:
                    conn.commit()
                self.events.put(("progress", (index, f"Indexed {index:,}/{total:,}: {path.name}")))

            # Remove stale rows for files that no longer exist under this indexed root.
            stale = set(existing) - seen
            if stale:
                conn.executemany("DELETE FROM files WHERE path = ?", [(p,) for p in stale])

            conn.commit()
            self.events.put((
                "done",
                f"Index updated for {selected_type}: {updated:,} changed/new, "
                f"{unchanged:,} unchanged, {failed:,} unreadable, "
                f"{len(stale):,} removed from this file-type scope.",
            ))
        except StopRequested:
            conn.commit()
            self.events.put(("done", "Indexing stopped. Completed work has been saved."))
        except Exception as exc:
            self.events.put(("error", f"Indexing failed: {exc}"))
        finally:
            conn.close()

    def start_search(self) -> None:
        query = self.query_text.get("1.0", "end").strip()
        if not query:
            messagebox.showerror("Search query", "Enter what you are looking for.")
            return

        root = Path(self.root_var.get()).expanduser()
        if not root.is_dir():
            messagebox.showerror("Invalid folder", "Choose an existing folder first.")
            return
        if self.worker and self.worker.is_alive():
            return

        try:
            start = None
            end = None
            if self.date_from_var.get().strip():
                year = int(self.date_from_var.get().strip())
                start = datetime(year, 1, 1).timestamp()
            if self.date_to_var.get().strip():
                year = int(self.date_to_var.get().strip())
                end = datetime(year + 1, 1, 1).timestamp() - 1
        except ValueError:
            messagebox.showerror("Date filter", "From year and To year must be four-digit years.")
            return

        options = {
            "start": start,
            "end": end,
            "max_results": max(1, int(self.max_results_var.get())),
            "ollama_top": max(0, int(self.ollama_top_var.get())),
            "use_ollama": bool(self.use_ollama_var.get()),
            "ollama_url": self.ollama_url_var.get().strip(),
            "model": self.model_var.get().strip(),
            "file_type": self.file_type_var.get().strip() or "All supported files",
        }

        self.clear_results()
        self.set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)
        self.status_var.set("Searching local index…")
        self.worker = threading.Thread(
            target=self._search_worker,
            args=(root, query, options),
            daemon=True,
        )
        self.worker.start()

    def _search_worker(self, root: Path, query: str, options: dict[str, Any]) -> None:
        conn = db_connect()
        try:
            start = options["start"]
            end = options["end"]
            sql = "SELECT path, extension, size, modified, status, content FROM files WHERE root = ?"
            params: list[Any] = [str(root)]
            if start is not None:
                sql += " AND modified >= ?"
                params.append(start)
            if end is not None:
                sql += " AND modified <= ?"
                params.append(end)

            selected_type = options.get("file_type", "All supported files")
            selected_exts = FILE_TYPE_PAESETS.get(selected_type)
            if selected_exts:
                placeholders = ",".join("?" for _ in selected_exts)
                sql += f" AND extension IN ({placeholders})"
                params.extend(sorted(selected_exts))

            terms = query_terms(query)
            if terms:
                # Restrict returned rows to files containing at least one useful
                # query term. SQLite still performs the content scan locally,
                # but we avoid loading the whole document index into Python.
                searchable_terms = terms[:10]
                clauses = []
                for term in searchable_terms:
                    clauses.append("(path LIKE ? OR content LIKE ?)")
                    wildcard = f"%{term}%"
                    params.extend([wildcard, wildcard])
                sql += " AND (" + " OR ".join(clauses) + ")"

            rows = conn.execute(sql, params).fetchall()
            if not rows:
                indexed_count = conn.execute(
                    "SELECT COUNT(*) FROM files WHERE root = ?",
                    (str(root),),
                ).fetchone()[0]
                if indexed_count == 0:
                    self.events.put(("done", "No indexed files found for this folder. Build / Update Index first."))
                else:
                    self.events.put(("done", "No indexed files matched this search."))
                return

            results: list[SearchResult] = []
            contents: dict[str, str] = {}
            for row_no, row in enumerate(rows, start=1):
                self.check_stop()
                path, ext, size, modified, status, content = row
                score, snippet = deterministic_score(query, path, content)
                if score <= 0:
                    continue
                result = SearchResult(
                    path=path,
                    extension=ext,
                    modified=modified,
                    size=size,
                    local_score=score,
                    final_score=score,
                    snippet=snippet,
                    status=status,
                )
                results.append(result)
                contents[path] = content

            results.sort(key=lambda item: item.local_score, reverse=True)
            max_results = options["max_results"]
            ollama_top = options["ollama_top"]
            working = results[:max(max_results, ollama_top)]

            if options["use_ollama"] and working and ollama_top > 0:
                ok, models, error = ollama_status(options["ollama_url"])
                if ok:
                    top_n = min(ollama_top, len(working))
                    self.events.put(("progress_mode", ("determinate", top_n)))
                    for idx, result in enumerate(working[:top_n], start=1):
                        self.check_stop()
                        self.events.put(("progress", (idx, f"Ollama {idx}/{top_n}: {Path(result.path).name}")))
                        try:
                            llm_score, reason = ollama_rank(
                                options["ollama_url"],
                                options["model"],
                                query,
                                result,
                                contents[result.path],
                            )
                            result.ollama_score = llm_score
                            result.ollama_reason = reason
                            result.final_score = round(result.local_score * 0.45 + llm_score * 0.55, 2)
                        except Exception as exc:
                            result.ollama_reason = f"Ollama check failed: {exc}"
                else:
                    self.events.put(("status", f"Ollama unavailable; showing local results only: {error}"))

            working.sort(key=lambda item: (item.final_score, item.local_score), reverse=True)
            final = working[:max_results]
            self.events.put(("results", final))
            self.events.put((
                "done",
                f"Search complete: {len(final):,} ranked results from {len(rows):,} "
                f"matching indexed files. Type filter: {options.get('file_type', 'All supported files')}."
            ))
        except StopRequested:
            self.events.put(("done", "Search stopped."))
        except ValueError as exc:
            self.events.put(("error", f"Filter error: {exc}"))
        except Exception as exc:
            self.events.put(("error", f"Search failed: {exc}"))
        finally:
            conn.close()

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "status":
                    self.status_var.set(str(payload))
                elif kind == "progress_mode":
                    mode, maximum = payload
                    self.progress.stop()
                    self.progress.configure(mode=mode, maximum=max(1, int(maximum)), value=0)
                elif kind == "progress":
                    value, text = payload
                    self.progress.configure(value=value)
                    self.status_var.set(text)
                elif kind == "results":
                    self._show_results(payload)
                elif kind == "done":
                    self.progress.stop()
                    self.progress.configure(value=0)
                    self.status_var.set(str(payload))
                    self.set_busy(False)
                elif kind == "error":
                    self.progress.stop()
                    self.progress.configure(value=0)
                    self.status_var.set(str(payload))
                    self.set_busy(False)
                    messagebox.showerror("Generic File Search", str(payload))
        except queue.Empty:
            pass
        self.after(120, self._poll_events)

    def _show_results(self, results: list[SearchResult]) -> None:
        self.clear_results()
        for index, result in enumerate(results, start=1):
            iid = str(index)
            self.result_map[iid] = result
            reason = result.ollama_reason or result.snippet
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    f"{result.final_score:.1f}",
                    "" if result.ollama_score is None else result.ollama_score,
                    datetime.fromtimestamp(result.modified).strftime("%Y-%m-%d"),
                    result.extension,
                    human_size(result.size),
                    result.path,
                    reason,
                ),
            )

    def clear_results(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.result_map.clear()

    def selected_result(self) -> SearchResult | None:
        selected = self.tree.selection()
        if not selected:
            return None
        return self.result_map.get(selected[0])

    def open_selected(self, _event: Any = None) -> None:
        result = self.selected_result()
        if result:
            try:
                open_path(result.path)
            except Exception as exc:
                messagebox.showerror("Open file", str(exc))

    def reveal_selected(self) -> None:
        result = self.selected_result()
        if result:
            try:
                reveal_path(result.path)
            except Exception as exc:
                messagebox.showerror("Show in folder", str(exc))

    def copy_selected_path(self) -> None:
        result = self.selected_result()
        if result:
            self.clipboard_clear()
            self.clipboard_append(result.path)

    def show_context_menu(self, event: tk.Event) -> None:
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.tk_popup(event.x_root, event.y_root)

    def check_ollama(self) -> None:
        self.status_var.set("Checking Ollama…")
        url = self.ollama_url_var.get().strip()

        def worker() -> None:
            ok, models, error = ollama_status(url)
            if ok:
                message = f"Connected to {url}\n\nModels:\n" + "\n".join(models)
                self.after(0, lambda: messagebox.showinfo("Ollama", message))
                self.after(0, lambda: self.status_var.set("Ollama connected"))
            else:
                self.after(0, lambda: messagebox.showerror("Ollama", error))
                self.after(0, lambda: self.status_var.set("Ollama unavailable"))

        threading.Thread(target=worker, daemon=True).start()



# v2.3.0 structured metadata-aware LLM search layer.
# Normal source checkout; no ZIP/runtime extraction required.
from v2_3_search_logic import install as _install_v2_3_search_logic
_install_v2_3_search_logic(sys.modules[__name__])

def main() -> int:
    app = GenericSearchApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
