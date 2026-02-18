"""Quotes Word template management service (separate from calendar template)."""

from __future__ import annotations

import os
import threading
import time
from pathlib import Path

from app.core.config import QUOTES_TEMPLATE_FALLBACK_PATH, QUOTES_TEMPLATE_PATH

_LOCK = threading.Lock()

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_MAX_TEMPLATE_BYTES = 20 * 1024 * 1024  # 20 MB


def _is_probably_docx(data: bytes) -> bool:
    return len(data) >= 4 and data[:2] == b"PK"


def ensure_template_exists() -> Path:
    target = Path(QUOTES_TEMPLATE_PATH)
    if target.exists():
        return target

    fallback = Path(QUOTES_TEMPLATE_FALLBACK_PATH)
    if not fallback.exists():
        raise FileNotFoundError(
            f"Quotes template not found. Expected current={target} or fallback={fallback}"
        )

    target.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = target.parent / f".{target.name}.{int(time.time() * 1000)}.tmp"
    tmp_path.write_bytes(fallback.read_bytes())
    os.replace(tmp_path, target)
    return target


def get_template_path() -> Path:
    return ensure_template_exists()


def get_template_info() -> dict:
    path = get_template_path()
    stat = path.stat()
    return {
        "path": str(path),
        "size_bytes": stat.st_size,
        "modified_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(stat.st_mtime)),
    }


def update_template_bytes(
    data: bytes,
    *,
    filename: str | None = None,
    content_type: str | None = None,
) -> dict:
    if not data:
        raise ValueError("Empty file.")
    if len(data) > _MAX_TEMPLATE_BYTES:
        raise ValueError(f"Template is too large (>{_MAX_TEMPLATE_BYTES} bytes).")

    if filename and not filename.lower().endswith(".docx"):
        raise ValueError("Only .docx files are supported.")
    if content_type and content_type not in ("application/octet-stream", DOCX_MIME):
        raise ValueError(f"Unsupported content type: {content_type}")
    if not _is_probably_docx(data):
        raise ValueError("File does not look like a .docx (zip) document.")

    target = Path(QUOTES_TEMPLATE_PATH)
    target.parent.mkdir(parents=True, exist_ok=True)

    with _LOCK:
        tmp_path = target.parent / f".{target.name}.{int(time.time() * 1000)}.tmp"
        tmp_path.write_bytes(data)
        os.replace(tmp_path, target)

    return get_template_info()

