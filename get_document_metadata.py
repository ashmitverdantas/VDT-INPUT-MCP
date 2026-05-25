"""
get_document_metadata.py
─────────────────────────
Handles the getDocumentMetadata MCP tool.

Returns lightweight metadata for a SharePoint document:
  - file format / extension
  - size in bytes
  - SHA-256 hash
  - whether text extraction is supported
  - for docx: list of content-control field names
  - for docx: list of brace {placeholder} tokens found
"""

from __future__ import annotations

import hashlib
import io
import logging
import re
import zipfile
from typing import Any

try:
    from sharepoint import download_file as sp_download
    _HAS_SP = True
except ImportError:
    _HAS_SP = False

import httpx

logger = logging.getLogger(__name__)

# ── JSON schema for MCP tool registration ─────────────────────────────────────
GET_METADATA_SCHEMA = {
    "type": "object",
    "properties": {
        "documentUrl": {
            "type": "string",
            "description": "SharePoint URL of the document to inspect.",
        }
    },
    "required": ["documentUrl"],
}

_SUPPORTED_TEXT_FORMATS = {".pdf", ".docx", ".doc", ".txt", ".csv", ".xlsx", ".xls"}

_EXTENSION_LABELS = {
    ".pdf":  "PDF document",
    ".docx": "Microsoft Word (Open XML)",
    ".doc":  "Microsoft Word (legacy)",
    ".xlsx": "Microsoft Excel (Open XML)",
    ".xls":  "Microsoft Excel (legacy)",
    ".txt":  "Plain text",
    ".csv":  "Comma-separated values",
    ".pptx": "Microsoft PowerPoint (Open XML)",
    ".png":  "PNG image",
    ".jpg":  "JPEG image",
    ".jpeg": "JPEG image",
}


async def _fetch_bytes(url: str) -> bytes:
    if url.startswith("http") and "sharepoint" in url.lower() and _HAS_SP:
        return await sp_download(url)
    if url.startswith("http"):
        async with httpx.AsyncClient(timeout=60, follow_redirects=True) as client:
            resp = await client.get(url)
            if resp.status_code != 200:
                raise ValueError(f"HTTP {resp.status_code}: {resp.text[:200]}")
            return resp.content
    import os
    with open(url, "rb") as f:
        return f.read()


def _inspect_docx(data: bytes) -> dict:
    """Extract SDT aliases and brace placeholders from a .docx."""
    sdt_aliases: list[str] = []
    brace_tokens: list[str] = []

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for name in zf.namelist():
                if name.endswith(".xml"):
                    xml = zf.read(name).decode("utf-8", errors="replace")
                    sdt_aliases += re.findall(r'w:alias w:val="([^"]+)"', xml)
                    brace_tokens += re.findall(r"\{([^}]{1,80})\}", xml)
    except Exception as exc:
        logger.warning("docx inspection failed: %s", exc)

    return {
        "content_controls": sorted(set(sdt_aliases)),
        "brace_placeholders": sorted(set(brace_tokens)),
    }


async def handle_get_document_metadata(params: dict[str, Any]) -> dict:
    url: str = params.get("documentUrl", "").strip()
    if not url:
        return {"success": False, "error": "documentUrl is required"}

    try:
        data = await _fetch_bytes(url)
    except Exception as exc:
        return {"success": False, "error": f"Could not fetch document: {exc}"}

    # Derive extension from URL
    url_path = url.split("?")[0]
    ext = ""
    if "." in url_path.split("/")[-1]:
        ext = "." + url_path.split("/")[-1].rsplit(".", 1)[-1].lower()

    sha256 = hashlib.sha256(data).hexdigest()
    size = len(data)

    meta: dict[str, Any] = {
        "success": True,
        "url": url,
        "size_bytes": size,
        "sha256": sha256,
        "extension": ext,
        "format_label": _EXTENSION_LABELS.get(ext, "Unknown"),
        "text_extraction_supported": ext in _SUPPORTED_TEXT_FORMATS,
    }

    if ext == ".docx":
        meta.update(_inspect_docx(data))

    return meta
