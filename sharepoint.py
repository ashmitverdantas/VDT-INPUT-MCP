"""
auth/sharepoint.py
──────────────────
Microsoft Graph API client for SharePoint:
  - OAuth2 client-credentials token (auto-refresh, cached)
  - File download (sharing links, team sites, OneDrive personal)
  - Folder listing with optional extension filter
"""

from __future__ import annotations

import base64
import json
import logging
import os
import re
import time
from typing import Optional
from urllib.parse import quote, unquote, urlparse

import httpx

logger = logging.getLogger(__name__)

# ── In-process token cache ─────────────────────────────────────────────────────
_TOKEN_CACHE: dict = {"token": None, "expires_at": 0.0}

GRAPH = "https://graph.microsoft.com/v1.0"


# ── Helpers ────────────────────────────────────────────────────────────────────

def _require(key: str) -> str:
    val = os.getenv(key, "")
    if not val:
        raise EnvironmentError(
            f"Required environment variable '{key}' is not set. "
            "Add it to your .env file or Azure App Service settings."
        )
    return val


# ── Token acquisition ──────────────────────────────────────────────────────────

async def get_access_token() -> str:
    """
    Returns a valid Bearer token for Microsoft Graph.
    Refreshes automatically when within 60 s of expiry.
    """
    now = time.time()
    if _TOKEN_CACHE["token"] and now < _TOKEN_CACHE["expires_at"] - 60:
        return _TOKEN_CACHE["token"]

    tenant_id     = _require("SHAREPOINT_TENANT_ID")
    client_id     = _require("SHAREPOINT_CLIENT_ID")
    client_secret = _require("SHAREPOINT_CLIENT_SECRET")

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
    }

    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(token_url, data=payload)

    if resp.status_code != 200:
        raise ValueError(
            f"Token request failed ({resp.status_code}). "
            f"Tenant: {tenant_id}  Client: {client_id}  "
            f"Response: {resp.text[:400]}"
        )

    data = resp.json()
    _TOKEN_CACHE["token"]      = data["access_token"]
    _TOKEN_CACHE["expires_at"] = now + data.get("expires_in", 3600)
    logger.info("SharePoint access token acquired / refreshed")
    return _TOKEN_CACHE["token"]


# ── URL classification ─────────────────────────────────────────────────────────

def classify_url(url: str) -> dict:
    """
    Parse a SharePoint / OneDrive URL and return its type + components.

    Returned keys
    ─────────────
    hostname        str   e.g. "hullinc.sharepoint.com"
    type            str   "sharing_link" | "onedrive_personal" | "team_site" | "unknown"
    site_path       str   team-site name (team_site only)
    user_principal  str   UPN slug (onedrive_personal only)
    file_path       str   path relative to drive root
    """
    parsed   = urlparse(url)
    hostname = parsed.hostname or ""
    path     = unquote(parsed.path)

    # Parse query string into a dict (handles both encoded and raw)
    params: dict[str, str] = {}
    if parsed.query:
        for part in unquote(parsed.query).split("&"):
            if "=" in part:
                k, v = part.split("=", 1)
                params[k] = v

    result: dict = {"hostname": hostname, "type": "unknown", "raw_url": url}

    # ① Sharing link ── /:w:/ /:b:/ /:x:/ etc.
    if re.match(r"/:[a-z]:/", path):
        result["type"] = "sharing_link"
        return result

    # ② OneDrive for Business — file path encoded in ?id= query param
    if "id" in params:
        id_path = unquote(params["id"])
        parts = id_path.strip("/").split("/")
        if len(parts) >= 2 and parts[0] == "personal":
            result["type"]            = "onedrive_personal"
            result["user_principal"]  = parts[1]
            file_parts = parts[2:]
            if file_parts and file_parts[0] == "Documents":
                file_parts = file_parts[1:]
            result["file_path"] = "/".join(file_parts)
            return result

    # ③ OneDrive personal — direct path
    personal = re.match(r"/personal/([^/]+)/(.*)", path)
    if personal:
        result["type"]           = "onedrive_personal"
        result["user_principal"] = personal.group(1)
        fp = personal.group(2)
        if fp.startswith("Documents/"):
            fp = fp[len("Documents/"):]
        result["file_path"] = fp
        return result

    # ④ Team site — /sites/SiteName/…
    team = re.match(r"/sites/([^/]+)/(.*)", path)
    if team:
        result["type"]      = "team_site"
        result["site_path"] = team.group(1)
        result["file_path"] = team.group(2)
        return result

    # ⑤ Bare file path with extension
    if re.search(r"\.\w{2,5}$", path):
        result["type"]      = "direct_path"
        result["file_path"] = path
        return result

    return result


# ── File download ──────────────────────────────────────────────────────────────

_LIBRARY_ALIASES: dict[str, list[str]] = {
    "shared documents": ["documents"],
    "documents":        ["shared documents"],
}


async def download_file(url: str) -> bytes:
    """
    Download a file from SharePoint / OneDrive via Microsoft Graph.

    Supported URL patterns
    ──────────────────────
    • Sharing links  (/:w:/, /:b:/, /:x:/)
    • Team sites     (/sites/SiteName/Shared%20Documents/…)
    • OneDrive       (/personal/user@org/Documents/…  or  ?id=… query)
    """
    token = await get_access_token()
    info  = classify_url(url)
    host  = info["hostname"]

    if not host:
        raise ValueError(f"Could not parse SharePoint hostname from URL: {url[:200]}")

    async with httpx.AsyncClient(follow_redirects=True, timeout=120) as client:
        headers = {"Authorization": f"Bearer {token}"}

        # ─── Sharing link ──────────────────────────────────────────────────────
        if info["type"] == "sharing_link":
            encoded  = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
            share_id = f"u!{encoded}"
            resp = await client.get(
                f"{GRAPH}/shares/{share_id}/driveItem/content",
                headers=headers,
            )
            if resp.status_code != 200:
                raise ValueError(
                    f"Sharing-link download failed ({resp.status_code}): {resp.text[:400]}"
                )
            return resp.content

        # ─── OneDrive personal ─────────────────────────────────────────────────
        if info["type"] == "onedrive_personal":
            upn       = info["user_principal"]
            file_path = info["file_path"]

            site_resp = await client.get(
                f"{GRAPH}/sites/{host}:/personal/{upn}",
                headers=headers,
            )
            if site_resp.status_code != 200:
                raise ValueError(
                    f"Could not resolve OneDrive site for '{upn}': {site_resp.text[:300]}"
                )
            site_id = site_resp.json()["id"]

            resp = await client.get(
                f"{GRAPH}/sites/{site_id}/drive/root:/{quote(file_path, safe='/')}:/content",
                headers=headers,
            )
            if resp.status_code == 200:
                return resp.content
            raise ValueError(
                f"Download failed for '{file_path}' ({resp.status_code}): {resp.text[:300]}"
            )

        # ─── Team site ─────────────────────────────────────────────────────────
        if info["type"] == "team_site":
            site_name = info["site_path"]
            file_path = info["file_path"]

            # Resolve site ID
            site_resp = await client.get(
                f"{GRAPH}/sites/{host}:/sites/{site_name}",
                headers=headers,
            )
            if site_resp.status_code != 200:
                raise ValueError(
                    f"Could not resolve site '{site_name}': {site_resp.text[:300]}"
                )
            site_id = site_resp.json()["id"]

            # Resolve drive (document library) — handle "Shared Documents" ↔ "Documents" alias
            path_parts   = file_path.split("/", 1)
            library_name = path_parts[0]
            rel_path     = path_parts[1] if len(path_parts) > 1 else ""

            drives_resp = await client.get(
                f"{GRAPH}/sites/{site_id}/drives", headers=headers
            )
            drives = drives_resp.json().get("value", [])

            candidates = [library_name.lower()] + _LIBRARY_ALIASES.get(library_name.lower(), [])
            drive_id   = next(
                (d["id"] for d in drives if d.get("name", "").lower() in candidates),
                None,
            )
            if not drive_id:
                available = [d.get("name") for d in drives]
                raise ValueError(
                    f"Drive '{library_name}' not found in site '{site_name}'. "
                    f"Available: {available}"
                )

            resp = await client.get(
                f"{GRAPH}/drives/{drive_id}/root:/{quote(rel_path, safe='/')}:/content",
                headers=headers,
            )
            if resp.status_code != 200:
                raise ValueError(
                    f"Graph download failed ({resp.status_code}): {resp.text[:400]}"
                )
            return resp.content

        raise ValueError(
            f"Unsupported URL type '{info['type']}' for: {url[:200]}"
        )


# ── Folder listing ─────────────────────────────────────────────────────────────

async def list_files_in_folder(
    site_id:     str,
    drive_id:    str,
    folder_path: str = "root",
    file_types:  Optional[list[str]] = None,
) -> list[dict]:
    """
    List files in a SharePoint document library folder.

    Args:
        site_id:     Graph site ID
        drive_id:    Document library drive ID
        folder_path: Path relative to drive root (use 'root' for top level)
        file_types:  Extension filter e.g. ['.pdf', '.docx'] — None = all

    Returns:
        List of dicts with keys: name, id, size_bytes, download_url,
        web_url, last_modified, extension, created_by
    """
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    if folder_path in ("root", "", "/"):
        graph_url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root/children"
    else:
        cleaned = folder_path.strip("/")
        graph_url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root:/{cleaned}:/children"

    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(graph_url, headers=headers)

    if resp.status_code != 200:
        raise ValueError(
            f"listFiles failed ({resp.status_code}): {resp.text[:400]}"
        )

    items = resp.json().get("value", [])
    lower_types = [ft.lower() for ft in file_types] if file_types else None

    files = []
    for item in items:
        if "folder" in item:
            continue
        name: str = item.get("name", "")
        ext = ("." + name.rsplit(".", 1)[-1].lower()) if "." in name else ""
        if lower_types and ext not in lower_types:
            continue
        files.append({
            "name":          name,
            "id":            item.get("id"),
            "size_bytes":    item.get("size", 0),
            "download_url":  item.get("@microsoft.graph.downloadUrl") or item.get("webUrl"),
            "web_url":       item.get("webUrl"),
            "last_modified": item.get("lastModifiedDateTime"),
            "extension":     ext,
            "created_by":    item.get("createdBy", {}).get("user", {}).get("displayName"),
        })
    return files
