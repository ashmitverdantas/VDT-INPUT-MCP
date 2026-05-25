"""
fill_letter_template.py
────────────────────────
Handles the fillLetterTemplate MCP tool.

Strategy
────────
The Verdantas template uses two placeholder mechanisms:
  1. Word Content Controls  (w:sdt with w:alias/w:tag)  — preferred, named fields
  2. Brace placeholders     e.g. {Client Name}          — fallback / older text

We replace both in one pass against the raw XML, then repack into a .docx
and return base64-encoded bytes.

Field → content-control alias mapping
──────────────────────────────────────
  letter_date         → (date SDT — matched by placeholder text)
  client_name         → ClientName
  company_name        → ClientCompany
  address             → ClientAddress
  email               → (brace placeholder {[name@email.com]} / {000.000.0000})
  subject_line        → SubjectLine
  project_number      → (brace placeholder {Project Number})
  letter_recipient    → Salutation
  cover_letter_body   → CoverLetterBody
  name_surname        → SignatoryNames
  job_title           → SignatoryTitles
  contact_information → SignatoryContacts
  attachment_1_title  → Attachment1Title
  attachment_1_content→ Attachment1Content
  attachment_2_title  → Attachment2Title
  attachment_2_content→ Attachment2Content
  attachment_3_content→ Section3Content
  attachment_4_content→ Section4Content
  attachment_5_content→ Section5Content
  attachment_6_content→ Section6Content
  attachment_7_content→ Section7Content
  attachment_8_content→ Section8Content
"""

from __future__ import annotations

import base64
import io
import logging
import os
import re
import shutil
import tempfile
import zipfile
from typing import Any

import httpx

# Optional: SharePoint download
try:
    from sharepoint import download_file as sp_download
    _HAS_SP = True
except ImportError:
    _HAS_SP = False

logger = logging.getLogger(__name__)

# ── Default template path (bundled alongside this module) ──────────────────────
DEFAULT_TEMPLATE = os.path.join(os.path.dirname(__file__), "Verdantas_Template_WithControls.docx")

# ── JSON schema for MCP tool registration ─────────────────────────────────────
FILL_LETTER_TEMPLATE_SCHEMA = {
    "type": "object",
    "properties": {
        "documentUrl": {
            "type": "string",
            "description": "SharePoint URL of the .docx template. Uses bundled template if omitted.",
        },
        "fields": {
            "type": "object",
            "description": "Key/value pairs of field names to replacement text.",
            "properties": {
                "letter_date":          {"type": "string"},
                "client_name":          {"type": "string"},
                "company_name":         {"type": "string"},
                "address":              {"type": "string"},
                "email":                {"type": "string"},
                "subject_line":         {"type": "string"},
                "project_number":       {"type": "string"},
                "letter_recipient":     {"type": "string"},
                "cover_letter_body":    {"type": "string"},
                "name_surname":         {"type": "string"},
                "job_title":            {"type": "string"},
                "contact_information":  {"type": "string"},
                "attachment_1_title":   {"type": "string"},
                "attachment_1_content": {"type": "string"},
                "attachment_2_title":   {"type": "string"},
                "attachment_2_content": {"type": "string"},
                "attachment_3_content": {"type": "string"},
                "attachment_4_content": {"type": "string"},
                "attachment_5_content": {"type": "string"},
                "attachment_6_content": {"type": "string"},
                "attachment_7_content": {"type": "string"},
                "attachment_8_content": {"type": "string"},
            },
        },
    },
    "required": [],
}

# ── Field → content-control alias mapping ──────────────────────────────────────
_ALIAS_MAP: dict[str, str] = {
    "client_name":          "ClientName",
    "company_name":         "ClientCompany",
    "address":              "ClientAddress",
    "subject_line":         "SubjectLine",
    "letter_recipient":     "Salutation",
    "cover_letter_body":    "CoverLetterBody",
    "name_surname":         "SignatoryNames",
    "job_title":            "SignatoryTitles",
    "contact_information":  "SignatoryContacts",
    "attachment_1_title":   "Attachment1Title",
    "attachment_1_content": "Attachment1Content",
    "attachment_2_title":   "Attachment2Title",
    "attachment_2_content": "Attachment2Content",
    "attachment_3_content": "Section3Content",
    "attachment_4_content": "Section4Content",
    "attachment_5_content": "Section5Content",
    "attachment_6_content": "Section6Content",
    "attachment_7_content": "Section7Content",
    "attachment_8_content": "Section8Content",
}

# ── Brace placeholder patterns (regex → field key) ────────────────────────────
# These match the literal placeholder text visible in the document
_BRACE_MAP: list[tuple[str, str]] = [
    (r"\{Project Number\}",                 "project_number"),
    (r"\{000\.000\.0000\}",                 "email"),           # phone placeholder
    (r"\{name@email\.com\}",               "email"),
    (r"\[name@email\.com\]",               "email"),
    (r"name@email\.com",                    "email"),
    (r"Click or tap to enter a date\.",     "letter_date"),
    (r"\{Subject Line\}",                   "subject_line"),
    (r"\{Letter Recipient\}",              "letter_recipient"),
    # Attachment title splits across runs — handled in XML rewrite
    (r"\{Attachment 1 Title\}",             "attachment_1_title"),
    (r"Attachment 1 Title\}",               "attachment_1_title"),
    (r"\{Attachment 1:",                    ""),                 # strip label
    (r"\{Attachment 2 Title",               "attachment_2_title"),
    (r"Attachment 2 Title\}",               "attachment_2_title"),
    (r"\{Attachment 2:",                    ""),                 # strip label
    # Section attachment body blocks
    (r"\{Attachment 3: Management Plan Content\}", "attachment_3_content"),
    (r"\{Attachment 4: Team Qualifications Content\}", "attachment_4_content"),
    (r"\{Attachment 5: Schedule Content\}", "attachment_5_content"),
    (r"\{Attachment 6: Compliance Matrix Content\}", "attachment_6_content"),
    (r"\{Attachment 7: Budget Approach Content\}", "attachment_7_content"),
    (r"\{Attachment 8: Appendices Content\}", "attachment_8_content"),
    # Misc
    (r"\{Management Plan\}",               "attachment_3_content"),
    (r"\{Team Qualifications\}",           "attachment_4_content"),
    (r"\{Attachments\}",                   ""),                 # strip label node
    (r"\{Client Name\}\. \{Designation\}", "client_name"),
    (r"\{Client Name\}",                   "client_name"),
    (r"\{Designation\}",                   ""),
    (r"\{Company Name\}",                  "company_name"),
    (r"\{Street Name",                     "address"),
    (r"City, ST Zip Code\}",               ""),
    (r"\{Letter Recipient\}",              "letter_recipient"),
    (r"\{Subject Line\}, \{Project Number\}", "subject_line"),
    (r"\{Name Surname\}",                  "name_surname"),
    (r"\{Job Title\}",                     "job_title"),
    (r"\{Contact Information\}",           "contact_information"),
    (r"\{email\}",                         "email"),
]


# ── XML helpers ────────────────────────────────────────────────────────────────

def _xml_escape(text: str) -> str:
    """Escape text for safe insertion into XML."""
    return (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _make_run(text: str, rpr_xml: str = "") -> str:
    """Wrap text in a minimal <w:r> run, preserving optional run-properties."""
    lines = text.split("\n")
    runs = []
    for i, line in enumerate(lines):
        if i > 0:
            runs.append(f"<w:r>{rpr_xml}<w:br/></w:r>")
        space = ' xml:space="preserve"' if line != line.strip() else ""
        runs.append(f"<w:r>{rpr_xml}<w:t{space}>{_xml_escape(line)}</w:t></w:r>")
    return "".join(runs)


def _replace_sdt_content(xml: str, alias: str, new_text: str) -> str:
    """
    Replace the sdtContent of a content control identified by its alias/tag.
    Preserves the sdtPr block and wraps replacement text in a minimal paragraph.
    Handles multi-line values by creating one <w:p> per line.
    """
    pattern = re.compile(
        r"(<w:sdt>\s*<w:sdtPr>(?:(?!<w:sdt>).)*?"
        + re.escape(f'w:val="{alias}"')
        + r"(?:(?!<w:sdt>).)*?</w:sdtPr>\s*(?:<w:sdtEndPr/>|<w:sdtEndPr>.*?</w:sdtEndPr>)?\s*)"
        r"<w:sdtContent>(.*?)</w:sdtContent>"
        r"(\s*</w:sdt>)",
        re.DOTALL,
    )

    def replacer(m: re.Match) -> str:
        before = m.group(1)
        original_content = m.group(2)
        after = m.group(3)

        # Extract run-properties from first run in original content
        rpr_match = re.search(r"<w:rPr>(.*?)</w:rPr>", original_content, re.DOTALL)
        rpr_xml = f"<w:rPr>{rpr_match.group(1)}</w:rPr>" if rpr_match else ""

        # Extract paragraph properties from original content
        ppr_match = re.search(r"<w:pPr>(.*?)</w:pPr>", original_content, re.DOTALL)
        ppr_xml = f"<w:pPr>{ppr_match.group(1)}</w:pPr>" if ppr_match else ""

        # Build one paragraph per newline for multi-line fields (e.g. address)
        lines = new_text.split("\n")
        paras = []
        for line in lines:
            space = ' xml:space="preserve"' if line != line.strip() else ""
            run = f"<w:r>{rpr_xml}<w:t{space}>{_xml_escape(line)}</w:t></w:r>"
            paras.append(f"<w:p>{ppr_xml}{run}</w:p>")

        new_content = f"<w:sdtContent>{''.join(paras)}</w:sdtContent>"
        return before + new_content + after

    new_xml, count = pattern.subn(replacer, xml)
    if count:
        logger.debug("SDT '%s' replaced (%d occurrence(s))", alias, count)
    else:
        logger.debug("SDT '%s' not found in XML", alias)
    return new_xml


def _replace_brace_placeholders(xml: str, fields: dict[str, str]) -> str:
    """Apply brace-placeholder substitutions."""
    for pattern, field_key in _BRACE_MAP:
        if not field_key:
            replacement = ""
        else:
            replacement = fields.get(field_key, "")
            if not replacement:
                continue  # field not supplied — leave as-is

        xml = re.sub(pattern, _xml_escape(replacement), xml)

    # Fix split email hyperlink pattern: <w:t>{</w:t>...<w:t>name@email.com}</w:t>
    if fields.get("email"):
        email_val = _xml_escape(fields["email"])
        # Replace the opening brace before hyperlink
        xml = re.sub(r"<w:t>\{</w:t>(\s*</w:r>\s*<w:hyperlink[^>]*>)", r"\1", xml)
        # Replace "name@email.com}" with just the email value
        xml = re.sub(r"<w:t>name@email\.com\}</w:t>", f"<w:t>{email_val}</w:t>", xml)

    # Fix {Sincerely, Verdantas} brace that wraps across lines
    xml = re.sub(r"\{Sincerely,\s*\n?\s*Verdantas\}", "Sincerely,\nVerdantas", xml)
    xml = re.sub(r"\{Sincerely,", "Sincerely,", xml)

    # Fix stray closing braces from split placeholders
    xml = re.sub(r"<w:t>\}</w:t>", "<w:t/>", xml)
    xml = re.sub(r"^\}", "", xml)

    return xml


# ── Template acquisition ───────────────────────────────────────────────────────

async def _get_template_bytes(document_url: str) -> bytes:
    """Fetch template bytes from URL, SharePoint, or local disk."""
    if document_url:
        if document_url.startswith("http") and "sharepoint" in document_url.lower():
            if not _HAS_SP:
                raise RuntimeError("sharepoint.py not available; cannot download from SharePoint")
            return await sp_download(document_url)
        elif document_url.startswith("http"):
            async with httpx.AsyncClient(timeout=60, follow_redirects=True) as client:
                resp = await client.get(document_url)
                if resp.status_code != 200:
                    raise ValueError(f"HTTP {resp.status_code} fetching template: {resp.text[:300]}")
                return resp.content
        elif os.path.isfile(document_url):
            with open(document_url, "rb") as f:
                return f.read()
        else:
            raise ValueError(f"Cannot resolve documentUrl: {document_url!r}")

    if os.path.isfile(DEFAULT_TEMPLATE):
        with open(DEFAULT_TEMPLATE, "rb") as f:
            return f.read()

    raise FileNotFoundError(
        "No documentUrl provided and no bundled template found at: "
        + DEFAULT_TEMPLATE
    )


# ── Main handler ───────────────────────────────────────────────────────────────

async def handle_fill_letter_template(params: dict[str, Any]) -> dict:
    """
    Download the template, fill all placeholders, return base64 docx.

    Params
    ──────
    documentUrl : str          SharePoint/HTTP/local path to the .docx template
    fields      : dict[str,str] Field key → replacement value
    """
    document_url: str = params.get("documentUrl", "")
    fields: dict[str, str] = params.get("fields", {})

    if not fields:
        return {
            "success": False,
            "error": "No fields provided. Pass at least one field to fill.",
        }

    # 1. Fetch template
    try:
        template_bytes = await _get_template_bytes(document_url)
    except Exception as exc:
        return {"success": False, "error": f"Failed to fetch template: {exc}"}

    # 2. Process the docx zip in memory
    try:
        in_buf = io.BytesIO(template_bytes)
        out_buf = io.BytesIO()

        with zipfile.ZipFile(in_buf, "r") as zin, zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                # Only rewrite XML part files (document, headers, footers, etc.)
                if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
                    xml = data.decode("utf-8", errors="replace")

                    # Pass 1: content-control (SDT) replacements
                    for field_key, alias in _ALIAS_MAP.items():
                        value = fields.get(field_key, "")
                        if value:
                            xml = _replace_sdt_content(xml, alias, value)

                    # Pass 2: brace-placeholder replacements
                    xml = _replace_brace_placeholders(xml, fields)

                    # Pass 3: date content control (matched by placeholder text)
                    if "letter_date" in fields and fields["letter_date"]:
                        date_val = _xml_escape(fields["letter_date"])
                        xml = re.sub(
                            r"<w:t[^>]*>Click or tap to enter a date\.</w:t>",
                            f"<w:t>{date_val}</w:t>",
                            xml,
                        )
                        # Also remove placeholder styling around it
                        xml = xml.replace(
                            '<w:rStyle w:val="PlaceholderText"/>',
                            "",
                        )

                    data = xml.encode("utf-8")

                zout.writestr(item, data)

        filled_bytes = out_buf.getvalue()

    except Exception as exc:
        logger.exception("Error processing docx template")
        return {"success": False, "error": f"Template processing failed: {exc}"}

    # 3. Return result
    b64 = base64.b64encode(filled_bytes).decode("ascii")
    filled_fields = [k for k in fields if fields[k]]

    return {
        "success": True,
        "document_base64": b64,
        "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "filename": "Verdantas_Letter_Filled.docx",
        "fields_filled": filled_fields,
        "message": f"Template filled successfully. {len(filled_fields)} field(s) replaced.",
    }
