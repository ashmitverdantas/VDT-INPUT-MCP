from __future__ import annotations

import argparse
import json
import logging
import os

import uvicorn
from dotenv import load_dotenv
from fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import JSONResponse

load_dotenv()

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("verdantas-mcp")

# ── Tool handlers (imported after logging is set up) ───────────────────────────
from fill_letter_template   import handle_fill_letter_template,   FILL_LETTER_TEMPLATE_SCHEMA
from get_document_metadata  import handle_get_document_metadata,   GET_METADATA_SCHEMA

# ── FastMCP instance ───────────────────────────────────────────────────────────
mcp = FastMCP("verdantas-mcp-server")

# ── Tool: fillLetterTemplate ──────────────────────────────────────────────────

@mcp.tool()
async def fillLetterTemplate(
    documentUrl:          str = "",
    letter_date:          str = "",
    client_name:          str = "",
    company_name:         str = "",
    address:              str = "",
    email:                str = "",
    subject_line:         str = "",
    project_number:       str = "",
    letter_recipient:     str = "",
    cover_letter_body:    str = "",
    name_surname:         str = "",
    job_title:            str = "",
    contact_information:  str = "",
    attachment_1_title:   str = "",
    attachment_1_content: str = "",
    attachment_2_title:   str = "",
    attachment_2_content: str = "",
    attachment_3_content: str = "",
    attachment_4_content: str = "",
    attachment_5_content: str = "",
    attachment_6_content: str = "",
    attachment_7_content: str = "",
    attachment_8_content: str = "",
) -> dict:
    """
    Download the Verdantas letter template (.docx) from SharePoint, replace all
    placeholder fields with the supplied values, and return the completed document
    as base64-encoded bytes.

    Omitted fields (left as empty string) are kept as their original placeholder
    text in the document.

    Args:
        documentUrl:          SharePoint URL of the .docx template (uses standard
                              Verdantas template if omitted).
        letter_date:          e.g. "May 25, 2026"
        client_name:          Full name and designation, e.g. "Jane Smith, PE"
        company_name:         Client company name
        address:              Street, city/state/zip, phone — use \n between lines
        email:                Client email address
        subject_line:         Subject of the letter
        project_number:       Project number / ID
        letter_recipient:     Name after "Dear", e.g. "Ms. Smith"
        cover_letter_body:    Full body text; use \n\n for paragraph breaks
        name_surname:         Signatory full name
        job_title:            Signatory job title
        contact_information:  Signatory contact (email / phone)
        attachment_1_title:   Title for Attachment 1
        attachment_1_content: Body content for Attachment 1
        attachment_2_title:   Title for Attachment 2
        attachment_2_content: Body content for Attachment 2
        attachment_3_content: Management Plan content
        attachment_4_content: Team Qualifications content
        attachment_5_content: Schedule content
        attachment_6_content: Compliance Matrix content
        attachment_7_content: Budget Approach content
        attachment_8_content: Appendices content
    """
    # Collect only non-empty fields into the dict the handler expects
    fields = {
        k: v for k, v in {
            "letter_date":          letter_date,
            "client_name":          client_name,
            "company_name":         company_name,
            "address":              address,
            "email":                email,
            "subject_line":         subject_line,
            "project_number":       project_number,
            "letter_recipient":     letter_recipient,
            "cover_letter_body":    cover_letter_body,
            "name_surname":         name_surname,
            "job_title":            job_title,
            "contact_information":  contact_information,
            "attachment_1_title":   attachment_1_title,
            "attachment_1_content": attachment_1_content,
            "attachment_2_title":   attachment_2_title,
            "attachment_2_content": attachment_2_content,
            "attachment_3_content": attachment_3_content,
            "attachment_4_content": attachment_4_content,
            "attachment_5_content": attachment_5_content,
            "attachment_6_content": attachment_6_content,
            "attachment_7_content": attachment_7_content,
            "attachment_8_content": attachment_8_content,
        }.items() if v
    }
    return await handle_fill_letter_template({"documentUrl": documentUrl, "fields": fields})




# ── Tool: getDocumentMetadata ──────────────────────────────────────────────────

@mcp.tool()
async def getDocumentMetadata(documentUrl: str) -> dict:
    """
    Fetch lightweight metadata for a SharePoint document: format, size, SHA-256 hash,
    and whether text extraction is supported.  Does NOT return document text.

    Args:
        documentUrl: SharePoint URL of the document to inspect.
    """
    return await handle_get_document_metadata({"documentUrl": documentUrl})


# ── Tool: listSupportedFormats ─────────────────────────────────────────────────

@mcp.tool()
async def listSupportedFormats() -> list:
    """List all file formats supported for text extraction."""
    return [
        {"extension": ".pdf",  "description": "PDF documents (text-based)"},
        {"extension": ".docx", "description": "Microsoft Word (.docx)"},
        {"extension": ".doc",  "description": "Legacy Word (.doc) — requires antiword"},
        {"extension": ".xlsx", "description": "Excel workbook (.xlsx)"},
        {"extension": ".xls",  "description": "Legacy Excel (.xls)"},
        {"extension": ".jpg",  "description": "JPEG image — OCR (requires pytesseract)"},
        {"extension": ".png",  "description": "PNG image  — OCR (requires pytesseract)"},
        {"extension": ".gif",  "description": "GIF image  — OCR (requires pytesseract)"},
        {"extension": ".tiff", "description": "TIFF image — OCR (requires pytesseract)"},
        {"extension": ".txt",  "description": "Plain text"},
        {"extension": ".csv",  "description": "CSV (returned as tab-separated rows)"},
    ]


# ── Health check endpoint ──────────────────────────────────────────────────────

@mcp.custom_route("/health", methods=["GET"])
async def health_check(request: Request) -> JSONResponse:
    """Liveness probe — returns 200 if the server is running."""
    return JSONResponse({
        "status":  "healthy",
        "service": "verdantas-mcp-server",
        "tools": [
            "fillLetterTemplate",
            "analyzeDocuments",
            "listSharePointFiles",
            "getDocumentMetadata",
            "listSupportedFormats",
        ],
        "env_check": {
            "tenant_id_set":     bool(os.getenv("SHAREPOINT_TENANT_ID")),
            "client_id_set":     bool(os.getenv("SHAREPOINT_CLIENT_ID")),
            "client_secret_set": bool(os.getenv("SHAREPOINT_CLIENT_SECRET")),
        },
    })


# ── ASGI app (SSE transport) ───────────────────────────────────────────────────
# IMPORTANT: must be created AFTER all @mcp.tool() and @mcp.custom_route() decorators
app = mcp.http_app()


# ── stdio transport ────────────────────────────────────────────────────────────

async def run_stdio() -> None:
    """Run the server over stdio for local MCP clients (Claude Desktop, etc.)."""
    import asyncio
    from mcp.server.stdio import stdio_server

    # FastMCP exposes the underlying mcp.server.Server as mcp._mcp_server
    server = mcp._mcp_server
    async with stdio_server() as (read_stream, write_stream):
        logger.info("Verdantas MCP Server running on stdio")
        await server.run(read_stream, write_stream, server.create_initialization_options())


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import asyncio

    parser = argparse.ArgumentParser(description="Verdantas MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
        help="Transport mode (default: stdio)",
    )
    parser.add_argument("--host", default="0.0.0.0", help="SSE bind host (default 0.0.0.0)")
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.getenv("WEBSITES_PORT", os.getenv("PORT", "8000"))),
        help="SSE port (default 8000 or $PORT)",
    )
    args = parser.parse_args()

    if args.transport == "sse":
        logger.info("Verdantas MCP Server — SSE at http://%s:%s/sse", args.host, args.port)
        uvicorn.run(app, host=args.host, port=args.port, log_level="info")
    else:
        asyncio.run(run_stdio())
