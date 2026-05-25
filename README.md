# Verdantas MCP Server

A [FastMCP](https://github.com/jlowin/fastmcp) server that lets any MCP-capable AI agent (Copilot, Claude, etc.) download a SharePoint document and fill all placeholders in the Verdantas letter template.

---

## Files

| File | Purpose |
|------|---------|
| `server.py` | FastMCP server — tool definitions and routing |
| `fill_letter_template.py` | Core handler: downloads template, fills SDT controls + brace placeholders, returns base64 docx |
| `get_document_metadata.py` | Handler: fetches lightweight metadata without returning document text |
| `sharepoint.py` | Microsoft Graph OAuth2 client: token refresh, file download, folder listing |
| `Verdantas_Template_WithControls.docx` | Bundled default template (used when no `documentUrl` is supplied) |
| `requirements.txt` | Python dependencies |
| `.env.example` | Copy to `.env` and fill in your Azure credentials |

---

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Configure credentials
cp .env.example .env
# Edit .env with your Azure App Registration values

# 3. Run (stdio — for Claude Desktop / Copilot Studio local testing)
python server.py --transport stdio

# 4. Run (SSE — for cloud / remote MCP clients)
python server.py --transport sse --port 8000
```

---

## MCP Tools

### `fillLetterTemplate`

Downloads the Verdantas letter template from SharePoint (or uses the bundled copy) and replaces all placeholders with your values.

**Parameters** — all optional; omitted fields are left as-is:

| Parameter | Description |
|-----------|-------------|
| `documentUrl` | SharePoint URL of the `.docx` template. Uses bundled template if blank. |
| `letter_date` | e.g. `"May 25, 2026"` |
| `client_name` | Full name + designation, e.g. `"Jane Smith, PE"` |
| `company_name` | Client company name |
| `address` | Multi-line address — use `\n` between lines |
| `email` | Client email address |
| `subject_line` | Letter subject |
| `project_number` | Project number / ID |
| `letter_recipient` | Name after "Dear", e.g. `"Ms. Smith"` |
| `cover_letter_body` | Full body text; use `\n\n` for paragraph breaks |
| `name_surname` | Signatory full name |
| `job_title` | Signatory job title |
| `contact_information` | Signatory contact details |
| `attachment_1_title` | Title for Attachment 1 |
| `attachment_1_content` | Body of Attachment 1 |
| `attachment_2_title` | Title for Attachment 2 |
| `attachment_2_content` | Body of Attachment 2 |
| `attachment_3_content` | Management Plan content |
| `attachment_4_content` | Team Qualifications content |
| `attachment_5_content` | Schedule content |
| `attachment_6_content` | Compliance Matrix content |
| `attachment_7_content` | Budget Approach content |
| `attachment_8_content` | Appendices content |

**Returns:**
```json
{
  "success": true,
  "document_base64": "<base64-encoded .docx bytes>",
  "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "filename": "Verdantas_Letter_Filled.docx",
  "fields_filled": ["letter_date", "client_name", ...],
  "message": "Template filled successfully. 12 field(s) replaced."
}
```

The calling agent should base64-decode `document_base64` and save or upload the result.

---

### `getDocumentMetadata`

Fetches metadata without returning document text. Useful for verifying a SharePoint file before processing.

**Parameters:** `documentUrl` (required)

**Returns:**
```json
{
  "success": true,
  "url": "...",
  "size_bytes": 213504,
  "sha256": "abc123...",
  "extension": ".docx",
  "format_label": "Microsoft Word (Open XML)",
  "text_extraction_supported": true,
  "content_controls": ["ClientName", "SubjectLine", ...],
  "brace_placeholders": ["Client Name", "Project Number", ...]
}
```

---

### `listSupportedFormats`

Returns the list of file formats supported for text extraction (no parameters).

---

## Placeholder Mechanism

The template uses two replacement strategies applied in order:

1. **Word Content Controls (SDT)** — Named fields embedded in the `.docx` XML (`w:sdt` with `w:alias`). These are the primary mechanism and are replaced first.

2. **Brace placeholders** — Raw `{Field Name}` tokens in the document text, replaced via regex as a fallback.

| Content Control Tag | Field |
|--------------------|-------|
| `ClientName` | `client_name` |
| `ClientCompany` | `company_name` |
| `ClientAddress` | `address` |
| `SubjectLine` | `subject_line` |
| `Salutation` | `letter_recipient` |
| `CoverLetterBody` | `cover_letter_body` |
| `SignatoryNames` | `name_surname` |
| `SignatoryTitles` | `job_title` |
| `SignatoryContacts` | `contact_information` |
| `Attachment1Title` / `Attachment1Content` | `attachment_1_title` / `attachment_1_content` |
| `Attachment2Title` / `Attachment2Content` | `attachment_2_title` / `attachment_2_content` |
| `Section3Content` – `Section8Content` | `attachment_3_content` – `attachment_8_content` |

---

## Azure App Registration (SharePoint access)

1. Go to **Azure Portal → App Registrations → New Registration**
2. Note the **Application (client) ID** and **Directory (tenant) ID**
3. Under **Certificates & Secrets**, create a **Client Secret**
4. Under **API Permissions**, add: `Microsoft Graph → Application → Sites.Read.All` (or `Sites.ReadWrite.All`)
5. **Grant admin consent**
6. Paste the three values into your `.env` file

---

## Copilot / Claude Integration

For **Copilot Studio** or **Claude** (via MCP), point your agent at the SSE endpoint:

```
http://your-server:8000/sse
```

The agent can then call `fillLetterTemplate` with the field values it has gathered from the user, receive the base64 docx, decode it, and save or email the resulting letter.
