Google Docs MCP Server
======================

A FastMCP server that exposes high-level tools for Google Docs (plus a Sheets charts helper). It handles OAuth, Drive copy, Docs batch updates (headers, footnotes, bullets, inline images), named ranges, and more.

Quick Start
-----------
1) Install
```
cd googledoc_mcp
pip install -r requirements.txt
```

2) Configure `.env` in `googledoc_mcp/`
```
GOOGLE_CREDENTIALS_PATH=credentials.json
GOOGLE_TOKEN_PATH=token.json
```

3) Enable APIs (once per GCP project)
- Google Docs API
- Google Drive API
- Google Sheets API (only if using the charts tool)

4) Create OAuth client (Desktop)
- Console → APIs & Services → Credentials → Create Credentials → OAuth client ID → Desktop App
- Download the JSON and save it as `googledoc_mcp/credentials.json` (or point `GOOGLE_CREDENTIALS_PATH` at it)

5) Run the server (Windows PowerShell)
```
cd "C:\Users\manus\OneDrive\Desktop\googledocs-mcp\googledoc_mcp"
python mcp_server.py
```
On first run a browser opens for consent. After approval, `token.json` is created automatically.

About token.json
----------------
- Contains approved scopes and refresh tokens after OAuth consent.
- Scope change or 403 “insufficient scopes”? Delete `token.json` and run again to re-consent.
- Keep it private; it grants access to your resources under the approved scopes.

Core Tools (What you can do)
----------------------------
- Create Docs from plain text or markdown
- Copy a Doc (Drive API)
- Add headers (default/first-page), footnotes
- Insert inline images (public https URLs)
- Create named ranges
- Apply paragraph bullets over a range
- Fetch a Doc by ID
- List charts in a Sheets file

Tools Overview (10 total)
-------------------------
- GOOGLEDOCS_CREATE_DOCUMENT
  - Purpose: Create a new Google Doc; optionally insert initial text at top.
  - Args: title (str, required), text (str, required)
  - Returns: { data: {documentId,title,revisionId,createdTime,modifiedTime}, error, successful }

- GOOGLEDOCS_COPY_DOCUMENT
  - Purpose: Duplicate an existing Doc (useful for templates).
  - Args: document_id (str, required), title (str, optional)
  - Returns: { data: {id,name,mimeType,parents}, error, successful }

- GOOGLEDOCS_CREATE_DOCUMENT_MARKDOWN
  - Purpose: Convert markdown (headers/bold/italic/paragraphs) into a new Doc.
  - Args: title (str, required), markdown_text (str, required)
  - Returns: { data: {documentId,title,revisionId,createdTime,modifiedTime}, error, successful }

- GOOGLEDOCS_CREATE_FOOTNOTE
  - Purpose: Insert a footnote reference at a location or end-of-segment.
  - Args: documentId (str, required), location (dict, optional), endOfSegmentLocation (dict, optional)
  - Returns: { data: {documentId,location,endOfSegmentLocation,replies}, error, successful }

- GOOGLEDOCS_CREATE_HEADER
  - Purpose: Add a header; supports DEFAULT or FIRST_PAGE.
  - Args: documentId (str, required), createHeader (dict, required: { type, sectionBreakLocation? })
  - Returns: { data: {documentId,createHeader,replies}, error, successful }

- GOOGLEDOCS_CREATE_NAMED_RANGE
  - Purpose: Define a named range over indices; validates/clamps bounds.
  - Args: documentId (str), name (str), rangeStartIndex (int), rangeEndIndex (int), rangeSegmentId (str, optional) — all required unless noted
  - Returns: { data: {documentId,name,range,replies}, error, successful }

- GOOGLEDOCS_CREATE_PARAGRAPH_BULLETS
  - Purpose: Apply bullet formatting to paragraphs within a text range.
  - Args: document_id (str, required), createParagraphBullets (dict, required: { range, bulletPreset? })
  - Returns: { data: {documentId,createParagraphBullets,replies}, error, successful }

- GOOGLEDOCS_GET_CHARTS_FROM_SPREADSHEET
  - Purpose: List all embedded charts in a Sheets file with chartId/spec.
  - Args: spreadsheet_id (str, required)
  - Returns: { data: {spreadsheetId,sheetsWithCharts:[{sheetTitle,charts:[{chartId,spec}]}]}, error, successful }

- GOOGLEDOCS_GET_DOCUMENT_BY_ID
  - Purpose: Retrieve a Docs file by ID; errors if not accessible or not found.
  - Args: id (str, required)
  - Returns: { data: {documentId,title,revisionId,body}, error, successful }

- GOOGLEDOCS_INSERT_INLINE_IMAGE
  - Purpose: Insert an inline image from a public https URL at a document index.
  - Args: documentId (str, required), location (dict, required: { index }), uri (str, required), objectSize (dict, optional)
  - Returns: { data: {documentId,location,uri,replies}, error, successful }

Usage Examples (payload shapes)
-------------------------------
Create a document
```
GOOGLEDOCS_CREATE_DOCUMENT
{
  "title": "My Doc",
  "text": "Hello world"
}
```

Insert inline image (public URL)
```
GOOGLEDOCS_INSERT_INLINE_IMAGE
{
  "documentId": "<doc-id>",
  "location": { "index": 1 },
  "uri": "https://www.google.com/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png",
  "objectSize": {
    "width":  { "magnitude": 120, "unit": "PT" },
    "height": { "magnitude": 40,  "unit": "PT" }
  }
}
```

Create bullets over a range
```
GOOGLEDOCS_CREATE_PARAGRAPH_BULLETS
{
  "document_id": "<doc-id>",
  "createParagraphBullets": {
    "range": { "startIndex": 1, "endIndex": 500 }
  }
}
```

Get charts from a spreadsheet
```
GOOGLEDOCS_GET_CHARTS_FROM_SPREADSHEET
{
  "spreadsheet_id": "<sheets-id>"
}
```

Troubleshooting
---------------
- PowerShell: `&&` not supported → run `cd` and `python` on separate lines.
- 403 insufficient scopes → delete `token.json`, ensure scopes in code, rerun and approve.
- 404 not found → verify the ID (copy between `/d/` and `/edit`) and sharing to the signed-in account.
- Inline image 400 → ensure the URL is a public https image that returns 200 with `Content-Type: image/*`.

Project Layout
--------------
- `googledoc_mcp/mcp_server.py`: FastMCP server and tools
- `googledoc_mcp/requirements.txt`: Python dependencies
- `googledoc_mcp/.env`: Env vars (credentials/token paths)
- `credentials.json`: OAuth client (you provide)
- `token.json`: Created after initial OAuth consent



