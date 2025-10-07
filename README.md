Google Docs MCP Server
======================

FastMCP server exposing rich Google Docs tools (plus a Sheets charts helper). Handles OAuth, Drive copy, Docs batch updates (headers/footers/footnotes/bullets/images/tables), named ranges, and more.

Quick Start (uv)
----------------
1) Install
```
uv pip install -r googledoc_mcp/requirements.txt
```

2) Configure `.env` in `googledoc_mcp/`
```
GOOGLE_CREDENTIALS_PATH=credentials.json
GOOGLE_TOKEN_PATH=token.json
```

3) Enable APIs (once per GCP project)
- Google Docs API
- Google Drive API
- Google Sheets API (only if using charts tools)

4) Create OAuth client (Desktop)
- Console → APIs & Services → Credentials → Create Credentials → OAuth client ID → Desktop App
- Save the JSON as `googledoc_mcp/credentials.json` (or point `GOOGLE_CREDENTIALS_PATH` at it)

5) Run the server
```
uv run googledoc_mcp/mcp_server.py
```
On first run a browser opens for consent. After approval, `googledoc_mcp/token.json` is created automatically.

Tools (32 total)
----------------
Creation & retrieval
- GOOGLEDOCS_CREATE_DOCUMENT
- GOOGLEDOCS_COPY_DOCUMENT
- GOOGLEDOCS_CREATE_DOCUMENT_MARKDOWN
- GOOGLEDOCS_GET_DOCUMENT_BY_ID
- GOOGLEDOCS_SEARCH_DOCUMENTS

Headers / Footers / Footnotes
- GOOGLEDOCS_CREATE_HEADER
- GOOGLEDOCS_CREATE_FOOTER
- GOOGLEDOCS_DELETE_HEADER
- GOOGLEDOCS_DELETE_FOOTER
- GOOGLEDOCS_CREATE_FOOTNOTE

Content editing
- GOOGLEDOCS_INSERT_TEXT_ACTION
- GOOGLEDOCS_DELETE_CONTENT_RANGE
- GOOGLEDOCS_REPLACE_ALL_TEXT
- GOOGLEDOCS_UPDATE_DOCUMENT_MARKDOWN
- GOOGLEDOCS_UPDATE_DOCUMENT_STYLE
- GOOGLEDOCS_UPDATE_EXISTING_DOCUMENT

Media & layout
- GOOGLEDOCS_INSERT_INLINE_IMAGE
- GOOGLEDOCS_REPLACE_IMAGE
- GOOGLEDOCS_INSERT_PAGE_BREAK

Bullets & named ranges
- GOOGLEDOCS_CREATE_PARAGRAPH_BULLETS
- GOOGLEDOCS_DELETE_PARAGRAPH_BULLETS
- GOOGLEDOCS_CREATE_NAMED_RANGE
- GOOGLEDOCS_DELETE_NAMED_RANGE

Tables
- GOOGLEDOCS_INSERT_TABLE_ACTION
- GOOGLEDOCS_INSERT_TABLE_COLUMN
- GOOGLEDOCS_DELETE_TABLE
- GOOGLEDOCS_DELETE_TABLE_COLUMN
- GOOGLEDOCS_DELETE_TABLE_ROW
- GOOGLEDOCS_UNMERGE_TABLE_CELLS
- GOOGLEDOCS_UPDATE_TABLE_ROW_STYLE (supports tableStartLocation/rowIndices)

Sheets helpers
- GOOGLEDOCS_GET_CHARTS_FROM_SPREADSHEET
- GOOGLEDOCS_LIST_SPREADSHEET_CHARTS_ACTION

Usage Examples (payload shapes)
--------------------------------
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
- PowerShell: `&&` not supported → run `cd` and `uv run` on separate lines.
- 403 insufficient scopes → delete `googledoc_mcp/token.json`, ensure scopes in code, rerun and approve.
- 404 not found → verify the ID (copy between `/d/` and `/edit`) and ensure the signed-in account has access.
- Inline image 400 → ensure the URL is a public https image returning `Content-Type: image/*`.

Project Layout
--------------
- `googledoc_mcp/mcp_server.py`: FastMCP server and tools
- `googledoc_mcp/requirements.txt`: Python dependencies
- `googledoc_mcp/.env`: Env vars (credentials/token paths)
- `credentials.json`: OAuth client (you provide)
- `googledoc_mcp/token.json`: Created after initial OAuth consent

Security Notes
--------------
- Keep credentials and tokens private; do not commit them
- Treat `token.json` as sensitive (refresh token inside)
