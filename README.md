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

Obtaining credentials.json and how token.json is created
--------------------------------------------------------
Follow these steps once to create `credentials.json`. After that, the server will guide you through OAuth and write `token.json` automatically on first run.

1) Create a Google Cloud project (or reuse one)
- Open the Google Cloud Console: `https://console.cloud.google.com/`
- Select an existing project or create a new one

2) Enable required APIs in that project
- APIs & Services → Library
- Enable: Google Docs API and Google Drive API (Sheets API is optional for charts tools)

3) Configure the OAuth consent screen (first time only)
- APIs & Services → OAuth consent screen
- Choose External (recommended) or Internal depending on your organization
- App name can be anything (e.g., "Docs MCP")
- Add a support email and developer contact email
- Scopes: The app will request scopes at runtime, but you can add these as references:
  - `https://www.googleapis.com/auth/documents`
  - `https://www.googleapis.com/auth/drive.file`
  - `https://www.googleapis.com/auth/drive.readonly`
  - `https://www.googleapis.com/auth/spreadsheets.readonly` (optional)
- Test users: add the Google account(s) you will log in with if your app is in testing mode
- Save/publish changes

4) Create OAuth client credentials (Desktop application)
- APIs & Services → Credentials → Create Credentials → OAuth client ID
- Application type: Desktop app
- Click Create
- Download the JSON; this is your `credentials.json`

5) Put `credentials.json` where the server expects it
- Place the file at `googledoc_mcp/credentials.json` (or set `GOOGLE_CREDENTIALS_PATH` in `googledoc_mcp/.env` to the absolute/relative path)

6) First run to generate `token.json`
- Run: `uv run googledoc_mcp/mcp_server.py`
- A browser window opens to the Google consent screen; sign in with a test/allowed user and grant access
- On success, the server saves `googledoc_mcp/token.json` (path controlled by `GOOGLE_TOKEN_PATH`)
- Subsequent runs reuse `token.json` to refresh access without prompting again

Troubleshooting credentials/token
- 403 or scope errors: delete `googledoc_mcp/token.json` and run again to re-consent
- Wrong project or missing API enablement: double-check which GCP project your credentials belong to and that Docs/Drive APIs are enabled
- Multiple accounts: ensure the browser flow uses the same Google account that owns/has access to the target Docs/Drive files
