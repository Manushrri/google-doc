Google Docs MCP Server
=======================

A FastMCP server exposing tools to work with Google Docs (and some Google Sheets helpers).

Prerequisites
-------------
- Python 3.10+
- A Google Cloud project with OAuth 2.0 Desktop Client credentials
- Google Docs API enabled (and Google Drive API). For Sheets-related tools, enable Google Sheets API.

Installation
------------
```
cd googledoc_mcp
pip install -r requirements.txt
```

Environment
-----------
Create a `.env` file inside `googledoc_mcp/`:
```
GOOGLE_CREDENTIALS_PATH=credentials.json
GOOGLE_TOKEN_PATH=token.json
```
- `credentials.json`: Downloaded OAuth client secret from Google Cloud Console (Desktop App).
- `token.json`: Will be created automatically after the first successful OAuth consent.

First Run & Authentication
--------------------------
```
cd googledoc_mcp
python mcp_server.py
```
- A browser window opens asking you to sign in and approve scopes.
- After consent, a `token.json` is created at the path specified in `.env`.
- If you change requested scopes later (e.g., add Sheets), delete `token.json` and run again to re-consent.

How to Create OAuth Credentials
-------------------------------
1. Go to Google Cloud Console → APIs & Services → Credentials.
2. Click “Create Credentials” → “OAuth client ID”.
3. Application type: “Desktop app”. Download the JSON file.
4. Save it as `googledoc_mcp/credentials.json` (or update `GOOGLE_CREDENTIALS_PATH`).

Enable Required APIs
--------------------
- Google Docs API
- Google Drive API
- Google Sheets API (only if using Sheets tools)

Enable each at: `https://console.developers.google.com/apis/library`

token.json Explained
--------------------
- Created after OAuth consent, stores refresh/access tokens and approved scopes.
- If scopes change or you get scope-related 403 errors, delete `token.json` and re-run to re-auth.
- Keep this file private; it grants access to your Google account under the approved scopes.

Running the Server (Windows PowerShell)
---------------------------------------
```
cd "C:\Users\manus\OneDrive\Desktop\googledocs-mcp\googledoc_mcp"
python mcp_server.py
```

Common Issues
-------------
- PowerShell error about `&&`: run the commands on separate lines (see above).
- 403 insufficient scopes: delete `token.json`, ensure required scopes in code, re-run.
- 404 not found: verify the ID has no hidden characters and that your signed-in account has access.

Tools (Examples)
----------------
- GOOGLEDOCS_CREATE_DOCUMENT(title, text)
- GOOGLEDOCS_COPY_DOCUMENT(document_id, title?)
- GOOGLEDOCS_CREATE_DOCUMENT_MARKDOWN(title, markdown_text)
- GOOGLEDOCS_CREATE_FOOTNOTE(documentId, location?, endOfSegmentLocation?)
- GOOGLEDOCS_CREATE_HEADER(documentId, createHeader)
- GOOGLEDOCS_INSERT_INLINE_IMAGE(documentId, location, uri, objectSize?)
- GOOGLEDOCS_GET_DOCUMENT_BY_ID(id)
- GOOGLEDOCS_CREATE_PARAGRAPH_BULLETS(document_id, createParagraphBullets)
- GOOGLEDOCS_GET_CHARTS_FROM_SPREADSHEET(spreadsheet_id)

Refer to each tool’s description in code for parameter details.


