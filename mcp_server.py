import os
import json
import re
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from typing import Optional, List, Dict, Any, Annotated

# Load .env file automatically from the same directory as this script
script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, '.env')
if os.path.exists(env_path):
    load_dotenv(env_path)
    print(f"Loaded environment variables from: {env_path}")
else:
    print(f"Warning: .env file not found at: {env_path}")

def get_env(var: str) -> str:
    """Fetch environment variable or raise error if missing."""
    value = os.getenv(var)
    if not value:
        # Try to load from .env file in script directory if not found
        script_dir = os.path.dirname(os.path.abspath(__file__))
        env_path = os.path.join(script_dir, '.env')
        if os.path.exists(env_path):
            load_dotenv(env_path, override=True)
            value = os.getenv(var)
        if not value:
            raise RuntimeError(f"Missing required environment variable: {var}")
    return value

def get_credentials() -> Credentials:
    """Get Google Docs API credentials."""
    creds = None
    token_path = get_env("GOOGLE_TOKEN_PATH")
    credentials_path = get_env("GOOGLE_CREDENTIALS_PATH")
    
    # Make paths relative to script directory if they're not absolute
    if not os.path.isabs(token_path):
        token_path = os.path.join(script_dir, token_path)
    if not os.path.isabs(credentials_path):
        credentials_path = os.path.join(script_dir, credentials_path)
    
    # Debug: Check if files exist
    if not os.path.exists(credentials_path):
        raise RuntimeError(f"Credentials file not found: {credentials_path}")
    if not os.path.exists(token_path):
        print(f"Token file not found: {token_path} (will be created on first auth)")
    
    # Load existing token if available
    if os.path.exists(token_path) and os.path.getsize(token_path) > 0:
        try:
            creds = Credentials.from_authorized_user_file(token_path)
        except Exception as e:
            print(f"Warning: Could not load token file: {e}")
            creds = None
    
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path, 
                [
                    'https://www.googleapis.com/auth/documents',
                    'https://www.googleapis.com/auth/drive.file',
                    'https://www.googleapis.com/auth/drive.readonly',
                    'https://www.googleapis.com/auth/spreadsheets.readonly'
                ]
            )
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    
    return creds

def get_docs_service():
    """Get Google Docs service object."""
    creds = get_credentials()
    return build('docs', 'v1', credentials=creds)

# Initialize MCP server
mcp = FastMCP("googledocs-mcp")

def docs_request(operation: str, document_id: str = None, **kwargs):
    """Helper for Google Docs API requests."""
    try:
        service = get_docs_service()
        
        if operation == "create":
            return service.documents().create(body=kwargs.get('body', {})).execute()
        elif operation == "get":
            return service.documents().get(documentId=document_id).execute()
        elif operation == "copy":
            # For copying documents, we need Drive API
            from googleapiclient.discovery import build as build_drive
            drive_service = build_drive('drive', 'v3', credentials=get_credentials())
            return drive_service.files().copy(
                fileId=document_id,
                body=kwargs.get('body', {})
            ).execute()
        elif operation == "batchUpdate":
            return service.documents().batchUpdate(
                documentId=document_id, 
                body=kwargs.get('body', {})
            ).execute()
        else:
            raise RuntimeError(f"Unknown operation: {operation}")
            
    except HttpError as e:
        raise RuntimeError(f"Google Docs API error: {e}")
    except Exception as e:
        raise RuntimeError(f"Google Docs request error: {str(e)}")

def _validate_required(params: Dict[str, Any], required: List[str]):
    """Raise ValueError if any required params are missing/blank.

    Treats empty strings, None, and empty lists as missing.
    """
    missing = []
    for key in required:
        value = params.get(key)
        if value is None:
            missing.append(key)
        elif isinstance(value, str) and value.strip() == "":
            missing.append(key)
        elif isinstance(value, (list, dict)) and len(value) == 0:
            missing.append(key)
    if missing:
        raise ValueError(f"Missing required parameter(s): {', '.join(missing)}")
    return None

def _markdown_to_google_docs_content(markdown_text: str) -> List[Dict[str, Any]]:
    """Convert markdown text to Google Docs content format.
    
    This is a basic implementation that handles:
    - Headers (# ## ###)
    - Bold text (**text**)
    - Italic text (*text*)
    - Line breaks
    - Basic paragraphs
    """
    lines = markdown_text.split('\n')
    content = []
    current_index = 1
    
    for line in lines:
        if not line.strip():
            # Empty line - add paragraph break
            content.append({
                "endIndex": current_index,
                "startIndex": current_index - 1,
                "paragraph": {
                    "elements": [{"endIndex": current_index, "startIndex": current_index - 1, "textRun": {"content": "\n"}}]
                }
            })
            current_index += 1
            continue
            
        # Check for headers
        if line.startswith('### '):
            # H3 header
            text = line[4:].strip()
            content.append({
                "endIndex": current_index + len(text),
                "startIndex": current_index - 1,
                "paragraph": {
                    "elements": [{
                        "endIndex": current_index + len(text),
                        "startIndex": current_index - 1,
                        "textRun": {
                            "content": text,
                            "textStyle": {"bold": True, "fontSize": {"magnitude": 14, "unit": "PT"}}
                        }
                    }]
                }
            })
            current_index += len(text) + 1
        elif line.startswith('## '):
            # H2 header
            text = line[3:].strip()
            content.append({
                "endIndex": current_index + len(text),
                "startIndex": current_index - 1,
                "paragraph": {
                    "elements": [{
                        "endIndex": current_index + len(text),
                        "startIndex": current_index - 1,
                        "textRun": {
                            "content": text,
                            "textStyle": {"bold": True, "fontSize": {"magnitude": 16, "unit": "PT"}}
                        }
                    }]
                }
            })
            current_index += len(text) + 1
        elif line.startswith('# '):
            # H1 header
            text = line[2:].strip()
            content.append({
                "endIndex": current_index + len(text),
                "startIndex": current_index - 1,
                "paragraph": {
                    "elements": [{
                        "endIndex": current_index + len(text),
                        "startIndex": current_index - 1,
                        "textRun": {
                            "content": text,
                            "textStyle": {"bold": True, "fontSize": {"magnitude": 18, "unit": "PT"}}
                        }
                    }]
                }
            })
            current_index += len(text) + 1
        else:
            # Regular paragraph - process bold and italic
            processed_text = line
            elements = []
            start_pos = 0
            
            # Simple bold and italic processing
            while True:
                # Find next bold or italic
                bold_match = re.search(r'\*\*(.*?)\*\*', processed_text[start_pos:])
                italic_match = re.search(r'\*(.*?)\*', processed_text[start_pos:])
                
                if not bold_match and not italic_match:
                    # No more formatting, add remaining text
                    remaining = processed_text[start_pos:]
                    if remaining:
                        elements.append({
                            "endIndex": current_index + len(remaining),
                            "startIndex": current_index - 1,
                            "textRun": {"content": remaining}
                        })
                        current_index += len(remaining)
                    break
                
                # Determine which comes first
                if bold_match and italic_match:
                    if bold_match.start() < italic_match.start():
                        match = bold_match
                        is_bold = True
                    else:
                        match = italic_match
                        is_bold = False
                elif bold_match:
                    match = bold_match
                    is_bold = True
                else:
                    match = italic_match
                    is_bold = False
                
                # Add text before the match
                before_text = processed_text[start_pos:start_pos + match.start()]
                if before_text:
                    elements.append({
                        "endIndex": current_index + len(before_text),
                        "startIndex": current_index - 1,
                        "textRun": {"content": before_text}
                    })
                    current_index += len(before_text)
                
                # Add the formatted text
                formatted_text = match.group(1)
                text_style = {"bold": True} if is_bold else {"italic": True}
                elements.append({
                    "endIndex": current_index + len(formatted_text),
                    "startIndex": current_index - 1,
                    "textRun": {
                        "content": formatted_text,
                        "textStyle": text_style
                    }
                })
                current_index += len(formatted_text)
                
                # Update start position
                start_pos += match.end()
            
            # Add paragraph with elements
            if elements:
                content.append({
                    "endIndex": current_index,
                    "startIndex": current_index - 1,
                    "paragraph": {"elements": elements}
                })
                current_index += 1
    
    return content

# -------------------- TOOLS --------------------

@mcp.tool(
    "GOOGLEDOCS_CREATE_DOCUMENT",
    description="Create Document. Creates a new Google Docs document using the provided title and, if non-empty, inserts the supplied text at the start of the body. Args: title (str): Name of the document (required). text (str): Initial body text (required; can be empty string). Returns: dict: { data: {documentId, title, revisionId, createdTime, modifiedTime}, error: str, successful: bool }.",
)
def GOOGLEDOCS_CREATE_DOCUMENT(
    title: Annotated[str, "The title of the document to create."],
    text: Annotated[str, "Initial text content for the document."]
):
    """Creates a new Google Docs document.

    Creates a new Google Docs document using the provided title as filename and inserts 
    the initial text at the beginning if non-empty, returning the document's id and 
    metadata (excluding body content).

    Args:
        title (str): The title of the document to create.
        text (str): Initial text content for the document.

    Returns:
        dict: Response containing data object with document metadata, error string, and success boolean.
    """
    err = _validate_required({"title": title, "text": text}, ["title", "text"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}
    
    try:
        body = {"title": title}
        if text.strip():
            body["body"] = {
                "content": [
                    {
                        "endIndex": 1,
                        "startIndex": 1,
                        "paragraph": {
                            "elements": [
                                {
                                    "endIndex": len(text) + 1,
                                    "startIndex": 1,
                                    "textRun": {"content": text}
                                }
                            ]
                        }
                    }
                ]
            }
    
        result = docs_request("create", body=body)
        
        return {
            "data": {
                "documentId": result.get("documentId"),
                "title": result.get("title"),
                "revisionId": result.get("revisionId"),
                "createdTime": result.get("createdTime"),
                "modifiedTime": result.get("modifiedTime")
            },
            "error": "",
            "successful": True
        }
        
    except Exception as e:
        return {
            "data": {},
            "error": f"Failed to create document: {str(e)}",
            "successful": False
        }




@mcp.tool(
    "GOOGLEDOCS_COPY_DOCUMENT",
    description="Copy Document. Duplicates an existing Google Docs file using the Drive API. Useful for templating. If no title is provided, Drive assigns a default (e.g., 'Copy of <title>'). Args: document_id (str): Source Docs file ID (required). title (str): New title (optional). Returns: dict: { data: {id, name, mimeType, parents}, error: str, successful: bool }.",
)
def GOOGLEDOCS_COPY_DOCUMENT(
    document_id: Annotated[str, "The ID of the Google Docs document to copy."],
    title: Annotated[Optional[str], "The title for the copied document. If not provided, will use 'Copy of [original title]'."] = None
):
    """Creates a copy of an existing Google Docs document.

    Creates a copy of an existing Google Docs document. Use this to duplicate a document, 
    for example, when using an existing document as a template. The copied document will 
    have a default title (e.g., 'Copy of [original title]') if no new title is provided, 
    and will be placed in the user's root Google Drive folder.

    Args:
        document_id (str): The ID of the Google Docs document to copy.
        title (str, optional): The title for the copied document. If not provided, 
            will use 'Copy of [original title]'.

    Returns:
        dict: Response containing data object with copied document information, error string, and success boolean.
    """
    err = _validate_required({"document_id": document_id}, ["document_id"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}
    
    try:
        # Prepare the copy request body
        copy_body = {}
        if title:
            copy_body["name"] = title
        
        # Execute the copy operation
        result = docs_request("copy", document_id=document_id, body=copy_body)
        
        return {
            "data": {
                "id": result.get("id"),
                "name": result.get("name"),
                "mimeType": result.get("mimeType"),
                "parents": result.get("parents", [])
            },
            "error": "",
            "successful": True
        }
        
    except Exception as e:
        return {
            "data": {},
            "error": f"Failed to copy document: {str(e)}",
            "successful": False
        }

@mcp.tool(
    "GOOGLEDOCS_CREATE_DOCUMENT_MARKDOWN",
    description="Create Document (Markdown). Converts provided markdown to formatted Google Docs content (basic headers, bold, italic, paragraphs) and creates a new document. Args: title (str): Document title (required). markdown_text (str): Markdown content to convert and insert (required). Returns: dict: { data: {documentId, title, revisionId, createdTime, modifiedTime}, error: str, successful: bool }.",
)
def GOOGLEDOCS_CREATE_DOCUMENT_MARKDOWN(
    title: Annotated[str, "The title of the document to create."],
    markdown_text: Annotated[str, "The markdown content to convert and insert into the document."]
):
    """Creates a new Google Docs document with markdown content.

    Creates a new Google Docs document, optionally initializing it with a title 
    and content provided as markdown text. The markdown will be converted to 
    formatted Google Docs content including headers, bold, italic, and paragraph formatting.

    Args:
        title (str): The title of the document to create.
        markdown_text (str): The markdown content to convert and insert into the document.

    Returns:
        dict: Response containing data object with document metadata, error string, and success boolean.
    """
    err = _validate_required({"title": title, "markdown_text": markdown_text}, ["title", "markdown_text"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}
    
    try:
        # Convert markdown to Google Docs content format
        content = _markdown_to_google_docs_content(markdown_text)
        
        # Create document body
        body = {"title": title}
        if content:
            body["body"] = {"content": content}
        
        result = docs_request("create", body=body)
        
        return {
            "data": {
                "documentId": result.get("documentId"),
                "title": result.get("title"),
                "revisionId": result.get("revisionId"),
                "createdTime": result.get("createdTime"),
                "modifiedTime": result.get("modifiedTime")
            },
            "error": "",
            "successful": True
        }
        
    except Exception as e:
        return {
            "data": {},
            "error": f"Failed to create markdown document: {str(e)}",
            "successful": False
        }

@mcp.tool(
    "GOOGLEDOCS_CREATE_FOOTNOTE",
    description="Create Footnote. Inserts a footnote reference at a specific index or at an end-of-segment location; automatically clamps out-of-range indices to a valid position. Args: documentId (str): Target Docs ID (required). location (dict): { index } insertion point (optional). endOfSegmentLocation (dict): end-of-segment location (optional). Returns: dict: { data: {documentId, location, endOfSegmentLocation, replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_CREATE_FOOTNOTE(
    documentId: Annotated[str, "The ID of the Google Docs document to add a footnote to."],
    location: Annotated[Optional[Dict[str, Any]], "The location where the footnote reference should be inserted. If not provided, footnote will be added at the beginning of the document."] = None,
    endOfSegmentLocation: Annotated[Optional[Dict[str, Any]], "Alternative location for the footnote reference. If both location and endOfSegmentLocation are provided, location takes precedence."] = None
):
    """Creates a new footnote in a Google document.

    Tool to create a new footnote in a Google document. Use this when you need to add a footnote 
    at a specific location or at the end of the document body.

    Args:
        documentId (str): The ID of the Google Docs document to add a footnote to.
        location (dict, optional): The location where the footnote reference should be inserted. 
            If not provided, footnote will be added at the beginning of the document.
        endOfSegmentLocation (dict, optional): Alternative location for the footnote reference. 
            If both location and endOfSegmentLocation are provided, location takes precedence.

    Returns:
        dict: Response containing data object with footnote information, error string, and success boolean.
    """
    err = _validate_required({"documentId": documentId}, ["documentId"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}
    
    try:
        # First, get the document to understand its structure and length
        doc = docs_request("get", document_id=documentId)
        
        # Get the document length
        body_content = doc.get("body", {}).get("content", [])
        if body_content:
            # Find the last element to get the document length
            last_element = body_content[-1]
            doc_length = last_element.get("endIndex", 1)
        else:
            doc_length = 1
        
        # Prepare the batch update request
        requests = []
        
        # Create footnote request - Google Docs API uses different field structure
        footnote_request = {
            "createFootnote": {}
        }
        
        # Add footnote reference location (required by API)
        if location:
            # Validate the location index
            if "index" in location and location["index"] >= doc_length:
                location["index"] = doc_length - 1  # Place at end of document
            footnote_request["createFootnote"]["location"] = location
        elif endOfSegmentLocation:
            footnote_request["createFootnote"]["endOfSegmentLocation"] = endOfSegmentLocation
        else:
            # If no location provided, use a valid location within the document
            valid_index = max(1, doc_length - 1)
            footnote_request["createFootnote"]["location"] = {"index": valid_index}
        
        requests.append(footnote_request)
        
        # Execute the batch update
        body = {"requests": requests}
        result = docs_request("batchUpdate", document_id=documentId, body=body)
        
        return {
            "data": {
                "documentId": documentId,
                "location": location,
                "endOfSegmentLocation": endOfSegmentLocation,
                "replies": result.get("replies", [])
            },
            "error": "",
            "successful": True
        }
        
    except Exception as e:
        return {
            "data": {},
            "error": f"Failed to create footnote: {str(e)}",
            "successful": False
        }

@mcp.tool(
    "GOOGLEDOCS_CREATE_HEADER",
    description="Create Header. Adds a header to the document (DEFAULT or FIRST_PAGE). If FIRST_PAGE is requested, the tool enables first-page headers automatically. Args: documentId (str): Docs ID (required). createHeader (dict): { type: 'DEFAULT'|'FIRST_PAGE', sectionBreakLocation? } (required). Returns: dict: { data: {documentId, createHeader, replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_CREATE_HEADER(
    documentId: Annotated[str, "The ID of the Google Docs document to add a header to."],
    createHeader: Annotated[Dict[str, Any], "The header configuration object containing type and optional section information."]
):
    """Creates a new header in a Google document.

    Tool to create a new header in a Google document. Use this tool when you need to add a header 
    to a document, optionally specifying the section it applies to.

    Args:
        documentId (str): The ID of the Google Docs document to add a header to.
        createHeader (dict): The header configuration object containing type and optional section information.

    Returns:
        dict: Response containing data object with header information, error string, and success boolean.
    """
    err = _validate_required({"documentId": documentId, "createHeader": createHeader}, ["documentId", "createHeader"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}
    
    try:
        # Prepare the batch update request
        requests = []
        
        # Create header request - ensure a valid type is always sent
        requested_type = createHeader.get("type") if isinstance(createHeader, dict) else None
        # Normalize common aliases to valid API enum values
        type_mapping = {
            "DEFAULT_HEADER": "DEFAULT",
            "FIRST_PAGE_HEADER": "FIRST_PAGE",
            "HEADER_FOOTER_TYPE_UNSPECIFIED": "DEFAULT",  # fall back to DEFAULT
        }
        normalized_type = type_mapping.get(requested_type, requested_type)
        if normalized_type not in ("DEFAULT", "FIRST_PAGE"):
            normalized_type = "DEFAULT"

        header_request = {
            "createHeader": {
                "type": normalized_type
            }
        }
        
        # Add sectionBreakLocation if provided
        if "sectionBreakLocation" in createHeader:
            header_request["createHeader"]["sectionBreakLocation"] = createHeader["sectionBreakLocation"]
        
        requests.append(header_request)
        
        # Execute the batch update
        body = {"requests": requests}
        result = docs_request("batchUpdate", document_id=documentId, body=body)
        
        return {
            "data": {
                "documentId": documentId,
                "createHeader": createHeader,
                "replies": result.get("replies", [])
            },
            "error": "",
            "successful": True
        }
        
    except Exception as e:
        error_str = str(e)
        # Check if the error is because header already exists
        if "already exists" in error_str:
            # Get the existing header ID from the document
            try:
                doc = docs_request("get", document_id=documentId)
                headers = doc.get("headers", {})
                if headers:
                    # Get the first header ID
                    header_id = list(headers.keys())[0]
                    return {
                        "data": {
                            "documentId": documentId,
                            "createHeader": createHeader,
                            "replies": [{
                                "createHeader": {
                                    "headerId": header_id
                                }
                            }]
                        },
                        "error": "",
                        "successful": True
                    }
            except:
                pass
        
        return {
            "data": {},
            "error": f"Failed to create header: {error_str}",
            "successful": False
        }

@mcp.tool(
    "GOOGLEDOCS_CREATE_FOOTER",
    description="Create Footer. Tool to create a new footer in a Google document. Use when you need to add a footer, optionally specifying its type and the section it applies to. Args: document_id (str): Docs ID (required). createFooter (dict): { type: 'DEFAULT'|'FIRST_PAGE', sectionBreakLocation? } (required). Returns: dict: { data: {documentId, createFooter, replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_CREATE_FOOTER(
    document_id: Annotated[str, "The ID of the Google Docs document to add a footer to."],
    createFooter: Annotated[Dict[str, Any], "The footer configuration object containing type and optional section information."]
):
    """Creates a new footer in a Google document.

    Tool to create a new footer in a Google document. Use this when you need to add a footer 
    to a document, optionally specifying the section it applies to.

    Args:
        document_id (str): The ID of the Google Docs document to add a footer to.
        createFooter (dict): The footer configuration object containing type and optional section information.

    Returns:
        dict: Response containing data object with footer information, error string, and success boolean.
    """
    err = _validate_required({"document_id": document_id, "createFooter": createFooter}, ["document_id", "createFooter"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}
    
    try:
        # Prepare the batch update request
        requests = []
        
        # Create footer request - ensure a valid type is always sent
        requested_type = createFooter.get("type") if isinstance(createFooter, dict) else None
        # Normalize common aliases to valid API enum values
        type_mapping = {
            "DEFAULT_FOOTER": "DEFAULT",
            "FIRST_PAGE_FOOTER": "FIRST_PAGE",
            "HEADER_FOOTER_TYPE_UNSPECIFIED": "DEFAULT",  # fall back to DEFAULT
        }
        normalized_type = type_mapping.get(requested_type, requested_type)
        if normalized_type not in ("DEFAULT", "FIRST_PAGE"):
            normalized_type = "DEFAULT"

        # Try using the exact structure that worked before
        footer_request = {
            "createFooter": {
                "type": normalized_type
            }
        }
        
        # Add sectionBreakLocation if provided
        if "sectionBreakLocation" in createFooter:
            footer_request["createFooter"]["sectionBreakLocation"] = createFooter["sectionBreakLocation"]
        
        requests.append(footer_request)
        
        # Execute the batch update
        body = {"requests": requests}
        result = docs_request("batchUpdate", document_id=document_id, body=body)
        
        return {
            "data": {
                "documentId": document_id,
                "createFooter": createFooter,
                "replies": result.get("replies", [])
            },
            "error": "",
            "successful": True
        }
        
    except Exception as e:
        error_str = str(e)
        # Check if the error is because footer already exists
        if "already exists" in error_str:
            # Get the existing footer ID from the document
            try:
                doc = docs_request("get", document_id=document_id)
                footers = doc.get("footers", {})
                if footers:
                    # Get the first footer ID
                    footer_id = list(footers.keys())[0]
                    return {
                        "data": {
                            "documentId": document_id,
                            "createFooter": createFooter,
                            "replies": [{
                                "createFooter": {
                                    "footerId": footer_id
                                }
                            }]
                        },
                        "error": "",
                        "successful": True
                    }
            except:
                pass
        
        return {
            "data": {},
            "error": f"Failed to create footer: {error_str}",
            "successful": False
        }

# New tool: Create Named Range
@mcp.tool(
    "GOOGLEDOCS_CREATE_NAMED_RANGE",
    description="Create Named Range. Defines a named range over a start/end index span; indices are validated against the document length and clamped if needed. Args: documentId (str): Docs ID (required). name (str): Named range label (required). rangeStartIndex (int): Inclusive start index (required). rangeEndIndex (int): Exclusive end index (required). rangeSegmentId (str): Segment ID for headers/footers (optional). Returns: dict: { data: {documentId, name, range, replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_CREATE_NAMED_RANGE(
    documentId: Annotated[str, "The ID of the Google Docs document to add the named range to."],
    name: Annotated[str, "The name to assign to the range."],
    rangeStartIndex: Annotated[int, "The start index of the range (inclusive)."],
    rangeEndIndex: Annotated[int, "The end index of the range (exclusive)."],
    rangeSegmentId: Annotated[Optional[str], "The segmentId for the range (omit for body)."] = None,
):
    """Creates a named range in a Google document.

    Creates a new named range in a Google document. Use this to assign a name to a specific
    part of the document for easier reference or programmatic manipulation.

    Args:
        documentId (str): The ID of the target Google Docs document.
        name (str): The name to assign to the range.
        rangeStartIndex (int): Start index (inclusive) of the range.
        rangeEndIndex (int): End index (exclusive) of the range.
        rangeSegmentId (str, optional): Segment ID if targeting headers/footers; omit for body.

    Returns:
        dict: Response containing data object with named range info, error string, and success boolean.
    """
    err = _validate_required(
        {"documentId": documentId, "name": name, "rangeStartIndex": rangeStartIndex, "rangeEndIndex": rangeEndIndex},
        ["documentId", "name", "rangeStartIndex", "rangeEndIndex"],
    )
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        # Fetch document to determine valid index bounds
        doc = docs_request("get", document_id=documentId)
        body_content = doc.get("body", {}).get("content", [])
        if body_content:
            last_element = body_content[-1]
            doc_length = last_element.get("endIndex", 1)
        else:
            doc_length = 1

        # Normalize indices to valid bounds
        start_index = max(1, int(rangeStartIndex))
        end_index = max(1, int(rangeEndIndex))
        if end_index > doc_length:
            end_index = doc_length
        if start_index >= end_index:
            return {
                "data": {},
                "error": f"Invalid range: start_index ({start_index}) must be < end_index ({end_index}).",
                "successful": False,
            }

        create_named_range = {
            "createNamedRange": {
                "name": name,
                "range": {
                    "startIndex": start_index,
                    "endIndex": end_index,
                },
            }
        }
        if rangeSegmentId:
            create_named_range["createNamedRange"]["range"]["segmentId"] = rangeSegmentId

        body = {"requests": [create_named_range]}
        result = docs_request("batchUpdate", document_id=documentId, body=body)

        return {
            "data": {
                "documentId": documentId,
                "name": name,
                "range": {
                    "startIndex": start_index,
                    "endIndex": end_index,
                    **({"segmentId": rangeSegmentId} if rangeSegmentId else {}),
                },
                "replies": result.get("replies", []),
            },
            "error": "",
            "successful": True,
        }

    except Exception as e:
        return {
            "data": {},
            "error": f"Failed to create named range: {str(e)}",
            "successful": False,
        }

# Add bullets to paragraphs
@mcp.tool(
    "GOOGLEDOCS_CREATE_PARAGRAPH_BULLETS",
    description="Create Paragraph Bullets. Applies bullet formatting to paragraphs fully or partially covered by the provided text range; removes unspecified presets that are rejected by the API. Args: document_id (str): Docs ID (required). createParagraphBullets (dict): { range: {startIndex,endIndex[,segmentId]}, bulletPreset? } (required). Returns: dict: { data: {documentId, createParagraphBullets, replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_CREATE_PARAGRAPH_BULLETS(
    document_id: Annotated[str, "The ID of the Google Docs document to update."],
    createParagraphBullets: Annotated[Dict[str, Any], "The request object including range (startIndex/endIndex[/segmentId]) and optional bullet settings like bulletPreset."]
):
    """Adds bullets to paragraphs within a range.

    Args:
        document_id (str): Target document ID.
        createParagraphBullets (dict): Must include a valid range. May include bulletPreset.

    Returns:
        dict: Response with replies from Docs API.
    """
    err = _validate_required({"document_id": document_id, "createParagraphBullets": createParagraphBullets}, ["document_id", "createParagraphBullets"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        # Fetch document to clamp indices
        doc = docs_request("get", document_id=document_id)
        body_content = doc.get("body", {}).get("content", [])
        if body_content:
            doc_length = body_content[-1].get("endIndex", 1)
        else:
            doc_length = 1

        request_obj = dict(createParagraphBullets or {})
        rng = request_obj.get("range") or {}
        start_index = int(rng.get("startIndex", 1))
        end_index = int(rng.get("endIndex", start_index + 1))
        # Clamp to valid bounds
        start_index = max(1, start_index)
        end_index = max(start_index + 1, min(end_index, doc_length))

        new_range: Dict[str, Any] = {"startIndex": start_index, "endIndex": end_index}
        if "segmentId" in rng and rng["segmentId"]:
            new_range["segmentId"] = rng["segmentId"]

        request_obj["range"] = new_range
        
        # If preset is explicitly UNSPECIFIED, remove it so the API uses default.
        if "bulletPreset" in request_obj:
            if str(request_obj["bulletPreset"]).strip() == "BULLET_GLYPH_PRESET_UNSPECIFIED":
                del request_obj["bulletPreset"]

        body = {"requests": [{"createParagraphBullets": request_obj}]}
        result = docs_request("batchUpdate", document_id=document_id, body=body)

        return {
            "data": {
                "documentId": document_id,
                "createParagraphBullets": request_obj,
                "replies": result.get("replies", []),
            },
            "error": "",
            "successful": True,
        }

    except Exception as e:
        return {
            "data": {},
            "error": f"Failed to create paragraph bullets: {str(e)}",
            "successful": False,
        }

# -------------------- GOOGLE SHEETS TOOLS --------------------

@mcp.tool(
    "GOOGLEDOCS_GET_CHARTS_FROM_SPREADSHEET",
    description="Get Charts from Spreadsheet. Retrieves all embedded charts across sheets in a spreadsheet and returns each chart's ID and spec. Args: spreadsheet_id (str): Google Sheets ID (required). Returns: dict: { data: {spreadsheetId, sheetsWithCharts: [{sheetTitle, charts: [{chartId,spec}]}]}, error: str, successful: bool }.",
)
def GOOGLEDOCS_GET_CHARTS_FROM_SPREADSHEET(
    spreadsheet_id: Annotated[str, "The ID of the Google Sheets spreadsheet to inspect for charts."]
):
    """Lists charts in a Google Sheets spreadsheet.

    Args:
        spreadsheet_id (str): Target spreadsheet ID.

    Returns:
        dict: data with list of charts per sheet, error, successful.
    """
    err = _validate_required({"spreadsheet_id": spreadsheet_id}, ["spreadsheet_id"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        from googleapiclient.discovery import build as build_sheets
        creds = get_credentials()
        sheets_service = build_sheets('sheets', 'v4', credentials=creds)

        resp = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            includeGridData=False,
            fields='spreadsheetId,sheets(properties(title),charts(chartId,spec))'
        ).execute()

        sheets = resp.get('sheets', [])
        charts_summary = []
        for sh in sheets:
            sheet_title = sh.get('properties', {}).get('title')
            charts = sh.get('charts', []) or []
            extracted = []
            for ch in charts:
                extracted.append({
                    'chartId': ch.get('chartId'),
                    'spec': ch.get('spec')
                })
            if extracted:
                charts_summary.append({
                    'sheetTitle': sheet_title,
                    'charts': extracted
                })

        return {
            'data': {
                'spreadsheetId': spreadsheet_id,
                'sheetsWithCharts': charts_summary
            },
            'error': '',
            'successful': True
        }
    except Exception as e:
        return {
            'data': {},
            'error': f'Failed to get charts: {str(e)}',
            'successful': False
        }

# -------------------- GOOGLE DOCS UTILITIES --------------------

@mcp.tool(
    "GOOGLEDOCS_GET_DOCUMENT_BY_ID",
    description="Get Document by ID. Fetches an existing Google Docs document by its ID; returns 404-style error information if not found. Args: id (str): Document ID (required). Returns: dict: { data: {documentId, title, revisionId, body}, error: str, successful: bool }.",
)
def GOOGLEDOCS_GET_DOCUMENT_BY_ID(
    id: Annotated[str, "The Google Docs document ID to retrieve."]
):
    """Get a Google Docs document by ID.

    Args:
        id (str): The document ID to fetch.

    Returns:
        dict: Response containing data with document info, error string, and success boolean.
    """
    err = _validate_required({"id": id}, ["id"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        result = docs_request("get", document_id=id)
        return {
            "data": {
                "documentId": result.get("documentId"),
                "title": result.get("title"),
                "revisionId": result.get("revisionId"),
                "body": result.get("body", {})
            },
            "error": "",
            "successful": True
        }
    except Exception as e:
        return {
            "data": {},
            "error": f"Failed to get document: {str(e)}",
            "successful": False
        }

@mcp.tool(
    "GOOGLEDOCS_INSERT_PAGE_BREAK",
    description="Insert Page Break. Inserts a page break at a given location or at the end of a segment. Args: documentId (str): Docs ID (required). insertPageBreak (object): The request object as per Docs API; provide either location {index} or endOfSegmentLocation {segmentId} (required). Returns: dict: { data: {documentId, request, replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_INSERT_PAGE_BREAK(
    documentId: Annotated[str, "The ID of the Google Docs document to update."],
    insertPageBreak: Annotated[Dict[str, Any], "InsertPageBreak request object with either location or endOfSegmentLocation."]
):
    """Insert a page break into a Google Doc.

    Validates the provided index against the document length when using location.index.
    """
    err = _validate_required({"documentId": documentId, "insertPageBreak": insertPageBreak}, ["documentId", "insertPageBreak"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req = dict(insertPageBreak or {})
        # Clamp index if provided
        if "location" in req and isinstance(req["location"], dict) and "index" in req["location"]:
            doc = docs_request("get", document_id=documentId)
            body_content = doc.get("body", {}).get("content", [])
            doc_length = body_content[-1].get("endIndex", 1) if body_content else 1
            idx = int(req["location"]["index"])
            if idx >= doc_length:
                req["location"]["index"] = max(1, doc_length - 1)
            if idx < 1:
                req["location"]["index"] = 1

        result = docs_request("batchUpdate", document_id=documentId, body={"requests": [{"insertPageBreak": req}]})
        return {
            "data": {"documentId": documentId, "request": req, "replies": result.get("replies", [])},
            "error": "",
            "successful": True,
        }
    except Exception as e:
        return {"data": {}, "error": f"Failed to insert page break: {str(e)}", "successful": False}

@mcp.tool(
    "GOOGLEDOCS_INSERT_TABLE_ACTION",
    description="Insert Table in Google Doc. Adds a table at a specific index or end of a segment (body/header/footer). Args: documentId (str): Docs ID (required). rows (int): Number of rows (required). columns (int): Number of columns (required). index (int): Text index to insert at (optional). insertAtEndOfSegment (bool): If true, ignore index and insert at end of segment (optional). segmentId (str): Segment to target when inserting at end (optional). tabId (str): Ignored placeholder (optional). Returns: dict: { data: {documentId, request, replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_INSERT_TABLE_ACTION(
    documentId: Annotated[str, "The ID of the Google Docs document to update."],
    rows: Annotated[int, "Number of rows to create."],
    columns: Annotated[int, "Number of columns to create."],
    index: Annotated[Optional[int], "Insertion index if not inserting at end of segment." ] = None,
    insertAtEndOfSegment: Annotated[Optional[bool], "If true, insert at end of segment (body/header/footer)."] = None,
    segmentId: Annotated[Optional[str], "Segment ID when targeting headers/footers."] = None,
    tabId: Annotated[Optional[str], "Unused placeholder to match client signature."] = None,
):
    """Insert a table into a Google Doc at a location or end-of-segment."""
    err = _validate_required({"documentId": documentId, "rows": rows, "columns": columns}, ["documentId", "rows", "columns"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        insert_req: Dict[str, Any] = {"rows": int(rows), "columns": int(columns)}
        if insertAtEndOfSegment:
            eos: Dict[str, Any] = {}
            if segmentId:
                eos["segmentId"] = segmentId
            insert_req["endOfSegmentLocation"] = eos
        else:
            # Use provided index or clamp to end-1
            target_index = int(index) if index is not None else None
            if target_index is None:
                doc = docs_request("get", document_id=documentId)
                content = doc.get("body", {}).get("content", [])
                doc_length = content[-1].get("endIndex", 1) if content else 1
                target_index = max(1, doc_length - 1)
            insert_req["location"] = {"index": max(1, int(target_index))}

        result = docs_request("batchUpdate", document_id=documentId, body={"requests": [{"insertTable": insert_req}]})
        return {
            "data": {"documentId": documentId, "request": insert_req, "replies": result.get("replies", [])},
            "error": "",
            "successful": True,
        }
    except Exception as e:
        return {"data": {}, "error": f"Failed to insert table: {str(e)}", "successful": False}

@mcp.tool(
    "GOOGLEDOCS_INSERT_TABLE_COLUMN",
    description="Insert Table Column. Adds a column to an existing table using raw Docs API requests. Args: document_id (str): Docs ID (required). requests (array): Array of Docs API request objects (required), typically with insertTableColumn entries (e.g., {insertTableColumn:{tableCellLocation:{tableStartLocation:{index},rowIndex,columnIndex}, insertRight:true}}). Returns: dict: { data: {documentId, replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_INSERT_TABLE_COLUMN(
    document_id: Annotated[str, "The ID of the Google Docs document to update."],
    requests: Annotated[List[Dict[str, Any]], "Docs API batchUpdate requests array containing insertTableColumn operations."]
):
    """Insert a table column by passing through Docs API batchUpdate requests."""
    err = _validate_required({"document_id": document_id, "requests": requests}, ["document_id", "requests"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        body = {"requests": list(requests)}
        result = docs_request("batchUpdate", document_id=document_id, body=body)
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to insert table column: {str(e)}", "successful": False}

# -------------------- EXTRA SHEETS/DOCS TOOLS --------------------

@mcp.tool(
    "GOOGLEDOCS_LIST_SPREADSHEET_CHARTS_ACTION",
    description="List Charts from Spreadsheet. Retrieves chart ids and metadata from a Google Sheets spreadsheet for embedding into Google Docs. Args: spreadsheet_id (str): Sheets ID (required). fields_mask (str): Optional fields mask; defaults to sheets(properties(sheetId,title),charts(chartId,spec(title,altText))). Returns: dict: { data: {spreadsheetId,sheetsWithCharts}, error: str, successful: bool }.",
)
def GOOGLEDOCS_LIST_SPREADSHEET_CHARTS_ACTION(
    spreadsheet_id: Annotated[str, "The Google Sheets spreadsheet ID."],
    fields_mask: Annotated[Optional[str], "Optional fields mask for spreadsheets.get."] = None,
):
    """List charts in a spreadsheet with optional fields mask."""
    err = _validate_required({"spreadsheet_id": spreadsheet_id}, ["spreadsheet_id"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        from googleapiclient.discovery import build as build_sheets
        creds = get_credentials()
        sheets_service = build_sheets('sheets', 'v4', credentials=creds)

        fields = fields_mask or 'sheets(properties(sheetId,title),charts(chartId,spec(title,altText)))'
        # Always include spreadsheetId for reference
        if 'spreadsheetId' not in fields:
            fields = f'spreadsheetId,{fields}'

        resp = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            includeGridData=False,
            fields=fields
        ).execute()

        charts_summary: List[Dict[str, Any]] = []
        for sh in resp.get('sheets', []):
            props = sh.get('properties', {})
            title = props.get('title')
            charts = sh.get('charts', []) or []
            if charts:
                charts_summary.append({
                    'sheetTitle': title,
                    'charts': charts
                })

        return {
            'data': {
                'spreadsheetId': resp.get('spreadsheetId', spreadsheet_id),
                'sheetsWithCharts': charts_summary
            },
            'error': '',
            'successful': True
        }
    except Exception as e:
        return {'data': {}, 'error': f'Failed to list charts: {str(e)}', 'successful': False}


@mcp.tool(
    "GOOGLEDOCS_REPLACE_ALL_TEXT",
    description="Replace All Text in Document. Replaces all occurrences of a string with another across the document. Args: document_id (str): Docs ID (required). find_text (str): Text to find (required). replace_text (str): Replacement text (required). match_case (bool): Case sensitive match (required). search_by_regex (bool): If true, attempts regex (Docs replaceAllText does not support full regex; best-effort). tab_ids (array): Ignored/unused placeholder. Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_REPLACE_ALL_TEXT(
    document_id: Annotated[str, "The Google Docs document ID."],
    find_text: Annotated[str, "Text to find."],
    replace_text: Annotated[str, "Text to replace with."],
    match_case: Annotated[bool, "Whether match is case-sensitive."],
    search_by_regex: Annotated[Optional[bool], "Best-effort regex flag (Docs API has limited support)."] = None,
    tab_ids: Annotated[Optional[List[str]], "Unused placeholder for compatibility."] = None,
):
    """Replace all matching text throughout the document."""
    err = _validate_required({"document_id": document_id, "find_text": find_text, "replace_text": replace_text, "match_case": match_case}, ["document_id", "find_text", "replace_text", "match_case"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req = {
            'replaceAllText': {
                'containsText': {
                    'text': find_text,
                    'matchCase': bool(match_case)
                },
                'replaceText': replace_text
            }
        }
        result = docs_request('batchUpdate', document_id=document_id, body={'requests': [req]})
        return {"data": {"documentId": document_id, "replies": result.get('replies', [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to replace text: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_REPLACE_IMAGE",
    description="Replace Image in Document. Replaces a specific image with a new image from a URI. Args: document_id (str): Docs ID (required). replace_image (object): Docs replaceImage request body (required) e.g., {imageObjectId, uri, imageReplaceMethod?}. Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_REPLACE_IMAGE(
    document_id: Annotated[str, "The Google Docs document ID."],
    replace_image: Annotated[Dict[str, Any], "The replaceImage request object as per Docs API."]
):
    """Replace an existing image via Docs API replaceImage."""
    err = _validate_required({"document_id": document_id, "replace_image": replace_image}, ["document_id", "replace_image"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req = {"replaceImage": replace_image}
        result = docs_request('batchUpdate', document_id=document_id, body={'requests': [req]})
        return {"data": {"documentId": document_id, "replies": result.get('replies', [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to replace image: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_SEARCH_DOCUMENTS",
    description="Search Documents. Searches Google Drive for Google Docs using filters like name, date ranges, sharing, starred, and more. Args: query (str): Free-form Drive query (optional). created_after (str): RFC3339 time (optional). modified_after (str): RFC3339 time (optional). include_trashed (bool): Include trashed (optional). shared_with_me (bool): Shared-with-me only (optional). starred_only (bool): Starred only (optional). order_by (str): Drive orderBy (default 'modifiedTime desc'). max_results (int): Page size (default 10). Returns: dict: { data: {files}, error: str, successful: bool }.",
)
def GOOGLEDOCS_SEARCH_DOCUMENTS(
    query: Annotated[Optional[str], "Optional Drive query string." ] = None,
    created_after: Annotated[Optional[str], "RFC3339 createdTime > value."] = None,
    modified_after: Annotated[Optional[str], "RFC3339 modifiedTime > value."] = None,
    include_trashed: Annotated[Optional[bool], "Include trashed files."] = None,
    shared_with_me: Annotated[Optional[bool], "Only files shared with me."] = None,
    starred_only: Annotated[Optional[bool], "Only starred files."] = None,
    order_by: Annotated[Optional[str], "Drive orderBy param."] = 'modifiedTime desc',
    max_results: Annotated[Optional[int], "Max results (page size)." ] = 10,
):
    """Search Google Drive for Google Docs files with filters."""
    try:
        from googleapiclient.discovery import build as build_drive
        drive = build_drive('drive', 'v3', credentials=get_credentials())

        q_parts = ["mimeType='application/vnd.google-apps.document'"]
        if query:
            # Perform a name contains search if simple query provided
            q_parts.append(f"name contains '{query.replace("'", "\\'")}'")
        if created_after:
            q_parts.append(f"createdTime > '{created_after}'")
        if modified_after:
            q_parts.append(f"modifiedTime > '{modified_after}'")
        if shared_with_me:
            q_parts.append('sharedWithMe')
        if starred_only:
            q_parts.append('starred = true')
        if not include_trashed:
            q_parts.append('trashed = false')

        q = ' and '.join(q_parts)
        resp = drive.files().list(q=q, orderBy=order_by or 'modifiedTime desc', pageSize=int(max_results or 10), fields='files(id,name,mimeType,owners,createdTime,modifiedTime,starred)').execute()
        return {"data": {"files": resp.get('files', [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to search documents: {str(e)}", "successful": False}

@mcp.tool(
    "GOOGLEDOCS_INSERT_INLINE_IMAGE",
    description="Insert Inline Image. Inserts an image from a publicly accessible https URL at a given document index; optionally sets size in points. Args: documentId (str): Docs ID (required). location (dict): { index } insertion point (required). uri (str): Public image URL (required). objectSize (dict): { width:{magnitude,unit}, height:{magnitude,unit} } (optional). Returns: dict: { data: {documentId, location, uri, replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_INSERT_INLINE_IMAGE(
    documentId: Annotated[str, "The ID of the Google Docs document to insert the image into."],
    location: Annotated[Dict[str, Any], "The location where the image should be inserted. Usually a { 'index': number }."],
    uri: Annotated[str, "Publicly accessible image URL to insert."],
    objectSize: Annotated[Optional[Dict[str, Any]], "Optional object size with width/height in PT, e.g. { 'height': {'magnitude': 100, 'unit': 'PT'}, 'width': {'magnitude': 100, 'unit': 'PT'} }."] = None,
):
    """Insert an inline image into a Google Docs document.

    Validates the target index against the document length and inserts the image
    using Docs batchUpdate insertInlineImage.
    """
    err = _validate_required({"documentId": documentId, "location": location, "uri": uri}, ["documentId", "location", "uri"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        # Fetch document to clamp index
        doc = docs_request("get", document_id=documentId)
        body_content = doc.get("body", {}).get("content", [])
        if body_content:
            doc_length = body_content[-1].get("endIndex", 1)
        else:
            doc_length = 1

        image_location = dict(location or {})
        if "index" in image_location:
            if image_location["index"] >= doc_length:
                image_location["index"] = max(1, doc_length - 1)
            if image_location["index"] < 1:
                image_location["index"] = 1

        request: Dict[str, Any] = {
            "insertInlineImage": {
                "location": image_location,
                "uri": uri,
            }
        }
        if objectSize:
            request["insertInlineImage"]["objectSize"] = objectSize

        result = docs_request("batchUpdate", document_id=documentId, body={"requests": [request]})

        return {
            "data": {
                "documentId": documentId,
                "location": image_location,
                "uri": uri,
                "replies": result.get("replies", []),
            },
            "error": "",
            "successful": True,
        }
    except Exception as e:
        return {
            "data": {},
            "error": f"Failed to insert inline image: {str(e)}",
            "successful": False,
        }

# ---------------------- Additional Update/Formatting Tools ----------------------

@mcp.tool(
    "GOOGLEDOCS_UNMERGE_TABLE_CELLS",
    description="Unmerge Table Cells. Tool to unmerge previously merged cells in a table. Use this when you need to revert merged cells in a Google document table back to their individual cell states. Args: document_id (str): Docs ID (required). tableRange (object): Docs unmergeTableCells.tableRange object (required). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_UNMERGE_TABLE_CELLS(
    document_id: Annotated[str, "The Google Docs document ID."],
    tableRange: Annotated[Dict[str, Any], "Docs API tableRange identifying cells to unmerge (must include tableStartLocation)."],
):
    """Unmerge previously merged cells using Docs API unmergeTableCells."""
    err = _validate_required({"document_id": document_id, "tableRange": tableRange}, ["document_id", "tableRange"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req = {"unmergeTableCells": {"tableRange": tableRange}}
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": [req]})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to unmerge table cells: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_UPDATE_DOCUMENT_MARKDOWN",
    description="Update Document Markdown. Replaces the entire content of an existing Google Docs document with new markdown text; requires edit permissions for the document. Args: document_id (str): Docs ID (required). new_markdown_text (str): Markdown text to insert (required). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_UPDATE_DOCUMENT_MARKDOWN(
    document_id: Annotated[str, "The Google Docs document ID."],
    new_markdown_text: Annotated[str, "Markdown text to replace the entire document body."],
):
    """Replace entire body content with the provided markdown text as plain text."""
    err = _validate_required({"document_id": document_id, "new_markdown_text": new_markdown_text}, ["document_id", "new_markdown_text"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        doc = docs_request("get", document_id=document_id)
        body_content = doc.get("body", {}).get("content", [])
        doc_len = body_content[-1].get("endIndex", 1) if body_content else 1

        requests = [
            {
                "deleteContentRange": {
                    "range": {"startIndex": 1, "endIndex": max(1, doc_len - 1)}
                }
            },
            {"insertText": {"location": {"index": 1}, "text": new_markdown_text}},
        ]
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": requests})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to update document with markdown: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_UPDATE_DOCUMENT_STYLE",
    description="Update Document Style. Tool to update the overall document style, such as page size, margins, and default text direction. Use when you need to modify the global style settings of a Google document. Args: document_id (str): Docs ID (required). document_style (object): Docs DocumentStyle object (required). fields (str): Fields mask for properties to update (required). tab_id (str): Optional tabId (optional). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_UPDATE_DOCUMENT_STYLE(
    document_id: Annotated[str, "The Google Docs document ID."],
    document_style: Annotated[Dict[str, Any], "DocumentStyle to apply (e.g., pageSize, margins)."],
    fields: Annotated[str, "Comma-separated fields mask indicating which properties to update."],
    tab_id: Annotated[Optional[str], "Optional tabId for multi-tab documents."] = None,
):
    """Update document-level style via updateDocumentStyle."""
    err = _validate_required({"document_id": document_id, "document_style": document_style, "fields": fields}, ["document_id", "document_style", "fields"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req: Dict[str, Any] = {
            "updateDocumentStyle": {
                "documentStyle": document_style,
                "fields": fields,
            }
        }
        if tab_id:
            req["updateDocumentStyle"]["tabId"] = tab_id
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": [req]})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to update document style: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_UPDATE_EXISTING_DOCUMENT",
    description="Update existing document. Applies programmatic edits, such as text insertion, deletion, or formatting, to a specified Google Doc using the `batchupdate` API method. Args: document_id (str): Docs ID (required). editDocs (array): Array of raw Docs API request objects (required). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_UPDATE_EXISTING_DOCUMENT(
    document_id: Annotated[str, "The Google Docs document ID."],
    editDocs: Annotated[List[Dict[str, Any]], "Array of Docs API request objects to send to batchUpdate."],
):
    """Pass-through for arbitrary batchUpdate requests."""
    err = _validate_required({"document_id": document_id, "editDocs": editDocs}, ["document_id", "editDocs"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": editDocs})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to update existing document: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_UPDATE_TABLE_ROW_STYLE",
    description="Update Table Row Style. Tool to update the style of a table row in a Google document. Use when you need to modify the appearance of specific rows within a table, such as setting minimum row height or marking rows as headers. Args: documentId (str): Docs ID (required). updateTableRowStyle (object): Docs updateTableRowStyle request body (required). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_UPDATE_TABLE_ROW_STYLE(
    documentId: Annotated[str, "The Google Docs document ID."],
    updateTableRowStyle: Annotated[Dict[str, Any], "Docs API updateTableRowStyle request object. Accepts either the modern shape {tableStartLocation,rowIndices,tableRowStyle,fields} or a legacy shape using tableRange that will be translated."],
):
    """Update a table row style using Docs API updateTableRowStyle."""
    err = _validate_required({"documentId": documentId, "updateTableRowStyle": updateTableRowStyle}, ["documentId", "updateTableRowStyle"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        # Prefer passing through modern shape directly if provided
        req_payload: Dict[str, Any] = {}
        fields = updateTableRowStyle.get("fields", "")

        # Modern shape
        if "tableStartLocation" in updateTableRowStyle or "rowIndices" in updateTableRowStyle:
            if "tableStartLocation" in updateTableRowStyle:
                req_payload["tableStartLocation"] = updateTableRowStyle["tableStartLocation"]
            if "rowIndices" in updateTableRowStyle:
                req_payload["rowIndices"] = list(updateTableRowStyle.get("rowIndices", []))
            req_payload["tableRowStyle"] = dict(updateTableRowStyle.get("tableRowStyle", {}))
            req_payload["fields"] = fields
        else:
            # Legacy shape: { tableRange, tableRowStyle, fields }
            tableRange = updateTableRowStyle.get("tableRange")
            tableRowStyle = updateTableRowStyle.get("tableRowStyle", {})

            # If legacy provided, try to derive rowIndices from tableRange when possible
            if tableRange and isinstance(tableRange, dict):
                table_cell_loc = tableRange.get("tableCellLocation", {})
                start_loc = table_cell_loc.get("tableStartLocation")
                start_row = table_cell_loc.get("rowIndex")
                row_span = tableRange.get("rowSpan")
                if start_loc is not None and start_row is not None and row_span:
                    req_payload["tableStartLocation"] = start_loc
                    req_payload["rowIndices"] = list(range(int(start_row), int(start_row) + int(row_span)))
                # If we cannot derive, fall back to API expecting tableRowStyle only (may error)
            req_payload["tableRowStyle"] = dict(tableRowStyle)
            req_payload["fields"] = fields

        req = {"updateTableRowStyle": req_payload}
        result = docs_request("batchUpdate", document_id=documentId, body={"requests": [req]})
        return {"data": {"documentId": documentId, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to update table row style: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_INSERT_TEXT_ACTION",
    description="Insert Text into Document. Tool to insert a string of text at a specified location within a Google document. Use when you need to add new text content to an existing document. Args: document_id (str): Docs ID (required). insertion_index (int): Index where to insert text (required). text_to_insert (str): Text to insert (required). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_INSERT_TEXT_ACTION(
    document_id: Annotated[str, "The Google Docs document ID."],
    insertion_index: Annotated[int, "The index where to insert the text (0-based)."],
    text_to_insert: Annotated[str, "The text to insert into the document."],
):
    """Insert text at a specified location in a Google Docs document."""
    err = _validate_required({"document_id": document_id, "insertion_index": insertion_index, "text_to_insert": text_to_insert}, ["document_id", "insertion_index", "text_to_insert"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        # Get document length to validate index
        doc = docs_request("get", document_id=document_id)
        body_content = doc.get("body", {}).get("content", [])
        doc_length = body_content[-1].get("endIndex", 1) if body_content else 1
        
        # Clamp index to valid range
        clamped_index = max(1, min(insertion_index, doc_length - 1))
        
        req = {
            "insertText": {
                "location": {"index": clamped_index},
                "text": text_to_insert
            }
        }
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": [req]})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to insert text: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_DELETE_CONTENT_RANGE",
    description="Delete Content Range in Document. Tool to delete a range of content from a Google document. Use when you need to remove a specific portion of text or other structural elements within a document. Args: document_id (str): Docs ID (required). range (object): Range object with startIndex and endIndex (required). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_DELETE_CONTENT_RANGE(
    document_id: Annotated[str, "The Google Docs document ID."],
    range: Annotated[Dict[str, Any], "Range object with startIndex and endIndex to delete."],
):
    """Delete a range of content from a Google Docs document."""
    err = _validate_required({"document_id": document_id, "range": range}, ["document_id", "range"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req = {"deleteContentRange": {"range": range}}
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": [req]})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to delete content range: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_DELETE_FOOTER",
    description="Delete Footer. Tool to delete a footer from a Google document. Use when you need to remove a footer from a specific section or the default footer. Args: document_id (str): Docs ID (required). footer_id (str): Footer ID to delete (required). tab_id (str): Optional tab ID (optional). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_DELETE_FOOTER(
    document_id: Annotated[str, "The Google Docs document ID."],
    footer_id: Annotated[str, "The footer ID to delete."],
    tab_id: Annotated[Optional[str], "Optional tab ID for multi-tab documents."] = None,
):
    """Delete a footer from a Google Docs document."""
    err = _validate_required({"document_id": document_id, "footer_id": footer_id}, ["document_id", "footer_id"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req_body = {"deleteFooter": {"footerId": footer_id}}
        if tab_id:
            req_body["deleteFooter"]["tabId"] = tab_id
        
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": [req_body]})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to delete footer: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_DELETE_HEADER",
    description="Delete Header. Deletes the header from the specified section or the default header if no section is specified. Use this tool to remove a header from a Google document. Args: document_id (str): Docs ID (required). header_id (str): Header ID to delete (required). tab_id (str): Optional tab ID (optional). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_DELETE_HEADER(
    document_id: Annotated[str, "The Google Docs document ID."],
    header_id: Annotated[str, "The header ID to delete."],
    tab_id: Annotated[Optional[str], "Optional tab ID for multi-tab documents."] = None,
):
    """Delete a header from a Google Docs document."""
    err = _validate_required({"document_id": document_id, "header_id": header_id}, ["document_id", "header_id"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req_body = {"deleteHeader": {"headerId": header_id}}
        if tab_id:
            req_body["deleteHeader"]["tabId"] = tab_id
        
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": [req_body]})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to delete header: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_DELETE_NAMED_RANGE",
    description="Delete Named Range. Tool to delete a named range from a Google document. Use when you need to remove a previously defined named range by its id or name. Args: document_id (str): Docs ID (required). deleteNamedRange (object): Delete named range request object (required). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_DELETE_NAMED_RANGE(
    document_id: Annotated[str, "The Google Docs document ID."],
    deleteNamedRange: Annotated[Dict[str, Any], "Delete named range request object with namedRangeId or name."],
):
    """Delete a named range from a Google Docs document."""
    err = _validate_required({"document_id": document_id, "deleteNamedRange": deleteNamedRange}, ["document_id", "deleteNamedRange"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req = {"deleteNamedRange": deleteNamedRange}
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": [req]})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to delete named range: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_DELETE_PARAGRAPH_BULLETS",
    description="Delete Paragraph Bullets. Tool to remove bullets from paragraphs within a specified range in a Google document. Use when you need to clear bullet formatting from a section of a document. Args: document_id (str): Docs ID (required). range (object): Range object with startIndex and endIndex (required). tab_id (str): Optional tab ID (optional). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_DELETE_PARAGRAPH_BULLETS(
    document_id: Annotated[str, "The Google Docs document ID."],
    range: Annotated[Dict[str, Any], "Range object with startIndex and endIndex to remove bullets from."],
    tab_id: Annotated[Optional[str], "Optional tab ID for multi-tab documents."] = None,
):
    """Delete paragraph bullets from a specified range in a Google Docs document."""
    err = _validate_required({"document_id": document_id, "range": range}, ["document_id", "range"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req_body = {"deleteParagraphBullets": {"range": range}}
        if tab_id:
            req_body["deleteParagraphBullets"]["tabId"] = tab_id
        
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": [req_body]})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to delete paragraph bullets: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_DELETE_TABLE",
    description="Delete Table. Tool to delete an entire table from a Google document. Use when you have the document id and the specific start and end index of the table element to be removed. The table's range can be found by inspecting the document's content structure. Args: document_id (str): Docs ID (required). table_start_index (int): Start index of table (required). table_end_index (int): End index of table (required). segment_id (str): Optional segment ID (optional). tab_id (str): Optional tab ID (optional). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_DELETE_TABLE(
    document_id: Annotated[str, "The Google Docs document ID."],
    table_start_index: Annotated[int, "The start index of the table to delete."],
    table_end_index: Annotated[int, "The end index of the table to delete."],
    segment_id: Annotated[Optional[str], "Optional segment ID for multi-segment documents."] = None,
    tab_id: Annotated[Optional[str], "Optional tab ID for multi-tab documents."] = None,
):
    """Delete an entire table from a Google Docs document."""
    err = _validate_required({"document_id": document_id, "table_start_index": table_start_index, "table_end_index": table_end_index}, ["document_id", "table_start_index", "table_end_index"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        # Use deleteContentRange to delete the entire table
        req_body = {
            "deleteContentRange": {
                "range": {
                    "startIndex": table_start_index,
                    "endIndex": table_end_index
                }
            }
        }
        if segment_id:
            req_body["deleteContentRange"]["range"]["segmentId"] = segment_id
        if tab_id:
            req_body["deleteContentRange"]["tabId"] = tab_id
        
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": [req_body]})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to delete table: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_DELETE_TABLE_COLUMN",
    description="Delete Table Column. Tool to delete a column from a table in a Google document. Use this tool when you need to remove a specific column from an existing table within a document. Args: document_id (str): Docs ID (required). requests (array): Array of deleteTableColumn request objects (required). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_DELETE_TABLE_COLUMN(
    document_id: Annotated[str, "The Google Docs document ID."],
    requests: Annotated[List[Dict[str, Any]], "Array of deleteTableColumn request objects."],
):
    """Delete columns from a table in a Google Docs document."""
    err = _validate_required({"document_id": document_id, "requests": requests}, ["document_id", "requests"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        result = docs_request("batchUpdate", document_id=document_id, body={"requests": requests})
        return {"data": {"documentId": document_id, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to delete table column: {str(e)}", "successful": False}


@mcp.tool(
    "GOOGLEDOCS_DELETE_TABLE_ROW",
    description="Delete Table Row. Tool to delete a row from a table in a Google document. Use when you need to remove a specific row from an existing table. Args: documentId (str): Docs ID (required). tableCellLocation (object): Table cell location object (required). Returns: dict: { data: {documentId,replies}, error: str, successful: bool }.",
)
def GOOGLEDOCS_DELETE_TABLE_ROW(
    documentId: Annotated[str, "The Google Docs document ID."],
    tableCellLocation: Annotated[Dict[str, Any], "Table cell location object specifying which row to delete."],
):
    """Delete a row from a table in a Google Docs document."""
    err = _validate_required({"documentId": documentId, "tableCellLocation": tableCellLocation}, ["documentId", "tableCellLocation"])
    if err:
        return {"data": {}, "error": str(err), "successful": False}

    try:
        req = {"deleteTableRow": {"tableCellLocation": tableCellLocation}}
        result = docs_request("batchUpdate", document_id=documentId, body={"requests": [req]})
        return {"data": {"documentId": documentId, "replies": result.get("replies", [])}, "error": "", "successful": True}
    except Exception as e:
        return {"data": {}, "error": f"Failed to delete table row: {str(e)}", "successful": False}


# -------------------- MAIN --------------------

if __name__ == "__main__":
    mcp.run()
