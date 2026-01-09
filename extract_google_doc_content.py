#!/usr/bin/env python3
"""
Extract Google Doc content from links in column R and write to column S
从 R 列的 Google Doc 链接提取内容并写入 S 列

Requirements:
1. Install Google API libraries:
   pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib

2. Set up Google API credentials:
   - Go to https://console.cloud.google.com/
   - Create a new project or select existing one
   - Enable Google Sheets API and Google Docs API
   - Create credentials (OAuth 2.0 Client ID)
   - Download credentials.json and place it in the project directory

3. The first time you run this script, it will open a browser for authentication
"""

import os
import sys
import re
from typing import Optional

try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    import pickle
except ImportError:
    print("[ERROR] Missing Google API packages. Please install:", file=sys.stderr)
    print("  pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib", file=sys.stderr)
    sys.exit(1)


# Google API scopes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/documents.readonly'
]

# Spreadsheet ID from the URL
# URL format: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit
SPREADSHEET_ID = '13Ot6VJVZbEp6Hv27uDQnuMsFHsAOsgT1xZWqxAdG6KI'

# Column letters (R=18th column, S=19th column)
DOC_LINK_COL = "R"  # Google Doc links in column R
OUTPUT_COL = "S"    # Output content in column S

# Batch write size: write every N documents to balance speed and real-time progress
BATCH_SIZE = 5  # Write every 5 documents (set to 1 for immediate write, larger for faster)

# Token file to store credentials
TOKEN_FILE = 'google_api_token.pickle'
CREDENTIALS_FILE = 'credentials.json'


def get_credentials():
    """Get valid user credentials from storage or create new ones."""
    creds = None
    
    # Load existing token
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    
    # If no valid credentials, authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"[ERROR] {CREDENTIALS_FILE} not found!", file=sys.stderr)
                print(f"Please download credentials.json from Google Cloud Console", file=sys.stderr)
                print(f"Enable Google Sheets API and Google Docs API", file=sys.stderr)
                sys.exit(1)
            
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for next run
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds


def extract_doc_id(url: str) -> Optional[str]:
    """
    Extract Google Doc ID from URL.
    Supports: https://docs.google.com/document/d/<DOC_ID>/edit
    """
    if not url or not isinstance(url, str):
        return None
    
    url = url.strip()
    
    # Simple pattern matching
    m = re.search(r"/document/d/([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    
    # Also try drive.google.com format
    m = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    
    # If it looks like just an ID (alphanumeric, 25-60 chars)
    if re.match(r'^[a-zA-Z0-9_-]{25,60}$', url):
        return url
    
    return None


def read_doc_text(docs_service, doc_id: str) -> Optional[str]:
    """Get plain text content from a Google Doc."""
    try:
        doc = docs_service.documents().get(documentId=doc_id).execute()
        content = doc.get("body", {}).get("content", [])
        
        text = []
        for element in content:
            if "paragraph" in element:
                for run in element["paragraph"].get("elements", []):
                    if "textRun" in run:
                        text.append(run["textRun"].get("content", ""))
            elif "table" in element:
                # Extract text from tables
                table = element["table"]
                if "tableRows" in table:
                    for row in table["tableRows"]:
                        if "tableCells" in row:
                            row_texts = []
                            for cell in row["tableCells"]:
                                cell_text = []
                                if "content" in cell:
                                    for cell_elem in cell["content"]:
                                        if "paragraph" in cell_elem:
                                            for para_elem in cell_elem["paragraph"].get("elements", []):
                                                if "textRun" in para_elem:
                                                    cell_text.append(para_elem["textRun"].get("content", ""))
                                row_texts.append(" ".join(cell_text).strip())
                            if row_texts:
                                text.append("\t".join(row_texts))
                                text.append("\n")
        
        return "".join(text).strip()
    except HttpError as e:
        print(f"[ERROR] Failed to get document {doc_id}: {e}", file=sys.stderr)
        return None
    except Exception as e:
        print(f"[ERROR] Unexpected error getting document {doc_id}: {e}", file=sys.stderr)
        return None


def get_hyperlinks_from_column(sheets_service, spreadsheet_id: str, sheet_name: str, column: str, start_row: int = 2):
    """
    Get hyperlinks from a column in Google Sheets.
    Returns a list of URLs (or None if cell has no hyperlink).
    """
    # Get sheet ID first
    spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = None
    for sheet in spreadsheet['sheets']:
        if sheet['properties']['title'] == sheet_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    if sheet_id is None:
        raise ValueError(f"Sheet '{sheet_name}' not found")
    
    # Read a large range to get all data with hyperlinks
    # We'll use get with includeGridData to get hyperlink information
    range_str = f"{sheet_name}!{column}{start_row}:{column}"
    try:
        # First get values to know how many rows
        values_result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_str
        ).execute()
        num_rows = len(values_result.get("values", []))
        
        if num_rows == 0:
            return []
        
        # Now get the grid data with hyperlinks
        result = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[f"{sheet_name}!{column}{start_row}:{column}{start_row + num_rows - 1}"],
            includeGridData=True
        ).execute()
        
        hyperlinks = []
        rows_data = result['sheets'][0]['data'][0].get('rowData', [])
        
        for row_data in rows_data:
            if not row_data.get('values') or len(row_data['values']) == 0:
                hyperlinks.append(None)
                continue
            
            cell = row_data['values'][0]
            # Try to get hyperlink from different possible locations
            hyperlink = None
            
            # Check userEnteredFormat.textFormat.link
            if 'userEnteredFormat' in cell:
                format_info = cell['userEnteredFormat']
                if 'textFormat' in format_info and 'link' in format_info['textFormat']:
                    hyperlink = format_info['textFormat']['link'].get('uri')
            
            # Check hyperlink property directly
            if not hyperlink and 'hyperlink' in cell:
                hyperlink = cell['hyperlink']
            
            # Check hyperlinkDisplayType (this means cell has hyperlink)
            if not hyperlink and 'userEnteredFormat' in cell:
                if 'textFormat' in cell['userEnteredFormat']:
                    text_format = cell['userEnteredFormat']['textFormat']
                    # If hyperlinkDisplayType exists, try to reconstruct from formattedValue
                    if 'link' in text_format:
                        hyperlink = text_format['link'].get('uri')
            
            # Fallback: if cell value looks like a URL, use it
            if not hyperlink and 'formattedValue' in cell:
                value = cell['formattedValue']
                if value and ('http://' in value or 'https://' in value):
                    hyperlink = value
            
            hyperlinks.append(hyperlink)
        
        return hyperlinks
    
    except HttpError as e:
        print(f"[ERROR] Failed to read hyperlinks from column {column}: {e}", file=sys.stderr)
        raise


def main():
    """Main function to extract Google Doc content and update spreadsheet."""
    print("[INFO] Authenticating with Google API...")
    creds = get_credentials()
    
    # Build services
    print("[INFO] Building Google API services...")
    sheets_service = build('sheets', 'v4', credentials=creds)
    docs_service = build('docs', 'v1', credentials=creds)
    
    # Get sheet name (use first sheet by default)
    try:
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheet_name = spreadsheet['sheets'][0]['properties']['title']
        print(f"[INFO] Working with sheet: {sheet_name}")
    except Exception as e:
        print(f"[ERROR] Failed to get sheet name: {e}", file=sys.stderr)
        sys.exit(1)
    
    # Read all hyperlinks from column R (starting from row 2 to skip header)
    print(f"[INFO] Reading hyperlinks from column {DOC_LINK_COL}...")
    try:
        hyperlinks = get_hyperlinks_from_column(sheets_service, SPREADSHEET_ID, sheet_name, DOC_LINK_COL, start_row=2)
    except Exception as e:
        print(f"[ERROR] Failed to read hyperlinks: {e}", file=sys.stderr)
        sys.exit(1)
    
    print(f"[INFO] Found {len(hyperlinks)} rows in column {DOC_LINK_COL}")
    
    # Process each hyperlink and extract content, write in batches
    processed_count = 0
    updates = []  # Batch updates: [(row_index, value), ...]
    
    def flush_updates(updates_list):
        """Write accumulated updates to spreadsheet."""
        if not updates_list:
            return
        
        # Prepare batch update
        batch_data = []
        for row_idx, value in updates_list:
            batch_data.append({
                'range': f"{sheet_name}!{OUTPUT_COL}{row_idx}",
                'values': [[value]]
            })
        
        try:
            body = {
                'valueInputOption': 'RAW',
                'data': batch_data
            }
            sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=body
            ).execute()
            print(f"[INFO] Wrote batch of {len(updates_list)} rows to column {OUTPUT_COL}")
        except HttpError as e:
            print(f"[ERROR] Failed to write batch: {e}", file=sys.stderr)
            # Try individual writes as fallback
            for row_idx, value in updates_list:
                try:
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=SPREADSHEET_ID,
                        range=f"{sheet_name}!{OUTPUT_COL}{row_idx}",
                        valueInputOption="RAW",
                        body={"values": [[value]]}
                    ).execute()
                except HttpError as err:
                    print(f"[ERROR] Row {row_idx}: Failed to write: {err}", file=sys.stderr)
    
    for i, hyperlink in enumerate(hyperlinks, start=2):
        value_to_write = ""
        
        if hyperlink is None:
            # No hyperlink, write empty string
            value_to_write = ""
        else:
            link = hyperlink.strip() if hyperlink else ""
            if not link:
                value_to_write = ""
            else:
                print(f"[INFO] Row {i}: Processing URL: {link[:80]}...")
                
                # Extract document ID
                doc_id = extract_doc_id(link)
                if not doc_id:
                    print(f"[WARN] Row {i}: Could not extract document ID from URL")
                    print(f"[DEBUG] Full URL: {repr(link)}")
                    value_to_write = "INVALID LINK"
                else:
                    # Get document content
                    try:
                        text = read_doc_text(docs_service, doc_id)
                        if text:
                            print(f"[INFO] Row {i}: Extracted {len(text)} characters")
                            value_to_write = text
                            processed_count += 1
                        else:
                            print(f"[WARN] Row {i}: Failed to extract content")
                            value_to_write = "ERROR: Failed to extract content"
                    except Exception as e:
                        error_msg = f"ERROR: {str(e)}"
                        print(f"[ERROR] Row {i}: {error_msg}")
                        value_to_write = error_msg
        
        # Add to batch
        updates.append((i, value_to_write))
        
        # Flush batch when it reaches BATCH_SIZE
        if len(updates) >= BATCH_SIZE:
            flush_updates(updates)
            updates = []
    
    # Flush remaining updates
    if updates:
        flush_updates(updates)
    
    print(f"\n[INFO] Processed {processed_count} documents successfully")
    print("[INFO] Done!")


if __name__ == '__main__':
    main()
