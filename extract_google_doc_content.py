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
from typing import List, Optional, Tuple

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
SPREADSHEET_ID = '1u1W9nV26a8-nvEx8R_bN6tEm3PrK6Z3O6A2YYXlPSLw'

# Column indices (A=0, B=1, ..., R=17, S=18)
R_COLUMN = 17  # Column R (0-indexed)
S_COLUMN = 18  # Column S (0-indexed)

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


def extract_doc_id_from_url(url: str) -> Optional[str]:
    """
    Extract Google Doc ID from various URL formats.
    Examples:
    - https://docs.google.com/document/d/DOC_ID/edit
    - https://docs.google.com/document/d/DOC_ID/view
    - https://docs.google.com/document/d/DOC_ID
    """
    if not url or not isinstance(url, str):
        return None
    
    # Pattern to match Google Doc URLs
    patterns = [
        r'docs\.google\.com/document/d/([a-zA-Z0-9_-]+)',
        r'drive\.google\.com/file/d/([a-zA-Z0-9_-]+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    
    return None


def get_doc_content(docs_service, doc_id: str) -> Optional[str]:
    """Get the plain text content from a Google Doc."""
    try:
        doc = docs_service.documents().get(documentId=doc_id).execute()
        
        # Extract text from document elements
        content_parts = []
        last_element_type = None
        
        def extract_text_from_paragraph(para):
            """Extract text from a paragraph element."""
            if 'elements' not in para:
                return
            
            for elem in para['elements']:
                if 'textRun' in elem:
                    text = elem['textRun'].get('content', '')
                    content_parts.append(text)
        
        def extract_text_from_table(table):
            """Extract text from a table element."""
            if 'tableRows' not in table:
                return
            
            for row in table['tableRows']:
                if 'tableCells' not in row:
                    continue
                
                row_texts = []
                for cell in row['tableCells']:
                    if 'content' in cell:
                        cell_texts = []
                        for cell_elem in cell['content']:
                            if 'paragraph' in cell_elem:
                                para = cell_elem['paragraph']
                                if 'elements' in para:
                                    for para_elem in para['elements']:
                                        if 'textRun' in para_elem:
                                            cell_texts.append(para_elem['textRun'].get('content', ''))
                        row_texts.append(' '.join(cell_texts).strip())
                
                # Join cell texts with tab, then add newline
                if row_texts:
                    content_parts.append('\t'.join(row_texts))
                    content_parts.append('\n')
        
        # Process document body
        if 'body' in doc and 'content' in doc['body']:
            for element in doc['body']['content']:
                if 'paragraph' in element:
                    # Add newline between paragraphs (but not first one)
                    if last_element_type == 'paragraph':
                        content_parts.append('\n')
                    extract_text_from_paragraph(element['paragraph'])
                    last_element_type = 'paragraph'
                    
                elif 'table' in element:
                    # Add newline before table
                    if content_parts and content_parts[-1] != '\n':
                        content_parts.append('\n')
                    extract_text_from_table(element['table'])
                    last_element_type = 'table'
        
        # Join and clean up content
        content = ''.join(content_parts)
        # Remove excessive newlines (more than 2 consecutive)
        content = re.sub(r'\n{3,}', '\n\n', content)
        return content.strip()
    
    except HttpError as e:
        print(f"[ERROR] Failed to get document {doc_id}: {e}", file=sys.stderr)
        return None
    except Exception as e:
        print(f"[ERROR] Unexpected error getting document {doc_id}: {e}", file=sys.stderr)
        return None


def column_index_to_letter(column_index: int) -> str:
    """Convert 0-based column index to A1 notation letter(s)."""
    result = ""
    column_index += 1  # Convert to 1-based
    while column_index > 0:
        column_index -= 1
        result = chr(65 + (column_index % 26)) + result
        column_index //= 26
    return result


def get_column_data(sheets_service, sheet_name: str, column_index: int) -> List[Tuple[int, str]]:
    """Get all values from a specific column, returning (row_index, value) pairs."""
    try:
        # Get all values in the column using A1 notation
        column_letter = column_index_to_letter(column_index)
        range_name = f"{sheet_name}!{column_letter}:{column_letter}"
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        
        # Return list of (row_index, value) tuples, skipping header row (index 0)
        data = []
        for i, row in enumerate(values):
            if i == 0:  # Skip header row
                continue
            value = row[0] if row else ''
            data.append((i + 1, value))  # +1 because row indices in Sheets API start at 1
        
        return data
    
    except HttpError as e:
        print(f"[ERROR] Failed to read column data: {e}", file=sys.stderr)
        return []


def update_column_data(sheets_service, sheet_name: str, column_index: int, updates: List[Tuple[int, str]]):
    """Update specific cells in a column."""
    if not updates:
        return
    
    try:
        # Prepare update requests
        values = []
        column_letter = column_index_to_letter(column_index)
        for row_index, value in updates:
            # Convert row index to A1 notation (row_index is already 1-based from get_column_data)
            cell_range = f"{sheet_name}!{column_letter}{row_index}"
            values.append({
                'range': cell_range,
                'values': [[value]]  # Wrap in nested list for batch update
            })
        
        # Batch update
        body = {
            'valueInputOption': 'RAW',
            'data': values
        }
        
        result = sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()
        
        updated = result.get('totalUpdatedCells', 0)
        print(f"[INFO] Updated {updated} cells in column {column_letter}")
    
    except HttpError as e:
        print(f"[ERROR] Failed to update column data: {e}", file=sys.stderr)


def main():
    """Main function to extract Google Doc content and update spreadsheet."""
    print("[INFO] Authenticating with Google API...")
    creds = get_credentials()
    
    # Build services
    print("[INFO] Building Google API services...")
    sheets_service = build('sheets', 'v4', credentials=creds)
    docs_service = build('docs', 'v1', credentials=creds)
    
    # Get sheet name (you may need to adjust this)
    # Try to get the first sheet, or you can specify the sheet name
    try:
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheet_name = spreadsheet['sheets'][0]['properties']['title']
        print(f"[INFO] Working with sheet: {sheet_name}")
    except Exception as e:
        print(f"[ERROR] Failed to get sheet name: {e}", file=sys.stderr)
        sys.exit(1)
    
    # Get all URLs from column R
    print(f"[INFO] Reading URLs from column R...")
    r_column_data = get_column_data(sheets_service, sheet_name, R_COLUMN)
    
    print(f"[INFO] Found {len(r_column_data)} rows with data in column R")
    
    # Process each URL and extract content
    updates = []
    for row_index, url in r_column_data:
        if not url or not url.strip():
            print(f"[INFO] Row {row_index + 1}: Empty URL, skipping...")
            continue
        
        print(f"[INFO] Row {row_index + 1}: Processing URL: {url[:80]}...")
        
        # Extract document ID from URL
        doc_id = extract_doc_id_from_url(url)
        if not doc_id:
            print(f"[WARN] Row {row_index + 1}: Could not extract document ID from URL")
            continue
        
        # Get document content
        content = get_doc_content(docs_service, doc_id)
        if content:
            print(f"[INFO] Row {row_index + 1}: Extracted {len(content)} characters")
            updates.append((row_index, content))
        else:
            print(f"[WARN] Row {row_index + 1}: Failed to extract content")
    
    # Update column S with extracted content
    if updates:
        print(f"\n[INFO] Updating column S with {len(updates)} extracted contents...")
        update_column_data(sheets_service, sheet_name, S_COLUMN, updates)
        print("[INFO] Done!")
    else:
        print("[WARN] No content extracted, nothing to update")


if __name__ == '__main__':
    main()
