#!/usr/bin/env python3
"""
Script to examine the Madison Ultimate roster Google Sheet
Uses service account credentials to read and analyze the data
"""

import json
import pandas as pd
from google.oauth2 import service_account
import gspread
from pprint import pprint

# Configuration
SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def authenticate_google_sheets():
    """Authenticate with Google Sheets API using service account"""
    try:
        # Load service account credentials
        with open(CREDENTIALS_FILE, 'r') as f:
            creds_info = json.load(f)
        
        credentials = service_account.Credentials.from_service_account_info(
            creds_info,
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets.readonly',
                'https://www.googleapis.com/auth/drive.readonly'
            ]
        )
        
        # Create gspread client
        client = gspread.authorize(credentials)
        return client
    except Exception as e:
        print(f"Authentication failed: {e}")
        return None

def examine_roster(client):
    """Read and analyze the roster sheet"""
    try:
        # Open the spreadsheet
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        print(f"Successfully opened spreadsheet: {spreadsheet.title}")
        
        # List all worksheets
        worksheets = spreadsheet.worksheets()
        print(f"\nAvailable worksheets:")
        for i, ws in enumerate(worksheets):
            print(f"  {i+1}. {ws.title} ({ws.row_count} rows, {ws.col_count} cols)")
        
        # Focus on the Roster sheet
        try:
            roster_sheet = spreadsheet.worksheet('Roster')
        except gspread.WorksheetNotFound:
            print("\n'Roster' sheet not found. Available sheets:")
            for ws in worksheets:
                print(f"  - {ws.title}")
            return None
        
        print(f"\n=== ROSTER SHEET ANALYSIS ===")
        print(f"Dimensions: {roster_sheet.row_count} rows x {roster_sheet.col_count} columns")
        
        # Read metadata rows (1-5)
        print(f"\n=== METADATA ROWS ===")
        metadata = roster_sheet.get('A1:Z5')  # First 26 columns
        
        if len(metadata) >= 5:
            headers = metadata[0] if len(metadata) > 0 else []
            types = metadata[1] if len(metadata) > 1 else []
            sources = metadata[2] if len(metadata) > 2 else []
            notes = metadata[3] if len(metadata) > 3 else []
            repeat_headers = metadata[4] if len(metadata) > 4 else []
            
            print(f"Row 1 (Headers): {len([h for h in headers if h])} non-empty columns")
            print(f"Row 2 (Types): {len([t for t in types if t])} non-empty columns")
            print(f"Row 3 (Sources): {len([s for s in sources if s])} non-empty columns")
            print(f"Row 4 (Notes): {len([n for n in notes if n])} non-empty columns")
            print(f"Row 5 (Repeat Headers): {len([h for h in repeat_headers if h])} non-empty columns")
            
            # Show first few columns in detail
            print(f"\n=== FIRST 10 COLUMNS DETAIL ===")
            for i in range(min(10, len(headers))):
                if headers[i]:
                    print(f"Column {i+1}: {headers[i]}")
                    print(f"  Type: {types[i] if i < len(types) else 'N/A'}")
                    print(f"  Source: {sources[i] if i < len(sources) else 'N/A'}")
                    print(f"  Note: {notes[i][:100] if i < len(notes) and notes[i] else 'N/A'}...")
                    print()
        
        # Count actual data rows (from row 6 onwards)
        print(f"\n=== DATA ANALYSIS ===")
        
        # Get first name column (should be column A)
        first_names = roster_sheet.col_values(1)[5:]  # Skip metadata rows
        non_empty_names = [name for name in first_names if name and name.strip()]
        
        print(f"Total data rows with first names: {len(non_empty_names)}")
        
        if non_empty_names:
            print(f"Sample first names: {non_empty_names[:5]}")
        
        # Check for formulas in first data row
        print(f"\n=== FORMULA CHECK ===")
        try:
            first_data_row = roster_sheet.get('A6:J6', value_render_option='FORMULA')
            if first_data_row:
                row = first_data_row[0]
                for i, cell in enumerate(row):
                    if cell and str(cell).startswith('='):
                        print(f"Column {i+1}: {cell}")
        except Exception as e:
            print(f"Could not check formulas: {e}")
        
        return roster_sheet
        
    except Exception as e:
        print(f"Error examining roster: {e}")
        return None

def examine_excel_file(excel_path):
    """Read and analyze the Excel file"""
    try:
        print(f"\n=== EXCEL FILE ANALYSIS ===")
        print(f"Reading: {excel_path}")
        
        # Read all sheets
        excel_file = pd.ExcelFile(excel_path)
        print(f"Available sheets: {excel_file.sheet_names}")
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            print(f"\nSheet '{sheet_name}': {df.shape[0]} rows x {df.shape[1]} columns")
            
            if df.shape[0] > 0:
                print(f"Columns: {list(df.columns)}")
                
                # Show first few rows
                print("\nFirst 3 rows:")
                print(df.head(3).to_string())
        
        return True
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return False

def main():
    print("Madison Ultimate Roster Examination Tool")
    print("=" * 50)
    
    # First, try to examine the Excel file
    excel_path = "/Users/sal/Downloads/2025 Fall Coach Sheets (3).xlsx"
    print(f"\n1. EXAMINING EXCEL FILE")
    examine_excel_file(excel_path)
    
    # Then examine the Google Sheets
    print(f"\n2. EXAMINING GOOGLE SHEETS")
    
    # Authenticate
    client = authenticate_google_sheets()
    if not client:
        print("Failed to authenticate. Check your credentials file.")
        return
    
    print("✅ Authentication successful")
    
    # Examine the roster
    roster_sheet = examine_roster(client)
    
    if roster_sheet:
        print("\n✅ Roster examination complete")
    else:
        print("\n❌ Failed to examine roster")

if __name__ == "__main__":
    main()