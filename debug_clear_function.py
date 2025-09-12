#!/usr/bin/env python3
"""
Debug the clearRosterData function to see what's happening
"""

import json
from google.oauth2 import service_account
import gspread

SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def debug_clear_function():
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    client = gspread.authorize(credentials)
    
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    roster_sheet = spreadsheet.worksheet('Roster')
    
    print("=== DEBUGGING CLEAR ROSTER DATA LOGIC ===")
    
    # Get all columns to check for Manual/Formula sources
    all_data = roster_sheet.get('1:3')  # Get header, type, and SOURCE rows
    headers = all_data[0] if len(all_data) > 0 else []
    type_row_data = all_data[1] if len(all_data) > 1 else []
    source_row_data = all_data[2] if len(all_data) > 2 else []
    
    print(f"Total columns found: {len(source_row_data)}")
    print(f"Source row data (first 30 columns): {source_row_data[:30]}")
    
    # Check which columns should be preserved (based on SOURCE, not type)
    preserved_columns = []
    cleared_columns = []
    
    for col_index, column_source in enumerate(source_row_data):
        col_num = col_index + 1  # 1-based column numbers
        
        if column_source and (str(column_source).lower() == 'manual' or str(column_source).lower() == 'formula'):
            preserved_columns.append(f"Col {col_num}: '{column_source}'")
        else:
            cleared_columns.append(f"Col {col_num}: '{column_source}'")
    
    print(f"\nColumns that SHOULD BE PRESERVED (source = Manual/Formula):")
    if preserved_columns:
        for col in preserved_columns:
            print(f"  {col}")
    else:
        print("  None found")
    
    print(f"\nColumns that SHOULD BE CLEARED (first 15):")
    for col in cleared_columns[:15]:
        print(f"  {col}")
    
    print(f"\n=== FULL COLUMN ANALYSIS (first 30) ===")
    for i in range(min(30, len(headers), len(source_row_data))):
        header = headers[i] if i < len(headers) else ""
        col_type = type_row_data[i] if i < len(type_row_data) else ""
        col_source = source_row_data[i] if i < len(source_row_data) else ""
        col_num = i + 1
        should_preserve = col_source and (str(col_source).lower() == 'manual' or str(col_source).lower() == 'formula')
        print(f"Col {col_num:2}: '{str(header)[:25]:25}' | Type: '{str(col_type)[:12]:12}' | Source: '{str(col_source)[:15]:15}' | Preserve: {should_preserve}")
    
    # Check if there are any Manual or Formula columns in the entire sheet
    manual_formula_count = sum(1 for s in source_row_data if s and (str(s).lower() == 'manual' or str(s).lower() == 'formula'))
    print(f"\nTotal Manual/Formula SOURCE columns found: {manual_formula_count}")
    
    if manual_formula_count == 0:
        print("⚠️  WARNING: No Manual or Formula SOURCE columns found. The clearRosterData function may not be working as expected.")

if __name__ == "__main__":
    debug_clear_function()