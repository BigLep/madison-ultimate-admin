#!/usr/bin/env python3
"""
Debug why Manual/Formula columns are still being cleared
"""

import json
from google.oauth2 import service_account
import gspread

SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def debug_clear_issue():
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    client = gspread.authorize(credentials)
    
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    roster_sheet = spreadsheet.worksheet('Roster')
    
    print("=== DEBUGGING CLEAR ROSTER DATA ISSUE ===")
    
    # Get the source row (row 3) exactly as the Apps Script would
    source_row_data = roster_sheet.get('3:3')[0] if roster_sheet.get('3:3') else []
    
    print(f"Source row data (row 3): {source_row_data}")
    print(f"Total columns: {len(source_row_data)}")
    
    # Check each column that should be preserved
    preserved_columns = []
    
    for col_index, column_source in enumerate(source_row_data):
        col_num = col_index + 1  # 1-based column numbers
        
        # Exact same logic as Apps Script
        if column_source and (str(column_source).lower() == 'manual' or str(column_source).lower() == 'formula'):
            preserved_columns.append({
                'col': col_num,
                'source': column_source,
                'source_lower': str(column_source).lower()
            })
    
    print(f"\nColumns that SHOULD BE PRESERVED:")
    for col_info in preserved_columns:
        print(f"  Column {col_info['col']}: '{col_info['source']}' (lower: '{col_info['source_lower']}')")
    
    if not preserved_columns:
        print("  ❌ NO COLUMNS FOUND TO PRESERVE!")
        print("  This means the clearRosterData function will clear everything.")
        
        # Let's check what's actually in row 3
        print(f"\n=== RAW ROW 3 ANALYSIS ===")
        for i, value in enumerate(source_row_data[:15]):
            print(f"  Column {i+1:2}: '{value}' | Type: {type(value)} | Lower: '{str(value).lower()}'")
    
    # Check if there might be case sensitivity issues
    manual_matches = [i+1 for i, v in enumerate(source_row_data) if v and 'manual' in str(v).lower()]
    formula_matches = [i+1 for i, v in enumerate(source_row_data) if v and 'formula' in str(v).lower()]
    
    print(f"\nColumns containing 'manual' (case insensitive): {manual_matches}")
    print(f"Columns containing 'formula' (case insensitive): {formula_matches}")
    
    # Let's also check if the Apps Script might be looking at the wrong row
    print(f"\n=== CHECKING OTHER ROWS ===")
    row_1_data = roster_sheet.get('1:1')[0] if roster_sheet.get('1:1') else []
    row_2_data = roster_sheet.get('2:2')[0] if roster_sheet.get('2:2') else []
    row_4_data = roster_sheet.get('4:4')[0] if roster_sheet.get('4:4') else []
    
    print(f"Row 1 (Headers): {row_1_data[:10]}")  
    print(f"Row 2 (Types): {row_2_data[:10]}")
    print(f"Row 3 (Sources): {source_row_data[:10]}")
    print(f"Row 4 (Notes): {row_4_data[:10]}")

if __name__ == "__main__":
    debug_clear_issue()