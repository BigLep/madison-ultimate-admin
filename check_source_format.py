#!/usr/bin/env python3
"""
Check the exact source format in row 3
"""

import json
from google.oauth2 import service_account
import gspread

SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def check_sources():
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    client = gspread.authorize(credentials)
    
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    roster_sheet = spreadsheet.worksheet('Roster')
    
    headers = roster_sheet.row_values(1)
    sources = roster_sheet.row_values(3)
    
    print("=== SOURCE ROW ANALYSIS ===\n")
    print("Columns that should be preserved (Manual/Formula/Blank):")
    print("-" * 60)
    
    preserved = []
    cleared = []
    
    for i, (header, source) in enumerate(zip(headers, sources)):
        if not header:
            continue
        
        col = chr(65 + i) if i < 26 else f"{chr(65 + i//26 - 1)}{chr(65 + i%26)}"
        
        # Check preservation logic
        source_str = source.strip() if source else ''
        
        if source_str == 'Manual' or source_str == 'Formula' or source_str == '':
            preserved.append((col, header, source_str))
        else:
            cleared.append((col, header, source_str))
    
    print("\nPRESERVED columns:")
    for col, header, source in preserved:
        source_display = '<blank>' if not source else source
        print(f"  {col:3} - {header:40} [{source_display}]")
    
    print(f"\nCLEARED columns (will be regenerated):")
    for col, header, source in cleared[:10]:  # Show first 10
        print(f"  {col:3} - {header:40} [{source}]")
    if len(cleared) > 10:
        print(f"  ... and {len(cleared) - 10} more")
    
    print(f"\n=== SUMMARY ===")
    print(f"Total columns: {len(headers)}")
    print(f"Preserved: {len(preserved)}")
    print(f"Cleared: {len(cleared)}")
    
    # Check specific problem columns
    print(f"\n=== PROBLEM CHECK ===")
    if sources[0] == 'Data Source:':
        print("⚠️  Column A has 'Data Source:' as its source - this should be empty or a valid source")
    
    # Count how many have the old format
    old_format_count = 0
    for source in sources:
        if source and ('FinalForms' in source or 'AdditionalInfo' in source or 'MailingList' in source):
            old_format_count += 1
    
    if old_format_count > 0:
        print(f"⚠️  {old_format_count} columns use old format like 'FinalForms First Name' instead of 'Final Forms'")

if __name__ == "__main__":
    check_sources()