#!/usr/bin/env python3
"""
Check the Mailing List structure to find Group status column
"""

import json
from google.oauth2 import service_account
import gspread

SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def check_mailing_list():
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    client = gspread.authorize(credentials)
    
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    mailing_sheet = spreadsheet.worksheet('Mailing List')
    
    print("=== MAILING LIST STRUCTURE ===\n")
    
    # Get first 3 rows and all columns
    data = mailing_sheet.get('1:3')
    
    print("First 3 rows:")
    for i, row in enumerate(data, 1):
        print(f"Row {i}: {row}")
    
    print(f"\n=== HEADERS (Row 2) ===")
    if len(data) > 1:
        headers = data[1]
        for i, header in enumerate(headers):
            if header:
                col_letter = chr(65 + i) if i < 26 else f"{chr(65 + i//26 - 1)}{chr(65 + i%26)}"
                print(f"  {col_letter} - {header}")
    
    print(f"\n=== SAMPLE DATA (Row 3) ===")
    if len(data) > 2:
        sample = data[2]
        for i, value in enumerate(sample):
            if value:
                col_letter = chr(65 + i) if i < 26 else f"{chr(65 + i//26 - 1)}{chr(65 + i%26)}"
                header = headers[i] if i < len(headers) else 'N/A'
                print(f"  {col_letter} ({header}): {value}")
    
    # Look for Group status specifically
    group_status_col = None
    if len(data) > 1:
        for i, header in enumerate(data[1]):
            if header and 'group status' in header.lower():
                group_status_col = chr(65 + i) if i < 26 else f"{chr(65 + i//26 - 1)}{chr(65 + i%26)}"
                print(f"\n✅ Found 'Group status' in column {group_status_col}")
                
                # Show some sample values
                status_data = mailing_sheet.col_values(i + 1)[2:7]  # Skip headers, get first 5
                print(f"Sample values: {[s for s in status_data if s]}")
                break
    
    if not group_status_col:
        print(f"\n⚠️  Could not find 'Group status' column")

if __name__ == "__main__":
    check_mailing_list()