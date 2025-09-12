#!/usr/bin/env python3
"""
Quick debug script to check mailing list structure
"""

import json
from google.oauth2 import service_account
import gspread

# Configuration
SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def debug_mailing_list():
    # Authenticate
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    client = gspread.authorize(credentials)
    
    # Open spreadsheet
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    mailing_sheet = spreadsheet.worksheet('Mailing List')
    roster_sheet = spreadsheet.worksheet('Roster')
    
    print("=== MAILING LIST DEBUG ===")
    
    # Check first 5 rows
    data = mailing_sheet.get('A1:F5')
    print("First 5 rows, columns A-F:")
    for i, row in enumerate(data, 1):
        print(f"Row {i}: {row}")
    
    print(f"\n=== TEST EMAIL FROM ROSTER ===")
    # Get a sample email from roster row 15
    sample_email = roster_sheet.get('E15')[0][0] if roster_sheet.get('E15') else ""
    print(f"Row 15, Column E (Student Personal Email): '{sample_email}'")
    
    if sample_email:
        # Check if this email exists in mailing list
        email_column = mailing_sheet.col_values(1)  # Column A
        permissions_column = mailing_sheet.col_values(6)  # Column F
        
        print(f"\nLooking for '{sample_email}' in mailing list...")
        print(f"First 5 emails in mailing list: {email_column[:5]}")
        
        found_index = None
        for i, email in enumerate(email_column):
            if email == sample_email:
                found_index = i
                break
        
        if found_index is not None:
            permission = permissions_column[found_index] if found_index < len(permissions_column) else "N/A"
            print(f"✅ Found at index {found_index} (row {found_index + 1})")
            print(f"Permission: '{permission}'")
        else:
            print(f"❌ Not found in mailing list")
    
    print(f"\n=== FORMULA SIMULATION ===")
    print("Testing the fixed formula logic:")
    print(f"COUNTIF('Mailing List'!$A$2:$A, '{sample_email}') > 0")
    
    # Simulate the fixed formula
    email_range = email_column[1:]  # Skip first row (A2:A)
    count = email_range.count(sample_email)
    print(f"Count result: {count}")
    
    if count > 0:
        try:
            match_index = email_range.index(sample_email)
            permissions_range = permissions_column[1:]  # Skip first row (F2:F)
            permission = permissions_range[match_index] if match_index < len(permissions_range) else "N/A"
            print(f"Permission at matched index: '{permission}'")
            result = permission == "allowed"
            print(f"Final result: {result}")
        except:
            print("Error in lookup simulation")
    else:
        print("Final result: FALSE (not found)")

if __name__ == "__main__":
    debug_mailing_list()