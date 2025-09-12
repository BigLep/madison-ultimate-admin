#!/usr/bin/env python3
"""
Analyze the exact formula issue with mailing list lookups
"""

import json
from google.oauth2 import service_account
import gspread

SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def analyze_formula_issue():
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    client = gspread.authorize(credentials)
    
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    roster_sheet = spreadsheet.worksheet('Roster')
    mailing_sheet = spreadsheet.worksheet('Mailing List')
    
    print("=== FORMULA ANALYSIS FOR ROW 15 ===")
    
    # Get the email being tested
    test_email = roster_sheet.get('G15')[0][0] if roster_sheet.get('G15') else ""
    print(f"Email in G15: '{test_email}'")
    
    # Get current value and formula
    try:
        f15_data = roster_sheet.get('F15')
        current_value = f15_data[0][0] if f15_data and f15_data[0] else ""
        
        f15_formula_data = roster_sheet.get('F15', value_render_option='FORMULA')
        current_formula = f15_formula_data[0][0] if f15_formula_data and f15_formula_data[0] else ""
    except (IndexError, TypeError):
        current_value = ""
        current_formula = ""
    
    print(f"Current value in F15: '{current_value}'")
    print(f"Current formula in F15: {current_formula}")
    
    print(f"\n=== MAILING LIST ANALYSIS ===")
    
    # Get first 10 rows of mailing list to understand structure
    mailing_data = mailing_sheet.get('A1:F10')
    
    print("Mailing List structure (first 10 rows):")
    for i, row in enumerate(mailing_data, 1):
        print(f"Row {i}: {row}")
    
    print(f"\n=== TESTING FORMULA COMPONENTS ===")
    
    if test_email:
        # Test range A2:A
        a_column = mailing_sheet.col_values(1)  # Column A
        print(f"Column A full data: {a_column}")
        
        # Test A2:A specifically
        a2_onwards = a_column[1:] if len(a_column) > 1 else []
        print(f"A2:A data: {a2_onwards}")
        
        # Test COUNTIF equivalent
        count = a2_onwards.count(test_email)
        print(f"COUNTIF('Mailing List'!$A$2:$A, '{test_email}') = {count}")
        
        if count > 0:
            # Test INDEX/MATCH equivalent
            try:
                match_index = a2_onwards.index(test_email)
                print(f"MATCH('{test_email}', A2:A, 0) would return index: {match_index}")
                
                # Get F column data
                f_column = mailing_sheet.col_values(6)  # Column F
                print(f"Column F full data: {f_column}")
                
                f2_onwards = f_column[1:] if len(f_column) > 1 else []
                print(f"F2:F data: {f2_onwards}")
                
                if match_index < len(f2_onwards):
                    permissions = f2_onwards[match_index]
                    print(f"INDEX('Mailing List'!$F$2:$F, {match_index}) = '{permissions}'")
                    print(f"'{permissions}' == 'allowed' ? {permissions == 'allowed'}")
                else:
                    print(f"❌ Index {match_index} is out of bounds for F2:F (length: {len(f2_onwards)})")
            except ValueError:
                print(f"❌ Email '{test_email}' not found in A2:A")
        else:
            print("Email not found, formula should return FALSE")
    
    print(f"\n=== CHECKING FOR EMPTY CELLS OR FORMATTING ISSUES ===")
    
    # Check if there are empty cells or formatting issues in the ranges
    a_range_values = mailing_sheet.get('A2:A20')  # Get first 20 rows
    f_range_values = mailing_sheet.get('F2:F20')
    
    print("A2:A20 with empty cell analysis:")
    for i, row in enumerate(a_range_values, 2):
        value = row[0] if row else ""
        print(f"  A{i}: '{value}' (empty: {not value})")
    
    print("F2:F20 with empty cell analysis:")
    for i, row in enumerate(f_range_values, 2):
        value = row[0] if row else ""
        print(f"  F{i}: '{value}' (empty: {not value})")

if __name__ == "__main__":
    analyze_formula_issue()