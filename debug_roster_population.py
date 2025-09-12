#!/usr/bin/env python3
"""
Debug why roster data isn't populating
"""

import json
from google.oauth2 import service_account
import gspread

SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def debug_roster():
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    client = gspread.authorize(credentials)
    
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    roster_sheet = spreadsheet.worksheet('Roster')
    
    print("=== DEBUGGING ROSTER POPULATION ===\n")
    
    # Check row 3 (sources) in detail
    sources = roster_sheet.row_values(3)
    print(f"Row 3 (Sources) - first 10 columns:")
    for i, source in enumerate(sources[:10]):
        print(f"  Col {chr(65+i)}: '{source}'")
    
    print(f"\n=== CHECKING DATA ROWS (6-15) ===")
    
    # Check if there's any data in rows 6-15
    for row_num in range(6, 16):
        row_data = roster_sheet.row_values(row_num)
        non_empty = [cell for cell in row_data if cell]
        if non_empty:
            print(f"Row {row_num}: {len(non_empty)} non-empty cells")
            print(f"  First few values: {non_empty[:3]}")
        else:
            print(f"Row {row_num}: EMPTY")
    
    # Check if formulas exist
    print(f"\n=== CHECKING FOR FORMULAS (row 6) ===")
    formulas = roster_sheet.get('A6:J6', value_render_option='FORMULA')
    if formulas and formulas[0]:
        for i, formula in enumerate(formulas[0]):
            if formula and str(formula).startswith('='):
                print(f"  Col {chr(65+i)}: {formula[:60]}...")
    
    # Check Final Forms data
    print(f"\n=== FINAL FORMS DATA CHECK ===")
    try:
        ff_sheet = spreadsheet.worksheet('Final Forms')
        ff_sample = ff_sheet.get('A1:E3')
        print(f"First 3 rows, 5 columns:")
        for i, row in enumerate(ff_sample):
            print(f"  Row {i+1}: {row}")
    except Exception as e:
        print(f"  Error: {e}")
    
    # Check Additional Info data
    print(f"\n=== ADDITIONAL INFO DATA CHECK ===")
    try:
        ai_sheet = spreadsheet.worksheet('Additional Info')
        ai_sample = ai_sheet.get('A1:E3')
        print(f"First 3 rows, 5 columns:")
        for i, row in enumerate(ai_sample):
            print(f"  Row {i+1}: {row}")
    except Exception as e:
        print(f"  Error: {e}")

if __name__ == "__main__":
    debug_roster()