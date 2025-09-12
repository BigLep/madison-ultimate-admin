#!/usr/bin/env python3
"""
Test the exact formula to identify #N/A source
"""

import json
from google.oauth2 import service_account
import gspread

SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def test_formula():
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    client = gspread.authorize(credentials)
    
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    roster_sheet = spreadsheet.worksheet('Roster')
    
    print("=== TESTING ROW 15 FORMULA ===")
    
    # Get current formula in row 15, column F (Student Personal Email On Mailing List?)
    try:
        current_formula = roster_sheet.get('F15', value_render_option='FORMULA')[0][0]
        print(f"Current formula: {current_formula}")
        
        # Get the current value
        current_value = roster_sheet.get('F15')[0][0]
        print(f"Current value: '{current_value}' (type: {type(current_value)})")
        
        # Get the email being tested (column E)
        test_email = roster_sheet.get('E15')[0][0]
        print(f"Email being tested: '{test_email}'")
        
        # Test if column E is empty
        if not test_email:
            print("✅ Column E is empty, should return FALSE")
        else:
            print("Column E has value, testing mailing list lookup...")
            
            # Test the mailing list lookup components
            mailing_sheet = spreadsheet.worksheet('Mailing List')
            
            # Test COUNTIF part
            emails_from_A2 = mailing_sheet.get('A2:A')[1:]  # Skip header row
            email_count = sum(1 for row in emails_from_A2 if row and len(row) > 0 and row[0] == test_email)
            print(f"COUNTIF result: {email_count}")
            
            if email_count > 0:
                # Test INDEX/MATCH part
                try:
                    match_result = None
                    permissions_from_F2 = mailing_sheet.get('F2:F')
                    
                    # Find the match manually
                    for i, row in enumerate(emails_from_A2):
                        if row and len(row) > 0 and row[0] == test_email:
                            if i < len(permissions_from_F2) and len(permissions_from_F2[i]) > 0:
                                match_result = permissions_from_F2[i][0]
                            break
                    
                    print(f"INDEX/MATCH result: '{match_result}'")
                    
                    if match_result:
                        final_result = match_result == "allowed"
                        print(f"Final boolean result: {final_result}")
                    else:
                        print("⚠️  INDEX/MATCH returned empty - this could cause #N/A")
                        
                except Exception as e:
                    print(f"❌ Error in INDEX/MATCH simulation: {e}")
            else:
                print("Email not found, should return FALSE")
        
        # Now let's test with a corrected formula
        print(f"\n=== TESTING CORRECTED FORMULA ===")
        
        # Try a simpler version first
        test_formula_simple = f'=IF(E15="",FALSE,IF(AND(COUNTIF(\'Mailing List\'!$A$2:$A,E15)>0,INDEX(\'Mailing List\'!$F$2:$F,MATCH(E15,\'Mailing List\'!$A$2:$A,0))="allowed"),TRUE,FALSE))'
        
        print("Setting test formula in G15 for comparison...")
        roster_sheet.update('G15', test_formula_simple, value_input_option='USER_ENTERED')
        
        # Get result
        import time
        time.sleep(2)  # Wait for calculation
        test_result = roster_sheet.get('G15')[0][0]
        print(f"Test formula result: '{test_result}'")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_formula()