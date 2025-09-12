#!/usr/bin/env python3
"""
Script to debug and fix the mailing list column mapping issue
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
        with open(CREDENTIALS_FILE, 'r') as f:
            creds_info = json.load(f)
        
        credentials = service_account.Credentials.from_service_account_info(
            creds_info,
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive.readonly'
            ]
        )
        
        client = gspread.authorize(credentials)
        return client
    except Exception as e:
        print(f"Authentication failed: {e}")
        return None

def examine_mailing_list_structure(client):
    """Examine the current mailing list structure"""
    try:
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        mailing_list_sheet = spreadsheet.worksheet('Mailing List')
        
        print("=== CURRENT MAILING LIST STRUCTURE ===")
        
        # Get first 10 rows and 10 columns to examine structure
        data = mailing_list_sheet.get('A1:J10')
        
        print("Raw data (first 10 rows, 10 columns):")
        for i, row in enumerate(data):
            print(f"Row {i+1}: {row}")
        
        print("\n=== PROBLEM ANALYSIS ===")
        
        if len(data) > 0:
            first_row = data[0]
            if len(first_row) > 0 and "Members for group" in str(first_row[0]):
                print("❌ ISSUE FOUND: First row contains group header, not column headers")
                print("Expected: ['Email address', 'Nickname', 'Group status', ...]")
                print(f"Actual: {first_row}")
                
                if len(data) > 1:
                    second_row = data[1]
                    print(f"Second row: {second_row}")
                    if len(second_row) > 0 and "Email address" in str(second_row[0]):
                        print("✅ Real headers found in row 2")
                        return True  # Can fix this
                        
        return False
        
    except Exception as e:
        print(f"Error examining mailing list: {e}")
        return False

def fix_mailing_list_structure(client):
    """Fix the mailing list structure by removing the header row"""
    try:
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        mailing_list_sheet = spreadsheet.worksheet('Mailing List')
        
        print("\n=== FIXING MAILING LIST STRUCTURE ===")
        
        # Get all data
        all_data = mailing_list_sheet.get_all_values()
        
        if len(all_data) < 2:
            print("❌ Not enough data to fix")
            return False
        
        # Check if first row is the problematic header
        if "Members for group" in str(all_data[0][0]):
            print("Removing problematic first row...")
            
            # Clear the sheet
            mailing_list_sheet.clear()
            
            # Write data starting from row 2 (skip the problematic first row)
            fixed_data = all_data[1:]  # Skip first row
            
            if len(fixed_data) > 0:
                mailing_list_sheet.update('A1', fixed_data)
                print(f"✅ Updated mailing list with {len(fixed_data)} rows")
                
                # Show the new structure
                print("\nNew structure (first 3 rows):")
                for i, row in enumerate(fixed_data[:3]):
                    print(f"Row {i+1}: {row[:6]}")  # Show first 6 columns
                
                return True
        else:
            print("✅ Mailing list structure looks correct already")
            return True
            
    except Exception as e:
        print(f"Error fixing mailing list structure: {e}")
        return False

def test_mailing_list_lookup(client):
    """Test the mailing list lookup to see what's working"""
    try:
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        mailing_list_sheet = spreadsheet.worksheet('Mailing List')
        roster_sheet = spreadsheet.worksheet('Roster')
        
        print("\n=== TESTING MAILING LIST LOOKUP ===")
        
        # Get mailing list structure
        mailing_data = mailing_list_sheet.get('A1:F10')  # First 10 rows, columns A-F
        
        print("Mailing list data (A1:F3):")
        for i, row in enumerate(mailing_data[:3]):
            print(f"Row {i+1}: {row}")
        
        # Find column positions
        if len(mailing_data) > 0:
            headers = mailing_data[0]
            print(f"\nHeaders: {headers}")
            
            email_col_idx = None
            permissions_col_idx = None
            
            for i, header in enumerate(headers):
                if "email" in str(header).lower():
                    email_col_idx = i
                    print(f"Email column found at index {i}: '{header}'")
                if "permission" in str(header).lower():
                    permissions_col_idx = i
                    print(f"Permissions column found at index {i}: '{header}'")
            
            if email_col_idx is not None and permissions_col_idx is not None:
                print(f"\n✅ Column mapping: Email=Col {email_col_idx+1}, Permissions=Col {permissions_col_idx+1}")
                
                # Test with a sample email
                sample_emails = []
                if len(mailing_data) > 1:
                    for row in mailing_data[1:4]:  # Get first 3 email addresses
                        if len(row) > email_col_idx and row[email_col_idx]:
                            sample_emails.append(row[email_col_idx])
                
                print(f"Sample emails from mailing list: {sample_emails[:3]}")
                
                # Get a sample email from roster to test
                roster_emails = roster_sheet.col_values(15)  # Parent 1 Email column (O)
                sample_roster_email = None
                for email in roster_emails[5:]:  # Skip metadata rows
                    if email and "@" in email:
                        sample_roster_email = email
                        break
                
                if sample_roster_email:
                    print(f"Testing with roster email: {sample_roster_email}")
                    
                    # Check if it exists in mailing list
                    found = False
                    permissions = None
                    for row in mailing_data[1:]:
                        if len(row) > email_col_idx and row[email_col_idx] == sample_roster_email:
                            found = True
                            if len(row) > permissions_col_idx:
                                permissions = row[permissions_col_idx]
                            break
                    
                    print(f"Email found in mailing list: {found}")
                    if found:
                        print(f"Permissions: '{permissions}'")
                
                return True
            else:
                print("❌ Could not find email or permissions columns")
                return False
                
        return False
        
    except Exception as e:
        print(f"Error testing mailing list lookup: {e}")
        return False

def main():
    print("Madison Ultimate Mailing List Fix Tool")
    print("=" * 50)
    
    # Authenticate
    client = authenticate_google_sheets()
    if not client:
        print("Failed to authenticate. Check your credentials file.")
        return
    
    print("✅ Authentication successful")
    
    # Examine the current structure
    needs_fix = examine_mailing_list_structure(client)
    
    if needs_fix:
        # Fix the structure
        if fix_mailing_list_structure(client):
            print("✅ Mailing list structure fixed")
        else:
            print("❌ Failed to fix mailing list structure")
            return
    
    # Test the lookup functionality
    test_mailing_list_lookup(client)
    
    print("\n=== NEXT STEPS ===")
    print("1. Check the Roster sheet - the #N/A errors should be resolved")
    print("2. If still showing #N/A, regenerate the roster using the Google Apps Script")
    print("3. The mailing list lookup should now work correctly")

if __name__ == "__main__":
    main()