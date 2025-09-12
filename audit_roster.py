#!/usr/bin/env python3
"""
Comprehensive audit of the Madison Ultimate roster spreadsheet
Checks data integrity and mapping from all three sources
"""

import json
import pandas as pd
from google.oauth2 import service_account
import gspread
from datetime import datetime
import sys

SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
CREDENTIALS_FILE = '.google-service-account.json'

def authenticate():
    """Authenticate with Google Sheets API"""
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    return gspread.authorize(credentials)

def audit_roster_structure(client):
    """Audit the roster sheet structure and metadata"""
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    roster_sheet = spreadsheet.worksheet('Roster')
    
    print("=" * 80)
    print("ROSTER STRUCTURE AUDIT")
    print("=" * 80)
    
    # Get metadata rows
    headers = roster_sheet.row_values(1)
    types = roster_sheet.row_values(2)
    sources = roster_sheet.row_values(3)
    notes = roster_sheet.row_values(4)
    
    # Analyze column structure
    print(f"\nTotal columns defined: {len(headers)}")
    print(f"Columns with headers: {len([h for h in headers if h])}")
    print(f"Columns with types: {len([t for t in types if t])}")
    print(f"Columns with sources: {len([s for s in sources if s])}")
    
    # Map sources to columns
    source_map = {
        'Final Forms': [],
        'Additional Info': [],
        'Mailing List': [],
        'Formula': [],
        'Manual': [],
        'Blank/Empty': [],
        'Unknown': []
    }
    
    for i, (header, source) in enumerate(zip(headers, sources)):
        if not header:
            continue
            
        col_letter = chr(65 + i) if i < 26 else f"{chr(65 + i//26 - 1)}{chr(65 + i%26)}"
        
        if 'Final Forms' in source:
            source_map['Final Forms'].append((col_letter, header))
        elif 'Additional Info' in source:
            source_map['Additional Info'].append((col_letter, header))
        elif 'Mailing List' in source:
            source_map['Mailing List'].append((col_letter, header))
        elif source == 'Formula':
            source_map['Formula'].append((col_letter, header))
        elif source == 'Manual':
            source_map['Manual'].append((col_letter, header))
        elif not source or source.strip() == '':
            source_map['Blank/Empty'].append((col_letter, header))
        else:
            source_map['Unknown'].append((col_letter, header, source))
    
    print("\n" + "=" * 60)
    print("COLUMN SOURCE MAPPING")
    print("=" * 60)
    
    for source, cols in source_map.items():
        if cols:
            print(f"\n{source} ({len(cols)} columns):")
            for col_info in cols[:5]:  # Show first 5
                if len(col_info) == 2:
                    print(f"  {col_info[0]:3} - {col_info[1]}")
                else:
                    print(f"  {col_info[0]:3} - {col_info[1]} [{col_info[2]}]")
            if len(cols) > 5:
                print(f"  ... and {len(cols) - 5} more")
    
    return roster_sheet, headers, types, sources

def audit_data_consistency(roster_sheet, headers):
    """Check data consistency in the roster"""
    print("\n" + "=" * 60)
    print("DATA CONSISTENCY AUDIT")
    print("=" * 60)
    
    # Get all data (skip metadata rows)
    all_data = roster_sheet.get_all_values()[5:]  # Skip first 5 metadata rows
    
    print(f"\nTotal data rows: {len(all_data)}")
    
    # Count non-empty rows (based on first name)
    non_empty_rows = [row for row in all_data if row[0]]  # Column A is First Name
    print(f"Non-empty data rows: {len(non_empty_rows)}")
    
    # Check for specific data issues
    issues = []
    
    # Find key column indices
    col_indices = {}
    for i, header in enumerate(headers):
        if 'First Name' in header:
            col_indices['first_name'] = i
        elif 'Last Name' in header:
            col_indices['last_name'] = i
        elif 'Student Personal Email' == header:
            col_indices['student_email'] = i
        elif 'Parent 1 Email' in header:
            col_indices['parent1_email'] = i
        elif 'Parent 2 Email' in header:
            col_indices['parent2_email'] = i
    
    # Check each row for issues
    for row_num, row in enumerate(non_empty_rows, start=6):  # Start at row 6 (after metadata)
        row_issues = []
        
        # Check if name fields are populated
        if 'first_name' in col_indices and not row[col_indices['first_name']]:
            row_issues.append("Missing first name")
        
        if 'last_name' in col_indices and not row[col_indices['last_name']]:
            row_issues.append("Missing last name")
        
        # Check email fields
        if 'student_email' in col_indices:
            email = row[col_indices['student_email']]
            if email and '@seattleschools.org' not in email and '@' not in email:
                row_issues.append(f"Invalid student email: {email}")
        
        if row_issues:
            issues.append((row_num, row_issues))
    
    if issues:
        print(f"\n⚠️  Found {len(issues)} rows with potential issues:")
        for row_num, row_issues in issues[:10]:  # Show first 10
            print(f"  Row {row_num}: {', '.join(row_issues)}")
        if len(issues) > 10:
            print(f"  ... and {len(issues) - 10} more")
    else:
        print("\n✅ No obvious data consistency issues found")
    
    return all_data, non_empty_rows

def audit_formulas(roster_sheet, headers, sources):
    """Audit formula columns for errors"""
    print("\n" + "=" * 60)
    print("FORMULA AUDIT")
    print("=" * 60)
    
    # Find formula columns
    formula_cols = []
    for i, (header, source) in enumerate(zip(headers, sources)):
        if source == 'Formula' or 'Mailing List' in source:
            formula_cols.append((i, header))
    
    print(f"\nFound {len(formula_cols)} formula/lookup columns")
    
    # Check a sample of formula cells for #N/A or errors
    sample_rows = [6, 10, 15, 20]  # Sample rows to check
    
    for col_idx, col_name in formula_cols:
        col_letter = chr(65 + col_idx) if col_idx < 26 else f"{chr(65 + col_idx//26 - 1)}{chr(65 + col_idx%26)}"
        
        errors_found = []
        for row in sample_rows:
            try:
                cell_value = roster_sheet.cell(row, col_idx + 1).value
                if cell_value and ('#N/A' in str(cell_value) or '#ERROR' in str(cell_value) or '#REF' in str(cell_value)):
                    errors_found.append((row, cell_value))
            except:
                pass
        
        if errors_found:
            print(f"\n⚠️  Column {col_letter} ({col_name}):")
            for row, error in errors_found:
                print(f"    Row {row}: {error}")
        else:
            print(f"  ✅ Column {col_letter} ({col_name}): No errors in sample")

def audit_mailing_list_lookups(client, roster_sheet, headers):
    """Specifically audit mailing list lookup columns"""
    print("\n" + "=" * 60)
    print("MAILING LIST LOOKUP AUDIT")
    print("=" * 60)
    
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    mailing_sheet = spreadsheet.worksheet('Mailing List')
    
    # Get mailing list emails
    mailing_emails = mailing_sheet.col_values(1)[2:]  # Skip header rows
    mailing_emails_set = set(email.lower() for email in mailing_emails if email)
    
    print(f"\nMailing list contains {len(mailing_emails_set)} unique emails")
    
    # Find "On Mailing List" columns
    mailing_list_cols = []
    for i, header in enumerate(headers):
        if 'On Mailing List' in header or 'on mailing list' in header.lower():
            mailing_list_cols.append((i, header))
    
    print(f"Found {len(mailing_list_cols)} 'On Mailing List' columns")
    
    # Check a sample of values
    sample_data = roster_sheet.get('A6:AZ15')  # Get rows 6-15
    
    for col_idx, col_name in mailing_list_cols:
        print(f"\n  Checking {col_name} (column {chr(65 + col_idx)}):")
        
        # Find corresponding email column
        email_col_name = col_name.replace('On Mailing List?', '').replace('on mailing list?', '').strip()
        email_col_idx = None
        
        for i, h in enumerate(headers):
            if h and email_col_name in h and 'On Mailing List' not in h:
                email_col_idx = i
                break
        
        if email_col_idx is not None:
            mismatches = []
            for row_idx, row in enumerate(sample_data):
                if row_idx >= len(row) or col_idx >= len(row):
                    continue
                    
                email = row[email_col_idx] if email_col_idx < len(row) else ''
                lookup_result = row[col_idx] if col_idx < len(row) else ''
                
                if email and email.strip():
                    email_lower = email.lower().strip()
                    expected = 'TRUE' if email_lower in mailing_emails_set else 'FALSE'
                    
                    if lookup_result and lookup_result != expected:
                        mismatches.append((row_idx + 6, email, lookup_result, expected))
            
            if mismatches:
                print(f"    ⚠️  Found {len(mismatches)} mismatches:")
                for row, email, actual, expected in mismatches[:3]:
                    print(f"      Row {row}: {email} -> {actual} (expected {expected})")
            else:
                print(f"    ✅ All sample lookups match expected values")

def audit_data_sources(client):
    """Audit the actual data source files"""
    print("\n" + "=" * 60)
    print("DATA SOURCE AUDIT")
    print("=" * 60)
    
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    
    # Check Final Forms sheet
    try:
        ff_sheet = spreadsheet.worksheet('Final Forms')
        ff_data = ff_sheet.get_all_values()
        print(f"\n✅ Final Forms sheet: {len(ff_data)} rows, {len(ff_data[0]) if ff_data else 0} columns")
        if ff_data:
            print(f"   Headers: {ff_data[0][:5]}...")
    except:
        print("\n⚠️  Final Forms sheet not found or inaccessible")
    
    # Check Additional Info sheet
    try:
        ai_sheet = spreadsheet.worksheet('Additional Info')
        ai_data = ai_sheet.get_all_values()
        print(f"\n✅ Additional Info sheet: {len(ai_data)} rows, {len(ai_data[0]) if ai_data else 0} columns")
        if ai_data:
            print(f"   Headers: {ai_data[0][:5]}...")
    except:
        print("\n⚠️  Additional Info sheet not found or inaccessible")
    
    # Check Mailing List sheet
    try:
        ml_sheet = spreadsheet.worksheet('Mailing List')
        ml_data = ml_sheet.get_all_values()
        print(f"\n✅ Mailing List sheet: {len(ml_data)} rows, {len(ml_data[0]) if ml_data else 0} columns")
        
        # Check for header row issues
        if ml_data and ml_data[0][0] != 'Email address':
            print(f"   ⚠️  First row may not be proper headers: {ml_data[0][0]}")
        if len(ml_data) > 2:
            print(f"   Data starts at row 3: {ml_data[2][0] if ml_data[2] else 'empty'}")
    except:
        print("\n⚠️  Mailing List sheet not found or inaccessible")

def main():
    print("=" * 80)
    print("MADISON ULTIMATE ROSTER AUDIT")
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)
    
    # Authenticate
    client = authenticate()
    print("✅ Authentication successful")
    
    # Run audits
    roster_sheet, headers, types, sources = audit_roster_structure(client)
    all_data, non_empty_rows = audit_data_consistency(roster_sheet, headers)
    audit_formulas(roster_sheet, headers, sources)
    audit_mailing_list_lookups(client, roster_sheet, headers)
    audit_data_sources(client)
    
    # Summary
    print("\n" + "=" * 80)
    print("AUDIT SUMMARY")
    print("=" * 80)
    print(f"""
    • Total columns: {len(headers)}
    • Data rows: {len(non_empty_rows)}
    • Formula columns: {len([s for s in sources if s == 'Formula'])}
    • Manual columns: {len([s for s in sources if s == 'Manual'])}
    • Empty source columns: {len([s for s in sources if not s])}
    
    Review the detailed output above for any issues that need attention.
    """)

if __name__ == "__main__":
    main()