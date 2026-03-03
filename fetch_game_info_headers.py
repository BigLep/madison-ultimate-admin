#!/usr/bin/env python3
"""
Fetch Game Info sheet headers from the roster spreadsheet.
Uses .google-service-account.json in this directory.
"""
import json
from google.oauth2 import service_account
import gspread

# Spring 2026 roster spreadsheet
SPREADSHEET_ID = '1kV3Y_GST_Y-X9PZFXu9yFkCzGWvhk9f7G24Y8QNuayU'
CREDENTIALS_FILE = '.google-service-account.json'
GAME_INFO_SHEET_NAME = '📍Game Info'


def main():
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)
    credentials = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'],
    )
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    worksheet = spreadsheet.worksheet(GAME_INFO_SHEET_NAME)
    # First row = headers; get first 2 data rows for context
    rows = worksheet.get('A1:Z5')
    headers = rows[0] if rows else []
    print('Game Info sheet – column headers (row 1):')
    for i, h in enumerate(headers):
        if h:
            print(f'  {i}: "{h}"')
    print('\nAll headers as list (for code):')
    print([h for h in headers if h])
    if len(rows) > 1:
        print('\nFirst data row (row 2):', rows[1])


if __name__ == '__main__':
    main()
