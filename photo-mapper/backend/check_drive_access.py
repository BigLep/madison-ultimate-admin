#!/usr/bin/env python3
"""
Check Google Drive access and permissions for the service account.
"""

import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

CREDENTIALS_FILE = '.google-service-account.json'
GOOGLE_DRIVE_URL = 'https://drive.google.com/drive/folders/1ojyCLPVl_kzpZW8MOkY3wmDG2XQdLcUOAumoNNtjo19KaYkGRdbU_PGLAzufBiAHz_ATafR6?usp=drive_link'

def authenticate():
    """Authenticate with Google Drive API"""
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)

    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=[
            'https://www.googleapis.com/auth/drive.readonly',
            'https://www.googleapis.com/auth/drive.metadata.readonly'
        ]
    )
    return build('drive', 'v3', credentials=credentials)

def extract_folder_id(drive_url):
    """Extract folder ID from Google Drive URL"""
    if 'folders/' in drive_url:
        folder_id = drive_url.split('folders/')[1].split('?')[0].split('&')[0]
    elif '/open?id=' in drive_url:
        folder_id = drive_url.split('id=')[1].split('&')[0]
    else:
        raise ValueError(f"Could not extract folder ID from URL: {drive_url}")
    return folder_id

def check_folder_access(drive_service, folder_id):
    """Check if we can access the folder and get its details"""
    try:
        # Try to get folder metadata
        print(f"Checking access to folder: {folder_id}")

        folder = drive_service.files().get(
            fileId=folder_id,
            fields='id, name, mimeType, parents, permissions, shared'
        ).execute()

        print("✅ Successfully accessed folder:")
        print(f"  Name: {folder.get('name', 'Unknown')}")
        print(f"  ID: {folder.get('id')}")
        print(f"  Type: {folder.get('mimeType')}")
        print(f"  Shared: {folder.get('shared', 'Unknown')}")

        return True

    except HttpError as error:
        print(f"❌ Error accessing folder: {error}")

        if error.resp.status == 404:
            print("  → Folder not found or not accessible")
        elif error.resp.status == 403:
            print("  → Permission denied - service account needs access to folder")

        return False

def check_folder_contents(drive_service, folder_id):
    """Try to list folder contents"""
    try:
        print(f"\nListing contents of folder: {folder_id}")

        results = drive_service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            pageSize=10,
            fields="files(id, name, mimeType)"
        ).execute()

        files = results.get('files', [])
        print(f"Found {len(files)} files/folders:")

        for file in files:
            print(f"  - {file['name']} ({file['mimeType']})")

        return len(files)

    except HttpError as error:
        print(f"❌ Error listing folder contents: {error}")
        return 0

def main():
    print("="*60)
    print("GOOGLE DRIVE ACCESS CHECK")
    print("="*60)

    # Get service account email
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)

    service_account_email = creds_info['client_email']
    print(f"Service Account Email: {service_account_email}")

    # Authenticate
    print("\nAuthenticating with Google Drive...")
    drive_service = authenticate()
    print("✅ Authentication successful")

    # Extract folder ID
    folder_id = extract_folder_id(GOOGLE_DRIVE_URL)
    print(f"Folder ID: {folder_id}")

    # Check folder access
    if check_folder_access(drive_service, folder_id):
        file_count = check_folder_contents(drive_service, folder_id)

        if file_count == 0:
            print("\n⚠️  Folder is accessible but appears to be empty")

    else:
        print(f"\n❌ Cannot access the folder. To fix this:")
        print(f"1. Open the Google Drive folder in your browser:")
        print(f"   {GOOGLE_DRIVE_URL}")
        print(f"2. Click 'Share' button")
        print(f"3. Add this service account email as a viewer:")
        print(f"   {service_account_email}")
        print(f"4. Make sure 'Notify people' is unchecked")
        print(f"5. Click 'Share'")

if __name__ == "__main__":
    main()