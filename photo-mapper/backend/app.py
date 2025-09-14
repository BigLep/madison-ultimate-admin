#!/usr/bin/env python3
"""
Flask API backend for photo-to-player mapping application.
Reuses existing Google API integration and matching algorithms.
"""

from flask import Flask, jsonify, request
from flask_cors import CORS
import json
import pandas as pd
from google.oauth2 import service_account
import gspread
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import re
import sys
import os
from pathlib import Path

app = Flask(__name__)
CORS(app)

# Configuration
CREDENTIALS_FILE = '.google-service-account.json'
SPREADSHEET_ID = '1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8'
GOOGLE_DRIVE_URL = 'https://drive.google.com/drive/folders/1ojyCLPVl_kzpZW8MOkY3wmDG2XQdLcUOAumoNNtjo19KaYkGRdbU_PGLAzufBiAHz_ATafR6?usp=drive_link'

def authenticate_drive():
    """Authenticate with Google Drive API"""
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)

    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=[
            'https://www.googleapis.com/auth/drive.readonly',
            'https://www.googleapis.com/auth/drive.metadata.readonly',
            'https://www.googleapis.com/auth/drive.file'
        ]
    )
    return build('drive', 'v3', credentials=credentials)

def authenticate_sheets():
    """Authenticate with Google Sheets API"""
    with open(CREDENTIALS_FILE, 'r') as f:
        creds_info = json.load(f)

    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    return gspread.authorize(credentials)

def extract_folder_id(drive_url):
    """Extract folder ID from Google Drive URL"""
    if 'folders/' in drive_url:
        folder_id = drive_url.split('folders/')[1].split('?')[0].split('&')[0]
    elif '/open?id=' in drive_url:
        folder_id = drive_url.split('id=')[1].split('&')[0]
    else:
        raise ValueError(f"Could not extract folder ID from URL: {drive_url}")
    return folder_id

def get_photos_from_folder(drive_service, folder_id):
    """Get list of photos from Google Drive folder"""
    print(f"Fetching photos from folder ID: {folder_id}")

    photos = []
    page_token = None

    # Image MIME types to look for
    image_mimes = [
        'image/jpeg', 'image/jpg', 'image/png', 'image/gif',
        'image/bmp', 'image/webp', 'image/tiff'
    ]

    try:
        while True:
            query = f"'{folder_id}' in parents and trashed=false"

            results = drive_service.files().list(
                q=query,
                pageSize=1000,
                fields="nextPageToken, files(id, name, mimeType, webViewLink, webContentLink, thumbnailLink)",
                pageToken=page_token
            ).execute()

            files = results.get('files', [])

            for file in files:
                # Filter for image files
                if file['mimeType'] in image_mimes:
                    photos.append({
                        'id': file['id'],
                        'name': file['name'],
                        'mimeType': file['mimeType'],
                        'webViewLink': file.get('webViewLink', ''),
                        'webContentLink': file.get('webContentLink', ''),
                        'thumbnailLink': file.get('thumbnailLink', ''),
                        'directLink': f"https://drive.google.com/uc?id={file['id']}"
                    })

            page_token = results.get('nextPageToken')
            if not page_token:
                break

    except HttpError as error:
        print(f"An error occurred: {error}")
        raise

    print(f"Found {len(photos)} photos in the folder")
    return photos

def load_roster_from_sheets(sheets_client):
    """Load roster data from Google Sheets"""
    print("Loading roster data from Google Sheets...")

    spreadsheet = sheets_client.open_by_key(SPREADSHEET_ID)
    roster_sheet = spreadsheet.worksheet('Roster')

    # Get all values
    all_values = roster_sheet.get_all_values()

    # Find header row (should be row 1, index 0)
    headers = all_values[0]

    # Get data starting from row 7 (index 6) to skip metadata
    data_rows = all_values[6:]

    players = []
    for row in data_rows:
        if len(row) > 0 and row[1]:  # Check if first name exists (column B)
            try:
                player_info = {
                    'student_id': row[1] if len(row) > 1 else '',
                    'first_name': row[2] if len(row) > 2 else '',
                    'last_name': row[3] if len(row) > 3 else '',
                    'full_name': row[4] if len(row) > 4 else '',
                    'parent1_first': row[18] if len(row) > 18 else '',
                    'parent1_last': row[19] if len(row) > 19 else '',
                    'parent2_first': row[22] if len(row) > 22 else '',
                    'parent2_last': row[23] if len(row) > 23 else '',
                }

                # Skip empty rows
                if player_info['first_name'] and player_info['last_name']:
                    players.append(player_info)

            except IndexError:
                continue

    print(f"Loaded {len(players)} players from roster")
    return players

def normalize_name(name):
    """Normalize a name for matching (remove special chars, lowercase, etc.)"""
    if not name:
        return ""

    # Convert to lowercase and remove special characters
    normalized = re.sub(r'[^\w\s-]', '', str(name).lower())
    # Replace multiple spaces/dashes with single space
    normalized = re.sub(r'[\s\-]+', ' ', normalized).strip()
    return normalized

def generate_name_variations(player):
    """Generate various name combinations for matching"""
    variations = set()

    first = player['first_name']
    last = player['last_name']
    full = player['full_name']

    if first and last:
        # Basic combinations
        variations.add(f"{first} {last}")
        variations.add(f"{last} {first}")
        variations.add(f"{first}_{last}")
        variations.add(f"{last}_{first}")
        variations.add(f"{first}.{last}")
        variations.add(f"{last}.{first}")
        variations.add(f"{first}-{last}")
        variations.add(f"{last}-{first}")

        # Just first name
        variations.add(first)
        # Just last name
        variations.add(last)

    # Full name variations
    if full:
        variations.add(full)
        variations.add(full.replace(' ', '_'))
        variations.add(full.replace(' ', '.'))
        variations.add(full.replace(' ', '-'))

    # Parent name variations
    for parent_first, parent_last in [
        (player['parent1_first'], player['parent1_last']),
        (player['parent2_first'], player['parent2_last'])
    ]:
        if parent_first and parent_last:
            variations.add(f"{parent_first} {parent_last}")
            variations.add(f"{parent_first}_{parent_last}")
            variations.add(f"{parent_last}_{parent_first}")
            variations.add(parent_first)
            variations.add(parent_last)

    # Remove empty strings and normalize
    normalized_variations = set()
    for var in variations:
        normalized = normalize_name(var)
        if normalized:
            normalized_variations.add(normalized)

    return normalized_variations

def match_photo_to_player(photo_name, players):
    """Match a photo filename to a player"""
    # Remove file extension and normalize filename
    base_name = Path(photo_name).stem
    normalized_filename = normalize_name(base_name)

    matches = []

    for player in players:
        player_variations = generate_name_variations(player)

        # Check for exact matches
        if normalized_filename in player_variations:
            matches.append({
                'player': player,
                'confidence': 'high',
                'match_type': 'exact',
                'matched_variation': normalized_filename
            })
            continue

        # Check for partial matches (filename contains player name or vice versa)
        for variation in player_variations:
            if len(variation) > 2:  # Avoid matching very short names
                if variation in normalized_filename or normalized_filename in variation:
                    matches.append({
                        'player': player,
                        'confidence': 'medium',
                        'match_type': 'partial',
                        'matched_variation': variation
                    })
                    break

    # Sort by confidence and return best matches
    confidence_order = {'high': 3, 'medium': 2, 'low': 1}
    matches.sort(key=lambda x: confidence_order[x['confidence']], reverse=True)

    return matches

def create_photo_mappings(photos, players):
    """Create mappings between photos and players"""
    print("Creating photo-to-player mappings...")

    mappings = []

    for photo in photos:
        matches = match_photo_to_player(photo['name'], players)

        if matches:
            # Take the best match
            best_match = matches[0]
            mapping = {
                'photo_id': photo['id'],
                'filename': photo['name'],
                'thumbnail_url': photo['thumbnailLink'],
                'direct_link': photo['directLink'],
                'matched_player': best_match['player']['full_name'],
                'confidence': best_match['confidence'],
                'match_type': best_match['match_type'],
                'matched_variation': best_match['matched_variation'],
                'student_id': best_match['player']['student_id'],
                'alternative_matches': [
                    f"{m['player']['full_name']} ({m['confidence']})"
                    for m in matches[1:3]  # Show up to 2 alternatives
                ] if len(matches) > 1 else []
            }
            mappings.append(mapping)
        else:
            # No match found
            mapping = {
                'photo_id': photo['id'],
                'filename': photo['name'],
                'thumbnail_url': photo['thumbnailLink'],
                'direct_link': photo['directLink'],
                'matched_player': '',
                'confidence': 'medium',
                'match_type': 'no_match',
                'matched_variation': '',
                'student_id': '',
                'alternative_matches': []
            }
            mappings.append(mapping)

    print(f"Created {len(mappings)} photo mappings")
    return mappings

@app.route('/api/load-data', methods=['GET'])
def load_data():
    """Load photos from Google Drive, roster from Sheets, and create mappings"""
    try:
        print("Loading data for photo mapping...")

        # Authenticate with Google APIs
        drive_service = authenticate_drive()
        sheets_client = authenticate_sheets()

        # Get photos from Google Drive
        folder_id = extract_folder_id(GOOGLE_DRIVE_URL)
        photos = get_photos_from_folder(drive_service, folder_id)

        if not photos:
            return jsonify({'error': 'No photos found in the Google Drive folder'}), 404

        # Load roster from Google Sheets
        players = load_roster_from_sheets(sheets_client)

        if not players:
            return jsonify({'error': 'No players found in the roster sheet'}), 404

        # Create mappings using existing algorithm
        mappings = create_photo_mappings(photos, players)

        # Format roster for frontend autocomplete
        roster_for_frontend = [
            {
                'full_name': player['full_name'],
                'first_name': player['first_name'],
                'last_name': player['last_name'],
                'student_id': player['student_id']
            }
            for player in players
        ]

        response_data = {
            'mappings': mappings,
            'roster': roster_for_frontend,
            'stats': {
                'total_photos': len(photos),
                'total_players': len(players),
                'high_confidence_matches': len([m for m in mappings if m['confidence'] == 'high']),
                'medium_confidence_matches': len([m for m in mappings if m['confidence'] == 'medium'])
            }
        }

        return jsonify(response_data)

    except Exception as e:
        print(f"Error loading data: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/rename-files', methods=['POST'])
def rename_files():
    """Rename files in Google Drive based on selected player names"""
    print("=== RENAME FILES ENDPOINT CALLED ===")
    try:
        data = request.json
        print(f"Request data: {data}")
        renames = data.get('renames', [])  # List of {photo_id, current_name, new_name}

        print(f"Rename request received with {len(renames)} files to rename")
        for r in renames:
            print(f"  - {r.get('photo_id')}: '{r.get('current_name')}' -> '{r.get('new_name')}'")

        if not renames:
            print("No renames provided, returning error")
            return jsonify({'error': 'No rename operations provided'}), 400

        drive_service = authenticate_drive()
        results = []

        # Test if we can access files through folder listing instead
        folder_id = "1ojyCLPVl_kzpZW8MOkY3wmDG2XQdLcUOAumoNNtjo19KaYkGRdbU_PGLAzufBiAHz_ATafR6"
        try:
            # Try to list files in folder - this works for load-data
            folder_files = drive_service.files().list(
                q=f"'{folder_id}' in parents",
                fields='files(id,name)'
            ).execute()
            print(f"Can list {len(folder_files['files'])} files in folder via folder query")

            # Try direct file access
            test_photo_id = renames[0]['photo_id']
            file_metadata = drive_service.files().get(fileId=test_photo_id).execute()
            print(f"Successfully got metadata for file {test_photo_id}: {file_metadata.get('name')}")

        except Exception as e:
            print(f"Failed file access test: {str(e)}")
            return jsonify({
                'error': f'Service account cannot access individual files. Please create a new folder and move photos there, then share the new folder with the service account as Editor.',
                'details': str(e)
            }), 403

        for rename_op in renames:
            photo_id = rename_op.get('photo_id')
            current_name = rename_op.get('current_name')
            new_name = rename_op.get('new_name')

            if not all([photo_id, current_name, new_name]):
                results.append({
                    'photo_id': photo_id,
                    'status': 'error',
                    'message': 'Missing required fields'
                })
                continue

            # Skip if names are the same
            current_stem = Path(current_name).stem
            if current_stem.lower() == new_name.lower():
                results.append({
                    'photo_id': photo_id,
                    'status': 'skipped',
                    'message': 'Names are already the same'
                })
                continue

            try:
                # Get file extension
                file_extension = Path(current_name).suffix
                new_filename = f"{new_name}{file_extension}"

                # Alternative approach: Copy file with new name, then delete original
                # First, get the original file metadata
                original_file = drive_service.files().get(fileId=photo_id).execute()

                # Copy the file with the new name
                copied_file = drive_service.files().copy(
                    fileId=photo_id,
                    body={'name': new_filename},
                    supportsAllDrives=True
                ).execute()

                # Try to delete the original file
                try:
                    drive_service.files().delete(fileId=photo_id).execute()
                    print(f"Successfully renamed {current_name} -> {new_filename} (via copy/delete)")
                    results.append({
                        'photo_id': copied_file['id'],  # Use new file ID
                        'status': 'success',
                        'old_name': current_name,
                        'new_name': new_filename
                    })
                except Exception as delete_error:
                    print(f"Copied file but couldn't delete original {current_name}: {str(delete_error)}")
                    results.append({
                        'photo_id': copied_file['id'],
                        'status': 'partial_success',
                        'old_name': current_name,
                        'new_name': new_filename,
                        'message': f'File copied with new name, but original remains: {str(delete_error)}'
                    })

            except HttpError as e:
                error_msg = f'Drive API error: {str(e)}'
                print(f"Failed to rename {current_name}: {error_msg}")
                results.append({
                    'photo_id': photo_id,
                    'status': 'error',
                    'message': error_msg
                })
            except Exception as e:
                error_msg = f'Unexpected error: {str(e)}'
                print(f"Failed to rename {current_name}: {error_msg}")
                results.append({
                    'photo_id': photo_id,
                    'status': 'error',
                    'message': error_msg
                })

        successful = len([r for r in results if r['status'] == 'success'])
        skipped = len([r for r in results if r['status'] == 'skipped'])
        failed = len([r for r in results if r['status'] == 'error'])

        print(f"Rename operation completed: {successful} successful, {skipped} skipped, {failed} failed")

        return jsonify({
            'success': True,
            'results': results,
            'total': len(renames),
            'successful': successful,
            'skipped': skipped,
            'failed': failed
        })

    except Exception as e:
        print(f"Error in rename_files: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/export-csv', methods=['POST'])
def export_csv():
    """Export final mappings as CSV data"""
    try:
        data = request.json
        mappings = data.get('mappings', [])

        if not mappings:
            return jsonify({'error': 'No mappings provided'}), 400

        # Create CSV data
        csv_rows = [['PlayerName', 'DriveId', 'Filename', 'GoogleDriveLink', 'DirectImageLink', 'ThumbnailLink']]

        for mapping in mappings:
            player_name = mapping.get('playerName', '').strip()
            if player_name:
                drive_id = mapping.get('driveId', '')
                drive_view_link = f"https://drive.google.com/file/d/{drive_id}/view"
                direct_image_link = f"https://drive.google.com/uc?id={drive_id}"

                csv_rows.append([
                    player_name,
                    drive_id,
                    mapping.get('filename', ''),
                    drive_view_link,
                    direct_image_link,
                    mapping.get('thumbnailUrl', '')
                ])

        return jsonify({
            'csvData': csv_rows,
            'totalMappings': len(csv_rows) - 1  # Subtract header row
        })

    except Exception as e:
        print(f"Error exporting CSV: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({'status': 'healthy', 'service': 'photo-mapper-backend'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)