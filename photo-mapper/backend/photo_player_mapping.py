#!/usr/bin/env python3
"""
Create a mapping between photos in Google Drive and player names from roster CSV.
Uses Google Drive API to list photos and matches them to player names using various strategies.
"""

import json
import pandas as pd
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
import sys
import urllib.parse
from pathlib import Path

CREDENTIALS_FILE = '.google-service-account.json'
ROSTER_CSV_PATH = '/Users/sal/Downloads/2025 Fall Coach Sheets - Roster.csv'
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
    # Handle various Google Drive URL formats
    if 'folders/' in drive_url:
        # Format: https://drive.google.com/drive/folders/FOLDER_ID?...
        folder_id = drive_url.split('folders/')[1].split('?')[0].split('&')[0]
    elif '/open?id=' in drive_url:
        # Format: https://drive.google.com/open?id=FOLDER_ID
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
        sys.exit(1)

    print(f"Found {len(photos)} photos in the folder")
    return photos

def load_roster_data(csv_path):
    """Load and parse roster CSV data"""
    print(f"Loading roster data from: {csv_path}")

    # Read CSV, using row 0 as header, skip metadata rows, start data from row 6
    df = pd.read_csv(csv_path, header=0, skiprows=range(1, 6))  # Skip rows 1-5, keep header row 0

    # Filter out empty rows (where First Name is empty)
    df = df[df['First Name'].notna() & (df['First Name'] != '')]

    players = []
    for _, row in df.iterrows():
        player_info = {
            'full_name': row.get('Full Name', ''),
            'first_name': row.get('First Name', ''),
            'last_name': row.get('Last Name', ''),
            'student_id': row.get('StudentID', ''),
            'parent1_first': row.get('Parent 1 First Name', ''),
            'parent1_last': row.get('Parent 1 Last Name', ''),
            'parent2_first': row.get('Parent 2 First Name', ''),
            'parent2_last': row.get('Parent 2 Last Name', ''),
        }
        players.append(player_info)

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

def create_photo_mapping(photos, players):
    """Create mapping between photos and players"""
    print("Creating photo-to-player mappings...")

    mappings = []
    unmatched_photos = []

    for photo in photos:
        matches = match_photo_to_player(photo['name'], players)

        if matches:
            # Take the best match
            best_match = matches[0]
            mapping = {
                'photo_filename': photo['name'],
                'photo_id': photo['id'],
                'direct_link': photo['directLink'],
                'web_view_link': photo['webViewLink'],
                'thumbnail_link': photo['thumbnailLink'],
                'matched_player_name': best_match['player']['full_name'],
                'matched_first_name': best_match['player']['first_name'],
                'matched_last_name': best_match['player']['last_name'],
                'student_id': best_match['player']['student_id'],
                'confidence': best_match['confidence'],
                'match_type': best_match['match_type'],
                'matched_variation': best_match['matched_variation'],
                'alternative_matches': [
                    f"{m['player']['full_name']} ({m['confidence']})"
                    for m in matches[1:3]  # Show up to 2 alternatives
                ] if len(matches) > 1 else []
            }
            mappings.append(mapping)
        else:
            unmatched_photos.append(photo)

    print(f"Successfully matched {len(mappings)} photos")
    print(f"Could not match {len(unmatched_photos)} photos")

    return mappings, unmatched_photos

def save_results(mappings, unmatched_photos, players):
    """Save results to CSV files"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Save successful mappings
    if mappings:
        mappings_df = pd.DataFrame(mappings)
        mappings_file = f'photo_player_mappings_{timestamp}.csv'
        mappings_df.to_csv(mappings_file, index=False)
        print(f"Saved {len(mappings)} mappings to {mappings_file}")

    # Save unmatched photos
    if unmatched_photos:
        unmatched_df = pd.DataFrame(unmatched_photos)
        unmatched_file = f'unmatched_photos_{timestamp}.csv'
        unmatched_df.to_csv(unmatched_file, index=False)
        print(f"Saved {len(unmatched_photos)} unmatched photos to {unmatched_file}")

    # Save summary report
    summary = {
        'total_photos': len(mappings) + len(unmatched_photos),
        'matched_photos': len(mappings),
        'unmatched_photos': len(unmatched_photos),
        'total_players': len(players),
        'matched_players': len(set(m['student_id'] for m in mappings if m['student_id'])),
        'high_confidence_matches': len([m for m in mappings if m['confidence'] == 'high']),
        'medium_confidence_matches': len([m for m in mappings if m['confidence'] == 'medium']),
        'timestamp': datetime.now().isoformat()
    }

    summary_file = f'mapping_summary_{timestamp}.json'
    with open(summary_file, 'w') as f:
        json.dump(summary, f, indent=2)

    print(f"Saved summary report to {summary_file}")

    # Print summary to console
    print("\n" + "="*60)
    print("PHOTO MAPPING SUMMARY")
    print("="*60)
    print(f"Total photos found: {summary['total_photos']}")
    print(f"Successfully matched: {summary['matched_photos']} ({summary['matched_photos']/summary['total_photos']*100:.1f}%)")
    print(f"High confidence matches: {summary['high_confidence_matches']}")
    print(f"Medium confidence matches: {summary['medium_confidence_matches']}")
    print(f"Unmatched photos: {summary['unmatched_photos']}")
    print(f"Players in roster: {summary['total_players']}")
    print(f"Players with photos: {summary['matched_players']}")

def main():
    print("="*80)
    print("PHOTO TO PLAYER MAPPING TOOL")
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*80)

    # Authenticate with Google Drive
    print("\n1. Authenticating with Google Drive...")
    drive_service = authenticate()
    print("✅ Authentication successful")

    # Extract folder ID from URL
    print("\n2. Extracting folder ID from URL...")
    folder_id = extract_folder_id(GOOGLE_DRIVE_URL)
    print(f"✅ Folder ID: {folder_id}")

    # Get photos from folder
    print("\n3. Fetching photos from Google Drive...")
    photos = get_photos_from_folder(drive_service, folder_id)

    if not photos:
        print("❌ No photos found in the folder. Check folder permissions or ID.")
        sys.exit(1)

    # Load roster data
    print("\n4. Loading roster data...")
    players = load_roster_data(ROSTER_CSV_PATH)

    # Create mappings
    print("\n5. Creating photo-to-player mappings...")
    mappings, unmatched_photos = create_photo_mapping(photos, players)

    # Save results
    print("\n6. Saving results...")
    save_results(mappings, unmatched_photos, players)

    print("\n✅ Photo mapping process completed successfully!")
    print("\nNext steps:")
    print("1. Review the generated CSV files")
    print("2. Manually match any unmatched photos if needed")
    print("3. Use the direct_link column for embedding photos in applications")

if __name__ == "__main__":
    main()