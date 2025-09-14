# Photo Mapper Backend

Flask API backend for the photo-to-player mapping application.

## Files

- `app.py` - Main Flask application with API endpoints
- `photo_player_mapping.py` - Original standalone mapping script
- `check_drive_access.py` - Script to verify Google Drive access
- `.google-service-account.json` - Google service account credentials
- `requirements.txt` - Python dependencies

## API Endpoints

- `GET /api/load-data` - Load photos from Drive, roster from Sheets, and create mappings
- `POST /api/export-csv` - Export final mappings as CSV data
- `GET /health` - Health check endpoint

## Running

```bash
# Install dependencies
pip3 install -r requirements.txt

# Start server
python3 app.py
```

Server runs on http://localhost:5000

## Dependencies

- Flask web framework
- Google API clients (Drive, Sheets)
- pandas for data processing
- CORS support for frontend integration