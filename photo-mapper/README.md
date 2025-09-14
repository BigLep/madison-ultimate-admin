# 📸 Photo to Player Mapper

Interactive web application for mapping team photos to player roster using Google Drive and Google Sheets integration.

## 🏗️ Architecture

**Hybrid NextJS Frontend + Python Backend**

- **Frontend**: NextJS 14 with TypeScript, Tailwind CSS, and React components
- **Backend**: Flask API server with Google Drive/Sheets integration
- **Data Sources**: Google Drive photos + Google Sheets roster

### Why Python Backend?

The backend uses Python/Flask instead of Node.js for several reasons:

1. **Legacy Code Reuse**: The original photo mapping algorithm was developed in Python with pandas for data processing and Google API clients
2. **Google API Libraries**: Python has mature, well-documented Google API client libraries (`google-api-python-client`, `gspread`)
3. **Data Processing**: pandas is excellent for CSV/spreadsheet manipulation and the fuzzy matching algorithms
4. **Rapid Prototyping**: The backend was quickly adapted from existing Python scripts

### Technical Debt

- **TODO**: Migrate backend to TypeScript/Node.js for consistency with frontend
- This would provide better type safety across the full stack
- Simpler deployment with a single JavaScript runtime
- Shared utilities and types between frontend/backend

## 🚀 Quick Start

```bash
# From the photo-mapper directory
./start.sh
```

This starts both services:
- Frontend: http://localhost:3000
- Backend: http://localhost:5000
- Main App: http://localhost:3000/map-images-to-players

## 📁 Project Structure

```
photo-mapper/
├── backend/                    # Flask Python API
│   ├── app.py                 # Main API server
│   ├── .google-service-account.json # Credentials
│   └── requirements.txt
├── frontend/                   # NextJS React app
│   ├── app/
│   │   ├── map-images-to-players/
│   │   └── components/
│   └── package.json
├── check_drive_access.py       # Utility to test Drive permissions
├── start.sh                   # Startup script
└── README.md
```

## 🔧 Manual Setup

### Backend
```bash
cd backend
pip3 install -r requirements.txt
python3 app.py
```

### Frontend
```bash
cd frontend
npm install
npm run dev
```

## ✨ Features

- **Automatic Matching**: AI-powered photo-to-player matching using filenames and metadata
- **Interactive Review**: Click-to-zoom photos with editable player assignments
- **Smart Autocomplete**: Type-ahead player selection from roster
- **Confidence Levels**: High/medium confidence indicators for matches
- **Filter & Search**: Focus on uncertain matches that need review
- **CSV Export**: Generate clean mapping file for other applications

## 🔌 Data Sources

- **Photos**: Google Drive folder (53 team photos)
- **Roster**: Google Sheets with 91+ players and parent information
- **Matching**: Uses names, parent names, and filename patterns

## ⚠️ Important Limitation: File Renaming Not Supported

**File renaming functionality is NOT available** unless the service account owns the Google Drive folder and all files within it.

### Why Renaming Doesn't Work

Google Drive API permissions work at the individual file level, not just the folder level. Even if a service account has "Editor" permissions on a folder, it cannot modify files that were uploaded by other users. The service account needs to be the actual owner of each file to rename it.

### Error You'll See

If you try to implement renaming, you'll encounter this error:
```
403 Forbidden: The user has not granted the app write access to the file
```

### Potential Workarounds

1. **Create a new folder**: Have the service account create a new Google Drive folder, then upload photos to that folder (service account will own the files)
2. **Transfer ownership**: Transfer ownership of existing files to the service account (not recommended for shared folders)
3. **Copy/Delete approach**: Copy files to new names and delete originals (requires file ownership)

For most use cases, **CSV export of the final mappings is the recommended approach** instead of renaming files in place.

## 🎯 Workflow

1. **Load Data**: Fetch photos from Drive and roster from Sheets
2. **Auto Match**: Run algorithm to suggest player matches
3. **Review**: User reviews and corrects suggestions
4. **Export**: Download CSV with final PlayerName,DriveId,Filename mapping

## 🔑 Authentication

Uses Google Service Account for API access:
- Google Drive API (read photos)
- Google Sheets API (read roster)
- Credentials stored in `.google-service-account.json`

## 💡 Technologies

**Frontend**:
- NextJS 14, React 18, TypeScript
- Tailwind CSS, Headless UI
- Image modals, autocomplete, filtering

**Backend**:
- Flask, pandas, Google API clients
- Reuses existing photo mapping algorithms
- RESTful API with CORS support