'use client'

import { useState, useEffect } from 'react'
import PhotoMappingTable from '../components/PhotoMappingTable'
import LoadingSpinner from '../components/LoadingSpinner'

interface PhotoMapping {
  photo_id: string
  filename: string
  thumbnail_url: string
  direct_link: string
  matched_player: string
  confidence: 'high' | 'medium'
  match_type: string
  alternative_matches: string[]
  student_id: string
}

interface RosterPlayer {
  full_name: string
  first_name: string
  last_name: string
  student_id: string
}

export default function MapImagesToPlayers() {
  const [mappings, setMappings] = useState<PhotoMapping[]>([])
  const [roster, setRoster] = useState<RosterPlayer[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    loadData()
  }, [])

  const loadData = async () => {
    try {
      setLoading(true)
      setError(null)

      const response = await fetch('/api/load-data')
      if (!response.ok) {
        throw new Error(`Failed to load data: ${response.statusText}`)
      }

      const data = await response.json()
      setMappings(data.mappings)
      setRoster(data.roster)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred')
      console.error('Error loading data:', err)
    } finally {
      setLoading(false)
    }
  }

  const handleExportCSV = () => {
    // Create CSV content
    const csvData = [['PlayerName', 'DriveId', 'Filename', 'GoogleDriveLink', 'DirectImageLink', 'ThumbnailLink']]

    mappings.forEach(mapping => {
      const selectElement = document.querySelector(`select[data-photo-id="${mapping.photo_id}"]`) as HTMLSelectElement

      if (selectElement?.value && selectElement.value !== '') {
        const driveViewLink = `https://drive.google.com/file/d/${mapping.photo_id}/view`
        const directImageLink = `https://drive.google.com/uc?id=${mapping.photo_id}`
        csvData.push([selectElement.value, mapping.photo_id, mapping.filename, driveViewLink, directImageLink, mapping.thumbnail_url])
      }
    })

    if (csvData.length === 1) {
      alert('Please assign at least one player name before exporting.')
      return
    }

    // Convert to CSV string
    const csvContent = csvData.map(row =>
      row.map(field => `"${field.replace(/"/g, '""')}"`).join(',')
    ).join('\n')

    // Download CSV
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
    const link = document.createElement('a')
    const url = URL.createObjectURL(blob)
    link.setAttribute('href', url)
    link.setAttribute('download', `photo_player_mapping_${new Date().toISOString().slice(0,10)}.csv`)
    link.style.visibility = 'hidden'
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
  }

  const handleRenameFiles = async () => {
    const renames: Array<{photo_id: string, current_name: string, new_name: string}> = []

    console.log('Starting rename process, checking', mappings.length, 'mappings')

    mappings.forEach(mapping => {
      const selectElement = document.querySelector(`select[data-photo-id="${mapping.photo_id}"]`) as HTMLSelectElement
      const excludeCheckbox = document.querySelector(`input[data-photo-id="${mapping.photo_id}-exclude"]`) as HTMLInputElement
      const selectedPlayerName = selectElement?.value

      console.log(`Checking ${mapping.filename}:`, {
        selectedPlayer: selectedPlayerName,
        excluded: excludeCheckbox?.checked,
        currentFilename: mapping.filename
      })

      // Skip if excluded
      if (excludeCheckbox?.checked) {
        console.log(`  -> Skipping ${mapping.filename} (excluded)`)
        return
      }

      if (selectedPlayerName && selectedPlayerName !== '') {
        // Check if the current filename (without extension) is different from selected name
        const currentStem = mapping.filename.split('.').slice(0, -1).join('.')
        console.log(`  -> Comparing "${currentStem}" vs "${selectedPlayerName}"`)
        if (currentStem.toLowerCase() !== selectedPlayerName.toLowerCase()) {
          console.log(`  -> Adding to rename list: ${mapping.filename} -> ${selectedPlayerName}`)
          renames.push({
            photo_id: mapping.photo_id,
            current_name: mapping.filename,
            new_name: selectedPlayerName
          })
        } else {
          console.log(`  -> Names match, skipping`)
        }
      } else {
        console.log(`  -> No player selected, skipping`)
      }
    })

    console.log('Final renames list:', renames)

    if (renames.length === 0) {
      alert('No files need to be renamed. Either all selected names match current filenames or files are excluded from renaming.')
      return
    }

    const confirmMessage = `This will rename ${renames.length} file(s) in Google Drive to match the selected player names. This action cannot be undone. Continue?`
    if (!confirm(confirmMessage)) {
      return
    }

    try {
      setLoading(true)
      const response = await fetch('/api/rename-files', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ renames })
      })

      if (!response.ok) {
        throw new Error(`Failed to rename files: ${response.statusText}`)
      }

      const result = await response.json()
      alert(`Rename operation completed!\n\nSuccessful: ${result.successful}\nSkipped: ${result.skipped}\nFailed: ${result.failed}\n\nTotal processed: ${result.total}`)

      // Reload data to show updated filenames
      await loadData()
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred during renaming')
      console.error('Error renaming files:', err)
    } finally {
      setLoading(false)
    }
  }

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <LoadingSpinner />
      </div>
    )
  }

  if (error) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="text-center">
          <h1 className="text-2xl font-bold text-red-600 mb-4">Error Loading Data</h1>
          <p className="text-gray-600 mb-4">{error}</p>
          <button
            onClick={loadData}
            className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded"
          >
            Retry
          </button>
        </div>
      </div>
    )
  }

  const highConfidenceCount = mappings.filter(m => m.confidence === 'high').length
  const mediumConfidenceCount = mappings.filter(m => m.confidence === 'medium').length

  return (
    <div className="min-h-screen bg-gray-50 py-8">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        {/* Header */}
        <div className="bg-gradient-to-r from-blue-600 to-purple-600 text-white rounded-lg p-6 mb-6">
          <h1 className="text-3xl font-bold mb-2">📸 Photo to Player Mapper</h1>
          <div className="flex flex-wrap gap-6 text-sm">
            <span>Total Photos: {mappings.length}</span>
            <span className="text-green-200">High Confidence: {highConfidenceCount}</span>
            <span className="text-yellow-200">Medium Confidence: {mediumConfidenceCount}</span>
          </div>
          <p className="mt-2 text-blue-100">
            Review the suggested matches and adjust player names as needed.
            Use autocomplete to select from the roster.
          </p>
        </div>

        {/* Controls */}
        <div className="bg-white rounded-lg shadow-sm p-4 mb-6 flex flex-wrap items-center gap-4">
          <button
            onClick={handleExportCSV}
            className="bg-green-600 hover:bg-green-700 text-white font-semibold py-2 px-4 rounded-lg transition-colors"
          >
            📥 Export CSV Mapping
          </button>
          <div className="text-sm text-gray-600">
            Export the final player-to-photo mappings as a CSV file.
          </div>
        </div>

        {/* Photo Mapping Table */}
        <PhotoMappingTable mappings={mappings} roster={roster} />
      </div>
    </div>
  )
}