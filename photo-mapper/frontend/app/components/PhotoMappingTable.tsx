'use client'

import { useState } from 'react'
import SimpleSelect from './AutocompleteInput'
import ImageModal from './ImageModal'

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

interface PhotoMappingTableProps {
  mappings: PhotoMapping[]
  roster: RosterPlayer[]
}

type FilterType = 'all' | 'high' | 'medium'

export default function PhotoMappingTable({ mappings, roster }: PhotoMappingTableProps) {
  const [filter, setFilter] = useState<FilterType>('all')
  const [selectedImage, setSelectedImage] = useState<{ url: string; filename: string } | null>(null)

  const filteredMappings = mappings.filter(mapping => {
    if (filter === 'all') return true
    return mapping.confidence === filter
  })

  const getConfidenceBadge = (confidence: string) => {
    if (confidence === 'high') {
      return (
        <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">
          High
        </span>
      )
    }
    return (
      <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-yellow-100 text-yellow-800">
        Medium
      </span>
    )
  }

  const handleSelectAllExclude = (exclude: boolean) => {
    filteredMappings.forEach(mapping => {
      const checkbox = document.querySelector(`input[data-photo-id="${mapping.photo_id}-exclude"]`) as HTMLInputElement
      if (checkbox) {
        checkbox.checked = exclude
      }
    })
  }

  return (
    <div>
      {/* Filter Buttons */}
      <div className="mb-4 flex gap-2">
        <button
          onClick={() => setFilter('all')}
          className={`px-4 py-2 rounded-md font-medium ${
            filter === 'all'
              ? 'bg-blue-600 text-white'
              : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-50'
          }`}
        >
          All ({mappings.length})
        </button>
        <button
          onClick={() => setFilter('high')}
          className={`px-4 py-2 rounded-md font-medium ${
            filter === 'high'
              ? 'bg-green-600 text-white'
              : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-50'
          }`}
        >
          High Confidence ({mappings.filter(m => m.confidence === 'high').length})
        </button>
        <button
          onClick={() => setFilter('medium')}
          className={`px-4 py-2 rounded-md font-medium ${
            filter === 'medium'
              ? 'bg-yellow-600 text-white'
              : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-50'
          }`}
        >
          Medium Confidence ({mappings.filter(m => m.confidence === 'medium').length})
        </button>
      </div>


      {/* Table */}
      <div className="bg-white shadow-sm rounded-lg overflow-hidden">
        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Photo
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Filename
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Suggested Match
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Player Name
                </th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {filteredMappings.map((mapping, index) => (
                <tr
                  key={mapping.photo_id}
                  className={`${
                    mapping.confidence === 'high' ? 'bg-green-50' : 'bg-yellow-50'
                  } hover:bg-opacity-75 transition-colors`}
                >
                  {/* Photo */}
                  <td className="px-6 py-4">
                    <div className="flex-shrink-0 h-20 w-20">
                      <img
                        className="h-20 w-20 rounded-lg object-cover cursor-pointer hover:opacity-75 transition-opacity border-2 border-gray-200"
                        src={mapping.thumbnail_url}
                        alt={mapping.filename}
                        onClick={() => setSelectedImage({
                          url: mapping.direct_link,
                          filename: mapping.filename
                        })}
                        title="Click to view full size"
                      />
                    </div>
                  </td>

                  {/* Filename */}
                  <td className="px-6 py-4">
                    <div className="text-sm text-gray-900 font-mono break-all max-w-xs">
                      {mapping.filename}
                    </div>
                  </td>

                  {/* Suggested Match */}
                  <td className="px-6 py-4">
                    <div className="space-y-2">
                      <div className="flex items-center space-x-2">
                        <span className="text-sm font-medium text-gray-900">
                          {mapping.matched_player}
                        </span>
                        {getConfidenceBadge(mapping.confidence)}
                      </div>
                      <div className="text-xs text-gray-500">
                        Match type: {mapping.match_type}
                      </div>
                      {mapping.alternative_matches && mapping.alternative_matches.length > 0 && (
                        <div className="text-xs text-gray-500">
                          <span className="font-medium">Alternatives:</span>{' '}
                          {mapping.alternative_matches.slice(0, 2).join(', ')}
                        </div>
                      )}
                    </div>
                  </td>

                  {/* Player Name Input */}
                  <td className="px-6 py-4">
                    <div className="w-64">
                      <SimpleSelect
                        roster={roster}
                        defaultValue={mapping.matched_player}
                        photoId={mapping.photo_id}
                        filename={mapping.filename}
                      />
                    </div>
                  </td>

                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {filteredMappings.length === 0 && (
          <div className="text-center py-12">
            <div className="text-gray-500">
              No mappings match the current filter.
            </div>
          </div>
        )}
      </div>

      {/* Image Modal */}
      {selectedImage && (
        <ImageModal
          imageUrl={selectedImage.url}
          filename={selectedImage.filename}
          onClose={() => setSelectedImage(null)}
        />
      )}
    </div>
  )
}