'use client'

import { useEffect } from 'react'
import { XMarkIcon } from '@heroicons/react/24/outline'

interface ImageModalProps {
  imageUrl: string
  filename: string
  onClose: () => void
}

export default function ImageModal({ imageUrl, filename, onClose }: ImageModalProps) {
  useEffect(() => {
    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose()
      }
    }

    document.addEventListener('keydown', handleEscape)
    return () => document.removeEventListener('keydown', handleEscape)
  }, [onClose])

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-75 modal-overlay"
      onClick={onClose}
    >
      <div className="relative max-w-4xl max-h-full mx-4">
        {/* Close button */}
        <button
          onClick={onClose}
          className="absolute -top-10 right-0 text-white hover:text-gray-300 transition-colors"
          aria-label="Close modal"
        >
          <XMarkIcon className="h-8 w-8" />
        </button>

        {/* Image */}
        <img
          src={imageUrl}
          alt={filename}
          className="max-w-full max-h-screen object-contain rounded-lg shadow-2xl"
          onClick={(e) => e.stopPropagation()}
        />

        {/* Filename */}
        <div className="absolute bottom-0 left-0 right-0 bg-black bg-opacity-75 text-white p-4 rounded-b-lg">
          <p className="text-sm font-mono break-all">{filename}</p>
        </div>
      </div>
    </div>
  )
}