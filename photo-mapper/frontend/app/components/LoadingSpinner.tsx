export default function LoadingSpinner() {
  return (
    <div className="flex flex-col items-center justify-center space-y-4">
      <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
      <div className="text-center">
        <h2 className="text-xl font-semibold text-gray-700">Loading Photos and Roster</h2>
        <p className="text-gray-500 mt-2">
          Fetching photos from Google Drive and running matching algorithm...
        </p>
      </div>
    </div>
  )
}