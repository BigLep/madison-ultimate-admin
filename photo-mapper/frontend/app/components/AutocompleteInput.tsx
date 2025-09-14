'use client'

interface RosterPlayer {
  full_name: string
  first_name: string
  last_name: string
  student_id: string
}

interface SimpleSelectProps {
  roster: RosterPlayer[]
  defaultValue: string
  photoId: string
  filename: string
  onSelect?: (playerName: string) => void
}

export default function SimpleSelect({
  roster,
  defaultValue,
  photoId,
  filename,
  onSelect
}: SimpleSelectProps) {
  const handleChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const selectedValue = e.target.value
    onSelect?.(selectedValue)
  }

  // Sort roster alphabetically by full name
  const sortedRoster = [...roster].sort((a, b) =>
    a.full_name.localeCompare(b.full_name)
  )

  return (
    <select
      data-photo-id={photoId}
      data-filename={filename}
      defaultValue={defaultValue}
      onChange={handleChange}
      className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white text-gray-900"
    >
      <option value="">-- Select Player --</option>
      <option value="Unknown">Unknown</option>
      {sortedRoster.map((player) => (
        <option key={player.student_id || player.full_name} value={player.full_name}>
          {player.full_name}
          {player.student_id && ` (ID: ${player.student_id})`}
        </option>
      ))}
    </select>
  )
}