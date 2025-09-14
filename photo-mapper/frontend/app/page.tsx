import Link from 'next/link'

export default function Home() {
  return (
    <main className="flex min-h-screen flex-col items-center justify-center p-24">
      <div className="z-10 max-w-5xl w-full items-center justify-between font-mono text-sm lg:flex">
        <div className="fixed bottom-0 left-0 flex h-48 w-full items-end justify-center bg-gradient-to-t from-white via-white dark:from-black dark:via-black lg:static lg:h-auto lg:w-auto lg:bg-none">
          <div className="text-center">
            <h1 className="text-4xl font-bold mb-4">Photo to Player Mapper</h1>
            <p className="text-lg mb-8">Map team photos to roster players</p>
            <Link
              href="/map-images-to-players"
              className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-3 px-6 rounded-lg text-lg transition-colors"
            >
              Start Mapping Photos
            </Link>
          </div>
        </div>
      </div>
    </main>
  )
}