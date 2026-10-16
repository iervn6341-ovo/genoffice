import type { PackageArchive } from './zip'

const PRESENTATION_PATH = 'ppt/presentation.xml'

/** The first number shown by dynamic slidenum fields; OOXML defaults to 1. */
export function getFirstSlideNumber(archive: PackageArchive): number {
  const xml = archive.readText(PRESENTATION_PATH)
  const open = xml?.match(/<p:presentation\b[^>]*>/)?.[0]
  const raw = open?.match(/\bfirstSlideNum=["'](-?\d+)["']/)?.[1]
  const value = raw == null ? 1 : Number(raw)
  return Number.isSafeInteger(value) ? value : 1
}

/** Patch only p:presentation@firstSlideNum, leaving slide IDs and package relationships intact. */
export function setFirstSlideNumber(archive: PackageArchive, value: number): boolean {
  if (!Number.isSafeInteger(value) || value < -9999 || value > 9999)
    throw new Error('Starting slide number must be an integer from -9999 to 9999')
  const xml = archive.readText(PRESENTATION_PATH)
  const open = xml?.match(/<p:presentation\b[^>]*>/)?.[0]
  if (!xml || !open) throw new Error('Presentation part is missing')
  if (getFirstSlideNumber(archive) === value) return false
  const nextOpen = /\bfirstSlideNum=["'][^"']*["']/.test(open)
    ? open.replace(/\bfirstSlideNum=["'][^"']*["']/, `firstSlideNum="${value}"`)
    : open.replace(/>$/, ` firstSlideNum="${value}">`)
  archive.entries.set(
    PRESENTATION_PATH,
    Buffer.from(
      xml.replace(open, () => nextOpen),
      'utf8',
    ),
  )
  return true
}
