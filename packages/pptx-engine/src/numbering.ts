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

/** A native DrawingML field with a cached value, rather than PowerPoint's global slidenum. */
export const OFFSET_SLIDE_NUMBER_FIELD = 'genofficeSlideNumber'
export interface SlideNumberingRule {
  startSlide: number
  startNumber: number
  hideBefore: boolean
  enabled: boolean
  excludeTitleSlide?: boolean
}
const RULE_URI = '{581A1EBB-85C7-4B55-A5FD-43AA26531141}'
const RULE_RE = /<p:ext\b[^>]*uri="\{581A1EBB-85C7-4B55-A5FD-43AA26531141\}"[^>]*>[\s\S]*?<\/p:ext>/
export function getSlideNumberingRule(archive: PackageArchive): SlideNumberingRule | null {
  const xml = archive.readText(PRESENTATION_PATH) ?? ''
  const tag = xml.match(RULE_RE)?.[0].match(/<go:slideNumbering\b[^>]*\/>/)?.[0]
  if (!tag) return null
  const attr = (name: string) => tag.match(new RegExp(`\\b${name}="([^" ]*)"`))?.[1]
  const rule = {
    startSlide: Number(attr('startSlide')),
    startNumber: Number(attr('startNumber')),
    hideBefore: attr('hideBefore') === '1',
    enabled: attr('enabled') === '1',
    ...(attr('excludeTitleSlide') === '1' ? { excludeTitleSlide: true } : {}),
  }
  try {
    validateSlideNumberingRule(rule)
    return rule
  } catch {
    return null
  }
}
export function validateSlideNumberingRule(rule: SlideNumberingRule): void {
  if (
    !rule ||
    !Number.isInteger(rule.startSlide) ||
    rule.startSlide < 1 ||
    rule.startSlide > 99999 ||
    !Number.isInteger(rule.startNumber) ||
    rule.startNumber < 0 ||
    rule.startNumber > 9999 ||
    typeof rule.hideBefore !== 'boolean' ||
    typeof rule.enabled !== 'boolean' ||
    (rule.excludeTitleSlide !== undefined && typeof rule.excludeTitleSlide !== 'boolean')
  )
    throw new Error('Invalid slide numbering rule')
}
export function setSlideNumberingRule(
  archive: PackageArchive,
  rule: SlideNumberingRule | null,
): boolean {
  if (rule) validateSlideNumberingRule(rule)
  const xml = archive.readText(PRESENTATION_PATH)
  if (!xml) throw new Error('Presentation part is missing')
  let next = xml.replace(RULE_RE, '')
  if (rule) {
    const ext = `<p:ext uri="${RULE_URI}"><go:slideNumbering xmlns:go="urn:genoffice:slide-numbering" startSlide="${rule.startSlide}" startNumber="${rule.startNumber}" hideBefore="${rule.hideBefore ? 1 : 0}" enabled="${rule.enabled ? 1 : 0}"${rule.excludeTitleSlide ? ' excludeTitleSlide="1"' : ''}/></p:ext>`
    if (/<p:extLst\b[^>]*\/>/.test(next))
      next = next.replace(/<p:extLst\b[^>]*\/>/, '<p:extLst></p:extLst>')
    next = next.includes('</p:extLst>')
      ? next.replace('</p:extLst>', ext + '</p:extLst>')
      : next.replace('</p:presentation>', `<p:extLst>${ext}</p:extLst></p:presentation>`)
  }
  if (next === xml) return false
  archive.entries.set(PRESENTATION_PATH, Buffer.from(next, 'utf8'))
  return true
}
