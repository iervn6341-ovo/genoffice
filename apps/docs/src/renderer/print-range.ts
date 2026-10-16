/**
 * Parse a Word-style print range ("1,3,5-8") into sorted 0-based page indices.
 * Accepts CJK/ASCII separators and dashes: hyphen, en-dash, em-dash, tilde,
 * fullwidth tilde and wave dash (Word IME users often type ～/〜 instead of -);
 * returns null on any invalid or out-of-bounds part (same semantics as the slides print dialog).
 */
export function parsePrintRange(text: string, max: number): number[] | null {
  const out = new Set<number>()
  const tokenRe = /\d+\s*[-–—~～〜]\s*\d+|\d+/g
  const parts = text.match(tokenRe) ?? []
  if (parts.length === 0) return null
  if (text.replace(tokenRe, '').replace(/[,，、;；\s]/g, '') !== '') return null
  for (const part of parts) {
    const m = /^(\d+)\s*[-–—~～〜]\s*(\d+)$|^(\d+)$/.exec(part)
    if (!m) return null
    const a = Number(m[1] ?? m[3])
    const b = Number(m[2] ?? m[3])
    if (a < 1 || b > max || a > b) return null
    for (let i = a; i <= b; i++) out.add(i - 1)
  }
  return [...out].sort((x, y) => x - y)
}

/** What the range syntax needs to know about each preview page: its 0-based section index */
export interface PrintPageInfo {
  section: number
}

const DASH = '-–—~～〜'
const POINT = 'p?\\d+(?:s\\d+)?|s\\d+'
const SPAN_RE = new RegExp(`^(${POINT})(?:\\s*[${DASH}]\\s*(${POINT}))?$`, 'i')

/**
 * Word's full custom-range syntax over the preview pages: plain pages and ranges ("1,3,5-8"), a
 * page within a section ("p2s3" = 2nd page of section 3), a whole section ("s2") and ranges between
 * any two of those ("p2s1-p1s3", "s2-s4", "3-s2"). Section numbers count from 1 in document order.
 * Returns sorted 0-based page indices, or null on any invalid or out-of-bounds part.
 */
export function parsePrintRangeSections(
  text: string,
  pages: readonly PrintPageInfo[],
): number[] | null {
  const parts = text
    .split(/[,，、;；]/)
    .map((x) => x.trim())
    .filter((x) => x !== '')
  if (parts.length === 0) return null
  // page indices of each section (1-based key), in document order
  const bySection = new Map<number, number[]>()
  pages.forEach((p, i) => {
    const list = bySection.get(p.section + 1) ?? []
    list.push(i)
    bySection.set(p.section + 1, list)
  })
  const resolve = (token: string, asEnd: boolean): number | null => {
    const m = /^(?:p?(\d+))?(?:s(\d+))?$/i.exec(token.trim())
    if (!m || (m[1] === undefined && m[2] === undefined)) return null
    const page = m[1] === undefined ? null : Number(m[1])
    if (m[2] === undefined)
      return page !== null && page >= 1 && page <= pages.length ? page - 1 : null
    const list = bySection.get(Number(m[2]))
    if (!list) return null
    if (page === null) return (asEnd ? list[list.length - 1] : list[0]) ?? null
    return page >= 1 && page <= list.length ? (list[page - 1] ?? null) : null
  }
  const out = new Set<number>()
  for (const part of parts) {
    const m = SPAN_RE.exec(part)
    if (!m) return null
    const a = resolve(m[1]!, false)
    const b = m[2] === undefined ? resolve(m[1]!, true) : resolve(m[2], true)
    if (a === null || b === null || a > b) return null
    for (let i = a; i <= b; i++) out.add(i)
  }
  return [...out].sort((x, y) => x - y)
}
