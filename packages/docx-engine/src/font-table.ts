import type JSZip from 'jszip'
import { parseRels } from './parse-package'
import type { EmbeddedFont, EmbeddedFontRef, EmbeddedFontSlot, FontTableEntry } from './types'
import { attrsOf, childrenOf, findChild, nameOf, xmlParser, type XNode } from './xml-utils'

export const FONT_TABLE_PART_PATH = 'word/fontTable.xml'

/** word/fontTable.xml → substitution hints (altName / PANOSE / family / pitch) */
export function parseFontTable(xml: string): FontTableEntry[] {
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(xml) as XNode[]
  } catch {
    return []
  }
  const root = parsed.find((n) => nameOf(n) === 'w:fonts')
  if (!root) return []
  const out: FontTableEntry[] = []
  for (const node of childrenOf(root)) {
    if (nameOf(node) !== 'w:font') continue
    const name = attrsOf(node)['w:name']
    if (!name) continue
    const val = (tag: string) => attrsOf(findChild(node, tag) ?? {})['w:val'] || undefined
    const entry: FontTableEntry = { name }
    const altName = val('w:altName')
    if (altName) entry.altName = altName
    const panose = val('w:panose1')
    if (panose) entry.panose = panose
    const family = val('w:family')
    if (family) entry.family = family
    const charset = val('w:charset')
    if (charset) entry.charset = charset
    const csb0 = attrsOf(findChild(node, 'w:sig') ?? {})['w:csb0']
    if (csb0) entry.csb0 = csb0
    const pitch = val('w:pitch')
    if (pitch) entry.pitch = pitch
    const embedded = parseEmbedSlots(node)
    if (embedded) entry.embedded = embedded
    out.push(entry)
  }
  return out
}

const EMBED_SLOTS: ReadonlyArray<[tag: string, slot: EmbeddedFontSlot]> = [
  ['w:embedRegular', 'regular'],
  ['w:embedBold', 'bold'],
  ['w:embedItalic', 'italic'],
  ['w:embedBoldItalic', 'boldItalic'],
]

function parseEmbedSlots(font: XNode): FontTableEntry['embedded'] | undefined {
  let out: FontTableEntry['embedded']
  for (const [tag, slot] of EMBED_SLOTS) {
    const attrs = attrsOf(findChild(font, tag) ?? {})
    const rId = attrs['r:id']
    if (!rId) continue
    const ref: EmbeddedFontRef = { rId }
    if (attrs['w:fontKey']) ref.fontKey = attrs['w:fontKey']
    ;(out ??= {})[slot] = ref
  }
  return out
}

const SFNT_MAGICS = new Set([0x00010000, 0x4f54544f, 0x74727565])

export function isSfnt(bytes: Uint8Array): boolean {
  if (bytes.length < 12) return false
  const magic = ((bytes[0] << 24) | (bytes[1] << 16) | (bytes[2] << 8) | bytes[3]) >>> 0
  return SFNT_MAGICS.has(magic)
}

/**
 * Test whether the SFNT font file's OS/2 table explicitly declares a non-bold weight (< 600).
 * Word embeds regular fonts under w:embedBold for fonts that lack a real bold variant (e.g. 標楷體).
 * Calling code must not register such faces with descriptor weight: '700', or the browser
 * will treat the face as an exact match and suppress synthetic (faux) bolding.
 */
export function isSfntNonBold(bytes: Uint8Array): boolean {
  if (bytes.length < 12) return false
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const numTables = view.getUint16(4)
  if (bytes.length < 12 + numTables * 16) return false
  for (let i = 0; i < numTables; i++) {
    const o = 12 + i * 16
    if (view.getUint32(o) === 0x4f532f32) {
      // 'OS/2'
      const tableOffset = view.getUint32(o + 8)
      if (tableOffset + 64 <= bytes.length) {
        const usWeightClass = view.getUint16(tableOffset + 4)
        const fsSelection = view.getUint16(tableOffset + 62)
        const isBoldSelection = (fsSelection & 0x0020) !== 0
        if (!isBoldSelection && usWeightClass > 0 && usWeightClass < 600) return true
      }
      return false
    }
  }
  return false
}

export function isSfntNonItalic(bytes: Uint8Array): boolean {
  if (bytes.length < 12) return false
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const numTables = view.getUint16(4)
  if (bytes.length < 12 + numTables * 16) return false
  for (let i = 0; i < numTables; i++) {
    const o = 12 + i * 16
    if (view.getUint32(o) === 0x4f532f32) {
      // 'OS/2'
      const tableOffset = view.getUint32(o + 8)
      if (tableOffset + 64 <= bytes.length) {
        const fsSelection = view.getUint16(tableOffset + 62)
        const isItalicSelection = (fsSelection & 0x0001) !== 0
        if (!isItalicSelection) return true
      }
      return false
    }
  }
  return false
}

/**
 * Undo the ECMA-376 17.8.1 obfuscation: the first 32 bytes are XORed with the
 * fontKey GUID's 16 bytes taken in reverse order. A missing/zero key or a part
 * that already carries an sfnt magic is returned as is.
 */
export function deobfuscateOdttf(bytes: Uint8Array, fontKey: string | undefined): Uint8Array {
  const hex = fontKey?.replace(/[^0-9a-f]/gi, '') ?? ''
  if (hex.length !== 32 || /^0+$/.test(hex) || isSfnt(bytes)) return bytes
  const key = new Uint8Array(16)
  for (let i = 0; i < 16; i++) key[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16)
  const out = bytes.slice()
  const n = Math.min(32, out.length)
  for (let i = 0; i < n; i++) out[i] ^= key[15 - (i % 16)]
  return out
}

export const FONT_TABLE_RELS_PATH = 'word/_rels/fontTable.xml.rels'

/** fontTable embed slots → de-obfuscated faces; unparseable parts are skipped */
export async function readEmbeddedFonts(
  zip: JSZip,
  fontTable: readonly FontTableEntry[],
): Promise<EmbeddedFont[]> {
  if (!fontTable.some((e) => e.embedded)) return []
  const rels = await parseRels(zip, FONT_TABLE_RELS_PATH)
  const out: EmbeddedFont[] = []
  for (const entry of fontTable) {
    for (const [slot, ref] of Object.entries(entry.embedded ?? {}) as Array<
      [EmbeddedFontSlot, EmbeddedFontRef]
    >) {
      const rel = rels.get(ref.rId)
      if (!rel || rel.targetMode === 'External') continue
      const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`
      const file = zip.file(path)
      if (!file) continue
      const data = deobfuscateOdttf(await file.async('uint8array'), ref.fontKey)
      if (!isSfnt(data)) continue
      const isBoldSlot = slot === 'bold' || slot === 'boldItalic'
      const isItalicSlot = slot === 'italic' || slot === 'boldItalic'
      const bold = isBoldSlot && !isSfntNonBold(data)
      const italic = isItalicSlot && !isSfntNonItalic(data)
      const families = [entry.name]
      if (entry.altName && entry.altName !== entry.name) families.push(entry.altName)
      for (const family of families) {
        out.push({
          family,
          bold,
          italic,
          data,
        })
      }
    }
  }
  return out
}
