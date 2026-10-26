/**
 * The App-side half of the clipboard transfer: what the open document tells a
 * copy about itself, and how a paste imports styles / numbering into it.
 * Kept free of React so it is unit-testable; App.tsx hands it live accessors.
 */
import type { Transaction } from '@tiptap/pm/state'
import type {
  CustomNumberingLevel,
  NoteInfo,
  NumberingDef,
  ParsedDocFull,
  StyleInfo,
  StyleUpsert,
} from '@genoffice/docx-engine'
import {
  currentClipDocKey,
  type ClipDest,
  type ClipSource,
  type ClipStyleDef,
} from './clipboard-transfer'
import type { ClipHost } from './clipboard-extension'

export interface ClipHostAccess {
  parsed(): ParsedDocFull | null
  /** record a style definition for the save (styleUpserts) */
  addStyleUpsert(up: StyleUpsert): void
  /** the document stylesheet must be rebuilt after styles were added */
  styleSheetChanged(): void
  /** create a numbering definition from custom levels; returns the numId */
  createListDef(levels: CustomNumberingLevel[]): string | null
  /** document numbering table, including pending (not yet saved) definitions */
  numbering(): ReadonlyMap<string, NumberingDef>
  /** create a note with `note`'s body; returns its id (null when notes cannot be added) */
  addNote?(kind: 'footnote' | 'endnote', note: NoteInfo): string | null
  /** renumber note markers after a paste that brought notes */
  normalizeNotes?(tr: Transaction, kind: 'footnote' | 'endnote'): Transaction
}

/** a styles.xml element renamed to another styleId (collision import) */
export function withStyleId(xml: string, styleId: string): string {
  return xml.replace(/(<w:style\b[^>]*?\bw:styleId=")[^"]*(")/, `$1${styleId}$2`)
}

/** best-effort upsert from the display hints (when the source styles.xml has no element for it) */
export function upsertFromStyleInfo(info: StyleInfo, styleId: string): StyleUpsert {
  const d = info.display ?? {}
  const up: StyleUpsert = {
    styleId,
    type: info.type,
    name: info.name,
    basedOn: info.basedOn ?? null,
    next: info.next ?? null,
    quickFormat: info.qFormat ?? true,
  }
  const rPr: NonNullable<StyleUpsert['rPr']> = {}
  if (d.bold !== undefined) rPr.bold = d.bold
  if (d.italic !== undefined) rPr.italic = d.italic
  if (d.underline !== undefined) rPr.underline = d.underline
  if (d.strike !== undefined) rPr.strike = d.strike
  if (d.color && d.color !== 'auto') rPr.color = d.color
  if (d.sizeHalfPoints != null) rPr.sizeHalfPoints = d.sizeHalfPoints
  if (d.fontAscii ?? d.font) rPr.font = d.fontAscii ?? d.font
  if (Object.keys(rPr).length > 0) up.rPr = rPr
  const pPr: NonNullable<StyleUpsert['pPr']> = {}
  if (d.align === 'left' || d.align === 'center' || d.align === 'right' || d.align === 'justify')
    pPr.align = d.align
  if (d.spaceBeforeTwips != null) pPr.spaceBeforeTwips = d.spaceBeforeTwips
  if (d.spaceAfterTwips != null) pPr.spaceAfterTwips = d.spaceAfterTwips
  if (d.lineSpacing != null) pPr.lineSpacing = d.lineSpacing
  if (d.indentLeftTwips != null) pPr.indentLeftTwips = d.indentLeftTwips
  if (d.indentRightTwips != null) pPr.indentRightTwips = d.indentRightTwips
  if (d.indentFirstLineTwips != null) pPr.firstLineTwips = d.indentFirstLineTwips
  if (d.keepNext !== undefined) pPr.keepNext = d.keepNext
  if (d.keepLines !== undefined) pPr.keepLines = d.keepLines
  if (Object.keys(pPr).length > 0) up.pPr = pPr
  return up
}

export function createClipHost(access: ClipHostAccess): ClipHost {
  /** kinds of notes the current paste created: their markers need renumbering once it lands */
  const importedNoteKinds = new Set<'footnote' | 'endnote'>()
  const source = (): ClipSource => {
    const parsed = access.parsed()
    return {
      docKey: currentClipDocKey(),
      styles: parsed?.styles,
      styleXml: (id) => parsed?.extras?.styleXml?.[id],
      numbering: access.numbering(),
      footnotes: parsed?.footnotes,
      endnotes: parsed?.endnotes,
      blockXml: (i) => parsed?.blocks.find((b) => b.docxIndex === i)?.originalXml ?? undefined,
    }
  }
  const dest = (): ClipDest => {
    const parsed = access.parsed()
    return {
      docKey: currentClipDocKey(),
      styles: parsed?.styles,
      importStyle: (wantedId: string, def: ClipStyleDef): string => {
        if (!parsed) return wantedId
        const info: StyleInfo = { ...def.info, styleId: wantedId }
        parsed.styles.set(wantedId, info)
        const up: StyleUpsert = def.xml
          ? { styleId: wantedId, type: def.info.type, rawXml: withStyleId(def.xml, wantedId) }
          : upsertFromStyleInfo(def.info, wantedId)
        access.addStyleUpsert(up)
        access.styleSheetChanged()
        return wantedId
      },
      importNote: (kind, note) => {
        const id = access.addNote?.(kind, note) ?? null
        if (id) importedNoteKinds.add(kind)
        return id
      },
      numbering: access.numbering(),
      importNumbering: (def: NumberingDef): string | null => {
        const levels = Object.keys(def.levels)
          .map(Number)
          .sort((a, b) => a - b)
          .map((i): CustomNumberingLevel => {
            const l = def.levels[i]!
            return {
              numFmt: l.numFmt,
              lvlText: l.lvlText,
              indentLeft: l.indentLeft ?? 720 * (i + 1),
              hanging: l.hanging ?? 360,
              start: l.start ?? 1,
              ...(l.font ? { font: l.font } : {}),
              ...(l.color ? { color: l.color } : {}),
              ...(l.suff ? { suff: l.suff } : {}),
              ...(l.lvlJc ? { lvlJc: l.lvlJc } : {}),
              ...(l.isLgl ? { isLgl: true } : {}),
            }
          })
        return levels.length > 0 ? access.createListDef(levels) : null
      },
    }
  }
  const afterPaste = (tr: Transaction): Transaction | null => {
    if (importedNoteKinds.size === 0 || !access.normalizeNotes) return null
    let out = tr
    for (const kind of importedNoteKinds) out = access.normalizeNotes(out, kind)
    importedNoteKinds.clear()
    return out.docChanged ? out : null
  }
  return { source, dest, afterPaste }
}
