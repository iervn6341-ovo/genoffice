/**
 * Lossless GenOffice → GenOffice clipboard payload.
 *
 * ProseMirror's own clipboard round trip goes through HTML: renderHTML on
 * copy, parseHTML rules on paste. For this schema that loses most of what a
 * table, picture or styled paragraph is made of (cell borders / fills /
 * widths / spans live in attrs no parse rule reads back, nested tables and
 * floating objects have no parse rule at all), and anything that points INTO
 * the source document (style ids, numbering ids, note ids, docx anchors, rIds)
 * dangles in another document.
 *
 * So a copy also carries the slice itself as JSON in a data attribute on the
 * first element of the clipboard HTML (`data-genoffice-slice`). The HTML stays
 * a normal rich fragment for other programs; a GenOffice paste rebuilds the
 * slice from the JSON instead. Everything document-relative is made portable
 * here:
 *  - identity / anchor attrs (docxIndex, bookmarks, comment ranges, raw XML
 *    slices, revision metadata) are stripped — a pasted block is new content;
 *  - styles referenced by the slice travel with their definition and are
 *    imported into the destination when it lacks them (Word's "use
 *    destination styles" rule when a style of that name already exists);
 *  - numbering definitions travel and are mapped onto equal destination
 *    definitions or imported;
 *  - footnote / endnote bodies travel and are re-created in the destination.
 */
import type { Node as PmNode, Schema } from '@tiptap/pm/model'
import { Fragment, Slice } from '@tiptap/pm/model'
import {
  pendingMediaOfDataUrl,
  retargetDrawingToPending,
  type NoteInfo,
  type NumberingDef,
  type StyleInfo,
} from '@genoffice/docx-engine'

export const CLIP_SLICE_ATTR = 'data-genoffice-slice'
const CLIP_VERSION = 1
/** payloads larger than this are not attached (the HTML still pastes, less exactly) */
export const CLIP_PAYLOAD_MAX_CHARS = 24 * 1024 * 1024

type Json = Record<string, unknown>
interface JsonNode {
  type: string
  attrs?: Json
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Json }>
}

export interface ClipStyleDef {
  info: StyleInfo
  /** the exact `<w:style>` element of the source styles.xml, when known */
  xml?: string
}

export interface ClipPayload {
  v: number
  /** random per loaded document: equal keys mean "same document, nothing to import" */
  docKey: string
  slice: { content: JsonNode[]; openStart: number; openEnd: number }
  styles?: Record<string, ClipStyleDef>
  numbering?: Record<string, NumberingDef>
  notes?: { footnote?: Record<string, NoteInfo>; endnote?: Record<string, NoteInfo> }
}

/** what the copying document can tell about itself */
export interface ClipSource {
  docKey: string
  styles?: ReadonlyMap<string, StyleInfo>
  styleXml?: (styleId: string) => string | undefined
  numbering?: ReadonlyMap<string, NumberingDef>
  footnotes?: readonly NoteInfo[]
  endnotes?: readonly NoteInfo[]
  /** original `<w:p>…` XML of a source block (docx anchor → XML), for blocks that cannot be rebuilt from attrs */
  blockXml?: (docxIndex: number) => string | undefined
}

/** what the pasting document offers for importing */
export interface ClipDest {
  docKey: string
  styles?: ReadonlyMap<string, StyleInfo>
  /** add a style the document lacks; returns the id it ended up with */
  importStyle?: (wantedId: string, def: ClipStyleDef) => string
  numbering?: ReadonlyMap<string, NumberingDef>
  /** add a numbering definition (levels of `def`); returns the new numId */
  importNumbering?: (def: NumberingDef) => string | null
  /** add a note with the body of `note`; returns the new note id */
  importNote?: (kind: 'footnote' | 'endnote', note: NoteInfo) => string | null
}

// ── attrs that must not travel ───────────────────────────────────────────

const BLOCK_STRIP = [
  'docxIndex',
  'bookmarks',
  'hiddenBookmarks',
  'commentStarts',
  'commentEnds',
  'sdtShell',
  'moveRevision',
  'pPrChange',
  'paraMarkDel',
  'blockRevision',
  'aiChanged',
]
const STRIP: Record<string, string[]> = {
  docParagraph: BLOCK_STRIP,
  docHeading: BLOCK_STRIP,
  docListItem: BLOCK_STRIP,
  docTable: ['docxIndex', 'originalStructure', 'originalFormatting', 'sdtShell', 'blockRevision'],
  docTableRow: ['rawTrPr', 'rowRevision'],
  docTableCell: ['rawTcPr', 'cellRevision'],
  docTableHeader: ['rawTcPr', 'cellRevision'],
  docInlineImage: ['rawRPr'],
  docNoteRef: ['rawRPr'],
  docProtected: ['docxIndex', 'blockRevision'],
}
const MARK_STRIP: Record<string, string[]> = {
  docTextStyle: ['rawRPr', 'rawRPrW14Prefixes'],
  link: ['rId'],
}
/** annotation marks that never travel */
const DROP_MARKS = new Set(['comment', 'ins', 'rprChange'])

function eachNode(nodes: JsonNode[], visit: (n: JsonNode) => void): void {
  for (const n of nodes) {
    visit(n)
    if (n.content) eachNode(n.content, visit)
  }
}

/** protected block XML that holds relationships (r:embed / r:id / r:link) cannot be cloned into another package */
const XML_HAS_RELS = /\br:(?:embed|id|link|pict)=/

// ── copy side ────────────────────────────────────────────────────────────

function newDocKey(): string {
  return Math.random().toString(36).slice(2) + Date.now().toString(36)
}
const SESSION_KEY = newDocKey()
let generation = 0
/** call when a different document is loaded into this editor */
export function bumpClipDocGeneration(): void {
  generation++
}
export function currentClipDocKey(): string {
  return `${SESSION_KEY}:${generation}`
}

/** style ids the slice refers to (paragraph, run and table styles) */
function collectStyleIds(nodes: JsonNode[]): Set<string> {
  const ids = new Set<string>()
  eachNode(nodes, (n) => {
    const a = n.attrs
    if (a) {
      if (typeof a.styleId === 'string' && a.styleId) ids.add(a.styleId)
      if (typeof a.tblStyleId === 'string' && a.tblStyleId) ids.add(a.tblStyleId)
    }
    for (const m of n.marks ?? []) {
      if (m.type === 'docTextStyle' && typeof m.attrs?.styleId === 'string' && m.attrs.styleId) {
        ids.add(m.attrs.styleId)
      }
    }
  })
  return ids
}

export function buildClipPayload(slice: Slice, source: ClipSource): ClipPayload {
  const raw = slice.content.toJSON() as JsonNode[] | null
  const content: JsonNode[] = JSON.parse(JSON.stringify(raw ?? []))
  const payload: ClipPayload = {
    v: CLIP_VERSION,
    docKey: source.docKey,
    slice: { content, openStart: slice.openStart, openEnd: slice.openEnd },
  }

  // styles (with their basedOn chain) the slice points at
  if (source.styles) {
    const styles: Record<string, ClipStyleDef> = {}
    const add = (id: string, depth = 0): void => {
      if (styles[id] || depth > 12) return
      const info = source.styles!.get(id)
      if (!info) return
      const xml = source.styleXml?.(id)
      styles[id] = { info: JSON.parse(JSON.stringify(info)) as StyleInfo, ...(xml ? { xml } : {}) }
      if (info.basedOn) add(info.basedOn, depth + 1)
      if (info.next) add(info.next, depth + 1)
    }
    for (const id of collectStyleIds(content)) add(id)
    if (Object.keys(styles).length > 0) payload.styles = styles
  }

  // numbering definitions of the list items
  if (source.numbering) {
    const numbering: Record<string, NumberingDef> = {}
    eachNode(content, (n) => {
      const id = n.type === 'docListItem' ? n.attrs?.numId : null
      if (typeof id === 'string' && id && !numbering[id]) {
        const def = source.numbering!.get(id)
        if (def) numbering[id] = JSON.parse(JSON.stringify(def)) as NumberingDef
      }
    })
    if (Object.keys(numbering).length > 0) payload.numbering = numbering
  }

  // note bodies
  const notes: NonNullable<ClipPayload['notes']> = {}
  eachNode(content, (n) => {
    if (n.type !== 'docNoteRef') return
    const kind = n.attrs?.kind === 'endnote' ? 'endnote' : 'footnote'
    const id = String(n.attrs?.id ?? '')
    const list = kind === 'endnote' ? source.endnotes : source.footnotes
    const note = list?.find((item) => item.id === id)
    if (note) (notes[kind] ??= {})[id] = JSON.parse(JSON.stringify(note)) as NoteInfo
  })
  if (notes.footnote || notes.endnote) payload.notes = notes

  sanitizeNodes(content, source)
  return payload
}

/** strip identity attrs / annotation marks; resolve protected blocks that point at the source package */
function sanitizeNodes(nodes: JsonNode[], source: ClipSource): void {
  for (let i = nodes.length - 1; i >= 0; i--) {
    const n = nodes[i]!
    // deleted-by-revision text is not part of the copied document
    if (n.marks?.some((m) => m.type === 'del')) {
      nodes.splice(i, 1)
      continue
    }
    if (n.marks) {
      n.marks = n.marks
        .filter((m) => !DROP_MARKS.has(m.type))
        .map((m) => {
          const strip = MARK_STRIP[m.type]
          if (!strip || !m.attrs) return m
          const attrs = { ...m.attrs }
          for (const k of strip) delete attrs[k]
          return { ...m, attrs }
        })
      if (n.marks.length === 0) delete n.marks
    }
    if (n.type === 'docProtected' && n.attrs) {
      const idx = n.attrs.docxIndex
      // a block that is not an image / generated fragment is rebuilt from its XML
      if (typeof idx === 'number' && !n.attrs.genXml && n.attrs.blockType !== 'image') {
        const xml = source.blockXml?.(idx)
        if (xml && !XML_HAS_RELS.test(xml)) n.attrs.genXml = xml
      }
    }
    const strip = STRIP[n.type]
    if (strip && n.attrs) {
      // a missing attr takes the schema default (null / false), which is what "new content" means
      const attrs = { ...n.attrs }
      for (const k of strip) delete attrs[k]
      n.attrs = attrs
    }
    if (n.content) sanitizeNodes(n.content, source)
  }
}

export function encodeClipPayload(payload: ClipPayload): string | null {
  try {
    const json = JSON.stringify(payload)
    if (json.length > CLIP_PAYLOAD_MAX_CHARS) return null
    // UTF-8 safe base64 (attribute values must survive any clipboard HTML transport)
    const bytes = new TextEncoder().encode(json)
    let bin = ''
    const CHUNK = 0x8000
    for (let i = 0; i < bytes.length; i += CHUNK) {
      bin += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
    }
    return btoa(bin)
  } catch {
    return null
  }
}

export function decodeClipPayload(encoded: string): ClipPayload | null {
  try {
    const bin = atob(encoded)
    const bytes = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
    const parsed = JSON.parse(new TextDecoder().decode(bytes)) as ClipPayload
    if (!parsed || typeof parsed !== 'object' || parsed.v !== CLIP_VERSION) return null
    if (!parsed.slice || !Array.isArray(parsed.slice.content)) return null
    return parsed
  } catch {
    return null
  }
}

/** payload carried by clipboard HTML (searched in the parsed fragment), or null */
export function payloadFromDom(root: ParentNode): ClipPayload | null {
  const el = root.querySelector?.(`[${CLIP_SLICE_ATTR}]`)
  const raw = el?.getAttribute(CLIP_SLICE_ATTR)
  return raw ? decodeClipPayload(raw) : null
}

// ── paste side ───────────────────────────────────────────────────────────

const levelsEqual = (a: NumberingDef, b: NumberingDef): boolean => {
  const ka = Object.keys(a.levels)
  const kb = Object.keys(b.levels)
  if (ka.length !== kb.length) return false
  return ka.every((k) => {
    const x = a.levels[Number(k)]
    const y = b.levels[Number(k)]
    return (
      !!y &&
      x!.numFmt === y.numFmt &&
      x!.lvlText === y.lvlText &&
      (x!.start ?? 1) === (y.start ?? 1)
    )
  })
}

/**
 * Rebuild the copied slice for the destination document: import / map styles,
 * numbering and notes, then materialize the slice. Same-document pastes keep
 * every id as is (the definitions are already there).
 */
export function payloadToSlice(schema: Schema, payload: ClipPayload, dest: ClipDest): Slice {
  const content: JsonNode[] = JSON.parse(JSON.stringify(payload.slice.content))
  const sameDoc = payload.docKey === dest.docKey

  // styles → id map
  const styleMap = new Map<string, string>()
  if (!sameDoc && payload.styles) {
    for (const [id, def] of Object.entries(payload.styles)) {
      const existing = dest.styles?.get(id)
      if (existing) {
        if (
          existing.name.toLowerCase() === def.info.name.toLowerCase() ||
          existing.type !== def.info.type
        ) {
          styleMap.set(id, id) // Word: destination definition wins
          continue
        }
        // same id, different style: import under a fresh id
        if (dest.importStyle) {
          let fresh = `${id}Pasted`
          for (let n = 2; dest.styles?.has(fresh); n++) fresh = `${id}Pasted${n}`
          styleMap.set(id, dest.importStyle(fresh, def))
        }
        continue
      }
      if (dest.importStyle) styleMap.set(id, dest.importStyle(id, def))
    }
  }
  // a style the destination cannot resolve at all falls back to no style
  const known = (id: string): boolean => !!dest.styles?.has(id) || styleMap.has(id)

  // numbering → numId map
  const numMap = new Map<string, string | null>()
  if (!sameDoc && payload.numbering) {
    for (const [id, def] of Object.entries(payload.numbering)) {
      const existing = dest.numbering?.get(id)
      if (existing && levelsEqual(existing, def)) {
        numMap.set(id, id)
        continue
      }
      // an equal definition under another id
      const twin = [...(dest.numbering?.values() ?? [])].find((d) => levelsEqual(d, def))
      if (twin) {
        numMap.set(id, twin.numId)
        continue
      }
      numMap.set(id, dest.importNumbering?.(def) ?? null)
    }
  }

  // notes → new ids
  const noteMap = new Map<string, string | null>()
  const noteKey = (kind: string, id: string) => `${kind}:${id}`
  if (payload.notes) {
    for (const kind of ['footnote', 'endnote'] as const) {
      for (const [id, note] of Object.entries(payload.notes[kind] ?? {})) {
        noteMap.set(noteKey(kind, id), dest.importNote?.(kind, note) ?? null)
      }
    }
  }

  const fixNode = (n: JsonNode): JsonNode | null => {
    if (n.attrs) {
      const a = n.attrs
      if (typeof a.styleId === 'string' && a.styleId && !sameDoc) {
        const mapped = styleMap.get(a.styleId)
        if (mapped !== undefined) a.styleId = mapped
        else if (!known(a.styleId)) a.styleId = null
      }
      if (typeof a.tblStyleId === 'string' && a.tblStyleId && !sameDoc) {
        const mapped = styleMap.get(a.tblStyleId)
        if (mapped !== undefined) a.tblStyleId = mapped
        else if (!known(a.tblStyleId)) a.tblStyleId = null
      }
      if (n.type === 'docListItem' && typeof a.numId === 'string' && !sameDoc) {
        a.numId = numMap.has(a.numId) ? numMap.get(a.numId) : null
      }
      if (n.type === 'docInlineImage' && !sameDoc && !a.rule) {
        // the drawing names a relationship of the SOURCE package: re-point it at its bytes
        const media = typeof a.dataUrl === 'string' ? pendingMediaOfDataUrl(a.dataUrl) : null
        if (!media || typeof a.xml !== 'string' || !a.xml) return null
        a.xml = retargetDrawingToPending(a.xml, media)
      }
      if (n.type === 'docNoteRef') {
        const kind = a.kind === 'endnote' ? 'endnote' : 'footnote'
        const mapped = noteMap.get(noteKey(kind, String(a.id ?? '')))
        // a reference whose note could not be created would dangle: drop it
        if (mapped === null || (mapped === undefined && payload.notes)) return null
        if (mapped) a.id = mapped
      }
    }
    if (n.marks) {
      for (const m of n.marks) {
        if (m.type === 'docTextStyle' && typeof m.attrs?.styleId === 'string' && !sameDoc) {
          const mapped = styleMap.get(m.attrs.styleId)
          if (mapped !== undefined) m.attrs.styleId = mapped
          else if (!known(m.attrs.styleId)) m.attrs.styleId = null
        }
      }
    }
    if (n.content) {
      const kept: JsonNode[] = []
      for (const c of n.content) {
        const fixed = fixNode(c)
        if (fixed) kept.push(fixed)
      }
      n.content = kept
    }
    return n
  }
  const fixed: JsonNode[] = []
  for (const n of content) {
    const f = fixNode(n)
    if (f) fixed.push(f)
  }

  const nodes: PmNode[] = fixed.map((n) => schema.nodeFromJSON(n))
  return new Slice(Fragment.from(nodes), payload.slice.openStart, payload.slice.openEnd)
}
