import { decodeEntities } from './parse-xml-text'
import { firstNoteMarkRun, noteCustomMarkFromRun, noteCustomMarkXml } from './note-custom-mark'
import { patchParagraphTexts } from './text-patch'
import type { NoteInfo, NoteRun, TextOutline } from './types'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

/**
 * Footnotes / endnotes part handling (word/footnotes.xml, word/endnotes.xml).
 *
 * Both parts share one schema: a list of w:footnote / w:endnote elements.
 * Entries carrying a w:type attribute (separator, continuationSeparator, ...)
 * are structural and must be preserved; only typeless entries are real notes.
 */

export type NoteKind = 'footnote' | 'endnote'

const ROOT: Record<NoteKind, string> = { footnote: 'w:footnotes', endnote: 'w:endnotes' }
const ENTRY: Record<NoteKind, string> = { footnote: 'w:footnote', endnote: 'w:endnote' }

export const NOTE_PART_PATH: Record<NoteKind, string> = {
  footnote: 'word/footnotes.xml',
  endnote: 'word/endnotes.xml',
}

export const NOTE_REL_TYPE: Record<NoteKind, string> = {
  footnote: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnote: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
}

export const NOTE_CONTENT_TYPE: Record<NoteKind, string> = {
  footnote: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnote: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
}

const NOTE_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/**
 * Root attributes for a regenerated part. Entries are spliced back in as original bytes, so
 * the root has to keep declaring whatever prefixes those bytes use: Word puts w14:paraId on
 * every paragraph it writes, and a literal namespace list leaves that prefix unbound.
 *
 * `required` lists namespaces the rebuild itself emits (e.g. w14 on rebuilt comments): they
 * are appended when the reused original root does not declare them, since an original from
 * a non-Word producer may bind fewer prefixes than our generated markup uses.
 */
export function rootAttributes(
  originalXml: string | null,
  rootTag: string,
  fallback: string,
  required: Record<string, string> = {},
): string {
  const attrs = originalXml
    ? new RegExp(`<${rootTag}\\b([^>]*?)/?>`).exec(originalXml)?.[1]?.trim()
    : undefined
  const base = attrs && attrs.includes('xmlns:') ? attrs : fallback
  const missing = Object.entries(required)
    .filter(([prefix]) => !base.includes(`xmlns:${prefix}=`))
    .map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`)
    .join('')
  return base + missing
}

/**
 * The note's self-reference mark run (`<w:footnoteRef/>` / `<w:endnoteRef/>`).
 * Tolerates the spaced self-closing form (`<w:footnoteRef />`, .NET XmlWriter /
 * Open XML SDK output) and the open/close pair; the exact `<w:footnoteRef/>`
 * match hid every note number in such documents.
 */
const NOTE_REF_MARK_RE = /<w:(?:footnote|endnote)Ref\b\s*\/?>/

/** Structures outside the plain-text editor's patch contract stay read-only. */
const OPAQUE_NOTE_RE =
  /<(?:w:(?:fldSimple|fldChar|instrText|drawing|object|pict|tbl|hyperlink|sdt|ins|del|smartTag|customXml|altChunk|bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd)|m:oMath|mc:AlternateContent)\b/

/** real notes (separator entries excluded), in file order */
export function parseNotesXml(
  xml: string,
  kind: NoteKind,
  customIds: ReadonlySet<string> = new Set(),
): NoteInfo[] {
  return noteEntriesOf(xml, kind, customIds).map(({ id, text, xml: entryXml, markRun, mark }) => {
    const contentXml = markRun ? entryXml.replace(markRun, '') : entryXml
    const richParas = noteRichParas(contentXml)
    const hasFormat = richParas.some((paras) =>
      paras.some(
        (r) =>
          r.bold ||
          r.italic ||
          r.underline ||
          r.strike ||
          r.color ||
          r.sizeHalfPoints ||
          r.caps ||
          r.fontAscii ||
          r.textOutline,
      ),
    )
    const styleId = /<w:pStyle w:val="([^"]+)"/.exec(entryXml)?.[1]
    const spacing = noteDirectSpacing(entryXml)
    // Word draws the entry number only where a self-reference mark run exists
    // (empty-body notes keep their line but show no numeral — Word probe 2026-09-01)
    const noRefMark = !NOTE_REF_MARK_RE.test(entryXml) && !mark
    return {
      id,
      text,
      ...(mark ? { customMark: mark.text } : {}),
      ...(mark?.symbol ? { customMarkSymbol: mark.symbol } : {}),
      ...(hasFormat ? { richParas } : {}),
      ...(styleId ? { styleId } : {}),
      ...(/ADDIN\s+(?:ZOTERO_|CSL_)/.test(entryXml) ? { zoteroField: true as const } : {}),
      ...(OPAQUE_NOTE_RE.test(contentXml) ||
      /<w:sym\b/.test(contentXml) ||
      (customIds.has(id) && !mark)
        ? { protectedContent: true as const }
        : {}),
      ...(spacing ? { spacing } : {}),
      ...(noRefMark ? { noRefMark: true as const } : {}),
    }
  })
}

/** direct w:spacing attrs of the first note paragraph (they override the style chain) */
function noteDirectSpacing(entryXml: string): NoteInfo['spacing'] {
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(entryXml)?.[0]
  const sp = pPr && /<w:spacing [^>]*\/>/.exec(pPr)?.[0]
  if (!sp) return undefined
  const num = (attr: string): number | undefined => {
    const v = new RegExp(`w:${attr}="(-?\\d+)"`).exec(sp)?.[1]
    return v === undefined ? undefined : parseInt(v, 10)
  }
  const rule = /w:lineRule="(auto|atLeast|exact)"/.exec(sp)?.[1] as
    'auto' | 'atLeast' | 'exact' | undefined
  const before = num('before')
  const after = num('after')
  const line = num('line')
  if (before === undefined && after === undefined && line === undefined) return undefined
  return {
    ...(before !== undefined ? { beforeTwips: before } : {}),
    ...(after !== undefined ? { afterTwips: after } : {}),
    ...(line !== undefined ? { lineRawTwips: line, lineRule: rule ?? 'auto' } : {}),
  }
}

/** Display runs per paragraph (bold/italic/underline/strike, color, size); the footnote self-reference mark run is skipped */
function noteRichParas(entryXml: string): NoteRun[][] {
  const out: NoteRun[][] = []
  const pRe = /<w:p[\s>][\s\S]*?<\/w:p>|<w:p\/>/g
  let p: RegExpExecArray | null
  const flag = (rPr: string, tag: string) =>
    new RegExp(
      `<w:${tag}(?:\\s*/>|\\s(?![^>]*w:val=(?:"(?:0|false|none|off)"|'(?:0|false|none|off)'))[^>]*/>)`,
      'i',
    ).test(rPr)
  while ((p = pRe.exec(entryXml)) !== null) {
    const runs: NoteRun[] = []
    const rRe = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g
    let r: RegExpExecArray | null
    while ((r = rRe.exec(p[0])) !== null) {
      const inner = r[1]
      if (NOTE_REF_MARK_RE.test(inner)) continue
      const text = notePlainText(inner)
      if (!text) continue
      const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(inner)?.[0] ?? ''
      const run: NoteRun = { text }
      if (flag(rPr, 'b')) run.bold = true
      if (flag(rPr, 'i')) run.italic = true
      // w:u is not a boolean prop: underline needs an explicit non-"none" w:val
      const uVal = /<w:u\s[^>]*w:val="([^"]*)"/.exec(rPr)?.[1]
      if (uVal && uVal !== 'none') run.underline = true
      if (flag(rPr, 'strike')) run.strike = true
      const color = /<w:color [^>]*w:val="([0-9A-Fa-f]{6})"/.exec(rPr)?.[1]
      if (color) run.color = color.toUpperCase()
      const sz = /<w:sz [^>]*w:val="(\d+)"/.exec(rPr)?.[1]
      if (sz) run.sizeHalfPoints = parseInt(sz, 10)
      const rFonts = /<w:rFonts\b[^>]*>/.exec(rPr)?.[0]
      const fontAscii =
        rFonts && (/\bw:ascii="([^"]+)"/.exec(rFonts) ?? /\bw:hAnsi="([^"]+)"/.exec(rFonts))?.[1]
      if (fontAscii) run.fontAscii = fontAscii
      if (flag(rPr, 'caps')) run.caps = 'all'
      else if (flag(rPr, 'smallCaps')) run.caps = 'small'
      const outline = noteTextOutline(rPr)
      if (outline) run.textOutline = outline
      runs.push(run)
    }
    out.push(runs)
  }
  // Strip the first paragraph's leading space (spacer after the self-reference
  // mark); a spacer that was its own run empties out and is dropped entirely
  if (out[0]?.[0]) {
    out[0][0].text = out[0][0].text.replace(/^\s+/, '')
    if (out[0][0].text === '') out[0].shift()
  }
  return out
}

/** w14:textOutline with a solid srgb fill (theme colours need the full parser) */
function noteTextOutline(rPr: string): TextOutline | undefined {
  const m = /<w14:textOutline\b([^>]*)>([\s\S]*?)<\/w14:textOutline>/.exec(rPr)
  if (!m) return undefined
  const widthEmu = parseInt(/\bw14:w="(\d+)"/.exec(m[1])?.[1] ?? '', 10)
  const solid = /<w14:solidFill>([\s\S]*?)<\/w14:solidFill>/.exec(m[2])?.[1]
  const color = solid && /<w14:srgbClr w14:val="([0-9A-Fa-f]{6})"/.exec(solid)?.[1]
  if (!(widthEmu > 0) || !color) return undefined
  const alphaRaw = parseInt(/<w14:alpha w14:val="(\d+)"/.exec(solid)?.[1] ?? '', 10)
  return {
    color: color.toUpperCase(),
    widthPt: Math.round((widthEmu / 12700) * 100) / 100,
    ...(alphaRaw >= 0 && alphaRaw < 100000 ? { alpha: alphaRaw / 100000 } : {}),
  }
}

/** typeless (real) note entries with their exact XML slice + plain text */
function noteEntriesOf(
  xml: string,
  kind: NoteKind,
  customIds: ReadonlySet<string> = new Set(),
): Array<{
  id: string
  text: string
  xml: string
  markRun: string | null
  mark: ReturnType<typeof noteCustomMarkFromRun>
}> {
  const out: Array<{
    id: string
    text: string
    xml: string
    markRun: string | null
    mark: ReturnType<typeof noteCustomMarkFromRun>
  }> = []
  const entry = ENTRY[kind]
  const re = new RegExp(`<${entry}(\\s[^>]*)?>([\\s\\S]*?)</${entry}>`, 'g')
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const attrs = m[1] ?? ''
    if (/w:type=(?:"[^"]*"|'[^']*')/.test(attrs)) continue // separator / continuation entries
    const id = /w:id=(?:"([^"]+)"|'([^']+)')/.exec(attrs)?.slice(1, 3).find(Boolean)
    if (!id) continue
    const markRun = customIds.has(id) ? firstNoteMarkRun(m[0]) : null
    const mark = markRun ? noteCustomMarkFromRun(markRun) : null
    const paras: string[] = []
    const pRe = /<w:p[\s>][\s\S]*?<\/w:p>|<w:p\/>/g
    let p: RegExpExecArray | null
    while ((p = pRe.exec(m[2])) !== null)
      paras.push(notePlainText(markRun ? p[0].replace(markRun, '') : p[0]))
    // the first paragraph starts with the self-reference mark + a spacer
    if (paras.length > 0) paras[0] = paras[0].replace(/^\s+/, '')
    out.push({ id, text: paras.join('\n'), xml: m[0], markRun, mark })
  }
  return out
}

/** Exact typeless entries needed when a note moves between parts and across save/undo. */
export function noteEntryXmlById(xml: string, kind: NoteKind): Record<string, string> {
  return Object.fromEntries(noteEntriesOf(xml, kind).map((entry) => [entry.id, entry.xml]))
}

/** Prefix bindings inherited by note entries from their part root. */
export function notePartNamespaces(xml: string, kind: NoteKind): Record<string, string> {
  const root = new RegExp(`<${ROOT[kind]}\\b([^>]*)>`).exec(xml)?.[1] ?? ''
  return Object.fromEntries(
    [...root.matchAll(/\bxmlns:([\w.-]+)\s*=\s*["']([^"']+)["']/g)].map((match) => [
      match[1],
      match[2],
    ]),
  )
}

export interface NoteStyleMap {
  sourceText?: string
  targetText?: string
  sourceReference?: string
  targetReference?: string
}

/** Move only a simple note entry; unknown body structures and part relationships remain untouched. */
export function convertNoteEntryXml(
  xml: string,
  from: NoteKind,
  to: NoteKind,
  id: string,
  styles: NoteStyleMap = {},
): string | null {
  const sourceTag = ENTRY[from]
  const targetTag = ENTRY[to]
  const opening = new RegExp(`^<${sourceTag}\\b([^>]*)>`).exec(xml)
  if (!opening || !xml.endsWith(`</${sourceTag}>`)) return null
  if (OPAQUE_NOTE_RE.test(xml) || /\br:(?:id|embed|link)\s*=/.test(xml)) return null
  if (/\bw:type\s*=/.test(opening[1])) return null
  const idAttr = /\bw:id\s*=\s*(["'])([^"']*)\1/.exec(opening[1])
  if (!idAttr) return null
  const attrs = opening[1].replace(idAttr[0], `w:id=${idAttr[1]}${escapeXmlAttr(id)}${idAttr[1]}`)
  let moved = `<${targetTag}${attrs}>${xml.slice(opening[0].length, -`</${sourceTag}>`.length)}</${targetTag}>`
  if (from !== to) {
    const originalRef = from === 'footnote' ? 'w:footnoteRef' : 'w:endnoteRef'
    const replacementRef = to === 'footnote' ? 'w:footnoteRef' : 'w:endnoteRef'
    moved = moved.replace(new RegExp(`<${originalRef}\\b`, 'g'), `<${replacementRef}`)
    moved = moved.replace(new RegExp(`</${originalRef}>`, 'g'), `</${replacementRef}>`)
    const replaceStyle = (tag: 'w:pStyle' | 'w:rStyle', source?: string, target?: string) => {
      if (!source) return true
      const styleRe = new RegExp(
        `<${tag}\\b[^>]*\\bw:val=(["'])${source.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\1[^>]*>`,
        'g',
      )
      if (!styleRe.test(moved)) return true
      if (!target) return false
      moved = moved.replace(styleRe, (match) =>
        match.replace(/\bw:val=(["'])[^"']*\1/, `w:val="${escapeXmlAttr(target)}"`),
      )
      return true
    }
    if (
      !replaceStyle('w:pStyle', styles.sourceText, styles.targetText) ||
      !replaceStyle('w:rStyle', styles.sourceReference, styles.targetReference)
    )
      return null
  }
  return moved
}

function notePlainText(xml: string): string {
  const texts: string[] = []
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) texts.push(m[1])
  return decodeEntities(texts.join(''))
}

/** default separator entries required by Word when the part is created fresh */
function separatorEntries(kind: NoteKind): string {
  const entry = ENTRY[kind]
  return (
    `<${entry} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></${entry}>` +
    `<${entry} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></${entry}>`
  )
}

/** one rich display run → run XML (size/font/bold…; save-side of NoteRun) */
function noteRunXml(run: NoteRun): string {
  const props: string[] = []
  const fonts: string[] = []
  if (run.fontAscii) {
    fonts.push(
      `w:ascii="${escapeXmlAttr(run.fontAscii)}" w:hAnsi="${escapeXmlAttr(run.fontAscii)}"`,
    )
  }
  if (run.font) fonts.push(`w:eastAsia="${escapeXmlAttr(run.font)}"`)
  if (fonts.length > 0) props.push(`<w:rFonts ${fonts.join(' ')}/>`)
  if (run.bold) props.push('<w:b/>')
  if (run.italic) props.push('<w:i/>')
  if (run.underline) props.push('<w:u w:val="single"/>')
  if (run.strike) props.push('<w:strike/>')
  if (run.color) props.push(`<w:color w:val="${escapeXmlAttr(run.color)}"/>`)
  if (run.sizeHalfPoints) {
    props.push(`<w:sz w:val="${run.sizeHalfPoints}"/><w:szCs w:val="${run.sizeHalfPoints}"/>`)
  }
  const rPr = props.length > 0 ? `<w:rPr>${props.join('')}</w:rPr>` : ''
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlText(run.text)}</w:t></w:r>`
}

function noteEntryXml(kind: NoteKind, note: NoteInfo): string {
  const entry = ENTRY[kind]
  const refTag = kind === 'footnote' ? 'w:footnoteRef' : 'w:endnoteRef'
  const markXml = note.customMark
    ? noteCustomMarkXml({
        text: note.customMark,
        ...(note.customMarkSymbol ? { symbol: note.customMarkSymbol } : {}),
      })
    : `<${refTag}/>`
  if (note.richParas?.length) {
    // rich rebuild (P17): runs keep their measured size/font, the reference
    // mark + spacer shrink to the first run's size, and the paragraph pins
    // single spacing — a note rendered from measured source content must not
    // inflate past its source area via template docDefaults (after=120,
    // line=276) or a body-sized marker run
    const paras = note.richParas.map((runs, i) => {
      const sz = runs[0]?.sizeHalfPoints
      const szXml = sz ? `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` : ''
      const refRun =
        i === 0
          ? `<w:r><w:rPr><w:vertAlign w:val="superscript"/>${szXml}</w:rPr>${markXml}</w:r>` +
            `<w:r>${sz ? `<w:rPr>${szXml}</w:rPr>` : ''}<w:t xml:space="preserve"> </w:t></w:r>`
          : ''
      const pPr = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
      return `<w:p>${pPr}${refRun}${runs.map(noteRunXml).join('')}</w:p>`
    })
    return `<${entry} w:id="${escapeXmlAttr(note.id)}">${paras.join('')}</${entry}>`
  }
  const paras = note.text.split('\n').map((line, i) => {
    // OOXML convention: the note body starts with the self-reference mark
    const refRun =
      i === 0
        ? `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>${markXml}</w:r>` +
          '<w:r><w:t xml:space="preserve"> </w:t></w:r>'
        : ''
    const textRun =
      line === '' ? '' : `<w:r><w:t xml:space="preserve">${escapeXmlText(line)}</w:t></w:r>`
    return `<w:p>${refRun}${textRun}</w:p>`
  })
  return `<${entry} w:id="${escapeXmlAttr(note.id)}">${paras.join('')}</${entry}>`
}

/** Native entry for a newly authored note that has no original part entry yet. */
export function noteEntryXmlForModel(kind: NoteKind, note: NoteInfo): string {
  return noteEntryXml(kind, note)
}

/**
 * Regenerate the notes part from the full desired list. Structural entries
 * (separator/continuation) from the original part are kept byte-identical;
 * when there is no original part, standard separators are created.
 * Surgical: existing entries whose text is unchanged keep their original bytes (inline
 * formatting, images, and hyperlinks are preserved); entries with changed text first try
 * an in-paragraph w:t-level patch (formatting still preserved). An imported entry
 * cannot be rebuilt if the patch fails, because its unmodeled children would be lost.
 */
export function buildNotesXml(
  kind: NoteKind,
  notes: NoteInfo[],
  originalXml: string | null,
  originalNotes: readonly NoteInfo[] = [],
): string {
  const entry = ENTRY[kind]
  let structural = ''
  const originals = new Map<string, ReturnType<typeof noteEntriesOf>[number]>()
  if (originalXml) {
    const re = new RegExp(
      `<${entry}\\s[^>]*w:type=(?:"[^"]*"|'[^']*')[^>]*>[\\s\\S]*?</${entry}>`,
      'g',
    )
    structural = (originalXml.match(re) ?? []).join('')
    const customIds = new Set(
      originalNotes.filter((note) => note.customMark).map((note) => note.id),
    )
    for (const e of noteEntriesOf(originalXml, kind, customIds)) originals.set(e.id, e)
  }
  if (!structural) structural = separatorEntries(kind)
  const body = notes
    .map((n) => {
      const history = n.historyEntryXml
        ? noteEntriesOf(n.historyEntryXml, kind, n.customMark ? new Set([n.id]) : new Set())[0]
        : undefined
      const orig = history ?? originals.get(n.id)
      if (!orig) return noteEntryXml(kind, n)
      const sameMark =
        (orig.mark?.text ?? undefined) === n.customMark &&
        (orig.mark?.symbol?.font ?? undefined) === n.customMarkSymbol?.font &&
        (orig.mark?.symbol?.char ?? undefined) === n.customMarkSymbol?.char
      if (orig.text === n.text && sameMark) return orig.xml
      if (OPAQUE_NOTE_RE.test(orig.xml))
        throw new Error(`Cannot safely edit ${kind} ${n.id}: protected OOXML content`)
      let source = orig.xml
      if (!sameMark) {
        const oldRun = firstNoteMarkRun(source)
        const refTag = kind === 'footnote' ? 'w:footnoteRef' : 'w:endnoteRef'
        const nextMark = n.customMark
          ? noteCustomMarkXml({
              text: n.customMark,
              ...(n.customMarkSymbol ? { symbol: n.customMarkSymbol } : {}),
            })
          : `<${refTag}/>`
        const updatedRun = oldRun?.replace(
          /<w:(?:footnoteRef|endnoteRef|sym)\b[^>]*\/>|<w:t\b[^>]*>[\s\S]*?<\/w:t>/,
          nextMark,
        )
        if (!oldRun || !updatedRun || updatedRun === oldRun)
          throw new Error(`Cannot safely change ${kind} ${n.id}: unsupported reference mark`)
        source = source.replace(oldRun, updatedRun)
      }
      if (orig.text === n.text) return source
      const markerRun = n.customMark ? firstNoteMarkRun(source) : null
      const hasTextMark = markerRun?.includes('<w:t') ?? false
      let placeholder = '<w:r w:customMarkPlaceholder="genoffice"/>'
      while (source.includes(placeholder)) placeholder = placeholder.replace('/>', 'x="1"/>')
      const originalForPatch =
        hasTextMark && markerRun ? source.replace(markerRun, placeholder) : source
      const patched = patchParagraphTexts(originalForPatch, n.text, {
        stripFirstParaLeadingSpace: true,
      })
      if (patched)
        return hasTextMark && markerRun ? patched.replace(placeholder, markerRun) : patched
      throw new Error(`Cannot safely edit ${kind} ${n.id}: unsupported paragraph change`)
    })
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<${ROOT[kind]} ${rootAttributes(originalXml, ROOT[kind], NOTE_NS, Object.assign({}, ...notes.map((note) => note.historyEntryNamespaces ?? {})))}>${structural}${body}</${ROOT[kind]}>`
  )
}

/** next free numeric note id (separator ids -1/0 and existing notes considered) */
export function nextNoteId(notes: NoteInfo[]): string {
  let max = 0
  for (const note of notes) {
    const n = parseInt(note.id, 10)
    if (Number.isFinite(n)) max = Math.max(max, n)
  }
  return String(max + 1)
}
