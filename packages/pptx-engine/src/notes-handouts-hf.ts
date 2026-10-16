/** Native header/footer fields for notes and handout masters. Print rendering reads
 * these same archive parts; edits never flatten a dynamic a:fld into plain text. */
import type { OpenedPptx } from './index'
import { resolveTarget, type PackageArchive } from './zip'
import {
  appendRelationship,
  addContentTypeOverride,
  ensureNotesMaster,
  notesPathForSlide,
  unescapeXml,
  getSlideNotesParagraphs,
} from './notes'
import { editPlaceholderXml } from './headerfooter'
import { escapeXmlText } from './xml-utils'
import { randomUUID } from 'node:crypto'
import { XMLParser } from 'fast-xml-parser'
import { resolveColorNode } from './color'
import { parseClrMap, parseTheme, resolveFontRef, type Theme } from './theme'

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const HANDOUT_MASTER_REL = `${REL_BASE}/handoutMaster`
const THEME_REL = `${REL_BASE}/theme`
const HANDOUT_MASTER_CT =
  'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
type Kind = 'hdr' | 'ftr' | 'dt' | 'sldNum'
const colorParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

function masterTheme(
  archive: PackageArchive,
  path: string | undefined,
  masterXml: string | null,
): Theme | undefined {
  if (!path || !masterXml) return undefined
  for (const rel of archive.readRels(path).values()) {
    if (!rel.type.endsWith('/theme')) continue
    const xml = archive.readText(resolveTarget(path, rel.target))
    if (xml) {
      const theme = parseTheme(xml)
      theme.clrMap = parseClrMap(masterXml)
      return theme
    }
  }
  return undefined
}

function xmlColor(fragment: string | undefined, theme?: Theme): string | undefined {
  const fill = fragment?.match(/<a:solidFill\b[^>]*>[\s\S]*?<\/a:solidFill>/)?.[0]
  return fill ? resolveColorNode(colorParser.parse(fill)['a:solidFill'], theme) : undefined
}

export interface NotesHandoutsHeaderFooterOptions {
  headerEnabled: boolean
  header: string
  footerEnabled: boolean
  footer: string
  dateEnabled: boolean
  date: string
  dateAuto: boolean
  dateFormat: string
  pageNum: boolean
}

export type NotesHandoutsHeaderFooterState = NotesHandoutsHeaderFooterOptions

export interface PrintHeaderFooterField {
  text: string
  /** Dynamic OOXML field type, preserved for date and page-number evaluation. */
  field?: string
  fontSizePt?: number
  color?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontFamily?: string
  align?: 'left' | 'center' | 'right' | 'justify'
  box?: PrintPageBox
}

export interface PrintPageBox {
  x: number
  y: number
  cx: number
  cy: number
  rotationDeg?: number
  flipH?: boolean
  flipV?: boolean
}

export interface PrintNotesParagraph {
  align?: 'left' | 'center' | 'right' | 'justify'
  lineHeight?: number
  lineExactPt?: number
  spaceBeforePt?: number
  spaceAfterPt?: number
  spaceBeforePct?: number
  spaceAfterPct?: number
  rtl?: boolean
  level?: number
  marginLeftPt?: number
  marginRightPt?: number
  indentPt?: number
  tabStopsPt?: number[]
  tabStops?: Array<{ posPt: number; align?: string }>
  defaultTabPt?: number
  bullet?: {
    type: 'char' | 'number' | 'blip'
    char?: string
    numberType?: string
    startAt?: number
    fontFamily?: string
    color?: string
    sizePt?: number
    sizePct?: number
    imageDataUrl?: string
  }
  runs: Array<
    Pick<
      PrintHeaderFooterField,
      'text' | 'fontSizePt' | 'fontFamily' | 'color' | 'bold' | 'italic' | 'underline'
    > & { strike?: boolean; baseline?: number; highlight?: string; letterSpacingPt?: number }
  >
}

export interface PrintHeaderFooterPage {
  header?: PrintHeaderFooterField
  footer?: PrintHeaderFooterField
  date?: PrintHeaderFooterField
  pageNum?: PrintHeaderFooterField
  slideImage?: PrintPageBox
  notesBody?: PrintPageBox
  /** Master placeholder existence is the visibility gate; local notes may override geometry. */
  slideImageVisible?: boolean
  notesBodyVisible?: boolean
  notesParagraphs?: PrintNotesParagraph[]
  canvas?: { cx: number; cy: number }
  background?: string
  masterPath?: string
}

const FIELD_RE = /^datetime[\w.-]*$/
function dateField(type: string): string {
  return FIELD_RE.test(type) ? type : 'datetime1'
}

function shapeOf(xml: string, kind: Kind | 'sldImg' | 'body'): string | null {
  for (const match of xml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)) {
    if (new RegExp(`<p:ph\\b[^>]*\\btype=["']${kind}["']`).test(match[0])) return match[0]
  }
  return null
}

function fieldOf(shape: string | null, theme?: Theme): PrintHeaderFooterField | undefined {
  if (!shape) return undefined
  const text = [...shape.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)]
    .map((match) => unescapeXml(match[1]!))
    .join('')
  const field = /<a:fld\b[^>]*\btype=["']([^"']+)["']/.exec(shape)?.[1]
  const rPr = /<a:rPr\b[^>]*\/?\s*>/.exec(shape)?.[0]
  const rPrBlock = /<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/.exec(shape)?.[0]
  const fontSize = rPr?.match(/\bsz=["'](\d+)["']/)?.[1]
  const color = xmlColor(rPrBlock, theme)
  const box = shapeBox(shape)
  const alignment = /<a:pPr\b[^>]*\balgn=["']([^"']+)["']/.exec(shape)?.[1]
  const typeface = /<a:latin\b[^>]*\btypeface=["']([^"']+)["']/.exec(rPrBlock ?? '')?.[1]
  const fontFamily = typeface ? (resolveFontRef(typeface, theme) ?? typeface) : undefined
  return {
    text,
    ...(field ? { field } : {}),
    ...(fontSize ? { fontSizePt: Number(fontSize) / 100 } : {}),
    ...(color ? { color } : {}),
    ...(rPr && /\bb=["'](?:1|true)["']/.test(rPr) ? { bold: true } : {}),
    ...(rPr && /\bi=["'](?:1|true)["']/.test(rPr) ? { italic: true } : {}),
    ...(rPr && /\bu=["'](?:sng|dbl|heavy|words)["']/.test(rPr) ? { underline: true } : {}),
    ...(fontFamily ? { fontFamily } : {}),
    ...(alignment === 'ctr'
      ? { align: 'center' as const }
      : alignment === 'r'
        ? { align: 'right' as const }
        : alignment === 'just'
          ? { align: 'justify' as const }
          : {}),
    ...(box ? { box } : {}),
  }
}

function shapeBox(shape: string | null): PrintPageBox | undefined {
  if (!shape) return undefined
  const xfrmMatch = /<a:xfrm\b([^>]*)>([\s\S]*?)<\/a:xfrm>/.exec(shape)
  const attrs = xfrmMatch?.[1] ?? ''
  const xfrm = xfrmMatch?.[2]
  const off = /<a:off\b[^>]*\bx=["'](-?\d+)["'][^>]*\by=["'](-?\d+)["']/.exec(xfrm ?? '')
  const ext = /<a:ext\b[^>]*\bcx=["'](\d+)["'][^>]*\bcy=["'](\d+)["']/.exec(xfrm ?? '')
  return off && ext
    ? {
        x: Number(off[1]),
        y: Number(off[2]),
        cx: Number(ext[1]),
        cy: Number(ext[2]),
        ...(/\brot=["'](-?\d+)["']/.test(attrs)
          ? { rotationDeg: Number(/\brot=["'](-?\d+)["']/.exec(attrs)![1]) / 60000 }
          : {}),
        ...(/\bflipH=["'](?:1|true)["']/.test(attrs) ? { flipH: true } : {}),
        ...(/\bflipV=["'](?:1|true)["']/.test(attrs) ? { flipV: true } : {}),
      }
    : undefined
}

function enabledByMaster(xml: string, kind: Kind): boolean {
  const hf = /<p:hf\b[^>]*\/?\s*>/.exec(xml)?.[0]
  return !hf || !new RegExp(`\\b${kind}=["'](?:0|false)["']`).test(hf)
}

export function readMasterHeaderFooterFlags(
  archive: PackageArchive,
  path: string,
): Record<Kind, boolean> | null {
  if (!/^ppt\/(?:notesMasters|handoutMasters)\/[^/]+\.xml$/.test(path)) return null
  const xml = archive.readText(path)
  if (!xml) return null
  return {
    hdr: !!shapeOf(xml, 'hdr') && enabledByMaster(xml, 'hdr'),
    ftr: !!shapeOf(xml, 'ftr') && enabledByMaster(xml, 'ftr'),
    dt: !!shapeOf(xml, 'dt') && enabledByMaster(xml, 'dt'),
    sldNum: !!shapeOf(xml, 'sldNum') && enabledByMaster(xml, 'sldNum'),
  }
}

/** Toggle one native p:hf flag without modifying field runs or other flags. */
export function setMasterHeaderFooterFlag(
  archive: PackageArchive,
  path: string,
  kind: Kind,
  enabled: boolean,
): boolean {
  if (!/^ppt\/(?:notesMasters|handoutMasters)\/[^/]+\.xml$/.test(path)) return false
  const original = archive.readText(path)
  if (!original) return false
  let xml = original
  if (enabled && !shapeOf(xml, kind)) {
    const defaultValue = kind === 'dt' ? '1/1/2000' : kind === 'sldNum' ? '1' : ''
    const field = kind === 'dt' ? 'datetime1' : kind === 'sldNum' ? 'slidenum' : undefined
    xml = xml.replace('</p:spTree>', (end) => `${newShape(xml, kind, defaultValue, field)}${end}`)
  }
  const existing = /<p:hf\b[^>]*\/?\s*>/.exec(xml)?.[0]
  let hf = existing ?? '<p:hf/>'
  const attr = new RegExp(`\\b${kind}=["'][^"']*["']`)
  hf = attr.test(hf)
    ? hf.replace(attr, `${kind}="${enabled ? 1 : 0}"`)
    : hf.replace(/\s*\/?>$/, (end) => ` ${kind}="${enabled ? 1 : 0}"${end}`)
  xml = existing
    ? xml.replace(existing, () => hf)
    : xml.replace('</p:cSld>', (end) => `${end}${hf}`)
  if (xml === original) return false
  archive.entries.set(path, Buffer.from(xml, 'utf8'))
  return true
}

function allMasterPaths(archive: PackageArchive, kind: 'notes' | 'handout'): string[] {
  const directory = kind === 'notes' ? 'notesMasters/notesMaster' : 'handoutMasters/handoutMaster'
  return [...archive.entries.keys()].filter((path) =>
    new RegExp(`^ppt/${directory}\\d+\\.xml$`).test(path),
  )
}

function registeredMasterPath(
  archive: PackageArchive,
  kind: 'notes' | 'handout',
): string | undefined {
  const presentation = archive.readText('ppt/presentation.xml')
  const tag = kind === 'notes' ? 'notesMasterId' : 'handoutMasterId'
  const rid = presentation?.match(new RegExp(`<p:${tag}\\b[^>]*\\br:id=["']([^"']+)["']`))?.[1]
  const rel = rid ? archive.readRels('ppt/presentation.xml').get(rid) : undefined
  const path = rel ? resolveTarget('ppt/presentation.xml', rel.target) : undefined
  return path && archive.has(path) ? path : allMasterPaths(archive, kind)[0]
}

function notesPageMasterPath(archive: PackageArchive, notesPath: string): string | undefined {
  for (const rel of archive.readRels(notesPath).values()) {
    if (rel.type === `${REL_BASE}/notesMaster`) {
      const path = resolveTarget(notesPath, rel.target)
      if (archive.has(path)) return path
    }
  }
  return registeredMasterPath(archive, 'notes')
}

function setHfFlags(xml: string, opts: NotesHandoutsHeaderFooterOptions): string {
  const flags: Record<Kind, boolean> = {
    hdr: opts.headerEnabled,
    ftr: opts.footerEnabled,
    dt: opts.dateEnabled,
    sldNum: opts.pageNum,
  }
  const existing = /<p:hf\b[^>]*\/?\s*>/.exec(xml)?.[0]
  let next = existing ?? '<p:hf/>'
  for (const [kind, on] of Object.entries(flags) as Array<[Kind, boolean]>) {
    const attr = new RegExp(`\\b${kind}=["'][^"']*["']`)
    next = attr.test(next)
      ? next.replace(attr, `${kind}="${on ? 1 : 0}"`)
      : next.replace(/\s*\/?>$/, (end) => ` ${kind}="${on ? 1 : 0}"${end}`)
  }
  if (existing) return xml.replace(existing, () => next)
  if (/<p:clrMap\b[^>]*\/>/.test(xml))
    return xml.replace(/<p:clrMap\b[^>]*\/>/, (marker) => `${marker}${next}`)
  return xml.replace('</p:cSld>', (marker) => `${marker}${next}`)
}

function nextShapeId(xml: string): number {
  return (
    Math.max(
      1,
      ...[...xml.matchAll(/<p:cNvPr\b[^>]*\bid=["'](\d+)["']/g)].map((m) => Number(m[1])),
    ) + 1
  )
}

function newShape(xml: string, kind: Kind, value: string, field?: string): string {
  const id = nextShapeId(xml)
  const box = {
    hdr: [457200, 182880, 4114800, 320040],
    dt: [457200, 8595360, 1828800, 320040],
    ftr: [2286000, 8595360, 2286000, 320040],
    sldNum: [5486400, 8595360, 914400, 320040],
  }[kind]
  const run = field
    ? `<a:fld id="{${randomUUID()}}" type="${field}"><a:rPr lang="en-US" sz="1200"/><a:t>${escapeXmlText(value)}</a:t></a:fld>`
    : `<a:r><a:rPr lang="en-US" sz="1200"/><a:t>${escapeXmlText(value)}</a:t></a:r>`
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${kind} Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${kind}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${box[0]}" y="${box[1]}"/><a:ext cx="${box[2]}" cy="${box[3]}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${run}</a:p></p:txBody></p:sp>`
}

function values(
  opts: NotesHandoutsHeaderFooterOptions,
): Record<Kind, { text: string; field?: string; enabled: boolean }> {
  return {
    hdr: { text: opts.header, enabled: opts.headerEnabled },
    ftr: { text: opts.footer, enabled: opts.footerEnabled },
    dt: {
      text: opts.date,
      field: opts.dateAuto ? dateField(opts.dateFormat) : undefined,
      enabled: opts.dateEnabled,
    },
    sldNum: { text: '1', field: 'slidenum', enabled: opts.pageNum },
  }
}

function patchMaster(xml: string, opts: NotesHandoutsHeaderFooterOptions): string {
  let next = xml
  for (const [kind, wanted] of Object.entries(values(opts)) as Array<
    [Kind, ReturnType<typeof values>[Kind]]
  >) {
    if (!wanted.enabled) continue
    const shape = shapeOf(next, kind)
    if (shape) {
      const current = fieldOf(shape)
      if (current?.text === wanted.text && current.field === wanted.field) continue
      next = next.replace(shape, () => editPlaceholderXml(shape, wanted.text, wanted.field))
    } else {
      next = next.replace(
        '</p:spTree>',
        (marker) => `${newShape(next, kind, wanted.text, wanted.field)}${marker}`,
      )
    }
  }
  return setHfFlags(next, opts)
}

function patchLocalNotes(xml: string, opts: NotesHandoutsHeaderFooterOptions): string {
  let next = xml
  for (const [kind, wanted] of Object.entries(values(opts)) as Array<
    [Kind, ReturnType<typeof values>[Kind]]
  >) {
    // Master p:hf flags hide local placeholders. Keep disabled local field runs
    // byte-for-byte so re-enabling does not discard native IDs or cached values.
    if (!wanted.enabled) continue
    const shape = shapeOf(next, kind)
    if (!shape) continue
    const current = fieldOf(shape)
    if (current?.text === wanted.text && current.field === wanted.field) continue
    next = next.replace(shape, () => editPlaceholderXml(shape, wanted.text, wanted.field))
  }
  return next
}

export function ensureHandoutMaster(archive: PackageArchive): string {
  const existing = allMasterPaths(archive, 'handout')[0]
  if (existing) return existing
  let index = 1
  while (archive.has(`ppt/handoutMasters/handoutMaster${index}.xml`)) index++
  const path = `ppt/handoutMasters/handoutMaster${index}.xml`
  const emptyTree =
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:handoutMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld>${emptyTree}</p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:handoutMaster>`
  archive.entries.set(path, Buffer.from(xml, 'utf8'))
  addContentTypeOverride(archive, path, HANDOUT_MASTER_CT)
  const theme = [...archive.entries.keys()].find((part) => /^ppt\/theme\/theme\d+\.xml$/.test(part))
  if (theme) {
    let themeIndex = 1
    while (archive.has(`ppt/theme/theme${themeIndex}.xml`)) themeIndex++
    const handoutTheme = `ppt/theme/theme${themeIndex}.xml`
    archive.entries.set(handoutTheme, Buffer.from(archive.readText(theme)!, 'utf8'))
    const themeRels = archive.readText(`ppt/theme/_rels/${theme.split('/').at(-1)}.rels`)
    if (themeRels)
      archive.entries.set(
        `ppt/theme/_rels/${handoutTheme.split('/').at(-1)}.rels`,
        Buffer.from(themeRels, 'utf8'),
      )
    addContentTypeOverride(
      archive,
      handoutTheme,
      'application/vnd.openxmlformats-officedocument.theme+xml',
    )
    appendRelationship(archive, path, THEME_REL, `../${handoutTheme.slice(4)}`)
  }
  const rid = appendRelationship(
    archive,
    'ppt/presentation.xml',
    HANDOUT_MASTER_REL,
    `handoutMasters/handoutMaster${index}.xml`,
  )
  const presentation = archive.readText('ppt/presentation.xml')!
  const idList = `<p:handoutMasterIdLst><p:handoutMasterId r:id="${rid}"/></p:handoutMasterIdLst>`
  const next = presentation.includes('</p:notesMasterIdLst>')
    ? presentation.replace('</p:notesMasterIdLst>', (marker) => `${marker}${idList}`)
    : presentation.replace('</p:sldMasterIdLst>', (marker) => `${marker}${idList}`)
  archive.entries.set('ppt/presentation.xml', Buffer.from(next, 'utf8'))
  return path
}

/** Apply globally to native masters and any notes-page local overrides. */
export function applyNotesHandoutsHeaderFooter(
  opened: OpenedPptx,
  opts: NotesHandoutsHeaderFooterOptions,
): boolean {
  const archive = opened.archive
  const hadNotesMaster = allMasterPaths(archive, 'notes').length > 0
  const hadHandoutMaster = allMasterPaths(archive, 'handout').length > 0
  if (
    !hadNotesMaster &&
    !hadHandoutMaster &&
    !opts.headerEnabled &&
    !opts.footerEnabled &&
    !opts.dateEnabled &&
    !opts.pageNum
  )
    return false
  ensureNotesMaster(archive)
  ensureHandoutMaster(archive)
  let changed = !hadNotesMaster || !hadHandoutMaster
  for (const path of [...allMasterPaths(archive, 'notes'), ...allMasterPaths(archive, 'handout')]) {
    const xml = archive.readText(path)!
    const next = patchMaster(xml, opts)
    if (next !== xml) {
      archive.entries.set(path, Buffer.from(next, 'utf8'))
      changed = true
    }
  }
  for (const slide of opened.deck.slides) {
    const path = notesPathForSlide(archive, slide.path)
    const xml = path ? archive.readText(path) : null
    if (!path || !xml) continue
    const next = patchLocalNotes(xml, opts)
    if (next !== xml) {
      archive.entries.set(path, Buffer.from(next, 'utf8'))
      changed = true
    }
  }
  return changed
}

function masterState(
  archive: PackageArchive,
  path: string | undefined,
): NotesHandoutsHeaderFooterState | null {
  const xml = path ? archive.readText(path) : null
  if (!xml) return null
  const header = fieldOf(shapeOf(xml, 'hdr'))
  const footer = fieldOf(shapeOf(xml, 'ftr'))
  const date = fieldOf(shapeOf(xml, 'dt'))
  const page = fieldOf(shapeOf(xml, 'sldNum'))
  return {
    headerEnabled: !!header && enabledByMaster(xml, 'hdr'),
    header: header?.text ?? '',
    footerEnabled: !!footer && enabledByMaster(xml, 'ftr'),
    footer: footer?.text ?? '',
    dateEnabled: !!date && enabledByMaster(xml, 'dt'),
    date: date?.text ?? '',
    dateAuto: !!date?.field?.startsWith('datetime'),
    dateFormat: date?.field && FIELD_RE.test(date.field) ? date.field : 'datetime1',
    pageNum: !!page && enabledByMaster(xml, 'sldNum'),
  }
}

export function readNotesHandoutsHeaderFooter(opened: OpenedPptx): NotesHandoutsHeaderFooterState {
  return (
    masterState(opened.archive, registeredMasterPath(opened.archive, 'notes')) ??
    masterState(opened.archive, registeredMasterPath(opened.archive, 'handout')) ?? {
      headerEnabled: false,
      header: '',
      footerEnabled: false,
      footer: '',
      dateEnabled: false,
      date: '',
      dateAuto: true,
      dateFormat: 'datetime1',
      pageNum: false,
    }
  )
}

function resolvedPrintFields(
  masterXml: string | null,
  localXml?: string | null,
  canvas?: { cx: number; cy: number },
  theme?: Theme,
  masterPath?: string,
): PrintHeaderFooterPage {
  const result: PrintHeaderFooterPage = {
    ...(canvas ? { canvas } : {}),
    ...(masterPath ? { masterPath } : {}),
  }
  if (!masterXml) {
    result.slideImageVisible = true
    result.notesBodyVisible = true
    return result
  }
  // The local notes page owns content visibility. The master provides geometry
  // and styling, but removing its placeholder need not delete a local one.
  // PowerPoint synthesizes a default page if no notesSlide part exists.
  result.slideImageVisible = !localXml || !!shapeOf(localXml, 'sldImg')
  result.notesBodyVisible = !localXml || !!shapeOf(localXml, 'body')
  const slideImage = result.slideImageVisible
    ? shapeBox((localXml && shapeOf(localXml, 'sldImg')) || shapeOf(masterXml, 'sldImg'))
    : undefined
  const notesBody = result.notesBodyVisible
    ? shapeBox((localXml && shapeOf(localXml, 'body')) || shapeOf(masterXml, 'body'))
    : undefined
  if (slideImage) result.slideImage = slideImage
  if (notesBody) result.notesBody = notesBody
  const bg = xmlColor(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/.exec(masterXml)?.[0], theme)
  if (bg) result.background = bg
  const keys: Array<[Kind, 'header' | 'footer' | 'date' | 'pageNum']> = [
    ['hdr', 'header'],
    ['ftr', 'footer'],
    ['dt', 'date'],
    ['sldNum', 'pageNum'],
  ]
  for (const [kind, key] of keys) {
    if (!enabledByMaster(masterXml, kind)) continue
    const master = fieldOf(shapeOf(masterXml, kind), theme)
    const local = localXml ? fieldOf(shapeOf(localXml, kind), theme) : undefined
    const value = local
      ? { ...master, ...local, field: local.field, box: local.box ?? master?.box }
      : master
    if (value && (value.text || value.field)) result[key] = value
  }
  return result
}

/** Read effective fields for print; a local notes shape overrides its master shape. */
export function readNotesHandoutsPrintFields(opened: OpenedPptx): {
  notesPages: PrintHeaderFooterPage[]
  handout: PrintHeaderFooterPage
} {
  const archive = opened.archive
  const notesTag = /<p:notesSz\b([^>]*)\/?\s*>/.exec(
    archive.readText('ppt/presentation.xml') ?? '',
  )?.[1]
  const canvas = {
    cx: Number(notesTag?.match(/\bcx=["'](\d+)["']/)?.[1]) || 6858000,
    cy: Number(notesTag?.match(/\bcy=["'](\d+)["']/)?.[1]) || 9144000,
  }
  const handoutPath = registeredMasterPath(archive, 'handout')
  const handoutMaster = archive.readText(handoutPath ?? '')
  return {
    notesPages: opened.deck.slides.map((slide) => {
      const path = notesPathForSlide(archive, slide.path)
      const masterPath = path
        ? notesPageMasterPath(archive, path)
        : registeredMasterPath(archive, 'notes')
      const master = archive.readText(masterPath ?? '')
      const fields = resolvedPrintFields(
        master,
        path ? archive.readText(path) : null,
        canvas,
        masterTheme(archive, masterPath, master),
        masterPath,
      )
      fields.notesParagraphs = getSlideNotesParagraphs(archive, slide.path).map((paragraph) => ({
        ...(paragraph.align ? { align: paragraph.align } : {}),
        ...(paragraph.lineHeight ? { lineHeight: paragraph.lineHeight } : {}),
        ...(paragraph.lineExact != null ? { lineExactPt: paragraph.lineExact } : {}),
        ...(paragraph.spaceBefore != null ? { spaceBeforePt: paragraph.spaceBefore } : {}),
        ...(paragraph.spaceAfter != null ? { spaceAfterPt: paragraph.spaceAfter } : {}),
        ...(paragraph.spaceBeforePct != null ? { spaceBeforePct: paragraph.spaceBeforePct } : {}),
        ...(paragraph.spaceAfterPct != null ? { spaceAfterPct: paragraph.spaceAfterPct } : {}),
        ...(paragraph.rtl != null ? { rtl: paragraph.rtl } : {}),
        ...(paragraph.level != null ? { level: paragraph.level } : {}),
        ...(paragraph.marL != null ? { marginLeftPt: paragraph.marL / 12700 } : {}),
        ...(paragraph.marR != null ? { marginRightPt: paragraph.marR / 12700 } : {}),
        ...(paragraph.indent != null ? { indentPt: paragraph.indent / 12700 } : {}),
        ...(paragraph.tabStops?.length
          ? {
              tabStopsPt: paragraph.tabStops.map((stop) => stop.pos / 12700),
              tabStops: paragraph.tabStops.map((stop) => ({
                posPt: stop.pos / 12700,
                ...(stop.algn ? { align: stop.algn } : {}),
              })),
            }
          : {}),
        ...(paragraph.defTabSz != null ? { defaultTabPt: paragraph.defTabSz / 12700 } : {}),
        ...(paragraph.bullet && paragraph.bullet.type !== 'none'
          ? {
              bullet: {
                type: paragraph.bullet.type,
                ...(paragraph.bullet.char ? { char: paragraph.bullet.char } : {}),
                ...(paragraph.bullet.numType ? { numberType: paragraph.bullet.numType } : {}),
                ...(paragraph.bullet.startAt ? { startAt: paragraph.bullet.startAt } : {}),
                ...(paragraph.bullet.font ? { fontFamily: paragraph.bullet.font } : {}),
                ...(paragraph.bullet.color ? { color: paragraph.bullet.color } : {}),
                ...(paragraph.bullet.sizePt ? { sizePt: paragraph.bullet.sizePt } : {}),
                ...(paragraph.bullet.sizePct ? { sizePct: paragraph.bullet.sizePct } : {}),
                ...(paragraph.bullet.mediaRef && archive.entries.has(paragraph.bullet.mediaRef)
                  ? {
                      imageDataUrl: `data:${/\.jpe?g$/i.test(paragraph.bullet.mediaRef) ? 'image/jpeg' : /\.gif$/i.test(paragraph.bullet.mediaRef) ? 'image/gif' : 'image/png'};base64,${Buffer.from(archive.entries.get(paragraph.bullet.mediaRef)!).toString('base64')}`,
                    }
                  : {}),
              },
            }
          : {}),
        runs: paragraph.runs
          .filter((run) => !!run.text)
          .map((run) => ({
            text: run.text,
            ...(run.fontSize ? { fontSizePt: run.fontSize } : {}),
            ...(run.fontFamily ? { fontFamily: run.fontFamily } : {}),
            ...(run.color ? { color: run.color } : {}),
            ...(run.bold ? { bold: true } : {}),
            ...(run.italic ? { italic: true } : {}),
            ...(run.underline ? { underline: true } : {}),
            ...(run.strike ? { strike: true } : {}),
            ...(run.baseline ? { baseline: run.baseline } : {}),
            ...(run.highlight ? { highlight: run.highlight } : {}),
            ...(run.letterSpacing != null ? { letterSpacingPt: run.letterSpacing } : {}),
          })),
      }))
      return fields
    }),
    handout: resolvedPrintFields(
      handoutMaster,
      null,
      canvas,
      masterTheme(archive, handoutPath, handoutMaster),
      handoutPath,
    ),
  }
}
