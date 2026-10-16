/**
 * Slide master / layout edit view support (6.16).
 *
 * Exception to the fidelity rule: only layout/master parts the user actively edited
 * in master view are written back, reusing the slide byte-surgery path (scanSlide
 * anchors + patchSlideXml); untouched elements pass through with original bytes.
 * This module only enumerates and parses; the caller writes back via patchSlideXml
 * and then writes the archive entry.
 */
import type { PackageArchive } from './zip'
import { resolveTarget } from './zip'
import { parseClrMap, parseTheme } from './theme'
import { parseSlide, parseDecorations, type ParseContext } from './parse'
import { parsePlaceholderMap, parseMasterTextStyles } from './placeholder'
import type { Slide } from './types'

export interface MasterPartInfo {
  /** Path inside the zip, e.g. ppt/slideMasters/slideMaster1.xml */
  partPath: string
  kind: 'master' | 'layout' | 'notesMaster' | 'handoutMaster'
  /** <p:cSld name="…">, falling back to the file name */
  name: string
}

const partNum = (p: string) => parseInt(/(\d+)\.xml$/.exec(p)?.[1] ?? '0', 10)

function cSldName(xml: string | null | undefined, fallback: string): string {
  const m = xml ? /<p:cSld\s[^>]*name="([^"]*)"/.exec(xml) : null
  return m?.[1] || fallback
}

/** Enumerate parts for the master edit view: each master first, followed by its layouts. */
export function listMasterParts(archive: PackageArchive): MasterPartInfo[] {
  const masters = [...archive.entries.keys()]
    .filter((p) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(p))
    .sort((a, b) => partNum(a) - partNum(b))
  const out: MasterPartInfo[] = []
  for (const m of masters) {
    out.push({ partPath: m, kind: 'master', name: cSldName(archive.readText(m), baseName(m)) })
    const layouts = [...archive.readRels(m).values()]
      .filter((r) => r.type.endsWith('/slideLayout'))
      .map((r) => resolveTarget(m, r.target))
      .sort((a, b) => partNum(a) - partNum(b))
    for (const l of layouts) {
      out.push({ partPath: l, kind: 'layout', name: cSldName(archive.readText(l), baseName(l)) })
    }
  }
  for (const [directory, kind] of [
    ['notesMasters/notesMaster', 'notesMaster'],
    ['handoutMasters/handoutMaster', 'handoutMaster'],
  ] as const) {
    for (const path of [...archive.entries.keys()]
      .filter((p) => new RegExp(`^ppt/${directory}\\d+\\.xml$`).test(p))
      .sort((a, b) => partNum(a) - partNum(b))) {
      out.push({ partPath: path, kind, name: cSldName(archive.readText(path), baseName(path)) })
    }
  }
  return out
}

/** Notes and handout masters use the presentation's notes canvas, not p:sldSz. */
export function masterPartSize(
  archive: PackageArchive,
  partPath: string,
  slideSize: { cx: number; cy: number },
): { cx: number; cy: number } {
  if (!/^ppt\/(?:notesMasters|handoutMasters)\//.test(partPath)) return slideSize
  const tag = /<p:notesSz\b([^>]*)\/?\s*>/.exec(archive.readText('ppt/presentation.xml') ?? '')?.[1]
  const cx = Number(tag?.match(/\bcx=["'](\d+)["']/)?.[1])
  const cy = Number(tag?.match(/\bcy=["'](\d+)["']/)?.[1])
  return cx > 0 && cy > 0 ? { cx, cy } : slideSize
}

/** p:notesSz is the native canvas for both notes and handout masters. */
export function setNotesPageOrientation(
  archive: PackageArchive,
  orientation: 'portrait' | 'landscape',
): boolean {
  const path = 'ppt/presentation.xml'
  const xml = archive.readText(path)
  if (!xml) return false
  const tag = /<p:notesSz\b[^>]*\/?\s*>/.exec(xml)?.[0]
  if (!tag) return false
  const cx = Number(tag.match(/\bcx=["'](\d+)["']/)?.[1])
  const cy = Number(tag.match(/\bcy=["'](\d+)["']/)?.[1])
  if (!(cx > 0 && cy > 0)) return false
  if ((orientation === 'landscape') === cx > cy) return false
  const next = tag
    .replace(/\bcx=["']\d+["']/, `cx="${cy}"`)
    .replace(/\bcy=["']\d+["']/, `cy="${cx}"`)
  archive.entries.set(
    path,
    Buffer.from(
      xml.replace(tag, () => next),
      'utf8',
    ),
  )
  return true
}

/** Notes body formatting belongs to p:notesStyle, which notes pages inherit. */
export function setNotesMasterBodyStyle(
  archive: PackageArchive,
  path: string,
  patch: {
    fontFamily?: string
    fontSizePt?: number
    bold?: boolean
    italic?: boolean
    underline?: boolean
    color?: string
    align?: 'left' | 'center' | 'right' | 'justify'
    lineSpacingPct?: number
  },
): boolean {
  if (!/^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)) return false
  const original = archive.readText(path)
  if (!original) return false
  const existingStyle = /<p:notesStyle\b[^>]*>[\s\S]*?<\/p:notesStyle>/.exec(original)?.[0]
  let style = existingStyle ?? '<p:notesStyle><a:lvl1pPr><a:defRPr/></a:lvl1pPr></p:notesStyle>'
  if (!/<a:lvl1pPr\b/.test(style))
    style = style.replace('</p:notesStyle>', '<a:lvl1pPr><a:defRPr/></a:lvl1pPr></p:notesStyle>')
  let level = /<a:lvl1pPr\b[^>]*\/>|<a:lvl1pPr\b[^>]*>[\s\S]*?<\/a:lvl1pPr>/.exec(style)![0]
  if (level.endsWith('/>')) level = level.replace(/\/>$/, '><a:defRPr/></a:lvl1pPr>')
  const beforeLevel = /<a:lvl1pPr\b[^>]*\/>|<a:lvl1pPr\b[^>]*>[\s\S]*?<\/a:lvl1pPr>/.exec(style)![0]
  if (patch.align) {
    const algn = { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[patch.align]
    level = level.replace(/<a:lvl1pPr\b[^>]*>/, (tag) =>
      /\balgn=["'][^"']*["']/.test(tag)
        ? tag.replace(/\balgn=["'][^"']*["']/, `algn="${algn}"`)
        : tag.replace(/>$/, ` algn="${algn}">`),
    )
  }
  if (patch.lineSpacingPct != null) {
    const ln = `<a:lnSpc><a:spcPct val="${Math.round(patch.lineSpacingPct * 1000)}"/></a:lnSpc>`
    level = /<a:lnSpc\b[^>]*>[\s\S]*?<\/a:lnSpc>/.test(level)
      ? level.replace(/<a:lnSpc\b[^>]*>[\s\S]*?<\/a:lnSpc>/, ln)
      : level.replace(/(<a:lvl1pPr\b[^>]*>)/, `$1${ln}`)
  }
  if (!/<a:defRPr\b/.test(level)) level = level.replace('</a:lvl1pPr>', '<a:defRPr/></a:lvl1pPr>')
  const defMatch = /<a:defRPr\b[^>]*\/>|<a:defRPr\b[^>]*>[\s\S]*?<\/a:defRPr>/.exec(level)![0]
  let def = defMatch.endsWith('/>') ? defMatch.replace(/\/>$/, '></a:defRPr>') : defMatch
  for (const [key, value] of [
    ['sz', patch.fontSizePt == null ? undefined : String(Math.round(patch.fontSizePt * 100))],
    ['b', patch.bold == null ? undefined : patch.bold ? '1' : '0'],
    ['i', patch.italic == null ? undefined : patch.italic ? '1' : '0'],
    ['u', patch.underline == null ? undefined : patch.underline ? 'sng' : 'none'],
  ] as const) {
    if (value === undefined) continue
    def = def.replace(/<a:defRPr\b[^>]*>/, (tag) =>
      new RegExp(`\\b${key}=["'][^"']*["']`).test(tag)
        ? tag.replace(new RegExp(`\\b${key}=["'][^"']*["']`), `${key}="${value}"`)
        : tag.replace(/>$/, ` ${key}="${value}">`),
    )
  }
  if (patch.color && /^#[0-9A-Fa-f]{6}$/.test(patch.color)) {
    const fill = `<a:solidFill><a:srgbClr val="${patch.color.slice(1).toUpperCase()}"/></a:solidFill>`
    def = /<a:solidFill\b[^>]*>[\s\S]*?<\/a:solidFill>/.test(def)
      ? def.replace(/<a:solidFill\b[^>]*>[\s\S]*?<\/a:solidFill>/, fill)
      : def.replace('</a:defRPr>', `${fill}</a:defRPr>`)
  }
  if (patch.fontFamily) {
    const family = patch.fontFamily.replace(
      /[&<>"']/g,
      (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[char]!,
    )
    def = /<a:latin\b[^>]*\/>/.test(def)
      ? def.replace(/<a:latin\b[^>]*\/>/, `<a:latin typeface="${family}"/>`)
      : def.replace('</a:defRPr>', `<a:latin typeface="${family}"/></a:defRPr>`)
  }
  level = level.replace(defMatch, () => def)
  style = style.replace(beforeLevel, () => level)
  const next = existingStyle
    ? original.replace(existingStyle, () => style)
    : /<p:extLst\b/.test(original)
      ? original.replace(/<p:extLst\b/, `${style}<p:extLst`)
      : original.replace('</p:notesMaster>', `${style}</p:notesMaster>`)
  if (next === original) return false
  archive.entries.set(path, Buffer.from(next, 'utf8'))
  return true
}

type NotesPlaceholderKind = 'sldImg' | 'body'

export function readNotesMasterPlaceholders(
  archive: PackageArchive,
  path: string,
): Record<NotesPlaceholderKind, boolean> | null {
  if (!/^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)) return null
  const xml = archive.readText(path)
  if (!xml) return null
  return {
    sldImg: !!findNotesMasterShape(xml, 'sldImg'),
    body: !!findNotesMasterShape(xml, 'body'),
  }
}

function findNotesMasterShape(xml: string, kind: NotesPlaceholderKind): string | undefined {
  return [...xml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)]
    .map((m) => m[0])
    .find((shape) => new RegExp(`<p:ph\\b[^>]*\\btype=["']${kind}["']`).test(shape))
}

/** PowerPoint's Notes Master checkbox adds/removes the corresponding p:sp. */
export function setNotesMasterPlaceholderEnabled(
  archive: PackageArchive,
  path: string,
  kind: NotesPlaceholderKind,
  enabled: boolean,
): boolean {
  if (!/^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)) return false
  const xml = archive.readText(path)
  if (!xml) return false
  const existing = findNotesMasterShape(xml, kind)
  if (!!existing === enabled) return false
  if (!enabled) {
    archive.entries.set(path, Buffer.from(xml.replace(existing!, ''), 'utf8'))
    return true
  }
  const id =
    Math.max(
      1,
      ...[...xml.matchAll(/<p:cNvPr\b[^>]*\bid=["'](\d+)["']/g)].map((m) => Number(m[1])),
    ) + 1
  const size = masterPartSize(archive, path, { cx: 6858000, cy: 9144000 })
  const rect = kind === 'sldImg' ? [0.1, 0.125, 0.8, 0.3375] : [0.1, 0.48125, 0.8, 0.39375]
  const [x, y, w, h] = rect.map((v, i) => Math.round(v * (i % 2 ? size.cy : size.cx)))
  const ph = kind === 'sldImg' ? 'type="sldImg" idx="2"' : 'type="body" sz="quarter" idx="3"'
  const body = kind === 'sldImg' ? '<a:bodyPr anchor="ctr"/>' : '<a:bodyPr/>'
  const shape = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${kind === 'sldImg' ? 'Slide Image' : 'Notes Body'} Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${ph}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>${body}<a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
  archive.entries.set(
    path,
    Buffer.from(
      xml.replace('</p:spTree>', (end) => `${shape}${end}`),
      'utf8',
    ),
  )
  return true
}

function baseName(p: string): string {
  return p.slice(p.lastIndexOf('/') + 1, -4)
}

/** A part's own image rels (same logic as index.ts partMediaRels, duplicated to avoid a circular dependency). */
function partMedia(archive: PackageArchive, partPath: string): Map<string, string> {
  const media = new Map<string, string>()
  for (const rel of archive.readRels(partPath).values()) {
    if (rel.type.endsWith('/image')) media.set(rel.id, resolveTarget(partPath, rel.target))
  }
  return media
}

/**
 * Parse a layout/master part into an editable Slide (its spTree structure is isomorphic
 * to a slide's, so scanSlide applies directly).
 * - master: text styles come from its own <p:txStyles>; placeholders always have an
 *   explicit xfrm, so no geometry inheritance is needed.
 * - layout: placeholder geometry/text styles/background inherit from its master; the
 *   master decoration layer is rendered underneath.
 */
export function parseMasterPart(archive: PackageArchive, partPath: string): Slide | null {
  const xml = archive.readText(partPath)
  if (!xml) return null
  const isMaster = !partPath.includes('/slideLayouts/')

  let masterPath: string | undefined = isMaster ? partPath : undefined
  if (!isMaster) {
    for (const rel of archive.readRels(partPath).values()) {
      if (rel.type.endsWith('/slideMaster')) {
        masterPath = resolveTarget(partPath, rel.target)
        break
      }
    }
  }

  const ctx: ParseContext = {}
  if (masterPath) {
    for (const rel of archive.readRels(masterPath).values()) {
      if (rel.type.endsWith('/theme')) {
        const themeXml = archive.readText(resolveTarget(masterPath, rel.target))
        if (themeXml) ctx.theme = parseTheme(themeXml)
        break
      }
    }
  }
  ctx.mediaRels = partMedia(archive, partPath)
  ctx.tableStyles = archive.readText('ppt/tableStyles.xml') ?? undefined

  const masterXml = !isMaster && masterPath ? archive.readText(masterPath) : undefined
  if (ctx.theme)
    ctx.theme.clrMap = isMaster ? parseClrMap(xml) : parseClrMap(masterXml ?? undefined, xml)
  if (isMaster) {
    ctx.masterTextStyles = parseMasterTextStyles(xml, ctx.theme, ctx.mediaRels)
  } else if (masterXml) {
    if (masterPath) ctx.masterMediaRels = partMedia(archive, masterPath)
    ctx.masterPlaceholders = parsePlaceholderMap(masterXml, ctx.theme, ctx.masterMediaRels)
    ctx.masterTextStyles = parseMasterTextStyles(masterXml, ctx.theme, ctx.masterMediaRels)
    ctx.masterBg = masterXml
  }

  const slide = parseSlide({ path: partPath, slideXml: xml, masterPath, ctx })

  // Layout view renders master concrete shapes underneath (master decorations stay visible when editing a layout)
  if (!isMaster && masterXml && masterPath) {
    if (!/<p:sldLayout\b[^>]*showMasterSp=(?:"(?:0|false)"|'(?:0|false)')/.test(xml)) {
      const dctx: ParseContext = { theme: ctx.theme, mediaRels: partMedia(archive, masterPath) }
      const dec = parseDecorations(masterXml, dctx, {})
      if (dec.length) slide.decorations = dec
    }
  }
  return slide
}
