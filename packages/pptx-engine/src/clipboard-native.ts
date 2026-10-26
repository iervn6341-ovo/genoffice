/**
 * PowerPoint's native clipboard formats.
 *
 * PowerPoint (Mac and Windows) puts copied shapes on the system clipboard as an OPC zip
 * under the `Art::GVML ClipFormat` type (`com.microsoft.Art--GVML-ClipFormat` on macOS),
 * and copied tables as `Art::Table ClipFormat` (`com.microsoft.Art--Table-ClipFormat`).
 * Both are small packages: `clipboard/drawings/drawing1.xml` (+ rels), the source theme
 * as `clipboard/theme/theme1.xml` and any media/charts under `clipboard/`.
 *
 *  - GVML: `<a:graphic>` → lockedCanvas → one group whose children are DrawingML shapes in
 *    the `a:` namespace (`a:sp` with `a:txSp` text, `a:pic`, `a:cxnSp`, `a:grpSp`,
 *    `a:graphicFrame`). The group's chOff/chExt is the selection bounding box; the
 *    receiver centres it on the target slide.
 *  - Table: a bare `<a:graphic>` holding the `a:tbl`.
 *
 * Pasting into PowerPoint works from this package alone (verified against PowerPoint for
 * Mac, which rebuilds editable native shapes/tables), so writing it makes GenOffice
 * elements paste as editable objects instead of pictures; reading it lets a PowerPoint
 * copy come in as editable elements. Shapes are converted between the slide's `p:`
 * vocabulary (what ElementClipboardItem carries) and the clipboard's `a:` vocabulary.
 */
import JSZip from 'jszip'
import { collectPart } from './slide-transfer'
import { PackageArchive, relsPathFor, resolveTarget } from './zip'
import { escapeXmlAttr } from './xml-utils'
import type { ElementClipboardItem } from './index'

export const GVML_FORMAT = 'com.microsoft.Art--GVML-ClipFormat'
export const TABLE_CLIP_FORMAT = 'com.microsoft.Art--Table-ClipFormat'
/** Copied slides: a small pptx whose parts live under `clipboard/` instead of `ppt/` */
export const SLIDES_PACKAGE_FORMAT = 'com.microsoft.PowerPoint-14.0-Slides-Package'

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const LC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas'
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
const DRAWING_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

/** Namespaces slices commonly use inline-less (declared on the slide root, not on the element) */
const KNOWN_NS: Record<string, string> = {
  a: A_NS,
  r: R_NS,
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
  a15: 'http://schemas.microsoft.com/office/drawing/2012/main',
  asvg: 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
  adec: 'http://schemas.microsoft.com/office/drawing/2017/decorative',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  p15: 'http://schemas.microsoft.com/office/powerpoint/2012/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}

export interface Bbox {
  x: number
  y: number
  cx: number
  cy: number
}

// ── shared helpers ──────────────────────────────────────────────────────

/** Outer frame of one slice: `a:xfrm` (shape/pic/group) or `p:xfrm` (graphicFrame) */
export function itemFrame(xml: string): Bbox | null {
  const m =
    /<(?:a|p):xfrm\b[^>]*>\s*<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>\s*<a:ext\s+cx="(-?\d+)"\s+cy="(-?\d+)"\s*\/>/.exec(
      xml,
    )
  return m ? { x: +m[1]!, y: +m[2]!, cx: +m[3]!, cy: +m[4]! } : null
}

export function unionBbox(frames: readonly Bbox[]): Bbox | null {
  if (!frames.length) return null
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -Infinity
  let y1 = -Infinity
  for (const f of frames) {
    x0 = Math.min(x0, f.x)
    y0 = Math.min(y0, f.y)
    x1 = Math.max(x1, f.x + f.cx)
    y1 = Math.max(y1, f.y + f.cy)
  }
  return { x: x0, y: y0, cx: x1 - x0, cy: y1 - y0 }
}

/** Selection bounding box of a set of items (EMU, slide coordinates) */
export function itemsBbox(items: readonly ElementClipboardItem[]): Bbox | null {
  return unionBbox(items.map((i) => itemFrame(i.xml)).filter((f): f is Bbox => !!f))
}

/** mc:AlternateContent → the branch a consumer without extensions reads (Fallback, else the first Choice) */
function unwrapAlternate(xml: string): string {
  if (!xml.startsWith('<mc:AlternateContent')) return xml
  const fb = /<mc:Fallback\b[^>]*>([\s\S]*?)<\/mc:Fallback>/.exec(xml)
  if (fb) return fb[1]!.trim()
  const ch = /<mc:Choice\b[^>]*>([\s\S]*?)<\/mc:Choice>/.exec(xml)
  return ch ? ch[1]!.trim() : xml
}

/** The `<a:themeElements>` of a theme part wrapped as the clipboard theme */
export function clipboardThemeXml(themeXml: string | null | undefined): string {
  const elements = themeXml
    ? /<a:themeElements\b[\s\S]*?<\/a:themeElements>/.exec(themeXml)?.[0]
    : null
  const body = elements ?? FALLBACK_THEME_ELEMENTS
  return (
    XML_DECL +
    `<a:clipboardTheme xmlns:a="${A_NS}">${body}` +
    '<a:clrMap bg1="lt1" tx1="dk1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '</a:clipboardTheme>'
  )
}

const FALLBACK_THEME_ELEMENTS =
  '<a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
  '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements>'

/** xmlns declarations for every known prefix the markup uses (a is always declared) */
function namespaceDecls(xml: string, itemNs?: Readonly<Record<string, string>>): string {
  const namespaces = new Map<string, string>([['a', A_NS]])
  for (const [prefix, uri] of Object.entries({ ...KNOWN_NS, ...(itemNs ?? {}) })) {
    if (new RegExp(`<\\/?${prefix}:|\\s${prefix}:[A-Za-z]+=`).test(xml)) namespaces.set(prefix, uri)
  }
  return [...namespaces].map(([p, u]) => ` xmlns:${p}="${u}"`).join('')
}

// ── p: → a: (writing) ──────────────────────────────────────────────────

const GVML_TAGS = new Set([
  'sp',
  'nvSpPr',
  'cNvPr',
  'cNvSpPr',
  'spPr',
  'style',
  'pic',
  'nvPicPr',
  'cNvPicPr',
  'blipFill',
  'cxnSp',
  'nvCxnSpPr',
  'cNvCxnSpPr',
  'grpSp',
  'nvGrpSpPr',
  'cNvGrpSpPr',
  'grpSpPr',
  'graphicFrame',
  'nvGraphicFramePr',
  'cNvGraphicFramePr',
])

function renameToA(xml: string): string {
  return xml.replace(/<(\/?)p:([A-Za-z0-9]+)/g, (whole, slash: string, name: string) =>
    GVML_TAGS.has(name) ? `<${slash}a:${name}` : whole,
  )
}

/** `<p:sp>` → `<a:sp>`: text moves into `a:txSp`, the style reference after it */
function convertSp(sp: string): string {
  let body = sp
  const tx = /<p:txBody>[\s\S]*<\/p:txBody>/.exec(body)
  const style = /<p:style>[\s\S]*?<\/p:style>/.exec(body)
  if (tx) body = body.replace(tx[0], '')
  if (style) body = body.replace(style[0], '')
  const txSp = tx
    ? `<a:txSp><a:txBody>${tx[0].slice('<p:txBody>'.length, -'</p:txBody>'.length)}</a:txBody><a:useSpRect/></a:txSp>`
    : ''
  body = body.replace(/\s+useBgFill="[^"]*"/, '')
  // a:sp = nvSpPr, spPr, txSp, style
  return body.replace(/<\/p:sp>$/, `${txSp}${style ? style[0] : ''}</p:sp>`)
}

/** `<p:graphicFrame>` → `<a:graphicFrame>`: the transform follows the graphic */
function convertGraphicFrame(gf: string): string {
  const xfrm = /<p:xfrm\b[\s\S]*?<\/p:xfrm>/.exec(gf)
  let body = gf
  if (xfrm) body = body.replace(xfrm[0], '')
  const moved = xfrm
    ? xfrm[0].replace(/^<p:xfrm/, '<a:xfrm').replace(/<\/p:xfrm>$/, '</a:xfrm>')
    : ''
  return body.replace(/<\/p:graphicFrame>$/, `${moved}</p:graphicFrame>`)
}

/** Slice in `p:` vocabulary → the clipboard's `a:` vocabulary; null when it has no representation */
export function sliceToGvml(rawXml: string): string | null {
  let xml = unwrapAlternate(rawXml)
  if (/<(?:dgm|dsp):/.test(xml)) return null
  if (/<p:oleObj\b|<p:contentPart\b/.test(xml)) return null
  xml = xml.replace(/<p:nvPr\s*\/>|<p:nvPr>[\s\S]*?<\/p:nvPr>/g, '')
  xml = xml.replace(/<p:custDataLst>[\s\S]*?<\/p:custDataLst>/g, '')
  xml = xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, convertSp)
  xml = xml.replace(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g, convertGraphicFrame)
  // a picture carries no style reference in GVML
  xml = xml.replace(/(<p:pic\b[\s\S]*?)<p:style>[\s\S]*?<\/p:style>([\s\S]*?<\/p:pic>)/g, '$1$2')
  xml = renameToA(xml)
  return /<p:[A-Za-z]/.test(xml) ? null : xml
}

/** Where a source part lands inside the clipboard package */
function clipboardPath(sourcePath: string): string {
  return sourcePath.startsWith('ppt/')
    ? `clipboard/${sourcePath.slice(4)}`
    : `clipboard/${sourcePath}`
}

function relTarget(fromPart: string, toPath: string): string {
  const from = fromPart.slice(0, fromPart.lastIndexOf('/')).split('/').filter(Boolean)
  const to = toPath.split('/').filter(Boolean)
  let shared = 0
  while (shared < from.length && shared < to.length && from[shared] === to[shared]) shared += 1
  return [...Array.from({ length: from.length - shared }, () => '..'), ...to.slice(shared)].join(
    '/',
  )
}

interface PackageParts {
  /** zip path → bytes */
  files: Map<string, Uint8Array>
  /** zip path → content type override */
  overrides: Map<string, string>
  /** extension → content type default */
  defaults: Map<string, string>
}

function newParts(): PackageParts {
  return {
    files: new Map(),
    overrides: new Map([
      ['clipboard/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
    ]),
    defaults: new Map([
      ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
      ['xml', 'application/xml'],
    ]),
  }
}

/**
 * Land an item's dependency parts under `clipboard/`, rewriting each part's own rels to the
 * moved paths. Returns source path → clipboard path.
 */
function landParts(
  parts: PackageParts,
  items: readonly ElementClipboardItem[],
): Map<string, string> {
  const map = new Map<string, string>()
  const enc = new TextEncoder()
  for (const item of items) {
    const bundle = item.parts ?? {}
    for (const sourcePath of Object.keys(bundle)) {
      if (sourcePath.includes('/_rels/')) continue
      map.set(sourcePath, clipboardPath(sourcePath))
    }
    for (const [sourcePath, base64] of Object.entries(bundle)) {
      if (sourcePath.includes('/_rels/')) continue
      const dest = map.get(sourcePath)!
      parts.files.set(dest, new Uint8Array(Buffer.from(base64, 'base64')))
      const relsB64 = bundle[relsPathFor(sourcePath)]
      if (relsB64) {
        const relsXml = Buffer.from(relsB64, 'base64')
          .toString('utf8')
          .replaceAll(
            /Target="([^"]+)"(?![^>]*TargetMode="External")/g,
            (whole, target: string) => {
              if (/^[a-z]+:/i.test(target)) return whole
              const abs = resolveTarget(sourcePath, target)
              return `Target="${escapeXmlAttr(relTarget(dest, clipboardPath(abs)))}"`
            },
          )
        parts.files.set(relsPathFor(dest), enc.encode(relsXml))
      }
    }
    for (const [key, ct] of Object.entries(item.contentTypes ?? {})) {
      if (key.includes('/')) parts.overrides.set(clipboardPath(key), ct)
      else parts.defaults.set(key, ct)
    }
  }
  return map
}

async function zipPackage(parts: PackageParts): Promise<Uint8Array> {
  const zip = new JSZip()
  const ct =
    XML_DECL +
    `<Types xmlns="${CT_NS}">` +
    [...parts.defaults].map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join('') +
    [...parts.overrides]
      .map(([p, t]) => `<Override PartName="/${p}" ContentType="${t}"/>`)
      .join('') +
    '</Types>'
  zip.file('[Content_Types].xml', ct)
  zip.file(
    '_rels/.rels',
    XML_DECL +
      `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${DRAWING_REL}" Target="clipboard/drawings/drawing1.xml"/></Relationships>`,
  )
  for (const [path, bytes] of parts.files) zip.file(path, bytes)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

export interface NativeClipboardBuild {
  /** GVML package for the shapes (absent when none can be expressed) */
  gvml?: Uint8Array
  /** Table package, when exactly one table is copied */
  table?: Uint8Array
  /** Items GVML could not carry (SmartArt data models, OLE, …): the picture fallback covers them */
  skipped: number
}

/**
 * Build the clipboard packages for a set of copied elements. A lone table goes out in the
 * table format (as PowerPoint does); everything else is one GVML group.
 */
export async function buildNativeClipboard(
  items: readonly ElementClipboardItem[],
  themeXml?: string | null,
): Promise<NativeClipboardBuild> {
  const enc = new TextEncoder()
  const theme = clipboardThemeXml(themeXml)
  const isTable = (i: ElementClipboardItem) => /graphicData\s+uri="[^"]*\/table"/.test(i.xml)

  if (items.length === 1 && isTable(items[0]!)) {
    const graphic = /<a:graphic\b[\s\S]*<\/a:graphic>/.exec(items[0]!.xml)?.[0]
    if (!graphic) return { skipped: 1 }
    const parts = newParts()
    const drawing =
      XML_DECL +
      graphic.replace(/^<a:graphic\b[^>]*>/, `<a:graphic${namespaceDecls(graphic, items[0]!.ns)}>`)
    parts.files.set('clipboard/drawings/drawing1.xml', enc.encode(drawing))
    parts.files.set(
      'clipboard/drawings/_rels/drawing1.xml.rels',
      enc.encode(
        XML_DECL +
          `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${THEME_REL}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
    )
    parts.files.set('clipboard/theme/theme1.xml', enc.encode(theme))
    parts.overrides.set(
      'clipboard/drawings/drawing1.xml',
      'application/vnd.openxmlformats-officedocument.drawing+xml',
    )
    return { table: await zipPackage(parts), skipped: 0 }
  }

  const parts = newParts()
  const drawingRels: string[] = []
  const relIdByKey = new Map<string, string>()
  let nextRid = 1
  const landed = landParts(parts, items)
  const children: string[] = []
  const frames: Bbox[] = []
  const namespaces = new Map<string, string>([['a', A_NS]])
  let skipped = 0
  for (const item of items) {
    let xml = sliceToGvml(item.xml)
    if (xml == null) {
      skipped++
      continue
    }
    // Slide jumps point into the source deck — meaningless on a clipboard
    const slideRids = new Set(item.rels.filter((r) => r.type.endsWith('/slide')).map((r) => r.rid))
    if (slideRids.size) {
      xml = xml.replace(
        /<a:hlink(?:Click|Hover)\b[^>]*\br:id="(rId\d+)"[^>]*(?:\/>|>[\s\S]*?<\/a:hlink(?:Click|Hover)>)/g,
        (whole, rid: string) => (slideRids.has(rid) ? '' : whole),
      )
    }
    const ridMap = new Map<string, string>()
    for (const rel of item.rels) {
      if (slideRids.has(rel.rid)) continue
      const target = rel.external ? rel.target : landed.get(rel.target)
      if (!target) continue
      const key = `${rel.type} ${target}`
      let rid = relIdByKey.get(key)
      if (!rid) {
        rid = `rId${nextRid++}`
        relIdByKey.set(key, rid)
        drawingRels.push(
          rel.external
            ? `<Relationship Id="${rid}" Type="${rel.type}" Target="${escapeXmlAttr(target)}" TargetMode="External"/>`
            : `<Relationship Id="${rid}" Type="${rel.type}" Target="${escapeXmlAttr(relTarget('clipboard/drawings/drawing1.xml', target))}"/>`,
        )
      }
      ridMap.set(rel.rid, rid)
    }
    xml = xml.replace(
      /\br:(embed|link|id|pict|dm|lo|qs|cs)="(rId\d+)"/g,
      (whole, attr: string, rid: string) =>
        ridMap.has(rid) ? `r:${attr}="${ridMap.get(rid)}"` : whole,
    )
    for (const [prefix, uri] of Object.entries({ ...KNOWN_NS, ...(item.ns ?? {}) })) {
      if (new RegExp(`<\\/?${prefix}:|\\s${prefix}:[A-Za-z]+=`).test(xml))
        namespaces.set(prefix, uri)
    }
    const f = itemFrame(item.xml)
    if (f) frames.push(f)
    children.push(xml)
  }
  if (!children.length) return { skipped }
  const bbox = unionBbox(frames) ?? { x: 0, y: 0, cx: 0, cy: 0 }
  const themeRid = `rId${nextRid}`
  drawingRels.push(
    `<Relationship Id="${themeRid}" Type="${THEME_REL}" Target="../theme/theme1.xml"/>`,
  )
  const nsDecl = [...namespaces].map(([p, u]) => ` xmlns:${p}="${u}"`).join('')
  const drawing =
    XML_DECL +
    `<a:graphic${nsDecl}><a:graphicData uri="${LC_NS}"><lc:lockedCanvas xmlns:lc="${LC_NS}">` +
    '<a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr>' +
    `<a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${bbox.cx}" cy="${bbox.cy}"/><a:chOff x="${bbox.x}" y="${bbox.y}"/><a:chExt cx="${bbox.cx}" cy="${bbox.cy}"/></a:xfrm></a:grpSpPr>` +
    `${children.join('')}</lc:lockedCanvas></a:graphicData></a:graphic>`
  parts.files.set('clipboard/drawings/drawing1.xml', enc.encode(drawing))
  parts.files.set(
    'clipboard/drawings/_rels/drawing1.xml.rels',
    enc.encode(
      XML_DECL + `<Relationships xmlns="${REL_NS}">${drawingRels.join('')}</Relationships>`,
    ),
  )
  parts.files.set('clipboard/theme/theme1.xml', enc.encode(theme))
  parts.overrides.set(
    'clipboard/drawings/drawing1.xml',
    'application/vnd.openxmlformats-officedocument.drawing+xml',
  )
  return { gvml: await zipPackage(parts), skipped }
}

// ── a: → p: (reading) ──────────────────────────────────────────────────

/** Inverse of renameToA for the vocabulary above */
function renameToP(xml: string): string {
  return xml.replace(/<(\/?)a:([A-Za-z0-9]+)/g, (whole, slash: string, name: string) =>
    GVML_TAGS.has(name) ? `<${slash}p:${name}` : whole,
  )
}

function convertTxSp(sp: string): string {
  const tx =
    /<a:txSp>\s*(<a:txBody>[\s\S]*?<\/a:txBody>)\s*(?:<a:useSpRect\s*\/>)?\s*<\/a:txSp>/.exec(sp)
  if (!tx) return sp
  const body = tx[1]!.replace(/^<a:txBody>/, '<p:txBody>').replace(/<\/a:txBody>$/, '</p:txBody>')
  const without = sp.replace(tx[0], '')
  return without.replace(/<\/a:sp>$/, `${body}</a:sp>`)
}

function addNvPr(xml: string): string {
  return xml.replace(/<\/p:nv(SpPr|PicPr|CxnSpPr|GrpSpPr|GraphicFramePr)>/g, '<p:nvPr/></p:nv$1>')
}

function convertFrameBack(gf: string): string {
  // nvGraphicFramePr, graphic, xfrm → nvGraphicFramePr, xfrm, graphic
  const at = gf.lastIndexOf('<a:xfrm')
  const graphicEnd = gf.lastIndexOf('</a:graphic>')
  if (at < 0 || graphicEnd < 0 || at < graphicEnd) return gf
  const close = gf.indexOf('</a:xfrm>', at)
  if (close < 0) return gf
  const xfrm = gf.slice(at, close + '</a:xfrm>'.length)
  const body = gf.slice(0, at) + gf.slice(close + '</a:xfrm>'.length)
  const moved = xfrm.replace(/^<a:xfrm/, '<p:xfrm').replace(/<\/a:xfrm>$/, '</p:xfrm>')
  return body.replace('</a:nvGraphicFramePr>', `</a:nvGraphicFramePr>${moved}`)
}

/** Split the children of a lockedCanvas group into top-level element slices */
function topLevelElements(xml: string): string[] {
  const out: string[] = []
  let i = 0
  while (i < xml.length) {
    const open = /<a:(sp|pic|cxnSp|grpSp|graphicFrame)\b/.exec(xml.slice(i))
    if (!open) break
    const start = i + open.index
    const tag = open[1]!
    // depth-aware close (grpSp nests)
    let depth = 0
    const re = new RegExp(`<a:${tag}\\b[^>]*?(/?)>|</a:${tag}>`, 'g')
    re.lastIndex = start
    let m: RegExpExecArray | null
    let end = -1
    while ((m = re.exec(xml))) {
      if (m[0].startsWith('</')) depth--
      else if (m[1] !== '/') depth++
      if (depth === 0) {
        end = m.index + m[0].length
        break
      }
    }
    if (end < 0) break
    out.push(xml.slice(start, end))
    i = end
  }
  return out
}

export interface ParsedNativeClipboard {
  items: ElementClipboardItem[]
  /** Selection frame in the package's own coordinates (EMU) */
  bbox: Bbox
  /** Source theme (clipboardTheme part), for callers that want to compare */
  themeXml?: string
}

async function openPackage(
  bytes: Uint8Array,
): Promise<{ archive: PackageArchive; drawingPath: string } | null> {
  let archive: PackageArchive
  try {
    archive = await PackageArchive.open(bytes)
  } catch {
    return null
  }
  const rootRels = archive.readText('_rels/.rels') ?? ''
  const target = /Type="[^"]*\/drawing"[^>]*Target="([^"]+)"/.exec(rootRels)?.[1]
  const drawingPath = target ? target.replace(/^\//, '') : 'clipboard/drawings/drawing1.xml'
  return archive.has(drawingPath) ? { archive, drawingPath } : null
}

/** Dependencies of the rIds a slice references → rels + parts (renamed `clipboard/…` → `ppt/…`) */
function bundleDependencies(
  archive: PackageArchive,
  drawingPath: string,
  xml: string,
): Pick<ElementClipboardItem, 'rels' | 'parts' | 'contentTypes'> {
  const drawingRels = archive.readRels(drawingPath)
  const rels: ElementClipboardItem['rels'] = []
  const parts: Record<string, string> = {}
  const contentTypes: Record<string, string> = {}
  const seen = new Set<string>()
  const rename = (p: string) =>
    p.startsWith('clipboard/') ? `ppt/${p.slice('clipboard/'.length)}` : p
  for (const m of xml.matchAll(/\br:(?:embed|link|id|pict|dm|lo|qs|cs)="(rId\d+)"/g)) {
    const rid = m[1]!
    if (seen.has(rid)) continue
    seen.add(rid)
    const rel = drawingRels.get(rid)
    if (!rel) continue
    if (rel.targetMode === 'External') {
      rels.push({ rid, type: rel.type, target: rel.target, external: true })
      continue
    }
    const abs = resolveTarget(drawingPath, rel.target)
    const bundle: Record<string, string> = {}
    const cts: Record<string, string> = {}
    collectPart(archive, abs, bundle, cts)
    for (const [k, v] of Object.entries(bundle)) parts[rename(k)] = v
    for (const [k, v] of Object.entries(cts)) contentTypes[k.includes('/') ? rename(k) : k] = v
    rels.push({ rid, type: rel.type, target: rename(abs) })
  }
  return {
    rels,
    ...(Object.keys(parts).length ? { parts, contentTypes } : {}),
  }
}

const NEXT_ID_BASE = 1

/**
 * Read a GVML package into clipboard items (`p:` vocabulary, ids unique inside the package).
 */
export async function parseGvml(bytes: Uint8Array): Promise<ParsedNativeClipboard | null> {
  const pkg = await openPackage(bytes)
  if (!pkg) return null
  const { archive, drawingPath } = pkg
  const drawing = archive.readText(drawingPath)
  if (!drawing) return null
  const canvas = /<lc:lockedCanvas\b[^>]*>([\s\S]*)<\/lc:lockedCanvas>/.exec(drawing)
  if (!canvas) return null
  const inner = canvas[1]!
  const grp =
    /<a:grpSpPr>\s*<a:xfrm>[\s\S]*?<a:chOff\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>\s*<a:chExt\s+cx="(\d+)"\s+cy="(\d+)"/.exec(
      inner,
    )
  const nsDecls: Record<string, string> = {}
  for (const m of drawing.matchAll(/\bxmlns:([A-Za-z0-9]+)="([^"]+)"/g)) nsDecls[m[1]!] = m[2]!
  const slices = topLevelElements(inner.replace(/^[\s\S]*?<\/a:grpSpPr>/, ''))
  const items: ElementClipboardItem[] = []
  let nextId = NEXT_ID_BASE
  for (const raw of slices) {
    let xml = raw
    xml = xml.replace(/<a:sp\b[\s\S]*?<\/a:sp>/g, convertTxSp)
    xml = xml.replace(/<a:graphicFrame\b[\s\S]*?<\/a:graphicFrame>/g, convertFrameBack)
    xml = renameToP(xml)
    xml = addNvPr(xml)
    // ids inside one clipboard package are unique; make sure the root has one
    xml = xml.replace(/<p:cNvPr\s+id="(\d+)"/, (whole, id: string) =>
      id === '0' ? `<p:cNvPr id="${nextId++}"` : whole,
    )
    const deps = bundleDependencies(archive, drawingPath, xml)
    items.push({ xml, ...deps, ns: nsDecls })
  }
  if (!items.length) return null
  const frames = items.map((i) => itemFrame(i.xml)).filter((f): f is Bbox => !!f)
  const bbox: Bbox = grp
    ? { x: +grp[1]!, y: +grp[2]!, cx: +grp[3]!, cy: +grp[4]! }
    : (unionBbox(frames) ?? { x: 0, y: 0, cx: 0, cy: 0 })
  const themeXml = archive.readText('clipboard/theme/theme1.xml') ?? undefined
  return { items, bbox, ...(themeXml ? { themeXml } : {}) }
}

/** Read a table clip package into a graphicFrame item */
export async function parseTableClip(bytes: Uint8Array): Promise<ParsedNativeClipboard | null> {
  const pkg = await openPackage(bytes)
  if (!pkg) return null
  const drawing = pkg.archive.readText(pkg.drawingPath)
  const graphic = drawing ? /<a:graphic\b[\s\S]*<\/a:graphic>/.exec(drawing)?.[0] : null
  if (!graphic || !/<a:tbl>/.test(graphic)) return null
  const cols = [...graphic.matchAll(/<a:gridCol\b[^>]*\bw="(\d+)"/g)].map((m) => +m[1]!)
  const rows = [...graphic.matchAll(/<a:tr\b[^>]*\bh="(\d+)"/g)].map((m) => +m[1]!)
  const cx = cols.reduce((a, b) => a + b, 0)
  const cy = rows.reduce((a, b) => a + b, 0)
  const xml =
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>${graphic.replace(/^<a:graphic\b[^>]*>/, `<a:graphic>`)}</p:graphicFrame>`
  const styleId = /<a:tableStyleId>([^<]+)<\/a:tableStyleId>/.exec(graphic)?.[1]
  const themeXml = pkg.archive.readText('clipboard/theme/theme1.xml') ?? undefined
  return {
    items: [
      {
        xml,
        rels: [],
        ...(styleId ? { tableStyles: [{ id: styleId }] } : {}),
        ns: { a: A_NS, r: R_NS },
      },
    ],
    bbox: { x: 0, y: 0, cx, cy },
    ...(themeXml ? { themeXml } : {}),
  }
}
