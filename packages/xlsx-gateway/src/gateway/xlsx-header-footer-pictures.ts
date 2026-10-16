import { ensureRelationshipNamespace } from './xlsx-namespace'

export class HeaderFooterPictureError extends Error {}

export interface HeaderFooterPictureEdit {
  readonly position: string
  /// null removes the slot; undefined keeps its existing image while changing size.
  readonly image?: { readonly mediaType: string; readonly base64: string } | null | undefined
  readonly widthPt?: number | undefined
  readonly heightPt?: number | undefined
}

interface MutablePicturePackage {
  paths(): Promise<readonly string[]>
  has(path: string): Promise<boolean>
  readText(path: string): Promise<string>
  write(path: string, content: string): void
  add(path: string, content: string): void
  addBinary(path: string, bytes: Uint8Array): void
}

const VML_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const RELS_EMPTY =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
const VML_EMPTY =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
  ' xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></xml>'
const PICTURE_SHAPE_TYPE =
  '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75"' +
  ' o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">' +
  '<v:stroke joinstyle="miter"/><v:formulas>' +
  [
    'if lineDrawn pixelLineWidth 0',
    'sum @0 1 0',
    'sum 0 0 @1',
    'prod @2 1 2',
    'prod @3 21600 pixelWidth',
    'prod @3 21600 pixelHeight',
    'sum @0 0 1',
    'prod @6 1 2',
    'prod @7 21600 pixelWidth',
    'sum @8 21600 0',
    'prod @7 21600 pixelHeight',
    'sum @10 21600 0',
  ]
    .map((eqn) => `<v:f eqn="${eqn}"/>`)
    .join('') +
  '</v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>' +
  '<o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'
const SLOT = /^[LCR][HF](?:EVEN|FIRST)?$/
const SHAPE = /<v:shape\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/v:shape>/g
const IMAGE_DATA = /<v:imagedata\b(?:[^>"']|"[^"]*"|'[^']*')*\/>/
const IMAGE_DATA_ALL = new RegExp(IMAGE_DATA.source, 'g')
const REL_TAG = /<Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|><\/Relationship>)/g
const MIME_EXTENSION: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
  'image/x-emz': 'emz',
  'image/x-wmz': 'wmz',
}

function attribute(tag: string, name: string): string | null {
  const match = new RegExp(`\\b${name}\\s*=\\s*(["'])([\\s\\S]*?)\\1`).exec(tag)
  return match?.[2] ?? null
}

function setAttribute(tag: string, name: string, value: string): string {
  const pattern = new RegExp(`\\b${name}\\s*=\\s*(["'])([\\s\\S]*?)\\1`)
  const match = pattern.exec(tag)
  if (match) return tag.replace(match[0], () => `${name}=${match[1]}${value}${match[1]}`)
  return tag.replace(/\s*\/?>(?=$)/, (end) => ` ${name}="${value}"${end}`)
}

function nextRid(xml: string): string {
  const ids = [...xml.matchAll(/\bId\s*=\s*["']rId(\d+)["']/g)].map((match) => Number(match[1]))
  return `rId${Math.max(0, ...ids) + 1}`
}

async function nextPath(pkg: MutablePicturePackage, extension: string): Promise<string> {
  for (let index = 1; index < 100_000; index += 1) {
    const path = `xl/media/hfPicture${index}.${extension}`
    if (!(await pkg.has(path))) return path
  }
  throw new HeaderFooterPictureError('No free header/footer media path.')
}

async function nextVmlPath(pkg: MutablePicturePackage): Promise<string> {
  for (let index = 1; index < 100_000; index += 1) {
    const path = `xl/drawings/vmlDrawing${index}.vml`
    if (!(await pkg.has(path))) return path
  }
  throw new HeaderFooterPictureError('No free header/footer VML path.')
}

function appendRel(xml: string, id: string, type: string, target: string): string {
  if (!xml.includes('</Relationships>'))
    throw new HeaderFooterPictureError('Invalid relationships part.')
  return xml.replace(
    '</Relationships>',
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`,
  )
}

function relativePath(basePath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = basePath.split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (part === '..') parts.pop()
    else if (part !== '.') parts.push(part)
  }
  return parts.join('/')
}

function relationshipTarget(xml: string, id: string, type: string): string | null {
  for (const match of xml.matchAll(REL_TAG)) {
    if (attribute(match[0], 'Id') === id && attribute(match[0], 'Type') === type) {
      return attribute(match[0], 'Target')
    }
  }
  return null
}

function ensureLegacyDrawingHF(xml: string, rid: string): string {
  if (/<legacyDrawingHF\b/.test(xml)) return xml
  const bound = ensureRelationshipNamespace(xml)
  const element = `<legacyDrawingHF r:id="${rid}"/>`
  const following =
    /<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b/.exec(bound)
  if (following) return bound.slice(0, following.index) + element + bound.slice(following.index)
  const end = bound.lastIndexOf('</worksheet>')
  if (end < 0) throw new HeaderFooterPictureError('Worksheet has no closing element.')
  return bound.slice(0, end) + element + bound.slice(end)
}

function imageBytes(image: NonNullable<HeaderFooterPictureEdit['image']>): {
  bytes: Uint8Array
  extension: string
} {
  const extension = MIME_EXTENSION[image.mediaType]
  if (
    !extension ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(image.base64)
  ) {
    throw new HeaderFooterPictureError('Unsupported header/footer picture data.')
  }
  const bytes = Uint8Array.from(Buffer.from(image.base64, 'base64'))
  if (bytes.length === 0 || bytes.length > 2 * 1024 * 1024) {
    throw new HeaderFooterPictureError('Header/footer picture exceeds the 2 MiB limit.')
  }
  const magic = Buffer.from(bytes)
  if (
    (extension === 'png' && !magic.subarray(0, 8).equals(Buffer.from('89504e470d0a1a0a', 'hex'))) ||
    (extension === 'jpeg' && !(magic[0] === 0xff && magic[1] === 0xd8 && magic[2] === 0xff)) ||
    (extension === 'gif' && !['GIF87a', 'GIF89a'].includes(magic.subarray(0, 6).toString('ascii')))
  ) {
    throw new HeaderFooterPictureError('Header/footer picture media type does not match its bytes.')
  }
  return { bytes, extension }
}

function sizedShape(shape: string, widthPt: number, heightPt: number): string {
  if (
    !Number.isFinite(widthPt) ||
    !Number.isFinite(heightPt) ||
    widthPt <= 0 ||
    heightPt <= 0 ||
    widthPt > 2_000 ||
    heightPt > 2_000
  ) {
    throw new HeaderFooterPictureError('Invalid header/footer picture size.')
  }
  const opening = /<v:shape\b(?:[^>"']|"[^"]*"|'[^']*')*>/.exec(shape)?.[0]
  if (!opening) throw new HeaderFooterPictureError('Invalid VML picture shape.')
  const prior = attribute(opening, 'style') ?? ''
  const other = prior.replace(/(?:^|;)\s*(?:width|height)\s*:[^;]*/gi, '').replace(/^;+|;+$/g, '')
  const style = `${other ? `${other};` : ''}width:${widthPt}pt;height:${heightPt}pt`
  return shape.replace(opening, () => setAttribute(opening, 'style', style))
}

function replaceShapeImage(shape: string, rid: string): string {
  const found = IMAGE_DATA.exec(shape)
  if (!found) throw new HeaderFooterPictureError('Existing VML picture has no image data.')
  let next = found[0]
  if (attribute(next, 'o:relid') !== null) next = setAttribute(next, 'o:relid', rid)
  if (attribute(next, 'r:id') !== null) next = setAttribute(next, 'r:id', rid)
  if (attribute(next, 'o:relid') === null && attribute(next, 'r:id') === null) {
    next = setAttribute(next, 'o:relid', rid)
  }
  return shape.replace(found[0], () => next)
}

function nextShapeId(vml: string): number {
  const numbers = [...vml.matchAll(/_x0000_s(\d+)/g)].map((match) => Number(match[1]))
  return Math.max(1024, ...numbers) + 1
}

function shapeFor(
  position: string,
  rid: string,
  widthPt: number,
  heightPt: number,
  spid: number,
): string {
  return `<v:shape id="${position}" o:spid="_x0000_s${spid}" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:${widthPt}pt;height:${heightPt}pt;z-index:1"><v:imagedata o:relid="${rid}"/><o:lock v:ext="edit" rotation="t"/></v:shape>`
}

function ensureVmlNamespaces(xml: string): string {
  const opening = /<xml\b(?:[^>"']|"[^"]*"|'[^']*')*>/.exec(xml)?.[0]
  if (!opening) throw new HeaderFooterPictureError('Invalid header/footer VML root.')
  let next = opening
  if (attribute(next, 'xmlns:v') === null)
    next = setAttribute(next, 'xmlns:v', 'urn:schemas-microsoft-com:vml')
  if (attribute(next, 'xmlns:o') === null)
    next = setAttribute(next, 'xmlns:o', 'urn:schemas-microsoft-com:office:office')
  return next === opening ? xml : xml.replace(opening, () => next)
}

function ensureContentTypes(xml: string, extensions: ReadonlySet<string>): string {
  let result = xml
  const contentTypes: Record<string, string> = {
    vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
    png: 'image/png',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    emf: 'image/x-emf',
    wmf: 'image/x-wmf',
    bmp: 'image/bmp',
    webp: 'image/webp',
    svg: 'image/svg+xml',
    emz: 'image/x-emz',
    wmz: 'image/x-wmz',
  }
  for (const extension of extensions) {
    if (new RegExp(`<Default\\b[^>]*\\bExtension=["']${extension}["']`).test(result)) continue
    result = result.replace(
      '</Types>',
      `<Default Extension="${extension}" ContentType="${contentTypes[extension]}"/></Types>`,
    )
  }
  return result
}

/** Surgically updates only selected VML header/footer slots. Other shapes, relationships and media survive. */
export async function applyHeaderFooterPictureEdits(
  pkg: MutablePicturePackage,
  worksheetPath: string,
  edits: readonly HeaderFooterPictureEdit[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (edits.length === 0) return
  const positions = new Set<string>()
  for (const edit of edits) {
    if (!SLOT.test(edit.position) || positions.has(edit.position)) {
      throw new HeaderFooterPictureError('Invalid or duplicate header/footer picture slot.')
    }
    positions.add(edit.position)
  }
  const worksheetRelsPath = worksheetPath.replace(/^(.*\/)([^/]+)$/, '$1_rels/$2.rels')
  let worksheetXml = await pkg.readText(worksheetPath)
  let worksheetRels = (await pkg.has(worksheetRelsPath))
    ? await pkg.readText(worksheetRelsPath)
    : RELS_EMPTY
  const legacyTag =
    /<legacyDrawingHF\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|><\/legacyDrawingHF>)/.exec(
      worksheetXml,
    )?.[0]
  const oldRid = legacyTag ? attribute(legacyTag, 'r:id') : null
  const target = oldRid ? relationshipTarget(worksheetRels, oldRid, VML_REL) : null
  if (legacyTag && (!oldRid || !target))
    throw new HeaderFooterPictureError('Header/footer VML relationship is missing.')
  let vmlPath = target ? relativePath(worksheetPath, target) : null
  if (vmlPath && !(await pkg.has(vmlPath)))
    throw new HeaderFooterPictureError('Header/footer VML part is missing.')
  if (vmlPath === null && edits.every((edit) => edit.image === null)) return
  let vml = vmlPath ? await pkg.readText(vmlPath) : VML_EMPTY
  let vmlRelsPath = vmlPath?.replace(/^(.*\/)([^/]+)$/, '$1_rels/$2.rels') ?? null
  let vmlRels =
    vmlRelsPath && (await pkg.has(vmlRelsPath)) ? await pkg.readText(vmlRelsPath) : RELS_EMPTY
  const replacedRelationships = new Set<string>()
  const extensions = new Set<string>()
  if (!vmlPath) {
    vmlPath = await nextVmlPath(pkg)
    vmlRelsPath = vmlPath.replace(/^(.*\/)([^/]+)$/, '$1_rels/$2.rels')
    const rid = nextRid(worksheetRels)
    worksheetRels = appendRel(
      worksheetRels,
      rid,
      VML_REL,
      `../drawings/${vmlPath.split('/').pop()}`,
    )
    worksheetXml = ensureLegacyDrawingHF(worksheetXml, rid)
    pkg.write(worksheetPath, worksheetXml)
    if (await pkg.has(worksheetRelsPath)) pkg.write(worksheetRelsPath, worksheetRels)
    else pkg.add(worksheetRelsPath, worksheetRels)
    touchedEntries.add(worksheetPath)
    touchedEntries.add(worksheetRelsPath)
    extensions.add('vml')
  }
  if (edits.some((edit) => edit.image !== null)) vml = ensureVmlNamespaces(vml)
  for (const edit of edits) {
    const found = [...vml.matchAll(SHAPE)].find((match) => {
      const opening = /<v:shape\b(?:[^>"']|"[^"]*"|'[^']*')*>/.exec(match[0])?.[0]
      return opening && attribute(opening, 'id')?.toUpperCase() === edit.position
    })
    if (found && edit.image !== undefined) {
      const imageData = IMAGE_DATA.exec(found[0])?.[0]
      const oldRid = imageData && (attribute(imageData, 'o:relid') ?? attribute(imageData, 'r:id'))
      if (oldRid) replacedRelationships.add(oldRid)
    }
    if (edit.image === null) {
      if (found) vml = vml.replace(found[0], '')
      continue
    }
    if (edit.widthPt === undefined || edit.heightPt === undefined) {
      throw new HeaderFooterPictureError('Header/footer picture dimensions are required.')
    }
    let rid: string | null = null
    if (edit.image !== undefined) {
      const { bytes, extension } = imageBytes(edit.image)
      const mediaPath = await nextPath(pkg, extension)
      pkg.addBinary(mediaPath, bytes)
      touchedEntries.add(mediaPath)
      extensions.add(extension)
      rid = nextRid(vmlRels)
      vmlRels = appendRel(vmlRels, rid, IMAGE_REL, `../media/${mediaPath.split('/').pop()}`)
    }
    if (found) {
      const replacement = sizedShape(
        rid ? replaceShapeImage(found[0], rid) : found[0],
        edit.widthPt,
        edit.heightPt,
      )
      vml = vml.replace(found[0], () => replacement)
    } else {
      if (!rid)
        throw new HeaderFooterPictureError(
          'Cannot size a header/footer picture that does not exist.',
        )
      if (!/\bid=["']_x0000_t75["']/.test(vml)) {
        vml = vml.replace('</xml>', `${PICTURE_SHAPE_TYPE}</xml>`)
      }
      vml = vml.replace(
        '</xml>',
        `${shapeFor(edit.position, rid, edit.widthPt, edit.heightPt, nextShapeId(vml))}</xml>`,
      )
    }
  }
  for (const rid of replacedRelationships) {
    const stillUsed = [...vml.matchAll(IMAGE_DATA_ALL)].some(
      (match) => attribute(match[0], 'o:relid') === rid || attribute(match[0], 'r:id') === rid,
    )
    if (stillUsed) continue
    vmlRels = vmlRels.replace(REL_TAG, (tag) =>
      attribute(tag, 'Id') === rid && attribute(tag, 'Type') === IMAGE_REL ? '' : tag,
    )
  }
  if (!vml.includes('</xml>')) throw new HeaderFooterPictureError('Invalid header/footer VML part.')
  if (await pkg.has(vmlPath)) pkg.write(vmlPath, vml)
  else pkg.add(vmlPath, vml)
  touchedEntries.add(vmlPath)
  if (vmlRelsPath && vmlRels !== RELS_EMPTY) {
    if (await pkg.has(vmlRelsPath)) pkg.write(vmlRelsPath, vmlRels)
    else pkg.add(vmlRelsPath, vmlRels)
    touchedEntries.add(vmlRelsPath)
  }
  if (extensions.size > 0) {
    const prior = await pkg.readText('[Content_Types].xml')
    const updated = ensureContentTypes(prior, extensions)
    if (updated !== prior) {
      pkg.write('[Content_Types].xml', updated)
      touchedEntries.add('[Content_Types].xml')
    }
  }
}
