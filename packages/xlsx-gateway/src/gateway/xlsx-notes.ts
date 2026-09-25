/// Legacy-comment (note) writer: replaces a worksheet's full comment set —
/// the comments part, the note shapes of the VML drawing Excel needs to show
/// them, the worksheet rels, the `<legacyDrawing>` element, and the
/// [Content_Types].xml entries. Non-note VML shapes (checkboxes, buttons)
/// survive a rewrite untouched.

import { ensureRelationshipNamespace } from './xlsx-namespace'

export class NoteEditError extends Error {}

export interface SheetNote {
  readonly row: number
  readonly column: number
  readonly author: string
  readonly text: string
}

interface MutableNotePackage {
  paths(): Promise<readonly string[]>
  has(path: string): Promise<boolean>
  readText(path: string): Promise<string>
  write(path: string, content: string): void
  add(path: string, content: string): void
  remove(path: string): void
}

const COMMENTS_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
const VML_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
const CONTENT_TYPES_PATH = '[Content_Types].xml'

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function columnName(index: number): string {
  let label = ''
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) {
    label = String.fromCharCode(65 + (i % 26)) + label
  }
  return label
}

function worksheetRelsPath(worksheetPath: string): string {
  return worksheetPath.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
}

function relTarget(relsXml: string, type: string): string | null {
  const pattern = new RegExp(`<Relationship\\b[^>]*Type="${type}"[^>]*/?>`)
  const found = pattern.exec(relsXml)
  if (!found) return null
  const target = / Target="([^"]*)"/.exec(found[0])
  return target?.[1] ?? null
}

/// "../comments1.xml" or "/xl/comments1.xml" → package path.
function resolveRelTarget(worksheetPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const base = worksheetPath.split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (part === '..') base.pop()
    else if (part !== '.') base.push(part)
  }
  return base.join('/')
}

function nextFreeRid(relsXml: string): string {
  const ids = [...relsXml.matchAll(/ Id="rId(\d+)"/g)].map((match) => Number(match[1]))
  return `rId${ids.length === 0 ? 1 : Math.max(...ids) + 1}`
}

async function nextFreePath(
  pkg: MutableNotePackage,
  template: (index: number) => string,
): Promise<string> {
  for (let index = 1; index < 10_000; index += 1) {
    const candidate = template(index)
    if (!(await pkg.has(candidate))) return candidate
  }
  throw new NoteEditError('No free part name for the comments part.')
}

function decodeXml(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (_, entity: string) => {
    if (entity[0] === '#') {
      const code = entity[1] === 'x' ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10)
      return String.fromCodePoint(code)
    }
    return { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }[entity]!
  })
}

/** A comment already in the file, with the plain text the importer derived from it */
interface ExistingComment {
  readonly ref: string
  readonly author: string
  readonly text: string
  readonly xml: string
}

/// The same reading as the sidecar's read_comments: every <t> of the comment
/// concatenated (rich runs and all), author by authorId, line breaks as \n.
function parseExistingComments(commentsXml: string): ExistingComment[] {
  const authorsBlock = /<authors\b[^>]*>([\s\S]*?)<\/authors>/.exec(commentsXml)?.[1] ?? ''
  const authors = [
    ...authorsBlock.matchAll(/<author\b[^>]*\/>|<author\b[^>]*>([\s\S]*?)<\/author>/g),
  ].map((match) => decodeXml(match[1] ?? ''))
  return [...commentsXml.matchAll(/<comment\b[^>]*>[\s\S]*?<\/comment>/g)].map((match) => {
    const xml = match[0]
    const authorId = Number(/\bauthorId="(\d+)"/.exec(xml)?.[1] ?? -1)
    const text = [...xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
      .map((t) => decodeXml(t[1] ?? ''))
      .join('')
      .replace(/\r\n?/g, '\n')
    return {
      ref: (/\bref="([^"]*)"/.exec(xml)?.[1] ?? '').replace(/\$/g, '').toUpperCase(),
      author: authors[authorId] ?? '',
      text,
      xml,
    }
  })
}

/**
 * The full comment set. A note whose author and text are unchanged keeps its
 * original <comment> (rich runs — the bold author line, colors — intact); the
 * editor only carries plain text, so rebuilding it would flatten every note on
 * the sheet whenever any one of them was edited.
 */
function buildCommentsXml(
  notes: readonly SheetNote[],
  existing: readonly ExistingComment[],
): string {
  const authors: string[] = []
  const authorId = (author: string): number => {
    const found = authors.indexOf(author)
    if (found !== -1) return found
    authors.push(author)
    return authors.length - 1
  }
  const comments = notes
    .map((note) => {
      const ref = `${columnName(note.column)}${note.row + 1}`
      const kept = existing.find(
        (comment) =>
          comment.ref === ref && comment.author === note.author && comment.text === note.text,
      )
      const id = authorId(note.author)
      if (kept) return kept.xml.replace(/\bauthorId="\d+"/, `authorId="${id}"`)
      return (
        `<comment ref="${ref}" authorId="${id}">` +
        `<text><t xml:space="preserve">${escapeXml(note.text)}</t></text></comment>`
      )
    })
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<authors>${authors.map((author) => `<author>${escapeXml(author)}</author>`).join('')}</authors>` +
    `<commentList>${comments}</commentList></comments>`
  )
}

const VML_HEADER =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
  ' xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
  '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"' +
  ' path="m,l,21600r21600,l21600,xe">' +
  '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>' +
  '</v:shapetype>'

function noteShape(note: SheetNote, index: number, shapeNumber: number): string {
  // Anchor: from one column right of the cell, spanning ~3 columns / 4 rows.
  const anchor = [note.column + 1, 15, note.row, 2, note.column + 4, 15, note.row + 4, 2].join(',')
  return (
    `<v:shape id="_x0000_s${shapeNumber}" type="#_x0000_t202"` +
    ' style="position:absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:60pt;' +
    `z-index:${index + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
    '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>' +
    '<v:path o:connecttype="none"/>' +
    '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
    '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
    `<x:Anchor>${anchor}</x:Anchor>` +
    '<x:AutoFill>False</x:AutoFill>' +
    `<x:Row>${note.row}</x:Row><x:Column>${note.column}</x:Column></x:ClientData>` +
    '</v:shape>'
  )
}

/// The note shapes of a VML drawing by cell ("row,column"), as Excel sized, placed and
/// showed or hid them.
function noteShapesByCell(vmlXml: string): Map<string, string> {
  const shapes = new Map<string, string>()
  for (const match of vmlXml.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)) {
    const shape = match[0]
    if (!shape.includes('ObjectType="Note"')) continue
    const row = /<x:Row>\s*(\d+)\s*<\/x:Row>/.exec(shape)?.[1]
    const column = /<x:Column>\s*(\d+)\s*<\/x:Column>/.exec(shape)?.[1]
    if (row !== undefined && column !== undefined)
      shapes.set(`${Number(row)},${Number(column)}`, shape)
  }
  return shapes
}

/// Note shapes for the new set: a cell that already had a note keeps its own shape
/// (an edited note keeps its box, like Excel); new notes get generated ones whose
/// ids avoid every shape id still in the drawing.
function noteShapesFor(notes: readonly SheetNote[], existingVml: string | null): string {
  const kept = existingVml === null ? new Map<string, string>() : noteShapesByCell(existingVml)
  const shapes = notes.map((note) => kept.get(`${note.row},${note.column}`) ?? null)
  const used = new Set<number>()
  const idSource = (existingVml === null ? '' : stripNoteShapes(existingVml)) + shapes.join('')
  for (const match of idSource.matchAll(/_x0000_s(\d+)/g)) used.add(Number(match[1]))
  let next = 1025
  return notes
    .map((note, index) => {
      const shape = shapes[index]
      if (shape) return shape
      while (used.has(next)) next += 1
      used.add(next)
      return noteShape(note, index, next)
    })
    .join('')
}

/// Drops every Note-typed shape, keeping other legacy objects verbatim.
function stripNoteShapes(vmlXml: string): string {
  return vmlXml.replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (shape) =>
    shape.includes('ObjectType="Note"') ? '' : shape,
  )
}

function ensureContentTypeOverride(contentTypes: string, partName: string): string {
  if (contentTypes.includes(`PartName="/${partName}"`)) return contentTypes
  return contentTypes.replace(
    '</Types>',
    `<Override PartName="/${partName}" ContentType="${COMMENTS_CONTENT_TYPE}"/></Types>`,
  )
}

function ensureVmlDefault(contentTypes: string): string {
  if (/<Default\b[^>]*Extension="vml"/.test(contentTypes)) return contentTypes
  return contentTypes.replace(
    '</Types>',
    '<Default Extension="vml"' +
      ' ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>',
  )
}

function appendRel(relsXml: string, id: string, type: string, target: string): string {
  return relsXml.replace(
    '</Relationships>',
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`,
  )
}

function removeRel(relsXml: string, type: string): string {
  return relsXml.replace(new RegExp(`<Relationship\\b[^>]*Type="${type}"[^>]*/?>`), '')
}

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '</Relationships>'

/// CT_Worksheet order: legacyDrawing follows drawing and precedes these.
const AFTER_LEGACY_DRAWING =
  /<legacyDrawingHF\b|<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b/

function ensureLegacyDrawingElement(worksheetXml: string, rid: string): string {
  if (/<legacyDrawing\b/.test(worksheetXml)) return worksheetXml
  worksheetXml = ensureRelationshipNamespace(worksheetXml)
  const element = `<legacyDrawing r:id="${rid}"/>`
  const anchor = AFTER_LEGACY_DRAWING.exec(worksheetXml)
  if (anchor) {
    return worksheetXml.slice(0, anchor.index) + element + worksheetXml.slice(anchor.index)
  }
  const end = worksheetXml.lastIndexOf('</worksheet>')
  if (end === -1) throw new NoteEditError('Worksheet has no closing element.')
  return worksheetXml.slice(0, end) + element + worksheetXml.slice(end)
}

/// Replaces the worksheet's whole comment set (empty list removes it).
export async function applySheetNotes(
  pkg: MutableNotePackage,
  worksheetPath: string,
  notes: readonly SheetNote[],
  touchedEntries: Set<string>,
): Promise<void> {
  const relsPath = worksheetRelsPath(worksheetPath)
  const hasRels = await pkg.has(relsPath)
  let relsXml = hasRels ? await pkg.readText(relsPath) : EMPTY_RELS
  const existingCommentsTarget = relTarget(relsXml, COMMENTS_REL_TYPE)
  const existingCommentsPath =
    existingCommentsTarget === null ? null : resolveRelTarget(worksheetPath, existingCommentsTarget)
  const existingVmlTarget = relTarget(relsXml, VML_REL_TYPE)
  const existingVmlPath =
    existingVmlTarget === null ? null : resolveRelTarget(worksheetPath, existingVmlTarget)
  let relsChanged = false

  if (notes.length === 0) {
    if (existingCommentsPath === null) return
    pkg.remove(existingCommentsPath)
    touchedEntries.add(existingCommentsPath)
    relsXml = removeRel(relsXml, COMMENTS_REL_TYPE)
    const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
    const stripped = contentTypes.replace(
      new RegExp(
        `<Override\\b[^>]*PartName="/${existingCommentsPath.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`,
      ),
      '',
    )
    if (stripped !== contentTypes) {
      pkg.write(CONTENT_TYPES_PATH, stripped)
      touchedEntries.add(CONTENT_TYPES_PATH)
    }
    if (existingVmlPath !== null && (await pkg.has(existingVmlPath))) {
      const remaining = stripNoteShapes(await pkg.readText(existingVmlPath))
      if (/<v:shape\b/.test(remaining)) {
        pkg.write(existingVmlPath, remaining)
      } else {
        pkg.remove(existingVmlPath)
        relsXml = removeRel(relsXml, VML_REL_TYPE)
        const worksheetXml = await pkg.readText(worksheetPath)
        const withoutLegacy = worksheetXml.replace(/<legacyDrawing\b[^>]*\/>/, '')
        if (withoutLegacy !== worksheetXml) {
          pkg.write(worksheetPath, withoutLegacy)
          touchedEntries.add(worksheetPath)
        }
      }
      touchedEntries.add(existingVmlPath)
    }
    pkg.write(relsPath, relsXml)
    touchedEntries.add(relsPath)
    return
  }

  // Comments part: rewrite in place, or allocate a fresh one.
  let commentsPath = existingCommentsPath
  if (commentsPath === null) {
    commentsPath = await nextFreePath(pkg, (index) => `xl/comments${index}.xml`)
    const rid = nextFreeRid(relsXml)
    const target = `../${commentsPath.replace(/^xl\//, '')}`
    relsXml = appendRel(relsXml, rid, COMMENTS_REL_TYPE, target)
    relsChanged = true
    pkg.add(commentsPath, buildCommentsXml(notes, []))
  } else {
    const existing = (await pkg.has(commentsPath))
      ? parseExistingComments(await pkg.readText(commentsPath))
      : []
    pkg.write(commentsPath, buildCommentsXml(notes, existing))
  }
  touchedEntries.add(commentsPath)

  // VML part: keep foreign shapes and the boxes of cells that keep a note.
  const existingVml =
    existingVmlPath !== null && (await pkg.has(existingVmlPath))
      ? await pkg.readText(existingVmlPath)
      : null
  const shapes = noteShapesFor(notes, existingVml)
  if (existingVmlPath !== null && existingVml !== null) {
    const vml = stripNoteShapes(existingVml)
    const end = vml.lastIndexOf('</xml>')
    if (end === -1) throw new NoteEditError(`${existingVmlPath} is not a VML drawing.`)
    pkg.write(existingVmlPath, vml.slice(0, end) + shapes + vml.slice(end))
    touchedEntries.add(existingVmlPath)
  } else {
    const vmlPath = await nextFreePath(pkg, (index) => `xl/drawings/vmlDrawing${index}.vml`)
    const rid = nextFreeRid(relsXml)
    relsXml = appendRel(relsXml, rid, VML_REL_TYPE, `../drawings/${vmlPath.split('/').pop()}`)
    relsChanged = true
    pkg.add(vmlPath, `${VML_HEADER}${shapes}</xml>`)
    touchedEntries.add(vmlPath)
    const worksheetXml = await pkg.readText(worksheetPath)
    const withLegacy = ensureLegacyDrawingElement(worksheetXml, rid)
    if (withLegacy !== worksheetXml) {
      pkg.write(worksheetPath, withLegacy)
      touchedEntries.add(worksheetPath)
    }
  }

  const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
  const updatedTypes = ensureVmlDefault(ensureContentTypeOverride(contentTypes, commentsPath))
  if (updatedTypes !== contentTypes) {
    pkg.write(CONTENT_TYPES_PATH, updatedTypes)
    touchedEntries.add(CONTENT_TYPES_PATH)
  }

  if (relsChanged) {
    if (hasRels) pkg.write(relsPath, relsXml)
    else pkg.add(relsPath, relsXml)
    touchedEntries.add(relsPath)
  }
}
