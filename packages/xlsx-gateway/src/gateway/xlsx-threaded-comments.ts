/// Threaded comments (Excel 365 "Comments", as opposed to legacy notes). A worksheet relates to
/// `xl/threadedComments/threadedCommentN.xml` (a flat list: root comments and replies, replies
/// pointing at their root through `parentId`) and the workbook to `xl/persons/person.xml` (who
/// wrote what). Every thread keeps a legacy `<comment>` whose author is `tc={rootId}` so older
/// readers show a placeholder note; that placeholder is owned by this module and never surfaces
/// as a note.
///
/// Anything this module does not understand stays verbatim: a comment keeps its original element
/// (mentions, extLst, unknown attributes) until its text or done flag is edited, the part's root
/// start tag and every person element are kept as read.

import {
  appendRel,
  applySheetNotes,
  decodeXml,
  EMPTY_RELS,
  escapeXml,
  nextFreePath,
  nextFreeRid,
  NoteEditError,
  relTarget,
  removeRel,
  resolveRelTarget,
  worksheetRelsPath,
  type MutableNotePackage,
  type SheetNote,
} from './xlsx-notes'

export interface ThreadComment {
  /// `{GUID}` as in the file
  readonly id: string
  readonly personId: string
  /// ISO-8601 date-time exactly as stored (`dT`)
  readonly time: string
  readonly text: string
  /// Set on a reply (and on an orphaned comment whose parent is missing)
  readonly parentId?: string | undefined
  /// The element as read; absent once the text or done flag was edited
  readonly raw?: string | undefined
}

export interface SheetThread {
  /// Id of the root comment
  readonly id: string
  readonly row: number
  readonly column: number
  readonly done: boolean
  /// Root comment first, replies after it in file order
  readonly comments: readonly ThreadComment[]
}

export interface ThreadPerson {
  readonly id: string
  readonly displayName: string
  readonly userId: string
  readonly providerId: string
  /// The element as read
  readonly raw?: string | undefined
}

export const THREADED_REL_TYPE =
  'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment'
export const PERSON_REL_TYPE = 'http://schemas.microsoft.com/office/2017/10/relationships/person'
const THREADED_CONTENT_TYPE = 'application/vnd.ms-excel.threadedcomments+xml'
const PERSON_CONTENT_TYPE = 'application/vnd.ms-excel.person+xml'
const THREADED_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments'
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const CONTENT_TYPES_PATH = '[Content_Types].xml'
const WORKBOOK_PATH = 'xl/workbook.xml'
const WORKBOOK_RELS_PATH = 'xl/_rels/workbook.xml.rels'
const DEFAULT_PERSONS_PATH = 'xl/persons/person.xml'

export const THREAD_PLACEHOLDER_AUTHOR_PREFIX = 'tc='

const PLACEHOLDER_INTRO =
  '[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, ' +
  'any edits to it will get removed if the file is opened in a newer version of Excel. ' +
  'Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    '

/// The legacy note Excel keeps beside a thread so old readers see something.
export function placeholderNoteFor(thread: SheetThread): SheetNote {
  const root = thread.comments[0]
  return {
    row: thread.row,
    column: thread.column,
    author: `${THREAD_PLACEHOLDER_AUTHOR_PREFIX}${thread.id}`,
    text: `${PLACEHOLDER_INTRO}${(root?.text ?? '').replace(/\n/g, '\n    ')}`,
  }
}

export function isThreadPlaceholderAuthor(author: string): boolean {
  return author.startsWith(THREAD_PLACEHOLDER_AUTHOR_PREFIX)
}

function attributesOf(openTag: string): Map<string, string> {
  const attributes = new Map<string, string>()
  for (const match of openTag.matchAll(/([\w:.-]+)\s*=\s*"([^"]*)"/g)) {
    attributes.set(match[1]!, decodeXml(match[2]!))
  }
  return attributes
}

function parseAddress(reference: string): { row: number; column: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(reference.split(':')[0] ?? '')
  if (!match) return null
  let column = 0
  for (const letter of match[1]!.toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, column: column - 1 }
}

function columnName(index: number): string {
  let label = ''
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) {
    label = String.fromCharCode(65 + (i % 26)) + label
  }
  return label
}

const COMMENT_ELEMENT =
  /<(?:[\w.-]+:)?threadedComment\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:[\w.-]+:)?threadedComment>)/g
const PERSON_ELEMENT = /<(?:[\w.-]+:)?person\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:[\w.-]+:)?person>)/g

export interface ParsedThreads {
  readonly threads: SheetThread[]
  /// Comments whose `ref` is not a single cell (kept verbatim at the end of the part)
  readonly unparsed: readonly string[]
  readonly rootStartTag: string | null
}

export function parseThreadedComments(xml: string): ParsedThreads {
  const rootStartTag = /<(?:[\w.-]+:)?ThreadedComments\b[^>]*>/.exec(xml)?.[0] ?? null
  const threads: { thread: SheetThread; comments: ThreadComment[] }[] = []
  const byRoot = new Map<string, { comments: ThreadComment[] }>()
  const unparsed: string[] = []
  const replies: { comment: ThreadComment; ref: { row: number; column: number } }[] = []
  for (const match of xml.matchAll(COMMENT_ELEMENT)) {
    const raw = match[0]
    const openTag = /^<[^>]*>/.exec(raw)![0]
    const attributes = attributesOf(openTag)
    const id = attributes.get('id')
    const ref = parseAddress(attributes.get('ref') ?? '')
    if (!id || !ref) {
      unparsed.push(raw)
      continue
    }
    const text = decodeXml(
      /<(?:[\w.-]+:)?text\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?text>/.exec(raw)?.[1] ?? '',
    ).replace(/\r\n?/g, '\n')
    const comment: ThreadComment = {
      id,
      personId: attributes.get('personId') ?? '',
      time: attributes.get('dT') ?? '',
      text,
      parentId: attributes.get('parentId'),
      raw,
    }
    if (comment.parentId === undefined) {
      const entry = { comments: [comment] }
      byRoot.set(id, entry)
      threads.push({
        thread: {
          id,
          row: ref.row,
          column: ref.column,
          done: attributes.get('done') === '1' || attributes.get('done') === 'true',
          comments: [],
        },
        comments: entry.comments,
      })
    } else replies.push({ comment, ref })
  }
  for (const { comment, ref } of replies) {
    const root = byRoot.get(comment.parentId!)
    if (root) root.comments.push(comment)
    else {
      // a reply whose root is gone stays a thread of its own, parentId intact
      const entry = { comments: [comment] }
      byRoot.set(comment.id, entry)
      threads.push({
        thread: { id: comment.id, row: ref.row, column: ref.column, done: false, comments: [] },
        comments: entry.comments,
      })
    }
  }
  return {
    threads: threads.map(({ thread, comments }) => ({ ...thread, comments })),
    unparsed,
    rootStartTag,
  }
}

export function parsePersons(xml: string): ThreadPerson[] {
  const persons: ThreadPerson[] = []
  for (const match of xml.matchAll(PERSON_ELEMENT)) {
    const raw = match[0]
    const attributes = attributesOf(/^<[^>]*>/.exec(raw)![0])
    const id = attributes.get('id')
    if (!id) continue
    persons.push({
      id,
      displayName: attributes.get('displayName') ?? '',
      userId: attributes.get('userId') ?? '',
      providerId: attributes.get('providerId') ?? 'None',
      raw,
    })
  }
  return persons
}

function commentXml(thread: SheetThread, comment: ThreadComment, index: number): string {
  if (comment.raw !== undefined) return comment.raw
  const ref = `${columnName(thread.column)}${thread.row + 1}`
  const parent = comment.parentId === undefined ? '' : ` parentId="${escapeXml(comment.parentId)}"`
  const done = index === 0 && comment.parentId === undefined ? ` done="${thread.done ? 1 : 0}"` : ''
  return (
    `<threadedComment ref="${ref}" dT="${escapeXml(comment.time)}" ` +
    `personId="${escapeXml(comment.personId)}" id="${escapeXml(comment.id)}"${parent}${done}>` +
    `<text>${escapeXml(comment.text)}</text></threadedComment>`
  )
}

/// A root comment whose done flag flipped was already stripped of `raw` by the editor; this also
/// guards the stored element against a stale `done` attribute.
function rootWithDone(thread: SheetThread, comment: ThreadComment): ThreadComment {
  if (comment.raw === undefined) return comment
  const attributes = attributesOf(/^<[^>]*>/.exec(comment.raw)![0])
  const stored = attributes.get('done') === '1' || attributes.get('done') === 'true'
  return stored === thread.done ? comment : { ...comment, raw: undefined }
}

export function buildThreadedCommentsXml(
  threads: readonly SheetThread[],
  rootStartTag: string | null,
  unparsed: readonly string[] = [],
): string {
  const body = threads
    .flatMap((thread) =>
      thread.comments.map((comment, index) =>
        commentXml(thread, index === 0 ? rootWithDone(thread, comment) : comment, index),
      ),
    )
    .concat(unparsed)
    .join('')
  const start = rootStartTag ?? `<ThreadedComments xmlns="${THREADED_NS}" xmlns:x="${MAIN_NS}">`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `${start}${body}</${/^<([\w.-]+:)?ThreadedComments/.exec(start)?.[1] ?? ''}ThreadedComments>`
  )
}

export function buildPersonsXml(
  persons: readonly ThreadPerson[],
  rootStartTag: string | null,
): string {
  const body = persons
    .map(
      (person) =>
        person.raw ??
        `<person displayName="${escapeXml(person.displayName)}" id="${escapeXml(person.id)}" ` +
          `userId="${escapeXml(person.userId)}" providerId="${escapeXml(person.providerId)}"/>`,
    )
    .join('')
  const start = rootStartTag ?? `<personList xmlns="${THREADED_NS}" xmlns:x="${MAIN_NS}">`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `${start}${body}</${/^<([\w.-]+:)?personList/.exec(start)?.[1] ?? ''}personList>`
  )
}

interface ReadablePackage {
  has(path: string): Promise<boolean>
  readText(path: string): Promise<string>
}

export interface SheetThreadSet {
  readonly threads: SheetThread[]
  readonly persons: ThreadPerson[]
  /// The legacy note texts the placeholders replaced are never notes
  readonly placeholderRefs: readonly string[]
}

async function personsPathOf(pkg: ReadablePackage): Promise<string | null> {
  if (!(await pkg.has(WORKBOOK_RELS_PATH))) return null
  const target = relTarget(await pkg.readText(WORKBOOK_RELS_PATH), PERSON_REL_TYPE)
  return target === null ? null : resolveRelTarget(WORKBOOK_PATH, target)
}

export async function readPersons(pkg: ReadablePackage): Promise<ThreadPerson[]> {
  const path = await personsPathOf(pkg)
  if (path === null || !(await pkg.has(path))) return []
  return parsePersons(await pkg.readText(path))
}

/// The threads of one worksheet; empty when it has none.
export async function readSheetThreads(
  pkg: ReadablePackage,
  worksheetPath: string,
): Promise<SheetThread[]> {
  const relsPath = worksheetRelsPath(worksheetPath)
  if (!(await pkg.has(relsPath))) return []
  const target = relTarget(await pkg.readText(relsPath), THREADED_REL_TYPE)
  if (target === null) return []
  const path = resolveRelTarget(worksheetPath, target)
  if (!(await pkg.has(path))) return []
  return parseThreadedComments(await pkg.readText(path)).threads
}

function ensureOverride(contentTypes: string, partName: string, contentType: string): string {
  if (contentTypes.includes(`PartName="/${partName}"`)) return contentTypes
  return contentTypes.replace(
    '</Types>',
    `<Override PartName="/${partName}" ContentType="${contentType}"/></Types>`,
  )
}

function dropOverride(contentTypes: string, partName: string): string {
  const escaped = partName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return contentTypes.replace(new RegExp(`<Override\\b[^>]*PartName="/${escaped}"[^>]*/>`), '')
}

/// Replaces the worksheet's complete comment set: the notes, plus the threads and their legacy
/// placeholders. Persons are rewritten from `persons` (the full list the editor holds, which
/// starts from the file's own elements) only when the sheet has threads.
export async function applySheetThreads(
  pkg: MutableNotePackage,
  worksheetPath: string,
  notes: readonly SheetNote[],
  threads: readonly SheetThread[],
  persons: readonly ThreadPerson[],
  touchedEntries: Set<string>,
): Promise<void> {
  const occupied = new Set(notes.map((note) => `${note.row},${note.column}`))
  for (const thread of threads) {
    if (thread.comments.length === 0) throw new NoteEditError('A thread has no comments.')
    if (occupied.has(`${thread.row},${thread.column}`))
      throw new NoteEditError(
        `${columnName(thread.column)}${thread.row + 1} has both a note and a threaded comment.`,
      )
    occupied.add(`${thread.row},${thread.column}`)
  }
  await applySheetNotes(
    pkg,
    worksheetPath,
    [...notes, ...threads.map(placeholderNoteFor)],
    touchedEntries,
  )

  const relsPath = worksheetRelsPath(worksheetPath)
  const hasRels = await pkg.has(relsPath)
  let relsXml = hasRels ? await pkg.readText(relsPath) : EMPTY_RELS
  const existingTarget = relTarget(relsXml, THREADED_REL_TYPE)
  const existingPath =
    existingTarget === null ? null : resolveRelTarget(worksheetPath, existingTarget)

  if (threads.length === 0) {
    if (existingPath === null) return
    if (await pkg.has(existingPath)) pkg.remove(existingPath)
    touchedEntries.add(existingPath)
    relsXml = removeRel(relsXml, THREADED_REL_TYPE)
    pkg.write(relsPath, relsXml)
    touchedEntries.add(relsPath)
    const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
    const stripped = dropOverride(contentTypes, existingPath)
    if (stripped !== contentTypes) {
      pkg.write(CONTENT_TYPES_PATH, stripped)
      touchedEntries.add(CONTENT_TYPES_PATH)
    }
    return
  }

  let partPath = existingPath
  let rootStartTag: string | null = null
  let unparsed: readonly string[] = []
  if (partPath !== null && (await pkg.has(partPath))) {
    const parsed = parseThreadedComments(await pkg.readText(partPath))
    rootStartTag = parsed.rootStartTag
    unparsed = parsed.unparsed
  }
  const xml = buildThreadedCommentsXml(threads, rootStartTag, unparsed)
  if (partPath === null) {
    partPath = await nextFreePath(pkg, (index) => `xl/threadedComments/threadedComment${index}.xml`)
    relsXml = appendRel(
      relsXml,
      nextFreeRid(relsXml),
      THREADED_REL_TYPE,
      `../threadedComments/${partPath.split('/').pop()}`,
    )
    if (hasRels) pkg.write(relsPath, relsXml)
    else pkg.add(relsPath, relsXml)
    touchedEntries.add(relsPath)
    pkg.add(partPath, xml)
  } else if (await pkg.has(partPath)) pkg.write(partPath, xml)
  else pkg.add(partPath, xml)
  touchedEntries.add(partPath)

  // persons
  let personsPath = await personsPathOf(pkg)
  let personsStart: string | null = null
  if (personsPath !== null && (await pkg.has(personsPath))) {
    personsStart =
      /<(?:[\w.-]+:)?personList\b[^>]*>/.exec(await pkg.readText(personsPath))?.[0] ?? null
  }
  const personsXml = buildPersonsXml(persons, personsStart)
  let contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
  contentTypes = ensureOverride(contentTypes, partPath, THREADED_CONTENT_TYPE)
  if (personsPath === null) {
    personsPath = (await pkg.has(DEFAULT_PERSONS_PATH))
      ? await nextFreePath(pkg, (index) => `xl/persons/person${index}.xml`)
      : DEFAULT_PERSONS_PATH
    const workbookRels = (await pkg.has(WORKBOOK_RELS_PATH))
      ? await pkg.readText(WORKBOOK_RELS_PATH)
      : EMPTY_RELS
    const updated = appendRel(
      workbookRels,
      nextFreeRid(workbookRels),
      PERSON_REL_TYPE,
      personsPath.replace(/^xl\//, ''),
    )
    if (await pkg.has(WORKBOOK_RELS_PATH)) pkg.write(WORKBOOK_RELS_PATH, updated)
    else pkg.add(WORKBOOK_RELS_PATH, updated)
    touchedEntries.add(WORKBOOK_RELS_PATH)
    pkg.add(personsPath, personsXml)
  } else if (await pkg.has(personsPath)) pkg.write(personsPath, personsXml)
  else pkg.add(personsPath, personsXml)
  touchedEntries.add(personsPath)
  contentTypes = ensureOverride(contentTypes, personsPath, PERSON_CONTENT_TYPE)
  pkg.write(CONTENT_TYPES_PATH, contentTypes)
  touchedEntries.add(CONTENT_TYPES_PATH)
}

/// New thread helpers shared by the editor and the tests.
export function newGuid(random: () => number = Math.random): string {
  const hex = (count: number): string =>
    Array.from({ length: count }, () => Math.floor(random() * 16).toString(16)).join('')
  const variant = (8 + Math.floor(random() * 4)).toString(16)
  return `{${hex(8)}-${hex(4)}-4${hex(3)}-${variant}${hex(3)}-${hex(12)}}`.toUpperCase()
}
