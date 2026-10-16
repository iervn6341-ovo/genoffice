import type { Node as PmNode } from '@tiptap/pm/model'
import {
  convertNoteEntryXml,
  noteEntryXmlForModel,
  type NoteInfo,
  type ParsedDoc,
} from '@genoffice/docx-engine'

export type NoteKind = 'footnote' | 'endnote'

export interface NoteMarkerEdit {
  pos: number
  kind: NoteKind
  id: string
  rawRPr: string | null
}

export interface NoteConversionPlan {
  before: { footnotes: NoteInfo[]; endnotes: NoteInfo[] }
  after: { footnotes: NoteInfo[]; endnotes: NoteInfo[] }
  markers: NoteMarkerEdit[]
}

/** One-note conversion follows reference order, like Word Mac, while retaining exact entry XML for history. */
export function planNoteConversion(
  doc: PmNode,
  parsed: ParsedDoc,
  footnotes: readonly NoteInfo[],
  endnotes: readonly NoteInfo[],
  kind: NoteKind,
  id: string,
): NoteConversionPlan | null {
  const destination: NoteKind = kind === 'footnote' ? 'endnote' : 'footnote'
  const sourceNotes = kind === 'footnote' ? footnotes : endnotes
  const targetNotes = destination === 'footnote' ? footnotes : endnotes
  const source = sourceNotes.find((note) => note.id === id)
  if (!source || source.protectedContent || source.noRefMark) return null

  const refs: Array<{ kind: NoteKind; id: string; pos: number; rawRPr: string | null }> = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'docNoteRef')
      refs.push({
        kind: node.attrs.kind as NoteKind,
        id: String(node.attrs.id),
        pos,
        rawRPr: typeof node.attrs.rawRPr === 'string' ? node.attrs.rawRPr : null,
      })
    return true
  })
  const sourceRefs = refs.filter((ref) => ref.kind === kind && ref.id === id)
  if (sourceRefs.length !== 1) return null
  const targetRefs = refs.filter((ref) => ref.kind === destination)
  if (
    targetRefs.length !== targetNotes.length ||
    targetRefs.some(
      (ref) =>
        targetRefs.filter((other) => other.id === ref.id).length !== 1 ||
        !targetNotes.some((note) => note.id === ref.id),
    )
  )
    return null
  // Reindexing an opaque target entry could invalidate unmodeled internal references.
  if (targetNotes.some((note) => note.protectedContent)) return null

  const baseline = (noteKind: NoteKind, noteId: string) =>
    (noteKind === 'footnote' ? parsed.footnotes : parsed.endnotes).find(
      (note) => note.id === noteId,
    )
  const rawFor = (noteKind: NoteKind, note: NoteInfo): string | null => {
    if (note.historyEntryXml) return note.historyEntryXml
    const original = parsed.internal.noteEntryXml?.[noteKind]?.[note.id]
    const originalNote = baseline(noteKind, note.id)
    if (originalNote) {
      if (
        !original ||
        originalNote.text !== note.text ||
        originalNote.customMark !== note.customMark
      )
        return null
      return original
    }
    return noteEntryXmlForModel(noteKind, note)
  }
  const namespacesFor = (noteKind: NoteKind, note: NoteInfo) =>
    note.historyEntryNamespaces ?? parsed.internal.notePartNamespaces?.[noteKind] ?? {}
  const destinationNamespaces = parsed.internal.notePartNamespaces?.[destination] ?? {}
  if (
    Object.entries(namespacesFor(kind, source)).some(
      ([prefix, uri]) =>
        destinationNamespaces[prefix] !== undefined && destinationNamespaces[prefix] !== uri,
    )
  )
    return null
  const sourceRaw = rawFor(kind, source)
  if (!sourceRaw) return null
  const targetRaw = new Map<string, string>()
  for (const note of targetNotes) {
    const raw = rawFor(destination, note)
    if (!raw) return null
    targetRaw.set(note.id, raw)
  }

  const styleId = (name: string, type: 'paragraph' | 'character') =>
    [...parsed.styles.values()].find(
      (style) => style.type === type && style.name.toLowerCase() === name,
    )?.styleId
  const sourceName = kind === 'footnote' ? 'footnote' : 'endnote'
  const targetName = destination === 'footnote' ? 'footnote' : 'endnote'
  const styles = {
    sourceText: styleId(`${sourceName} text`, 'paragraph'),
    targetText: styleId(`${targetName} text`, 'paragraph'),
    sourceReference: styleId(`${sourceName} reference`, 'character'),
    targetReference: styleId(`${targetName} reference`, 'character'),
  }
  if (
    (/<w:pStyle\b/.test(sourceRaw) && (!styles.sourceText || !styles.targetText)) ||
    (/<w:rStyle\b/.test(sourceRaw) && (!styles.sourceReference || !styles.targetReference))
  )
    return null

  const convertedReferenceRPr = (rawRPr: string | null): string | null | undefined => {
    if (!rawRPr || !styles.sourceReference) return rawRPr
    const style = /<w:rStyle\b[^>]*\bw:val=(['"])([^'"]+)\1/.exec(rawRPr)
    if (!style || style[2] !== styles.sourceReference) return rawRPr
    if (!styles.targetReference) return undefined
    return rawRPr.replace(style[0], style[0].replace(style[2], styles.targetReference))
  }

  const orderedRefs = [...targetRefs, ...sourceRefs].sort((a, b) => a.pos - b.pos)
  const mapped = orderedRefs.map((ref, index) => {
    const moved = ref.kind === kind
    const rawRPr = moved ? convertedReferenceRPr(ref.rawRPr) : ref.rawRPr
    if (rawRPr === undefined) return null
    const original = moved ? source : targetNotes.find((note) => note.id === ref.id)!
    const raw = moved ? sourceRaw : targetRaw.get(ref.id)!
    const nextId = String(index + 1)
    const converted = convertNoteEntryXml(
      raw,
      moved ? kind : destination,
      destination,
      nextId,
      moved ? styles : undefined,
    )
    if (!converted) return null
    return {
      note: {
        ...original,
        id: nextId,
        historyEntryXml: converted,
        historyEntryNamespaces: namespacesFor(moved ? kind : destination, original),
        ...(moved && styles.targetText ? { styleId: styles.targetText } : {}),
      },
      marker: { pos: ref.pos, kind: destination, id: nextId, rawRPr },
    }
  })
  if (mapped.some((entry) => entry === null)) return null
  const beforeSource = sourceNotes.map((note) =>
    note.id === id
      ? {
          ...note,
          historyEntryXml: sourceRaw,
          historyEntryNamespaces: namespacesFor(kind, note),
        }
      : note,
  )
  const beforeTarget = targetNotes.map((note) => ({
    ...note,
    historyEntryXml: targetRaw.get(note.id)!,
    historyEntryNamespaces: namespacesFor(destination, note),
  }))
  const afterSource = sourceNotes.filter((note) => note.id !== id)
  const afterTarget = mapped.map((entry) => entry!.note)
  return {
    before:
      kind === 'footnote'
        ? { footnotes: beforeSource, endnotes: beforeTarget }
        : { footnotes: beforeTarget, endnotes: beforeSource },
    after:
      kind === 'footnote'
        ? { footnotes: afterSource, endnotes: afterTarget }
        : { footnotes: afterTarget, endnotes: afterSource },
    markers: mapped.map((entry) => entry!.marker),
  }
}
