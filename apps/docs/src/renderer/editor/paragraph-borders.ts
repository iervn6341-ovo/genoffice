import type { Editor } from '@tiptap/core'

/**
 * Paragraph borders as Word's Home → Borders menu edits them. The model keeps the drawn sides in
 * one string: t(op) b(ottom) l(eft) r(ight) and h = the rule between two paragraphs that share the
 * border (w:between). Each Bottom / Top / Left / Right entry toggles its own side, so the entries
 * combine (Bottom then Top gives both) and clicking an active one removes it.
 */
export type BorderSide = 't' | 'b' | 'l' | 'r' | 'h'

const ORDER: BorderSide[] = ['t', 'b', 'l', 'r', 'h']

/** normalized "tblrh" subset in a fixed order; null when nothing is drawn */
export function bordersOf(sides: Iterable<string>): string | null {
  const set = new Set(sides)
  const out = ORDER.filter((s) => set.has(s)).join('')
  return out || null
}

export function hasSide(borders: string | null | undefined, side: BorderSide): boolean {
  return !!borders && borders.includes(side)
}

export function toggleSide(borders: string | null | undefined, side: BorderSide): string | null {
  const set = new Set((borders ?? '').split(''))
  if (set.has(side)) set.delete(side)
  else set.add(side)
  return bordersOf(set)
}

export type BorderPreset = 'all' | 'outside' | 'inside' | 'none'

/**
 * The borders a preset leaves behind. "All" = box + between, "Outside" = the box only (the
 * between rule is dropped), "Inside" = only the between rule is added to what is there.
 */
export function presetBorders(
  borders: string | null | undefined,
  preset: BorderPreset,
): string | null {
  switch (preset) {
    case 'all':
      return 'tblrh'
    case 'outside':
      return 'tblr'
    case 'inside':
      return bordersOf([...(borders ?? '').split(''), 'h'])
    case 'none':
      return null
  }
}

export type BorderLines = Partial<
  Record<BorderSide, { color?: string; szPt?: number; spacePt?: number; val?: string }>
>

export function parseBorderLines(raw: unknown): BorderLines {
  if (typeof raw !== 'string' || !raw) return {}
  try {
    const v = JSON.parse(raw) as BorderLines
    return v && typeof v === 'object' ? v : {}
  } catch {
    return {}
  }
}

/**
 * Word's "Horizontal Line": an empty paragraph carrying only a bottom border, inserted after the
 * current paragraph, with the caret moved to a fresh paragraph below it.
 */
export function insertHorizontalLine(editor: Editor): void {
  const { $to } = editor.state.selection
  const after = $to.after($to.depth)
  editor
    .chain()
    .focus()
    .insertContentAt(after, [
      { type: 'docParagraph', attrs: { borders: 'b' } },
      { type: 'docParagraph' },
    ])
    .run()
  // the second inserted paragraph: right after the line paragraph
  const line = editor.state.doc.nodeAt(after)
  if (line) editor.commands.setTextSelection(after + line.nodeSize + 1)
}
