import { describe, expect, it } from 'vitest'

import { markdownHeading, typedTableAt } from '../src/renderer/editor/markdown-shortcut'

describe('markdownHeading', () => {
  it('reads the level from the typed hashes', () => {
    expect(markdownHeading('#')).toEqual({ level: 1, markerLength: 1 })
    expect(markdownHeading('##')).toEqual({ level: 2, markerLength: 2 })
    expect(markdownHeading('###### ')).toEqual({ level: 6, markerLength: 7 })
  })

  it('drops the spaces after the marker and keeps the title', () => {
    const found = markdownHeading('### My title')
    expect(found).toEqual({ level: 3, markerLength: 4 })
    expect('### My title'.slice(found!.markerLength)).toBe('My title')
  })

  it('ignores everything that is not a marker', () => {
    expect(markdownHeading('####### seven')).toBeNull()
    expect(markdownHeading('#hashtag')).toBeNull()
    expect(markdownHeading(' ## indented')).toBeNull()
    expect(markdownHeading('plain')).toBeNull()
    expect(markdownHeading('')).toBeNull()
  })
})

describe('typedTableAt', () => {
  // "\n" in a text stands for a line break (Shift+Enter) inside the paragraph
  const para = (text: string) => ({
    type: { name: 'docParagraph' },
    content: { size: text.length },
    textBetween: () => text,
    nodeSize: text.length + 2,
  })
  const docOf = (...texts: string[]) => ({
    childCount: texts.length,
    child: (i: number) => para(texts[i]!),
  })

  it('needs two consecutive table-row paragraphs', () => {
    expect(typedTableAt(docOf('| a | b |'), 0)).toBeNull()
    expect(typedTableAt(docOf('| a | b |', 'plain'), 0)).toBeNull()
    expect(typedTableAt(docOf('x', '| a | b |', '| -- | -- |'), 1)).toMatchObject({
      markdown: '| a | b |\n| -- | -- |\n|  |  |',
      needsBodyRow: true,
    })
  })

  it('covers the whole block from any of its rows and reports its range', () => {
    const doc = docOf('intro', '| a | b |', '| -- | -- |', '| 1 | 2 |', 'outro')
    const found = typedTableAt(doc, 3)
    expect(found).toMatchObject({
      needsBodyRow: false,
      markdown: '| a | b |\n| -- | -- |\n| 1 | 2 |',
    })
    expect(found!.from).toBe(7) // after "intro" (5 + 2)
    expect(found!.to).toBe(7 + 11 + 13 + 11)
  })

  it('reads a paragraph of line-broken rows (Shift+Enter) as a table too', () => {
    const doc = docOf('| Hi| Hi |\n| -- | -- |\n|hi2 | 2 |')
    expect(typedTableAt(doc, 0)).toMatchObject({
      markdown: '| Hi| Hi |\n| -- | -- |\n|hi2 | 2 |',
      needsBodyRow: false,
    })
    // one line alone, or a line that is not a row, is not a table
    expect(typedTableAt(docOf('| a | b |'), 0)).toBeNull()
    expect(typedTableAt(docOf('| a | b |\nnot a row'), 0)).toBeNull()
  })
})
