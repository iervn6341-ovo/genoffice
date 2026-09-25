/**
 * C5: Word's Find (Match case off) folds ß/SS, final sigma, Turkish İ and
 * composed/decomposed accents; highlight offsets must map back exactly.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { findFoldedSpans, foldForSearch } from '../src/renderer/editor/find-fold'
import { findMatches } from '../src/renderer/components/FindPanel'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

const spansText = (text: string, query: string, matchCase = false) =>
  findFoldedSpans(text, query, matchCase).map((s) => text.slice(s.from, s.to))

describe('findFoldedSpans (C5)', () => {
  it('ß matches SS / ss and vice versa, with source offsets', () => {
    expect(spansText('Die Straße und die STRASSE', 'strasse')).toEqual(['Straße', 'STRASSE'])
    expect(spansText('Die Straße und die STRASSE', 'STRAßE')).toEqual(['Straße', 'STRASSE'])
    const [span] = findFoldedSpans('xßy', 'SS', false)
    expect(span).toEqual({ from: 1, to: 2 })
  })

  it('Greek final sigma matches medial / capital sigma', () => {
    expect(spansText('ΟΔΥΣΣΕΥΣ οδυσσευς', 'ΟΔΥΣΣΕΥΣ')).toEqual(['ΟΔΥΣΣΕΥΣ', 'οδυσσευς'])
    expect(spansText('λόγος', 'ΛΌΓΟΣ')).toEqual(['λόγος'])
  })

  it('Turkish İ matches i without shifting later offsets', () => {
    expect(spansText('İstanbul istanbul', 'istanbul')).toEqual(['İstanbul', 'istanbul'])
    const text = 'İİ abc'
    expect(findFoldedSpans(text, 'ABC', false)).toEqual([{ from: 3, to: 6 }])
  })

  it('composed and decomposed accents match each other, never a bare base letter', () => {
    const decomposed = 'café'
    expect(spansText(decomposed, 'café')).toEqual([decomposed])
    expect(spansText('café', 'café')).toEqual(['café'])
    expect(findFoldedSpans(decomposed, 'cafe', false)).toEqual([])
    expect(spansText('CAFÉ', 'café')).toEqual(['CAFÉ'])
  })

  it('match case keeps case but still normalizes', () => {
    expect(spansText('Straße STRASSE', 'STRASSE', true)).toEqual(['STRASSE'])
    expect(spansText('café CAFÉ', 'café', true)).toEqual(['café'])
  })

  it('maps every folded unit back to its segment', () => {
    const f = foldForSearch('aßb', false)
    expect(f.folded).toBe('assb')
    expect(f.srcStart).toEqual([0, 1, 1, 2])
    expect(f.srcEnd).toEqual([1, 2, 2, 3])
  })

  it('whole-word predicate sees source offsets', () => {
    const text = 'Straßen Straße'
    const spans = findFoldedSpans(
      text,
      'strasse',
      false,
      (from, to) => !/\p{L}/u.test(text[to] ?? ''),
    )
    expect(spans.map((s) => text.slice(s.from, s.to))).toEqual(['Straße'])
    expect(spans[0]!.from).toBe(8)
  })
})

describe('findMatches editor positions (C5)', () => {
  it('highlights cover the source characters across marks', async () => {
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'Die ' },
              { type: 'text', text: 'Stra', marks: [{ type: 'bold' }] },
              { type: 'text', text: 'ße ist lang' },
            ],
          },
        ],
      } as never,
    })
    editors.push(editor)
    const matches = findMatches(editor, 'STRASSE', { matchCase: false, wholeWord: true })
    expect(matches).toHaveLength(1)
    const [m] = matches
    expect(editor.state.doc.textBetween(m!.from, m!.to)).toBe('Straße')
  })
})
