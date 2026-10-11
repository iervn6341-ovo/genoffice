import { describe, expect, it } from 'vitest'

import {
  healPipeTables,
  isSourceMarkupHtml,
  looksLikeMarkdown,
  markdownPasteHtml,
  normalizeTableRow,
} from '../src/renderer/editor/markdown-paste'

describe('looksLikeMarkdown', () => {
  it('recognizes unambiguous constructs on their own', () => {
    expect(looksLikeMarkdown('## Heading\n\nBody text.')).toBe(true)
    expect(looksLikeMarkdown('```js\nconsole.log(1)\n```')).toBe(true)
    expect(looksLikeMarkdown('See [the docs](https://example.com) for details.')).toBe(true)
    expect(looksLikeMarkdown('This is **important** to know.')).toBe(true)
    expect(looksLikeMarkdown('| A | B |\n| --- | --- |\n| 1 | 2 |')).toBe(true)
  })

  it('requires repetition for list and quote markers', () => {
    expect(looksLikeMarkdown('- one item only')).toBe(false)
    expect(looksLikeMarkdown('- one\n- two')).toBe(true)
    expect(looksLikeMarkdown('1. first\n2. second')).toBe(true)
    expect(looksLikeMarkdown('> single quoted line')).toBe(false)
    expect(looksLikeMarkdown('> line one\n> line two')).toBe(true)
  })

  it('does not treat Python dunder identifiers as bold markers', () => {
    expect(looksLikeMarkdown('def __init__(self):\n    self.__name__ = "x"')).toBe(false)
    expect(looksLikeMarkdown('call __main__ before __exit__ runs')).toBe(false)
  })

  it('leaves ordinary prose alone', () => {
    expect(looksLikeMarkdown('Meeting notes from Monday.\nBudget is $100.')).toBe(false)
    expect(looksLikeMarkdown('Use the # channel for questions.')).toBe(false)
    expect(looksLikeMarkdown('1. A single numbered sentence.')).toBe(false)
    expect(looksLikeMarkdown('a * b * c = abc')).toBe(false)
  })
})

describe('markdownPasteHtml', () => {
  it('converts headings, emphasis, lists, and code fences', () => {
    const html = markdownPasteHtml('## Title\n\nSome **bold** text.\n\n- one\n- two')
    expect(html).toContain('<h2>')
    expect(html).toContain('<strong>bold</strong>')
    expect(html).toContain('<ul>')
    expect(html).toContain('<li>one</li>')
    const fenced = markdownPasteHtml('```\nlet x = 1\n```')
    expect(fenced).toContain('<pre><code>')
  })

  it('normalizes Windows line endings before conversion', () => {
    const html = markdownPasteHtml('# A\r\n\r\n- x\r\n- y')
    expect(html).toContain('<h1>')
    expect(html).toContain('<li>y</li>')
  })

  it('returns null for plain prose so the literal paste path runs', () => {
    expect(markdownPasteHtml('Just an ordinary sentence.')).toBeNull()
    expect(markdownPasteHtml('Line one.\nLine two.')).toBeNull()
  })

  it('converts GFM tables', () => {
    const html = markdownPasteHtml('| A | B |\n| --- | --- |\n| 1 | 2 |')
    expect(html).toContain('<table>')
    expect(html).toContain('<td>1</td>')
  })
})

describe('healPipeTables', () => {
  it('adds the missing separator under a header row', () => {
    expect(healPipeTables('| Hi | Hi2|\n| 21 | 21 |')).toBe(
      '| Hi | Hi2|\n| --- | --- |\n| 21 | 21 |',
    )
  })

  it('leaves real tables, single rows and code fences alone', () => {
    const real = '| A | B |\n| -- | -- |\n| 1 | 2 |'
    expect(healPipeTables(real)).toBe(real)
    expect(healPipeTables('| lone |')).toBe('| lone |')
    const fenced = '```\n| a | b |\n| c | d |\n```'
    expect(healPipeTables(fenced)).toBe(fenced)
  })

  it('converts the sample from the bug report: valid table, gap, header-less table', () => {
    const html = markdownPasteHtml(
      '| Hi | Hi2|\n| -- | -- |\n| 21 | 21 |\n\n\n\n| Hi | Hi2|\n| 21 | 21 |',
    )
    expect(html?.match(/<table>/g)?.length).toBe(2)
  })
})

describe('healPipeTables: sloppy delimiter rows', () => {
  it('completes a half-typed separator so the table converts', () => {
    expect(healPipeTables('| Hi | Hi2|\n| -- | |\n| 1 | 2 |')).toBe(
      '| Hi | Hi2|\n| --- | --- |\n| 1 | 2 |',
    )
    expect(markdownPasteHtml('| Hi | Hi2|\n| -- | |')).toContain('<table>')
  })

  it('keeps alignment colons and valid rows untouched', () => {
    const ok = '| a | b |\n| :-- | --: |\n| 1 | 2 |'
    expect(healPipeTables(ok)).toBe(ok)
  })
})

describe('isSourceMarkupHtml', () => {
  it('treats editor/terminal soup as source and rich pages as rich', () => {
    expect(isSourceMarkupHtml('<div style="color:#ccc"><span>| a | b |</span></div>')).toBe(true)
    expect(isSourceMarkupHtml('<pre><code>## x</code></pre>')).toBe(true)
    expect(isSourceMarkupHtml('<h2>Title</h2><p>text</p>')).toBe(false)
    expect(isSourceMarkupHtml('<p>see <a href="x">this</a></p>')).toBe(false)
  })
})

describe('normalizeTableRow (CJK input methods)', () => {
  it('turns full-width bars and dashes into plain Markdown', () => {
    expect(normalizeTableRow('｜ Hi ｜ Hi2｜')).toBe('| Hi | Hi2|')
    expect(normalizeTableRow('｜ －－ ｜ —— ｜')).toBe('| -- | -- |')
    expect(normalizeTableRow('plain text')).toBe('plain text')
    expect(normalizeTableRow('｜ only a start')).toBe('｜ only a start')
  })

  it('makes a pasted full-width table convert', () => {
    expect(markdownPasteHtml('｜ A ｜ B ｜\n｜ -- ｜ -- ｜\n｜ 1 ｜ 2 ｜')).toContain('<table>')
  })
})
