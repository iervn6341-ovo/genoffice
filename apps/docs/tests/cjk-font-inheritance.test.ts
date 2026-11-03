import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Block, ParsedDocFull, StyleInfo } from '@genoffice/docx-engine'
import {
  cssFontFamily,
  cssLatinOnlyFontFamily,
  setUserFontSubstitutions,
} from '../src/renderer/line-metrics'
import { DEFAULT_FONT_SUBSTITUTIONS } from '@genoffice/ui'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

/**
 * Traditional Chinese contracts (LibreOffice comparison, 2026-10-01): headings in
 * BiauKai / PMingLiU printed regular instead of bold, and a Times New Roman title
 * under a BiauKai Normal mixed Songti SC and Noto glyphs instead of BiauKai.
 */

const families = (chain: string) => chain.split(',').map((f) => f.trim().replace(/^'|'$/g, ''))

describe('Traditional Chinese face chains', () => {
  it('標楷體 resolves to the Kai faces (DFKai-SB / BiauKai) ahead of the Ming tail', () => {
    const f = families(cssFontFamily('標楷體'))
    expect(f.slice(0, 4)).toEqual(['標楷體', 'DFKai-SB', 'BiauKaiTC', 'Kaiti TC'])
    expect(f.indexOf('DFKai-SB')).toBeLessThan(f.indexOf('GenOffice MingLiU'))
  })

  it('標楷體 under the default substitution (→ BiauKai) still falls back to the installed Kai faces', () => {
    // A Word file embeds only the characters it typed (no 第 章 探 討 緒 in one real case): the
    // embedded 標楷體 is a partial face, and the glyphs it lacks must stay Kai, not Noto / Songti
    setUserFontSubstitutions(DEFAULT_FONT_SUBSTITUTIONS)
    try {
      for (const name of ['標楷體', 'DFKai-SB']) {
        const f = families(cssFontFamily(name))
        expect(f.slice(0, 2)).toEqual(['BiauKai', name])
        expect(f).toEqual(expect.arrayContaining(['DFKai-SB', 'BiauKaiTC', 'Kaiti TC']))
        expect(f.indexOf('DFKai-SB')).toBeLessThan(f.indexOf('Noto Serif CJK SC'))
        expect(f.indexOf('BiauKaiTC')).toBeLessThan(f.indexOf('Noto Serif CJK SC'))
      }
      // other substitutions are untouched
      expect(families(cssFontFamily('新細明體'))).not.toContain('DFKai-SB')
    } finally {
      setUserFontSubstitutions([])
    }
  })

  it('新細明體 / PMingLiU reach Songti TC through the weight-normal alias (synthetic bold)', () => {
    for (const name of ['新細明體', 'PMingLiU']) {
      const f = families(cssFontFamily(name))
      expect(f).toContain('GenOffice Songti TC')
      expect(f.indexOf('GenOffice Songti TC')).toBeLessThan(f.indexOf('Songti TC'))
    }
  })
})

describe('Latin-only runs keep the inherited East Asian face', () => {
  it('chain = Latin face, then the inherited --doc-ea-chain with the Latin chain as backstop', () => {
    const chain = cssLatinOnlyFontFamily('Times New Roman')
    expect(chain.startsWith("'Times New Roman'")).toBe(true)
    expect(chain).toContain('var(--doc-ea-chain,')
  })

  const markOf = (run: Record<string, unknown>) => {
    const block = { type: 'paragraph', runs: [{ text: '軟體開發', ...run }] } as unknown as Block
    const doc = blocksToPmDoc([block]) as {
      content: Array<{
        content?: Array<{ marks?: Array<{ type: string; attrs: Record<string, unknown> }> }>
      }>
    }
    return doc.content[0]!.content![0]!.marks!.find((m) => m.type === 'docTextStyle')!.attrs
  }

  it('a run with only w:ascii (parse fills font from it) is flagged eaInherited', () => {
    expect(
      markOf({ font: 'Times New Roman', fontAscii: 'Times New Roman', bold: true }).eaInherited,
    ).toBe(true)
  })

  it('a run declaring w:eastAsia, or a CJK-named Latin slot, is not', () => {
    expect(
      markOf({ font: '標楷體', fontAscii: 'Times New Roman', eastAsiaFont: '標楷體' }).eaInherited,
    ).toBeNull()
    expect(markOf({ font: '標楷體', fontAscii: '標楷體' }).eaInherited).toBeNull()
  })

  it('docStyleCss publishes the Normal East Asian chain as --doc-ea-chain', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('a1', {
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: { font: '標楷體', fontAscii: 'Arial' },
    } as unknown as StyleInfo)
    const css = docStyleCss({ styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull)
    expect(css).toMatch(/--doc-ea-chain:'標楷體','DFKai-SB'/)
  })

  it('computeFormatState resolves inherited East Asian font on eaInherited CJK runs', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('a1', {
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: { font: '標楷體', fontAscii: 'Arial' },
    } as unknown as StyleInfo)

    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: '軟體開發專案合約書',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: {
                      font: 'Times New Roman',
                      fontAscii: 'Times New Roman',
                      eaInherited: true,
                    },
                  },
                ],
              },
              {
                type: 'text',
                text: 'XXX',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: {
                      font: 'Times New Roman',
                      fontAscii: 'Times New Roman',
                      eaInherited: true,
                    },
                  },
                ],
              },
            ],
          },
        ],
      },
    })

    // Caret at CJK text -> resolves inherited East Asian font
    editor.commands.setTextSelection(3)
    const fsCjk = computeFormatState(editor, styles)
    expect(fsCjk.fontFamily).toBe('標楷體')

    // Caret at Latin text ('XXX') -> resolves Times New Roman
    editor.commands.setTextSelection(11)
    const fsLatin = computeFormatState(editor, styles)
    expect(fsLatin.fontFamily).toBe('Times New Roman')

    editor.destroy()
  })

  it('computeFormatState resolves Normal East Asian font on unformatted CJK text', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('a1', {
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: { font: '標楷體', fontAscii: 'Arial' },
    } as unknown as StyleInfo)

    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [{ type: 'text', text: '本專案總金額為新臺幣' }],
          },
        ],
      },
    })

    editor.commands.setTextSelection(2)
    const fs = computeFormatState(editor, styles)
    expect(fs.fontFamily).toBe('標楷體')

    editor.destroy()
  })
})
