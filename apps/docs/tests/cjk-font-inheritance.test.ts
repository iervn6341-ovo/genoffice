import { describe, expect, it } from 'vitest'
import type { Block, ParsedDocFull, StyleInfo } from '@genoffice/docx-engine'
import { cssFontFamily, cssLatinOnlyFontFamily } from '../src/renderer/line-metrics'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { blocksToPmDoc } from '../src/renderer/editor/convert'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

/**
 * Traditional Chinese contracts (LibreOffice comparison, 2026-10-01): headings in
 * 標楷體 / 新細明體 printed regular instead of bold, and a Times New Roman title
 * under a 標楷體 Normal mixed Songti SC and Noto glyphs instead of 標楷體.
 */

const families = (chain: string) => chain.split(',').map((f) => f.trim().replace(/^'|'$/g, ''))

describe('Traditional Chinese face chains', () => {
  it('標楷體 resolves to the Kai faces (DFKai-SB / BiauKai) ahead of the Ming tail', () => {
    const f = families(cssFontFamily('標楷體'))
    expect(f.slice(0, 4)).toEqual(['標楷體', 'DFKai-SB', 'BiauKaiTC', 'Kaiti TC'])
    expect(f.indexOf('DFKai-SB')).toBeLessThan(f.indexOf('GenOffice MingLiU'))
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
})
