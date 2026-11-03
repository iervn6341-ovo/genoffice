import { describe, expect, it } from 'vitest'
import { buildBlankDocx, parseDocx, saveDocx, type StyleInfo } from '@genoffice/docx-engine'
import {
  BUILTIN_PARA_STYLES,
  builtinStyle,
  displayOfUpsert,
  styleInfoOfUpsert,
} from '../src/renderer/editor/builtin-styles'
import { listStyles } from '../src/renderer/components/StylesPane'

describe('built-in paragraph styles (Styles gallery)', () => {
  it('follow Word defaults', () => {
    expect(builtinStyle('Title')?.upsert.rPr).toMatchObject({
      font: 'Calibri Light',
      sizeHalfPoints: 56,
    })
    expect(builtinStyle('Quote')?.upsert.pPr).toMatchObject({
      align: 'center',
      indentLeftTwips: 864,
    })
    expect(builtinStyle('IntenseQuote')?.upsert.rPr).toMatchObject({
      italic: true,
      color: '4472C4',
    })
    expect(builtinStyle('NoSpacing')?.upsert.pPr?.spaceAfterTwips).toBe(0)
    expect(builtinStyle('Nope')).toBeUndefined()
  })

  it('display hints come from the same properties the upsert writes', () => {
    const d = displayOfUpsert(builtinStyle('Quote')!.upsert)
    expect(d).toMatchObject({
      italic: true,
      color: '404040',
      align: 'center',
      spaceBeforeTwips: 200,
      spaceAfterTwips: 160,
      indentLeftTwips: 864,
      indentRightTwips: 864,
    })
    const info = styleInfoOfUpsert(builtinStyle('Title')!.upsert)
    expect(info).toMatchObject({
      styleId: 'Title',
      name: 'Title',
      type: 'paragraph',
      basedOn: 'Normal',
    })
    expect(info.display).toMatchObject({ sizeHalfPoints: 56, fontAscii: 'Calibri Light' })
  })

  it('what the editor shows matches what a saved file parses back to', async () => {
    const blank = await parseDocx(await buildBlankDocx())
    const saved = await saveDocx(blank, [{ kind: 'original', docxIndex: 0 }], {
      styleUpserts: BUILTIN_PARA_STYLES.filter((s) => !blank.styles.has(s.styleId)).map(
        (s) => s.upsert,
      ),
    })
    const reparsed = await parseDocx(saved)
    for (const b of BUILTIN_PARA_STYLES) {
      const info = reparsed.styles.get(b.styleId)
      expect(info, b.styleId).toBeDefined()
      expect(info!.name).toBe(b.name)
      const mine = styleInfoOfUpsert(b.upsert).display!
      // every hint the gallery shows at once must survive the round trip
      for (const key of [
        'sizeHalfPoints',
        'bold',
        'italic',
        'color',
        'align',
        'spaceBeforeTwips',
        'spaceAfterTwips',
        'indentLeftTwips',
        'indentRightTwips',
      ] as const) {
        if (mine[key] !== undefined)
          expect(info!.display?.[key], `${b.styleId}.${key}`).toEqual(mine[key])
      }
    }
  })
})

describe('Styles pane list', () => {
  const info = (over: Partial<StyleInfo> & { styleId: string }): StyleInfo => ({
    name: over.styleId,
    type: 'paragraph',
    ...over,
  })
  it('lists Normal, headings, quick styles and the undefined built-ins, hiding the rest', () => {
    const styles = new Map<string, StyleInfo>([
      ['Normal', info({ styleId: 'Normal', name: 'Normal', isDefault: true, qFormat: true })],
      ['Heading2', info({ styleId: 'Heading2', name: 'heading 2', headingLevel: 2 })],
      ['Heading1', info({ styleId: 'Heading1', name: 'heading 1', headingLevel: 1 })],
      ['Hidden', info({ styleId: 'Hidden', semiHidden: true, qFormat: true })],
      ['Plain', info({ styleId: 'Plain' })],
      ['Emph', info({ styleId: 'Emph', name: 'Emphasis', type: 'character', qFormat: true })],
      ['TableGrid', info({ styleId: 'TableGrid', type: 'table', qFormat: true })],
    ])
    const ids = listStyles(styles, {}).map((s) => s.styleId)
    expect(ids.slice(0, 3)).toEqual(['Normal', 'Heading1', 'Heading2'])
    expect(ids).toContain('Emph')
    expect(ids).toContain('Title') // built-in the file lacks
    for (const hidden of ['Hidden', 'Plain', 'TableGrid']) expect(ids).not.toContain(hidden)
  })
  it("a document's own definition wins over the built-in one", () => {
    const own = info({
      styleId: 'Title',
      name: 'Title',
      qFormat: true,
      display: { sizeHalfPoints: 80 },
    })
    const title = listStyles(new Map([['Title', own]]), {}).find((s) => s.styleId === 'Title')
    expect(title?.display?.sizeHalfPoints).toBe(80)
  })

  it('filters styles correctly by mode (recommended, inUse, inDoc, all)', () => {
    const docStyles = new Map<string, StyleInfo>([
      ['Normal', info({ styleId: 'Normal', name: 'Normal', isDefault: true, qFormat: true })],
      ['Custom', info({ styleId: 'Custom', name: 'CustomStyle', qFormat: false })],
    ])
    // inDoc only lists styles present in the document
    const inDocIds = listStyles(docStyles, {}, 'inDoc').map((s) => s.styleId)
    expect(inDocIds).toContain('Normal')
    expect(inDocIds).toContain('Custom')
    expect(inDocIds).not.toContain('Caption') // built-in not in doc

    // all includes built-ins even if not in doc
    const allIds = listStyles(docStyles, {}, 'all').map((s) => s.styleId)
    expect(allIds).toContain('Caption')
    expect(allIds).toContain('Bibliography')
    expect(allIds).toContain('Custom')

    // inUse only lists styles in use
    const inUseIds = listStyles(docStyles, {}, 'inUse', new Set(['Normal', 'Caption'])).map(
      (s) => s.styleId,
    )
    expect(inUseIds).toContain('Normal')
    expect(inUseIds).toContain('Caption')
    expect(inUseIds).not.toContain('Custom')
  })
})

describe('localizedStyleName', () => {
  it('localizes styles to Traditional Chinese matching Word', async () => {
    const { localizedStyleName } = await import('../src/renderer/editor/builtin-styles')
    expect(localizedStyleName('Normal', 'Normal', 'zh-TW')).toBe('內文')
    expect(localizedStyleName('Heading1', 'Heading 1', 'zh-TW')).toBe('標題 1')
    expect(localizedStyleName('Quote', 'Quote', 'zh-TW')).toBe('引文')
    expect(localizedStyleName('IntenseQuote', 'Intense Quote', 'zh-TW')).toBe('鮮明引文')
    expect(localizedStyleName('Caption', 'caption', 'zh-TW')).toBe('標號')
    expect(localizedStyleName('Bibliography', 'bibliography', 'zh-TW')).toBe('書目')
    expect(localizedStyleName('TOC1', 'toc 1', 'zh-TW')).toBe('目錄 1')
    expect(localizedStyleName('TOCHeading', 'TOC Heading', 'zh-TW')).toBe('目錄標題')
    expect(localizedStyleName('TableofFigures', 'table of figures', 'zh-TW')).toBe('圖表目錄')
    expect(localizedStyleName('Header', 'header', 'zh-TW')).toBe('頁首')
    expect(localizedStyleName('Footer', 'footer', 'zh-TW')).toBe('頁尾')
    expect(localizedStyleName('Hyperlink', 'Hyperlink', 'zh-TW')).toBe('超連結')
  })
})

