import { describe, expect, it } from 'vitest'

import { StylesheetEditor } from '@genoffice/xlsx-gateway/gateway/xlsx-styles'

/** An RTL cell as Excel writes it: right-aligned, explicit reading order, shrink to fit. */
const STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="2"><xf/><xf applyAlignment="1"><alignment horizontal="right" readingOrder="2" shrinkToFit="1" justifyLastLine="1" relativeIndent="2"/></xf></cellXfs>
</styleSheet>`

const CARRIED = ['readingOrder', 'shrinkToFit', 'justifyLastLine', 'relativeIndent']

const NAMED_STYLE = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="2"><xf/><xf applyAlignment="1"><alignment horizontal="center" vertical="justify" wrapText="1" indent="3"/></xf></cellStyleXfs>
  <cellXfs count="4"><xf xfId="0"/><xf xfId="1" applyAlignment="0"/><xf xfId="1" applyAlignment="1"/><xf xfId="1" applyAlignment="1"><alignment horizontal="left"/></xf></cellXfs>
</styleSheet>`

function xfAt(xml: string, index: number): string {
  const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)?.[1] ?? ''
  const xfs = [...cellXfs.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)]
  const xf = xfs[index]?.[0]
  // every negative assertion below would pass against '' for the wrong reason
  if (!xf) throw new Error(`no xf at index ${index} in: ${cellXfs}`)
  return xf
}

describe('StylesheetEditor alignment carries', () => {
  it('edits actual cell alignment without importing missing fields from a named style', () => {
    const editor = new StylesheetEditor(NAMED_STYLE)
    const edited = [1, 2, 3].map((base) =>
      editor.resolveStyle(base, { verticalAlignment: 'distributed' }),
    )
    const xml = editor.serialize()
    expect(xml).toContain(
      '<cellStyleXfs count="2"><xf/><xf applyAlignment="1"><alignment horizontal="center" vertical="justify" wrapText="1" indent="3"/></xf></cellStyleXfs>',
    )
    expect(xfAt(xml, edited[0]!)).toContain('xfId="1"')
    expect(xfAt(xml, edited[1]!)).toContain('xfId="1"')
    for (const index of edited) {
      const xf = xfAt(xml, index)
      expect(xf).toContain('applyAlignment="1"')
      expect(xf).toContain('vertical="distributed"')
      expect(xf).not.toContain('wrapText=')
      expect(xf).not.toContain('indent=')
      expect(xf).not.toContain('horizontal="center"')
    }
    expect(xfAt(xml, edited[0]!)).not.toContain('horizontal=')
    expect(xfAt(xml, edited[1]!)).not.toContain('horizontal=')
    expect(xfAt(xml, edited[2]!)).toContain('horizontal="left"')
  })

  it('retains the cell apply flag and partial child on an unrelated edit', () => {
    const editor = new StylesheetEditor(NAMED_STYLE)
    const edited = [1, 2, 3].map((base) => editor.resolveStyle(base, { fillColor: '#FF0000' }))
    const xml = editor.serialize()
    expect(xfAt(xml, edited[0]!)).toContain('applyAlignment="0"')
    expect(xfAt(xml, edited[0]!)).not.toContain('<alignment')
    expect(xfAt(xml, edited[1]!)).toContain('applyAlignment="1"')
    expect(xfAt(xml, edited[1]!)).not.toContain('<alignment')
    expect(xfAt(xml, edited[2]!)).toContain('applyAlignment="1"')
    expect(xfAt(xml, edited[2]!)).toContain('horizontal="left"')
    expect(xfAt(xml, edited[2]!)).not.toContain('wrapText=')
    expect(xfAt(xml, edited[2]!)).not.toContain('indent=')
  })

  it('keeps the unmodeled alignment attributes through an unrelated edit', () => {
    const editor = new StylesheetEditor(STYLES)
    const index = editor.resolveStyle(1, { fillColor: '#FF0000' })
    const xf = xfAt(editor.serialize(), index)
    for (const name of CARRIED) expect(xf).toContain(name)
    // the edit itself still applies, and the modeled attributes still resolve
    expect(xf).toContain('applyFill="1"')
    expect(xf).toContain('horizontal="right"')
  })

  it('keeps a vertical value the renderer approximates through an unrelated edit', () => {
    const distributed = STYLES.replace(
      '<alignment horizontal="right"',
      '<alignment horizontal="right" vertical="distributed"',
    )
    const editor = new StylesheetEditor(distributed)
    const index = editor.resolveStyle(1, { fillColor: '#FF0000' })
    const xf = xfAt(editor.serialize(), index)
    expect(xf).toContain('applyFill="1"')
    expect(xf).toContain('vertical="distributed"')
  })

  it('keeps them when the edit changes alignment itself', () => {
    const editor = new StylesheetEditor(STYLES)
    const index = editor.resolveStyle(1, { horizontalAlignment: 'center' })
    const xf = xfAt(editor.serialize(), index)
    expect(xf).toContain('horizontal="center"')
    expect(xf).toContain('readingOrder="2"')
  })

  it('preserves namespace-prefixed alignment attributes without adding core versions', () => {
    const prefixed = STYLES.replace(
      '<alignment horizontal="right" readingOrder="2" shrinkToFit="1" justifyLastLine="1" relativeIndent="2"/>',
      '<alignment horizontal="right" x14:readingOrder="2" x14:shrinkToFit="1"/>',
    )
    const editor = new StylesheetEditor(prefixed)
    const index = editor.resolveStyle(1, { fillColor: '#FF0000' })
    const xf = xfAt(editor.serialize(), index)
    // The fill edit changes the XF but does not reinterpret the vendor fields.
    expect(xf).toContain('applyFill="1"')
    expect(xf).toContain('x14:readingOrder="2"')
    expect(xf).toContain('x14:shrinkToFit="1"')
    expect(xf).not.toMatch(/(?<!:)\breadingOrder="/)
    expect(xf).not.toMatch(/(?<!:)\bshrinkToFit="/)
  })

  it('keeps a namespaced applyAlignment without adding a core apply flag', () => {
    const prefixed = STYLES.replace(
      '<xf applyAlignment="1"><alignment horizontal="right" readingOrder="2"',
      '<xf x14:applyAlignment="1"><alignment readingOrder="2"',
    )
    const editor = new StylesheetEditor(prefixed)
    const index = editor.resolveStyle(1, { fillColor: '#FF0000' })
    const xf = xfAt(editor.serialize(), index)
    expect(xf).toContain('applyFill="1"')
    expect(xf).toContain('x14:applyAlignment="1"')
    expect(xf).not.toMatch(/(?<!:)\bapplyAlignment="/)
    expect(xf).toContain('readingOrder="2"')
  })

  it('leaves a vendor child untouched on an unrelated fill edit', () => {
    const vendor = STYLES.replace(
      '<xf applyAlignment="1"><alignment horizontal="right" readingOrder="2" shrinkToFit="1" justifyLastLine="1" relativeIndent="2"/></xf>',
      '<xf><settings xmlns="urn:vendor" applyAlignment="1"><alignment readingOrder="2"/></settings></xf>',
    )
    const editor = new StylesheetEditor(vendor)
    const index = editor.resolveStyle(1, { fillColor: '#FF0000' })
    const xf = xfAt(editor.serialize(), index)
    expect(xf).toContain('applyFill="1"')
    expect(xf).toContain('<settings xmlns="urn:vendor" applyAlignment="1">')
    expect(xf).toContain('<alignment readingOrder="2"/></settings>')
  })

  it('leaves mc:AlternateContent untouched on an unrelated fill edit', () => {
    const alternate = STYLES.replace(
      '<xf applyAlignment="1"><alignment horizontal="right" readingOrder="2" shrinkToFit="1" justifyLastLine="1" relativeIndent="2"/></xf>',
      '<xf><mc:AlternateContent><mc:Choice Requires="x14">' +
        '<alignment horizontal="general" readingOrder="2"/>' +
        '</mc:Choice><mc:Fallback/></mc:AlternateContent></xf>',
    )
    const editor = new StylesheetEditor(alternate)
    const index = editor.resolveStyle(1, { fillColor: '#FF0000' })
    const xf = xfAt(editor.serialize(), index)
    expect(xf).toContain('applyFill="1"')
    expect(xf).toContain('<mc:AlternateContent>')
    expect(xf).toContain('<alignment horizontal="general" readingOrder="2"/>')
  })

  it('preserves an unapplied alignment child on an unrelated fill edit', () => {
    const carriedOnly = STYLES.replace(
      '<xf applyAlignment="1"><alignment horizontal="right" readingOrder="2" shrinkToFit="1" justifyLastLine="1" relativeIndent="2"/></xf>',
      '<xf><alignment readingOrder="2"/></xf>',
    )
    const editor = new StylesheetEditor(carriedOnly)
    const index = editor.resolveStyle(1, { fillColor: '#FF0000' })
    const xf = xfAt(editor.serialize(), index)
    expect(xf).toContain('<alignment readingOrder="2"/>')
    expect(xf).not.toContain('applyAlignment')
  })

  it('carries them when a modeled attribute keeps the element alive anyway', () => {
    // applyAlignment is absent, but horizontal survives, so the element is emitted
    // regardless and the cell already applies whatever it holds
    const notApplied = STYLES.replace('<xf applyAlignment="1">', '<xf>')
    const editor = new StylesheetEditor(notApplied)
    const index = editor.resolveStyle(1, { fillColor: '#FF0000' })
    const xf = xfAt(editor.serialize(), index)
    expect(xf).toContain('horizontal="right"')
    expect(xf).toContain('readingOrder="2"')
  })

  it('accepts applyAlignment spelled the other xsd:boolean way', () => {
    const spelled = STYLES.replace('applyAlignment="1"', 'applyAlignment="true"')
      // strip the modeled attribute so only the applyAlignment spelling can decide
      .replace('horizontal="right" readingOrder="2"', 'readingOrder="2"')
    const editor = new StylesheetEditor(spelled)
    const index = editor.resolveStyle(1, { fillColor: '#FF0000' })
    expect(xfAt(editor.serialize(), index)).toContain('readingOrder="2"')
  })

  it('adds nothing when the base has none of them', () => {
    const editor = new StylesheetEditor(STYLES)
    const index = editor.resolveStyle(0, { fillColor: '#FF0000' })
    const xf = xfAt(editor.serialize(), index)
    for (const name of CARRIED) expect(xf).not.toContain(name)
  })
})
