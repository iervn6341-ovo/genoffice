import { describe, expect, it } from 'vitest'
import type { StyleDisplay } from '@genoffice/docx-engine'
import {
  PARAGRAPH_SPACING_PRESETS,
  captureSpacing,
  spacingPreset,
  spacingUpsert,
  writeSpacing,
} from '../src/renderer/editor/paragraph-spacing'

describe('Design → Paragraph Spacing presets', () => {
  it("has Word's seven presets in order", () => {
    expect(PARAGRAPH_SPACING_PRESETS.map((p) => p.id)).toEqual([
      'default',
      'none',
      'compact',
      'tight',
      'open',
      'relaxed',
      'double',
    ])
    expect(spacingPreset('double')).toMatchObject({ after: 160, line: 2 })
    expect(spacingPreset('nope')).toBeUndefined()
  })

  it('writes a preset into the style display and puts the original back', () => {
    const display: StyleDisplay = {
      spaceAfterTwips: 200,
      spaceAfterAuto: true,
      lineSpacing: 1.15,
      bold: true,
    }
    const original = captureSpacing(display)
    writeSpacing(display, spacingPreset('none')!)
    expect(display).toMatchObject({
      spaceBeforeTwips: 0,
      spaceAfterTwips: 0,
      lineSpacing: 1,
      lineRule: 'auto',
      bold: true,
    })
    expect(display.spaceAfterAuto).toBeUndefined()
    writeSpacing(display, original)
    expect(display).toMatchObject({
      spaceAfterTwips: 200,
      spaceAfterAuto: true,
      lineSpacing: 1.15,
      bold: true,
    })
    expect(display.spaceBeforeTwips).toBeUndefined()
    expect(display.lineRule).toBeUndefined()
  })

  it('saves as a patch of the Normal style', () => {
    expect(spacingUpsert(spacingPreset('open')!)).toEqual({
      styleId: 'Normal',
      pPr: { spaceBeforeTwips: 0, spaceAfterTwips: 200, lineSpacing: 1.15 },
    })
  })
})
