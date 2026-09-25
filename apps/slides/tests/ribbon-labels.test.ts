import { describe, expect, it } from 'vitest'
import { labelFromTip } from '@genoffice/ui'
import { ribbonStrings } from '../src/renderer/i18n/strings-ribbon'

/** Wide-window clipboard captions are derived from tooltips (labelFromTip); every language must give a clean word. */
const ribbon = ribbonStrings as unknown as Record<string, Record<string, string>>
const langs = Object.keys(ribbon)

describe('slides ribbon captions', () => {
  it('covers all 20 UI languages', () => {
    expect(langs.length).toBe(20)
  })

  it.each(['ribbonCutTip', 'ribbonCopyTip', 'ribbonBrushTipDefault'])(
    '%s yields a short clean caption in every language',
    (key) => {
      for (const lang of langs) {
        const label = labelFromTip(ribbon[lang]![key]!)
        expect(label, `${lang}/${key}`).not.toBe('')
        expect(label.length, `${lang}/${key} "${label}"`).toBeLessThanOrEqual(30)
        expect(label, `${lang}/${key} "${label}"`).not.toMatch(/[(（):：⌘]/)
      }
    },
  )
})
