import { describe, expect, it } from 'vitest'
import { labelFromTip } from '@genoffice/ui'
import { ribbonStrings } from '../src/renderer/i18n/strings-ribbon'
import { appStrings } from '../src/renderer/i18n/strings-app'

/**
 * The wide-window ribbon captions are derived from tooltip strings (labelFromTip), so every
 * UI language must yield a short, clean word for each captioned button.
 */
type Table = Record<string, string>
const ribbon = ribbonStrings as unknown as Record<string, Table>
const app = appStrings as unknown as Record<string, Table>
const langs = Object.keys(ribbon)

describe('docs ribbon captions', () => {
  it('covers all 20 UI languages', () => {
    expect(langs.length).toBe(20)
  })

  it.each(['ribbonCutTip', 'ribbonCopyTip', 'ribbonPainterTip'])(
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

  it.each(['appFindPlaceholder', 'appReplace'])(
    'the Editing group reuses %s, which must be a bare word in every language',
    (key) => {
      for (const lang of Object.keys(app)) {
        const text = app[lang]![key]!
        expect(text.length, `${lang}/${key} "${text}"`).toBeLessThanOrEqual(16)
        expect(text, `${lang}/${key}`).not.toMatch(/[…:：(（]/)
      }
    },
  )
})
