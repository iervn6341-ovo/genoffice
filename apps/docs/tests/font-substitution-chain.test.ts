import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { DEFAULT_FONT_SUBSTITUTIONS } from '@genoffice/ui'
import {
  cssFontFamily,
  fontNameEquivalents,
  noteEmbeddedFontsChanged,
  setUserFontSubstitutions,
} from '../src/renderer/line-metrics'

/**
 * The substitution setting reads "when the system lacks font A, show A as B". The chain must follow
 * that literally: a declared face that resolves is drawn as itself (B only completes it), and the
 * names a face is known under on the machine all take part — never a silent swap to B.
 */

const families = (chain: string) => chain.split(',').map((f) => f.trim().replace(/^'|'$/g, ''))

/** canvas stub: a family "resolves" when it is in `installed`; its text then measures wider than the generic */
function stubCanvas(installed: Set<string>) {
  const proto = HTMLCanvasElement.prototype as unknown as { getContext: unknown }
  const original = proto.getContext
  proto.getContext = () => {
    let font = ''
    return {
      set font(v: string) {
        font = v
      },
      get font() {
        return font
      },
      measureText: () => {
        const first = /^\S+\s+"?([^",]+)"?/.exec(font)?.[1] ?? ''
        return { width: installed.has(first) ? 100 : 50 }
      },
    }
  }
  return () => {
    proto.getContext = original
  }
}

const NAMES = ['標楷體', 'DFKai-SB', 'BiauKai', '新細明體', 'PMingLiU', 'Calibri', 'Carlito GO']
let restore: (() => void) | null = null

function withInstalled(...names: string[]) {
  restore?.()
  restore = stubCanvas(new Set(names))
  noteEmbeddedFontsChanged(NAMES) // drops the availability cache
}

beforeEach(() => setUserFontSubstitutions(DEFAULT_FONT_SUBSTITUTIONS))
afterEach(() => {
  restore?.()
  restore = null
  setUserFontSubstitutions([])
  noteEmbeddedFontsChanged(NAMES)
})

describe('substitution rules do not override a face that resolves', () => {
  it('a face installed under its other name is drawn as itself, ahead of the substitute', () => {
    // kaiu.ttf answers to "DFKai-SB" only; "標楷體" does not resolve on its own
    withInstalled('DFKai-SB')
    const f = families(cssFontFamily('標楷體'))
    expect(f.slice(0, 3)).toEqual(['標楷體', 'DFKai-SB', 'BiauKai'])
  })

  it('an installed PMingLiU is not replaced by Songti TC', () => {
    withInstalled('PMingLiU')
    expect(families(cssFontFamily('新細明體')).slice(0, 3)).toEqual([
      '新細明體',
      'PMingLiU',
      'Songti TC',
    ])
  })

  it('without the face the substitute still comes first, as configured', () => {
    withInstalled()
    expect(families(cssFontFamily('新細明體'))[0]).toBe('Songti TC')
    expect(families(cssFontFamily('標楷體'))[0]).toBe('BiauKai')
  })

  it('a bundled metric-compatible stand-in stays first (same line breaks on every machine)', () => {
    withInstalled('Calibri')
    expect(families(cssFontFamily('Calibri'))[0]).toBe('Carlito GO')
  })

  it('a substitute the OS does not know still leaves the installed Kai faces in the chain', () => {
    // BiauKai does not exist on current macOS; the Kai faces that do must catch the glyphs
    withInstalled()
    const f = families(cssFontFamily('標楷體'))
    expect(f.indexOf('BiauKaiTC')).toBeGreaterThan(-1)
    expect(f.indexOf('BiauKaiTC')).toBeLessThan(f.indexOf('Noto Serif CJK SC'))
  })
})

describe('font name equivalents', () => {
  it('are looked up case-insensitively and exclude the declared name', () => {
    expect(fontNameEquivalents('dfkai-sb')).toContain('標楷體')
    expect(fontNameEquivalents('標楷體')).toEqual(expect.arrayContaining(['DFKai-SB', 'BiauKai']))
    expect(fontNameEquivalents('標楷體')).not.toContain('標楷體')
    expect(fontNameEquivalents('Some Unknown Face')).toEqual([])
  })
})
