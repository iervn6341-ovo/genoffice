// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest'
import {
  DEFAULT_FONT_SUBSTITUTIONS,
  fontNameEquivalents,
  isAppBundledFamily,
  normalizeFontSubstitutions,
  substituteIsUnavailable,
} from '../src/font-substitution'

describe('font-substitution', () => {
  it('provides default font substitutions', () => {
    expect(DEFAULT_FONT_SUBSTITUTIONS.length).toBeGreaterThanOrEqual(10)
    const pmingliu = DEFAULT_FONT_SUBSTITUTIONS.find((r) => r.original === '新細明體')
    expect(pmingliu).toBeDefined()
    expect(pmingliu?.substitute).toBe('Songti TC')
  })

  it('normalizes custom rules and strips duplicates and empty entries', () => {
    const raw = [
      { original: '  CustomFont  ', substitute: 'Arial' },
      { original: '', substitute: 'Times' },
      { original: 'customfont', substitute: 'DuplicateShouldBeIgnored' },
    ]
    const normalized = normalizeFontSubstitutions(raw)
    expect(normalized).toHaveLength(1)
    expect(normalized[0]).toEqual({
      original: 'CustomFont',
      substitute: 'Arial',
    })
  })
})

describe('substitution targets the machine cannot answer', () => {
  const proto = HTMLCanvasElement.prototype as unknown as { getContext: unknown }
  const original = proto.getContext
  /** canvas stub: a family resolves when installed, and then measures wider than the generic stack */
  const install = (...names: string[]) => {
    const set = new Set(names)
    proto.getContext = () => {
      let font = ''
      return {
        set font(v: string) {
          font = v
        },
        get font() {
          return font
        },
        measureText: () => ({
          width: set.has(/^\S+\s+"?([^",]+)"?/.exec(font)?.[1] ?? '') ? 100 : 50,
        }),
      }
    }
  }
  afterEach(() => {
    proto.getContext = original
  })

  it('flags a substitute nothing on the machine answers to (BiauKai on current macOS)', () => {
    install('Songti TC')
    expect(substituteIsUnavailable('BiauKai')).toBe(true)
    expect(substituteIsUnavailable('Songti TC')).toBe(false)
  })

  it('accepts a substitute that is installed under another name of the same face', () => {
    install('DFKai-SB')
    expect(substituteIsUnavailable('BiauKai')).toBe(false)
  })

  it('never flags the app-bundled stand-ins, which only exist inside the editors', () => {
    install()
    for (const name of ['Carlito GO', 'Caladea', 'Liberation Serif', 'GenOffice Songti TC']) {
      expect(isAppBundledFamily(name)).toBe(true)
      expect(substituteIsUnavailable(name)).toBe(false)
    }
    expect(isAppBundledFamily('Songti TC')).toBe(false)
  })

  it('shares the other-name table with the editors', () => {
    expect(fontNameEquivalents('標楷體')).toEqual(expect.arrayContaining(['DFKai-SB', 'BiauKai']))
    expect(fontNameEquivalents('dfkai-sb')).toContain('標楷體')
    expect(fontNameEquivalents('unknown face')).toEqual([])
  })
})
