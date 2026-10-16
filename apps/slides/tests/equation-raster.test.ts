// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { RenderSlide } from '@genoffice/pptx-render'
import { equationSvg, prepareEquationRasters } from '../src/renderer/equation-raster'

const fraction = '<math display="block"><mrow><mfrac><mi>a</mi><mi>b</mi></mfrac></mrow></math>'

afterEach(() => vi.unstubAllGlobals())

describe('native equation raster for editor and export', () => {
  it('encodes inline MathML in a transparent, self-contained SVG image', () => {
    const svg = equationSvg({
      mathML: fraction,
      width: 60,
      height: 42,
      fontSize: 24,
      color: '#2468AC',
    })
    expect(svg).toContain('width="120" height="84" viewBox="0 0 60 42"')
    expect(svg).toContain('<foreignObject width="60" height="42">')
    expect(svg).toContain('xmlns="http://www.w3.org/1998/Math/MathML" display="inline"')
    expect(svg).toContain('font-size:24px;color:#2468AC')
    expect(svg).not.toMatch(/(?:href|src)\s*=/i)
    expect(svg).not.toContain('display="block"')
  })

  it('waits for equation images inside both groups and table cells before export capture', async () => {
    const loaded: string[] = []
    class LoadedImage {
      onload: (() => void) | null = null
      onerror: (() => void) | null = null
      set src(value: string) {
        loaded.push(value)
        queueMicrotask(() => this.onload?.())
      }
    }
    vi.stubGlobal('Image', LoadedImage)
    const run = (mathML: string) => ({
      mathML,
      widthPx: 50,
      mathAscentPx: 30,
      mathDescentPx: 10,
      fontSizePx: 22,
      color: '#111111',
    })
    const slide = {
      nodes: [
        {
          type: 'group',
          children: [{ type: 'shape', text: { lines: [{ runs: [run(fraction)] }] } }],
        },
        {
          type: 'table',
          cells: [
            { text: { lines: [{ runs: [run(fraction.replace('<mi>a</mi>', '<mi>x</mi>'))] }] } },
          ],
        },
      ],
    } as unknown as RenderSlide
    await prepareEquationRasters(slide)
    expect(loaded).toHaveLength(2)
    expect(loaded.every((src) => src.startsWith('data:image/svg+xml;charset=utf-8,'))).toBe(true)
  })

  it('rejects export preparation when a native equation image fails to load', async () => {
    class FailedImage {
      onload: (() => void) | null = null
      onerror: (() => void) | null = null
      set src(_value: string) {
        queueMicrotask(() => this.onerror?.())
      }
    }
    vi.stubGlobal('Image', FailedImage)
    const slide = {
      nodes: [
        {
          type: 'shape',
          text: {
            lines: [
              {
                runs: [
                  {
                    mathML: fraction.replace('<mi>b</mi>', '<mi>z</mi>'),
                    widthPx: 50,
                    mathAscentPx: 30,
                    mathDescentPx: 10,
                    fontSizePx: 22,
                    color: '#111111',
                  },
                ],
              },
            ],
          },
        },
      ],
    } as unknown as RenderSlide
    await expect(prepareEquationRasters(slide)).rejects.toThrow(
      'Unable to rasterize native equation MathML',
    )
  })
})
