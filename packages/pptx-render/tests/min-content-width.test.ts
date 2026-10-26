import { describe, expect, it } from 'vitest'
import type { Paragraph, TextBody } from '@genoffice/pptx-engine'
import { makeViewport } from '../src/coords'
import { OpentypeMetrics, type OpentypeFontLike } from '../src/metrics'
import { layoutText, minContentWidthPx } from '../src/text-layout'
import { TABLE_COL_FIT_SLACK_PX, tableColFloors, tableMinScale } from '../src/table-floor'

const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000) // scale 1

// Latin advance 0.5em, bold 0.6em; CJK falls back to the heuristic (1em per full-width char)
const font: OpentypeFontLike = {
  unitsPerEm: 1000,
  ascender: 800,
  descender: -200,
  getAdvanceWidth: (t, s) => t.length * s * 0.5,
  charToGlyphIndex: (ch) => (ch.codePointAt(0)! < 0x2e80 ? 1 : 0),
}
const metrics = new OpentypeMetrics(() => font)
const PX = 96 / 72 // 1pt in px at scale 1

function body(paras: Paragraph[], extra: Partial<TextBody> = {}): TextBody {
  return {
    paragraphs: paras,
    anchor: 'top',
    insets: { l: 91440, t: 45720, r: 91440, b: 45720 }, // 9.6px each side
    autofit: 'none',
    wrap: true,
    ...extra,
  }
}
const para = (text: string, fontSize = 20, extra: Record<string, unknown> = {}): Paragraph =>
  ({ runs: [{ text, fontSize, ...extra }] }) as Paragraph
const INS = 2 * 9.6
const need = (b: TextBody) => minContentWidthPx({ body: b, metrics, vp })

describe('minContentWidthPx', () => {
  it('Latin text: the longest word plus the left/right insets', () => {
    const w = need(body([para('非 Redmine Issue')]))
    // "Redmine" = 7 chars × 0.5em × 20pt; the CJK char is 1em and shorter than that
    expect(w).toBeCloseTo(7 * 0.5 * 20 * PX + INS, 3)
  })

  it('CJK text: one full-width character, however long the line', () => {
    const w = need(body([para('元件區塊的顯示異常')]))
    expect(w).toBeCloseTo(20 * PX + INS, 3)
  })

  it('mixed text: the widest of the Latin words and the full-width characters', () => {
    // "RWD" = 3 × 0.5em = 1.5em is wider than one 1em character; a lone "ab" is not
    expect(need(body([para('元件 RWD 顯示')]))).toBeCloseTo(1.5 * 20 * PX + INS, 3)
    expect(need(body([para('元件 ab 顯示')]))).toBeCloseTo(20 * PX + INS, 3)
  })

  it('bold and size change the word width (the widest paragraph decides)', () => {
    const small = need(body([para('Redmine', 12)]))
    const big = need(body([para('Redmine', 12), para('Issue', 30)]))
    expect(big).toBeGreaterThan(small)
    expect(big).toBeCloseTo(5 * 0.5 * 30 * PX + INS, 3)
  })

  it('bold runs measure with their own (wider) advance', () => {
    const boldMetrics = {
      metrics: () => ({ ascent: 16, descent: 4, lineHeight: 20 }),
      measure: (t: string, st: { fontSizePx: number; bold: boolean }) =>
        t.length * st.fontSizePx * (st.bold ? 0.6 : 0.5),
    }
    const w = (bold: boolean) =>
      minContentWidthPx({
        body: body([para('Redmine', 20, { bold })]),
        metrics: boldMetrics as never,
        vp,
      })
    expect(w(true) - w(false)).toBeCloseTo(7 * 0.1 * 20 * PX, 3)
  })

  it('letter spacing widens every word', () => {
    const plain = need(body([para('Redmine')]))
    const spaced = need(body([para('Redmine', 20, { letterSpacing: 3 })]))
    expect(spaced - plain).toBeCloseTo(7 * 3 * PX, 3)
  })

  it('a dash is a break opportunity: only the longer half counts', () => {
    const w = need(body([para('Cloud-Edgefoo')]))
    expect(w).toBeCloseTo(7 * 0.5 * 20 * PX + INS, 3) // "Cloud-" = 6 chars, "Edgefoo" = 7
  })

  it('empty text, spaces only and wrap-off bodies require nothing', () => {
    expect(need(body([para('')]))).toBe(0)
    expect(need(body([para('   ')]))).toBe(0)
    expect(need(body([para('Redmine')], { wrap: false }))).toBe(0)
  })

  it('paragraph margins and a hanging bullet add to the first word', () => {
    const base = need(body([para('Redmine')]))
    const marL = need(body([{ ...para('Redmine'), marL: 190500 } as Paragraph])) // 20px
    expect(marL - base).toBeCloseTo(20, 3)
    const bullet = need(
      body([
        { ...para('Redmine'), marL: 95250, indent: -95250, bullet: { type: 'char', char: '•' } },
      ] as Paragraph[]),
    )
    // hanging indent: the bullet sits in the margin and the word starts at marL (+ any glyph overflow)
    expect(bullet).toBeGreaterThanOrEqual(base + 10 - 1e-6)
  })

  it('kinsoku keeps a closing mark with its character', () => {
    const w = need(body([para('異常。')]))
    expect(w).toBeCloseTo(2 * 20 * PX + INS, 3)
  })

  it('is exactly the width below which layoutText starts cutting a word', () => {
    const b = body([para('非 Redmine Issue 元件區塊的 RWD')])
    const min = need(b)
    const lines = (wpx: number) =>
      layoutText({ body: b, boxWidthPx: wpx, boxHeightPx: 200, metrics, vp }).lines.map((l) =>
        l.runs.map((r) => r.text).join(''),
      )
    const words = ['Redmine', 'Issue', 'RWD']
    const intact = (rows: string[]) =>
      words.every((w) => rows.some((r) => r.includes(w))) &&
      rows.every((r) => !/^[a-z]+$/i.test(r) || words.includes(r.trim()))
    expect(intact(lines(min + TABLE_COL_FIT_SLACK_PX))).toBe(true)
    // one pixel narrower: the widest word no longer fits and is cut
    const cut = lines(min - 1).join('|')
    expect(cut).not.toContain('Redmine')
  })
})

describe('table column floors', () => {
  it('floor = max(need + slack, 24) capped at the current width', () => {
    const f = tableColFloors([200, 50, 300], [100, 90, 0])
    expect(f[0]).toBeCloseTo(100 + TABLE_COL_FIT_SLACK_PX, 6)
    expect(f[1]).toBe(50) // already narrower than its need: stays as it is
    expect(f[2]).toBe(24) // empty column: the absolute floor
  })

  it('the table can shrink until the first column reaches its floor', () => {
    expect(tableMinScale([200, 400], [100, 100])).toBeCloseTo(100.5 / 200, 6)
    expect(tableMinScale([200, 400], [0, 0])).toBeCloseTo(24 / 200, 6)
    expect(tableMinScale([50], [90])).toBe(1)
  })
})
