/**
 * Home → Font controls applied to a selected shape (no text editing): character spacing,
 * text highlight and Change Case must reach the saved XML and survive a reopen.
 */
import { describe, it, expect } from 'vitest'
import PptxGenJS from 'pptxgenjs'
import JSZip from 'jszip'
import { openPptx, savePptx, setElementFont } from '../src/index'
import type { TextElement } from '../src/types'

async function deck(text: string): Promise<Uint8Array> {
  const p = new PptxGenJS()
  p.addSlide().addText(text, { x: 1, y: 1, w: 8, h: 1, fontSize: 24 })
  const buf = (await p.write({ outputType: 'nodebuffer' })) as Buffer
  return new Uint8Array(buf)
}

async function firstText(bytes: Uint8Array) {
  const opened = await openPptx(bytes)
  const slide = opened.deck.slides[0]!
  const el = slide.elements.find((e) => e.type === 'text' || e.type === 'shape') as TextElement
  return { opened, slide, el }
}

async function slideXml(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('ppt/slides/slide1.xml')!.async('string')
}

describe('setElementFont: Home tab font extras', () => {
  it('character spacing writes spc (1/100 pt) and reopens with the same value', async () => {
    const { opened, slide, el } = await firstText(await deck('Hello'))
    expect(setElementFont(slide, el.id, { letterSpacingPt: 3 })).toBe(true)
    const saved = await savePptx(opened)
    expect(await slideXml(saved)).toMatch(/spc="300"/)
    const again = await firstText(saved)
    expect(again.el.text!.paragraphs[0]!.runs[0]!.letterSpacing).toBe(3)
  })

  it('"Normal" spacing resets spc to 0 (PowerPoint\'s normal)', async () => {
    const { opened, slide, el } = await firstText(await deck('Hello'))
    setElementFont(slide, el.id, { letterSpacingPt: 3 })
    const once = await savePptx(opened)
    const again = await firstText(once)
    setElementFont(again.slide, again.el.id, { letterSpacingPt: 0 })
    const xml = await slideXml(await savePptx(again.opened))
    expect(xml).not.toMatch(/spc="(?!0")/)
    const third = await firstText(await savePptx(again.opened))
    expect(third.el.text!.paragraphs[0]!.runs[0]!.letterSpacing ?? 0).toBe(0)
  })

  it('highlight writes <a:highlight> and null removes it', async () => {
    const { opened, slide, el } = await firstText(await deck('Hello'))
    setElementFont(slide, el.id, { highlight: '#FFFF00' })
    const saved = await savePptx(opened)
    const xml = await slideXml(saved)
    expect(xml).toMatch(/<a:highlight><a:srgbClr val="FFFF00"\/><\/a:highlight>/)
    // CT_TextCharacterProperties order: fill, highlight, then the font slots
    const rpr = /<a:rPr\b[\s\S]*?<\/a:rPr>/.exec(xml)![0]
    const names = [...rpr.matchAll(/<(a:[A-Za-z]+)\b/g)].map((m) => m[1]).slice(1)
    const hi = names.indexOf('a:highlight')
    for (const late of ['a:latin', 'a:ea', 'a:cs', 'a:hlinkClick']) {
      const at = names.indexOf(late)
      if (at >= 0) expect(at, `${late} after a:highlight in ${rpr}`).toBeGreaterThan(hi)
    }
    for (const early of ['a:ln', 'a:solidFill']) {
      const at = names.indexOf(early)
      if (at >= 0) expect(at, `${early} before a:highlight in ${rpr}`).toBeLessThan(hi)
    }
    const again = await firstText(saved)
    expect(again.el.text!.paragraphs[0]!.runs[0]!.highlight?.toUpperCase()).toBe('#FFFF00')
    setElementFont(again.slide, again.el.id, { highlight: null })
    expect(await slideXml(await savePptx(again.opened))).not.toMatch(/a:highlight/)
  })

  it('Change Case rewrites the text and keeps it after reopen', async () => {
    const { opened, slide, el } = await firstText(await deck('hello WORLD. again'))
    setElementFont(slide, el.id, { textCase: 'sentence' })
    const again = await firstText(await savePptx(opened))
    expect(again.el.text!.paragraphs[0]!.runs.map((r) => r.text).join('')).toBe(
      'Hello world. Again',
    )
    setElementFont(again.slide, again.el.id, { textCase: 'upper' })
    const third = await firstText(await savePptx(again.opened))
    expect(third.el.text!.paragraphs[0]!.runs.map((r) => r.text).join('')).toBe(
      'HELLO WORLD. AGAIN',
    )
  })
})
