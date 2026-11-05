import { expect, it } from 'vitest'
import type { TextBody } from '@genoffice/pptx-engine'
import { textLayoutKey } from '../src/main/text-layout-key'
const original: TextBody = {
  autofit: 'resize',
  anchor: 'middle',
  paragraphs: [
    {
      align: 'center',
      runs: [
        {
          text: '2',
          fontFamily: 'Rockwell',
          fontSize: 11,
          field: 'slidenum',
          color: '#898989',
          colorInherited: true,
        },
      ],
    },
  ],
}
it('keeps geometry for field colors and DOM-authoritative false flags', () => {
  const changed = structuredClone(original),
    r = changed.paragraphs[0]!.runs[0]!
  r.color = '#FFFFFF'
  delete r.colorInherited
  r.bold = false
  r.italic = false
  r.underline = false
  r.strike = false
  expect(textLayoutKey(changed)).toBe(textLayoutKey(original))
})
it('refits changed content, font metrics and paragraph flow', () => {
  for (const edit of [
    (b: TextBody) => {
      b.paragraphs[0]!.runs[0]!.text = '22'
      delete b.paragraphs[0]!.runs[0]!.field
    },
    (b: TextBody) => (b.paragraphs[0]!.runs[0]!.fontSize = 24),
    (b: TextBody) => (b.paragraphs[0]!.runs[0]!.bold = true),
    (b: TextBody) => (b.paragraphs[0]!.lineHeight = 150),
    (b: TextBody) => (b.paragraphs[0]!.runs[0]!.fontFamily = 'Arial'),
  ]) {
    const changed = structuredClone(original)
    edit(changed)
    expect(textLayoutKey(changed)).not.toBe(textLayoutKey(original))
  }
})

it('keeps the box when a native field cache becomes its evaluated page number', () => {
  const before = structuredClone(original),
    after = structuredClone(original)
  before.paragraphs[0]!.runs[0]!.text = '‹#›'
  after.paragraphs[0]!.runs[0]!.text = '2'
  after.paragraphs[0]!.runs[0]!.color = '#FFFFFF'
  expect(textLayoutKey(after)).toBe(textLayoutKey(before))
})
