import JSZip from 'jszip'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, saveDocx } from '../src/index'

const dir = resolve(
  __dirname,
  '../../../docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/section-end-endnotes',
)

const keep = (parsed: Awaited<ReturnType<typeof parseDocx>>) =>
  parsed.blocks
    .filter((block) => !block.hidden && block.docxIndex !== null)
    .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! }))

describe('Word Mac section-end endnote probes', () => {
  it('reads the document-level position and keeps every endnote entry and reference', async () => {
    const source = await readFile(resolve(dir, 'word-mac-probe-settings-sectend-next-page.docx'))
    const parsed = await parseDocx(source)
    expect(parsed.endnoteSettings?.pos).toBe('sectEnd')
    expect(readSections(parsed)).toHaveLength(2)
    expect(parsed.endnotes).toHaveLength(16)
    // body-order numbering is continuous across sections (Word: i .. xvi)
    expect(parsed.noteNumbers?.['endnote:1']).toBe(1)
    expect(parsed.noteNumbers?.['endnote:16']).toBe(16)
  })

  it('saves untouched byte-for-byte for the note parts and switches the position through settings.xml only', async () => {
    const source = await readFile(resolve(dir, 'word-mac-probe-settings-sectend-next-page.docx'))
    const parsed = await parseDocx(source)
    const original = await JSZip.loadAsync(source)
    const same = await JSZip.loadAsync(await saveDocx(parsed, keep(parsed), {}))
    expect(await same.file('word/endnotes.xml')!.async('string')).toBe(
      await original.file('word/endnotes.xml')!.async('string'),
    )
    expect(await same.file('word/settings.xml')!.async('string')).toBe(
      await original.file('word/settings.xml')!.async('string'),
    )
    const moved = await saveDocx(parsed, keep(parsed), {
      noteProps: { endnote: { pos: 'docEnd' } },
    })
    const movedZip = await JSZip.loadAsync(moved)
    expect(await movedZip.file('word/settings.xml')!.async('string')).toContain(
      '<w:pos w:val="docEnd"/>',
    )
    expect(await movedZip.file('word/endnotes.xml')!.async('string')).toBe(
      await original.file('word/endnotes.xml')!.async('string'),
    )
    const reparsed = await parseDocx(moved)
    expect(reparsed.endnoteSettings?.pos).toBe('docEnd')
    expect(reparsed.endnotes).toHaveLength(16)
  })

  it('keeps a section-level position that Word ignores, without promoting it to a document setting', async () => {
    const source = await readFile(resolve(dir, 'word-mac-probe-sectpr-only-pos-ignored.docx'))
    const parsed = await parseDocx(source)
    expect(parsed.endnoteSettings?.pos).toBeUndefined()
    expect(readSections(parsed).map((section) => section.settings.endnotePr?.pos)).toEqual([
      'sectEnd',
      'sectEnd',
    ])
    const saved = await JSZip.loadAsync(await saveDocx(parsed, keep(parsed), {}))
    expect(await saved.file('word/document.xml')!.async('string')).toContain(
      '<w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr>',
    )
  })
})
