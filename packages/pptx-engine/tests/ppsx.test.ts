import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  createBlankPptx,
  openPptx,
  savePptx,
  savePptxToFile,
  setSlideNotes,
  getSlideNotes,
} from '../src/index'

describe('PowerPoint Show packages', () => {
  it('converts both ways, preserving notes and all parts other than the package type', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-ppsx-'))
    try {
      const opened = await openPptx(await createBlankPptx())
      setSlideNotes(opened, 0, 'Speaker notes')
      const source = await openPptx(await savePptx(opened))
      const path = join(dir, 'show.PPSX')
      await savePptxToFile(source, path)
      const show = await openPptx(await readFile(path))
      expect(show.archive.readText('[Content_Types].xml')).toContain(
        'presentationml.slideshow.main+xml',
      )
      expect(source.archive.readText('[Content_Types].xml')).toContain(
        'presentationml.presentation.main+xml',
      )
      for (const [name, bytes] of source.archive.entries) {
        if (name !== '[Content_Types].xml') expect(show.archive.entries.get(name)).toEqual(bytes)
      }
      setSlideNotes(show, 0, 'Updated notes')
      await savePptxToFile(show, path)
      const saved = await openPptx(await readFile(path))
      expect(getSlideNotes(saved.archive, saved.deck.slides[0]!.path)).toBe('Updated notes')
      expect(saved.archive.readText('[Content_Types].xml')).toContain(
        'presentationml.slideshow.main+xml',
      )
      const converted = join(dir, 'editable.pptx')
      await savePptxToFile(saved, converted)
      const deck = await openPptx(await readFile(converted))
      expect(deck.archive.readText('[Content_Types].xml')).toContain(
        'presentationml.presentation.main+xml',
      )
      expect(getSlideNotes(deck.archive, deck.deck.slides[0]!.path)).toBe('Updated notes')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it('a failed conversion leaves the current document type unchanged', async () => {
    const opened = await openPptx(await createBlankPptx())
    const before = opened.archive.readText('[Content_Types].xml')
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-ppsx-failure-'))
    try {
      await expect(savePptxToFile(opened, join(dir, 'missing', 'show.ppsx'))).rejects.toThrow()
      expect(opened.archive.readText('[Content_Types].xml')).toBe(before)
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })
})
