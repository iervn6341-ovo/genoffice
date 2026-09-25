/**
 * C8: Crop and Page Size rewrite the file in place and were not undoable. The
 * main process now keeps the bytes around each rewrite; the renderer's undo
 * stack carries a token (PageOpUndo) and swaps the file back through it.
 */
import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { PageOpHistory, sameBytes } from '../src/main/page-op-history'
import { cropPagesBytes, setPageSizeBytes } from '../src/main/save-pdf'
import { isPageOpUndo, type UndoEntry } from '../src/renderer/edit-state'

async function makePdf(sizes: [number, number][]): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  for (const size of sizes) doc.addPage(size)
  return doc.save({ useObjectStreams: false })
}

const pageSize = async (bytes: Uint8Array) => {
  const page = (await PDFDocument.load(bytes)).getPage(0)
  return { media: page.getSize(), crop: page.getCropBox() }
}

describe('PageOpHistory (C8)', () => {
  it('crop → undo → redo swaps exactly the recorded bytes', async () => {
    const history = new PageOpHistory()
    const before = await makePdf([[612, 792]])
    const after = await cropPagesBytes(before, [0], { l: 0.1, t: 0.1, r: 0.9, b: 0.9 })
    expect((await pageSize(after)).crop.width).toBeLessThan(612)
    const token = history.record('/a.pdf', before, after)!

    const undo = history.resolve('/a.pdf', token, 'undo', after)
    expect(undo.ok && sameBytes(undo.bytes, before)).toBe(true)
    expect((await pageSize((undo as { bytes: Uint8Array }).bytes)).crop.width).toBe(612)

    const redo = history.resolve('/a.pdf', token, 'redo', before)
    expect(redo.ok && sameBytes(redo.bytes, after)).toBe(true)
  })

  it('page size step round-trips the media box', async () => {
    const history = new PageOpHistory()
    const before = await makePdf([[612, 792]])
    const after = await setPageSizeBytes(before, 595.28, 841.89)
    const token = history.record('/b.pdf', before, after)!
    const undo = history.resolve('/b.pdf', token, 'undo', after)
    expect(undo.ok).toBe(true)
    expect((await pageSize((undo as { bytes: Uint8Array }).bytes)).media).toEqual({
      width: 612,
      height: 792,
    })
  })

  it('refuses when the file changed on disk or the step is unknown', async () => {
    const history = new PageOpHistory()
    const before = await makePdf([[100, 100]])
    const after = await makePdf([[200, 200]])
    const token = history.record('/c.pdf', before, after)!
    const other = await makePdf([[300, 300]])
    expect(history.resolve('/c.pdf', token, 'undo', other)).toMatchObject({ ok: false })
    expect(history.resolve('/c.pdf', token, 'redo', after)).toMatchObject({ ok: false })
    expect(history.resolve('/c.pdf', 'pageop-999', 'undo', after)).toMatchObject({ ok: false })
    expect(history.resolve('/other.pdf', token, 'undo', after)).toMatchObject({ ok: false })
  })

  it('keeps a bounded number of steps per file', () => {
    const history = new PageOpHistory()
    const b = new Uint8Array([1])
    const tokens = Array.from({ length: 7 }, () => history.record('/d.pdf', b, b)!)
    expect(history.resolve('/d.pdf', tokens[0]!, 'undo', b)).toMatchObject({ ok: false })
    expect(history.resolve('/d.pdf', tokens[6]!, 'undo', b)).toMatchObject({ ok: true })
  })

  it('undo entries distinguish page ops from edit snapshots', () => {
    const op: UndoEntry = { kind: 'pageOp', token: 't' }
    expect(isPageOpUndo(op)).toBe(true)
    expect(isPageOpUndo({ redactions: [] } as unknown as UndoEntry)).toBe(false)
    expect(isPageOpUndo(undefined)).toBe(false)
  })
})
