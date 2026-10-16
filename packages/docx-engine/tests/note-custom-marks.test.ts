import { readFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, wordSymbolForMark, type SaveBlock } from '../src/index'

const sourceUrl = new URL(
  '../../../docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/word-mac-custom-marks-source.docx',
  import.meta.url,
)

const visibleOrder = (
  doc: Awaited<ReturnType<typeof parseDocx>>,
): Array<Extract<SaveBlock, { kind: 'original' }>> =>
  doc.blocks
    .filter((block) => !block.hidden && block.docxIndex !== null)
    .map((block) => ({ kind: 'original', docxIndex: block.docxIndex! }))

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file(path)!.async('string')
}

describe('Word Mac custom note marks', () => {
  it('reads the native reference and note mark as one atom without including it in note text', async () => {
    const source = new Uint8Array(await readFile(sourceUrl))
    const parsed = await parseDocx(source)
    expect(parsed.footnotes[0]).toMatchObject({
      id: '1',
      text: 'Custom symbol footnote body.',
      customMark: '♠',
      customMarkSymbol: { font: 'Symbol', char: 'F0AA' },
    })
    expect(parsed.endnotes[0]).toMatchObject({
      id: '1',
      text: 'Custom symbol endnote body.',
      customMark: '♥',
      customMarkSymbol: { font: 'Symbol', char: 'F0A9' },
    })
    expect(parsed.footnotes[0].noRefMark).toBeUndefined()
    expect(parsed.endnotes[0].noRefMark).toBeUndefined()
    const marks = parsed.blocks.flatMap((block) => block.runs ?? []).filter((run) => run.noteRef)
    expect(marks.map((run) => run.noteRef)).toEqual([
      {
        kind: 'footnote',
        id: '1',
        customMark: '♠',
        customMarkSymbol: { font: 'Symbol', char: 'F0AA' },
      },
      {
        kind: 'endnote',
        id: '1',
        customMark: '♥',
        customMarkSymbol: { font: 'Symbol', char: 'F0A9' },
      },
    ])
    const saved = await saveDocx(parsed, visibleOrder(parsed), {
      footnotes: parsed.footnotes,
      endnotes: parsed.endnotes,
    })
    for (const path of ['word/footnotes.xml', 'word/endnotes.xml']) {
      const original = await part(source, path)
      const output = await part(saved, path)
      const tag = path.includes('footnotes') ? 'footnote' : 'endnote'
      const entry = new RegExp(`<w:${tag} w:id="1">[\\s\\S]*?<\\/w:${tag}>`).exec(original)?.[0]
      expect(entry).toBeTruthy()
      expect(output).toContain(entry!)
    }
  })

  it('keeps custom marks native through neighboring text edits and does not consume the next automatic number', async () => {
    const parsed = await parseDocx(new Uint8Array(await readFile(sourceUrl)))
    const paragraph = parsed.blocks.find((block) => block.runs?.some((run) => run.noteRef))!
    const blocks: SaveBlock[] = visibleOrder(parsed).map((block) =>
      block.docxIndex === paragraph.docxIndex
        ? {
            kind: 'generated',
            block: {
              ...paragraph,
              type: 'paragraph' as const,
              runs: [{ text: 'Before ' }, ...paragraph.runs!],
            },
          }
        : block,
    )
    blocks.push({
      kind: 'generated',
      block: {
        type: 'paragraph',
        runs: [{ text: 'Automatic ' }, { text: '1', noteRef: { kind: 'footnote', id: '2' } }],
      },
    })
    const saved = await saveDocx(parsed, blocks, {
      footnotes: [...parsed.footnotes, { id: '2', text: 'Automatic note body.' }],
      endnotes: parsed.endnotes,
    })
    const body = await part(saved, 'word/document.xml')
    expect(body).toContain('Before ')
    expect(body).toMatch(
      /<w:footnoteReference[^>]*w:customMarkFollows="1"[^>]*\/><w:sym w:font="Symbol" w:char="F0AA"\/>/,
    )
    expect(body).toMatch(
      /<w:endnoteReference[^>]*w:customMarkFollows="1"[^>]*\/><w:sym w:font="Symbol" w:char="F0A9"\/>/,
    )
    const reopened = await parseDocx(saved)
    expect(reopened.noteNumbers?.['footnote:2']).toBe(1)
    expect(reopened.footnotes.map((note) => note.text)).toEqual([
      'Custom symbol footnote body.',
      'Automatic note body.',
    ])
  })

  it('reverses a mark change after save while retaining both native note bodies', async () => {
    const original = await parseDocx(new Uint8Array(await readFile(sourceUrl)))
    const paragraph = original.blocks.find((block) => block.runs?.some((run) => run.noteRef))!
    const changedRun = paragraph.runs!.map((run) =>
      run.noteRef?.kind === 'footnote'
        ? {
            ...run,
            noteRef: { ...run.noteRef, customMark: '♥', customMarkSymbol: wordSymbolForMark('♥') },
          }
        : run,
    )
    const changeBlocks: SaveBlock[] = visibleOrder(original).map((block) =>
      block.docxIndex === paragraph.docxIndex
        ? {
            kind: 'generated',
            block: { ...paragraph, type: 'paragraph' as const, runs: changedRun },
          }
        : block,
    )
    const changed = await parseDocx(
      await saveDocx(original, changeBlocks, {
        footnotes: [
          { ...original.footnotes[0], customMark: '♥', customMarkSymbol: wordSymbolForMark('♥') },
        ],
        endnotes: original.endnotes,
      }),
    )
    expect(changed.footnotes[0].customMark).toBe('♥')
    expect(await part(changed.internal.originalBytes, 'word/footnotes.xml')).toContain(
      'w:char="F0A9"',
    )
    const changedParagraph = changed.blocks.find((block) => block.runs?.some((run) => run.noteRef))!
    const restoredRuns = changedParagraph.runs!.map((run) =>
      run.noteRef?.kind === 'footnote'
        ? {
            ...run,
            noteRef: { ...run.noteRef, customMark: '♠', customMarkSymbol: wordSymbolForMark('♠') },
          }
        : run,
    )
    const restored = await parseDocx(
      await saveDocx(
        changed,
        visibleOrder(changed).map((block) =>
          block.docxIndex === changedParagraph.docxIndex
            ? {
                kind: 'generated',
                block: { ...changedParagraph, type: 'paragraph' as const, runs: restoredRuns },
              }
            : block,
        ),
        {
          footnotes: [
            { ...changed.footnotes[0], customMark: '♠', customMarkSymbol: wordSymbolForMark('♠') },
          ],
          endnotes: changed.endnotes,
        },
      ),
    )
    expect(restored.footnotes[0].customMark).toBe('♠')
    expect(restored.endnotes[0].customMark).toBe('♥')
    expect(await part(restored.internal.originalBytes, 'word/footnotes.xml')).toContain(
      'w:char="F0AA"',
    )
  })

  it('edits Word Mac note text without changing the custom symbol or the other note', async () => {
    const original = await parseDocx(new Uint8Array(await readFile(sourceUrl)))
    const beforeEndnote = await part(original.internal.originalBytes, 'word/endnotes.xml')
    const saved = await saveDocx(original, visibleOrder(original), {
      footnotes: [{ ...original.footnotes[0], text: 'Edited custom body.' }],
      endnotes: original.endnotes,
    })
    const reopened = await parseDocx(saved)
    expect(reopened.footnotes[0]).toMatchObject({
      text: 'Edited custom body.',
      customMark: '♠',
      customMarkSymbol: { font: 'Symbol', char: 'F0AA' },
    })
    expect(await part(saved, 'word/footnotes.xml')).toContain(
      '<w:sym w:font="Symbol" w:char="F0AA"/>',
    )
    const originalEntry = /<w:endnote w:id="1">[\s\S]*?<\/w:endnote>/.exec(beforeEndnote)?.[0]
    expect(originalEntry).toBeTruthy()
    expect(await part(saved, 'word/endnotes.xml')).toContain(originalEntry!)
  })

  it('round-trips a typed Unicode custom mark and keeps it outside editable note text', async () => {
    const source = await parseDocx(
      await (
        await import('./helpers/build-docx')
      ).buildDocx({
        bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
      }),
    )
    const created = await parseDocx(
      await saveDocx(
        source,
        [
          ...visibleOrder(source),
          {
            kind: 'generated',
            block: {
              type: 'paragraph',
              runs: [
                { text: 'Custom ' },
                { text: '1', noteRef: { kind: 'footnote', id: '1', customMark: '§' } },
              ],
            },
          },
        ],
        { footnotes: [{ id: '1', text: 'First body.', customMark: '§' }] },
      ),
    )
    expect(created.footnotes[0]).toMatchObject({ text: 'First body.', customMark: '§' })
    const edited = await parseDocx(
      await saveDocx(created, visibleOrder(created), {
        footnotes: [{ ...created.footnotes[0], text: 'Second body.' }],
      }),
    )
    expect(edited.footnotes[0]).toMatchObject({ text: 'Second body.', customMark: '§' })
    expect(await part(edited.internal.originalBytes, 'word/footnotes.xml')).toContain(
      '<w:t xml:space="preserve">§</w:t>',
    )
  })

  it('protects a custom reference run mixed with other text instead of dropping its children', async () => {
    const { buildDocx } = await import('./helpers/build-docx')
    const mixedRun =
      '<w:p><w:r><w:footnoteReference w:customMarkFollows="1" w:id="1"/>' +
      '<w:sym w:font="Symbol" w:char="F0AA"/><w:t>TAIL</w:t></w:r></w:p>'
    const notes =
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:id="1"><w:p><w:r><w:sym w:font="Symbol" w:char="F0AA"/></w:r>' +
      '<w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>'
    const source = await parseDocx(
      await buildDocx({
        bodyXml: mixedRun,
        extraRels:
          '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
        extraParts: [
          {
            path: 'word/footnotes.xml',
            xml: notes,
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
          },
        ],
      }),
    )
    expect(source.blocks[0].type).toBe('passthrough')
    const saved = await saveDocx(source, visibleOrder(source), { footnotes: source.footnotes })
    expect(await part(saved, 'word/document.xml')).toContain(mixedRun)
    expect(await part(saved, 'word/footnotes.xml')).toContain('<w:t>Note body</w:t>')
  })
})
