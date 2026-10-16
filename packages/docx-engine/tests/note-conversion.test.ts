import { readFile } from 'node:fs/promises'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildNotesXml, convertNoteEntryXml } from '../src/notes'

const fixture = resolve(
  dirname(fileURLToPath(import.meta.url)),
  '../../../docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/conversion',
)

const parsed = async (name: string) =>
  parseDocx(new Uint8Array(await readFile(resolve(fixture, name))))

describe('native Word Mac note conversion', () => {
  it('moves a custom footnote into endnotes and reindexes an existing endnote without flattening either entry', async () => {
    const source = await parsed('word-mac-conversion-before.docx')
    const wordConverted = await parsed('word-mac-conversion-footnote-to-endnote.docx')
    const sourceReferences = source.blocks.flatMap((block) =>
      (block.runs ?? []).filter((run) => run.noteRef),
    )
    expect(sourceReferences.map((run) => [run.noteRef?.kind, run.rawRPr])).toEqual([
      ['footnote', '<w:rPr><w:rStyle w:val="af0"/></w:rPr>'],
      ['endnote', '<w:rPr><w:rStyle w:val="af3"/></w:rPr>'],
    ])
    expect(
      wordConverted.blocks
        .flatMap((block) => (block.runs ?? []).filter((run) => run.noteRef))
        .map((run) => run.rawRPr),
    ).toEqual(Array(2).fill('<w:rPr><w:rStyle w:val="af3"/></w:rPr>'))
    const sourceFoot = source.internal.noteEntryXml?.footnote['1']
    const sourceEnd = source.internal.noteEntryXml?.endnote['1']
    expect(sourceFoot).toBeTruthy()
    expect(sourceEnd).toBeTruthy()
    const styles = {
      sourceText: 'ae',
      targetText: 'af1',
      sourceReference: 'af0',
      targetReference: 'af3',
    }
    const moved = convertNoteEntryXml(sourceFoot!, 'footnote', 'endnote', '1', styles)
    const shifted = convertNoteEntryXml(sourceEnd!, 'endnote', 'endnote', '2')
    expect(moved).toMatch(/<w:endnote w:id="1">[\s\S]*?<w:pStyle w:val="af1"\/>/)
    expect(moved).toContain('<w:rStyle w:val="af3"/>')
    expect(moved).toContain('<w:sym w:font="Symbol" w:char="F0AA"/>')
    expect(shifted).toContain('w:id="2"')
    expect(shifted).toContain('w:char="F0A9"')
    const xml = buildNotesXml(
      'endnote',
      [
        { ...source.footnotes[0], id: '1', historyEntryXml: moved! },
        { ...source.endnotes[0], id: '2', historyEntryXml: shifted! },
      ],
      null,
    )
    expect(xml).toContain(moved!)
    expect(xml).toContain(shifted!)
    expect(wordConverted.endnotes.map(({ id, customMark }) => [id, customMark])).toEqual([
      ['1', '♠'],
      ['2', '♥'],
    ])
  })

  it('moves a custom endnote into footnotes with native style IDs and preserves the opposite source mark', async () => {
    const source = await parsed('word-mac-conversion-before.docx')
    const wordConverted = await parsed('word-mac-conversion-endnote-to-footnote.docx')
    expect(
      wordConverted.blocks
        .flatMap((block) => (block.runs ?? []).filter((run) => run.noteRef))
        .map((run) => run.rawRPr),
    ).toEqual(Array(2).fill('<w:rPr><w:rStyle w:val="af0"/></w:rPr>'))
    const moved = convertNoteEntryXml(
      source.internal.noteEntryXml!.endnote['1'],
      'endnote',
      'footnote',
      '2',
      {
        sourceText: 'af1',
        targetText: 'ae',
        sourceReference: 'af3',
        targetReference: 'af0',
      },
    )
    expect(moved).toMatch(/<w:footnote w:id="2">[\s\S]*?<w:pStyle w:val="ae"\/>/)
    expect(moved).toContain('<w:rStyle w:val="af0"/>')
    expect(moved).toContain('w:char="F0A9"')
    expect(wordConverted.footnotes.map(({ id, customMark }) => [id, customMark])).toEqual([
      ['1', '♠'],
      ['2', '♥'],
    ])
  })

  it('rejects relationship-bearing or field-containing entries instead of losing their XML', () => {
    const source = '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>'
    expect(
      convertNoteEntryXml(
        `${source}<w:r><w:drawing r:embed="rId1"/></w:r></w:p></w:footnote>`,
        'footnote',
        'endnote',
        '1',
      ),
    ).toBeNull()
    expect(
      convertNoteEntryXml(
        `${source}<w:r><w:instrText>DATE</w:instrText></w:r></w:p></w:footnote>`,
        'footnote',
        'endnote',
        '1',
      ),
    ).toBeNull()
  })

  it('carries inherited Word extension namespaces when the target part does not exist yet', async () => {
    const source = await parsed('word-mac-conversion-before.docx')
    const converted = convertNoteEntryXml(
      source.internal.noteEntryXml!.footnote['1'],
      'footnote',
      'endnote',
      '1',
      {
        sourceText: 'ae',
        targetText: 'af1',
        sourceReference: 'af0',
        targetReference: 'af3',
      },
    )
    const xml = buildNotesXml(
      'endnote',
      [
        {
          ...source.footnotes[0],
          historyEntryXml: converted!,
          historyEntryNamespaces: source.internal.notePartNamespaces!.footnote,
        },
      ],
      null,
    )
    expect(xml).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')
    expect(xml).toContain('w14:paraId=')
  })
})
