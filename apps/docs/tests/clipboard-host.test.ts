import { describe, expect, it, vi } from 'vitest'
import type { NumberingDef, ParsedDocFull, StyleInfo, StyleUpsert } from '@genoffice/docx-engine'
import {
  createClipHost,
  upsertFromStyleInfo,
  withStyleId,
  type ClipHostAccess,
} from '../src/renderer/editor/clipboard-host'

const info = (over: Partial<StyleInfo> = {}): StyleInfo => ({
  styleId: 'Callout',
  name: 'Callout',
  type: 'paragraph',
  basedOn: 'Normal',
  qFormat: true,
  display: {
    bold: true,
    color: 'C00000',
    sizeHalfPoints: 26,
    fontAscii: 'Georgia',
    align: 'center',
    spaceAfterTwips: 120,
    indentLeftTwips: 720,
    lineSpacing: 1.5,
  },
  ...over,
})

function fakeAccess(parsed: Partial<ParsedDocFull> | null, extra: Partial<ClipHostAccess> = {}) {
  const upserts: StyleUpsert[] = []
  const levelsSeen: unknown[] = []
  const access: ClipHostAccess = {
    parsed: () => parsed as ParsedDocFull | null,
    addStyleUpsert: (up) => upserts.push(up),
    styleSheetChanged: vi.fn(),
    createListDef: (levels) => {
      levelsSeen.push(levels)
      return '77'
    },
    numbering: () => new Map(),
    ...extra,
  }
  return { access, upserts, levelsSeen }
}

describe('style import', () => {
  it('re-ids a copied styles.xml element', () => {
    expect(
      withStyleId('<w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/></w:style>', 'A2'),
    ).toContain('w:styleId="A2"')
  })

  it('records the verbatim element for the save and teaches the style map + stylesheet', () => {
    const parsed = {
      styles: new Map<string, StyleInfo>(),
      extras: { styleXml: {} },
    } as unknown as ParsedDocFull
    const { access, upserts } = fakeAccess(parsed)
    const host = createClipHost(access)
    const id = host.dest().importStyle!('Callout', {
      info: info(),
      xml: '<w:style w:type="paragraph" w:styleId="Callout"><w:name w:val="Callout"/></w:style>',
    })
    expect(id).toBe('Callout')
    expect(parsed.styles.get('Callout')?.name).toBe('Callout')
    expect(upserts[0]).toMatchObject({ styleId: 'Callout', type: 'paragraph' })
    expect(upserts[0]!.rawXml).toContain('w:styleId="Callout"')
    expect(access.styleSheetChanged).toHaveBeenCalled()
  })

  it('without the source element, rebuilds the definition from the display hints', () => {
    const up = upsertFromStyleInfo(info(), 'Callout')
    expect(up).toMatchObject({
      styleId: 'Callout',
      name: 'Callout',
      basedOn: 'Normal',
      rPr: { bold: true, color: 'C00000', sizeHalfPoints: 26, font: 'Georgia' },
      pPr: { align: 'center', spaceAfterTwips: 120, indentLeftTwips: 720, lineSpacing: 1.5 },
    })
    expect(up.rawXml).toBeUndefined()
  })
})

describe('numbering import', () => {
  it('rebuilds the levels (with the marker font) as a new list definition', () => {
    const { access, levelsSeen } = fakeAccess(null)
    const def: NumberingDef = {
      numId: '5',
      abstractNumId: '3',
      levels: {
        0: {
          numFmt: 'bullet',
          lvlText: '',
          start: 1,
          indentLeft: 720,
          hanging: 360,
          font: 'Symbol',
        },
        1: { numFmt: 'lowerLetter', lvlText: '(%2)', start: 3, indentLeft: 1440 },
      },
      startOverrides: {},
    }
    expect(createClipHost(access).dest().importNumbering!(def)).toBe('77')
    expect(levelsSeen[0]).toEqual([
      { numFmt: 'bullet', lvlText: '', indentLeft: 720, hanging: 360, start: 1, font: 'Symbol' },
      { numFmt: 'lowerLetter', lvlText: '(%2)', indentLeft: 1440, hanging: 360, start: 3 },
    ])
  })
})

describe('notes', () => {
  it('renumbers the markers once, only after a paste that created notes', () => {
    const normalize = vi.fn((tr) => tr)
    const { access } = fakeAccess(null, { addNote: () => '9', normalizeNotes: normalize })
    const host = createClipHost(access)
    const fakeTr = { docChanged: true } as never
    expect(host.afterPaste!(fakeTr)).toBeNull() // nothing imported yet
    expect(host.dest().importNote!('footnote', { id: '2', text: 'body' })).toBe('9')
    expect(host.afterPaste!(fakeTr)).toBe(fakeTr)
    expect(normalize).toHaveBeenCalledTimes(1)
    expect(host.afterPaste!(fakeTr)).toBeNull() // consumed
  })
})
