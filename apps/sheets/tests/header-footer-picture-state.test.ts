import { describe, expect, it } from 'vitest'

import {
  headerFooterPicturePosition,
  hasHeaderFooterPictureCode,
  withHeaderFooterPictureCode,
  withoutHeaderFooterPictureCode,
} from '../src/renderer/header-footer-picture-codes'
import { pageSetupPatch } from '../src/renderer/page-setup-undo'
import { MARGIN_PRESETS, resolveEffectivePageSetup } from '../src/renderer/print-settings'
import { workbookPageSetupStateSchema } from '../src/shared/desktop-api'

const PNG = 'iVBORw0KGgoAAAANSUhEUg=='
const old = {
  id: 'hf-picture-0-lh',
  position: 'LH',
  widthPt: 80,
  heightPt: 20,
  mediaType: 'image/png',
  base64: PNG,
}
const oldWithoutBytes = {
  id: old.id,
  position: old.position,
  widthPt: 120,
  heightPt: old.heightPt,
  mediaType: old.mediaType,
}
const shape = {
  orientation: 'portrait' as const,
  paperSize: 9,
  scale: 100,
  fitToWidth: 0,
  fitToHeight: 0,
  fitToPage: false,
  margins: MARGIN_PRESETS.normal,
  printGridlines: false,
  printHeadings: false,
  printAreas: [] as readonly string[],
  printTitles: null,
  header: null,
  footer: null,
  firstPage: { enabled: false, parts: null },
  evenPages: { enabled: false, parts: null },
}

describe('header/footer picture slots and codes', () => {
  it('maps all variant/section combinations to native VML positions', () => {
    expect(headerFooterPicturePosition('odd', 'header-left')).toBe('LH')
    expect(headerFooterPicturePosition('first', 'footer-right')).toBe('RFFIRST')
    expect(headerFooterPicturePosition('even', 'header-center')).toBe('CHEVEN')
  })

  it('changes only actual &G codes, preserving literal &&G text', () => {
    expect(hasHeaderFooterPictureCode('Logo &&G')).toBe(false)
    expect(withHeaderFooterPictureCode('Logo &&G')).toBe('Logo &&G&G')
    expect(withoutHeaderFooterPictureCode('Logo &&G&G&G')).toBe('Logo &&G')
    expect(withoutHeaderFooterPictureCode('&"A&G Font,Regular"Text&G')).toBe(
      '&"A&G Font,Regular"Text',
    )
  })

  it('overlays a pending image for PDF and retains unrelated first/even slots', () => {
    const setup = resolveEffectivePageSetup(
      {
        headerFooterPictureEdits: [
          { position: 'LH', image: null },
          {
            position: 'RFFIRST',
            image: { mediaType: 'image/png', base64: PNG },
            widthPt: 42,
            heightPt: 12,
          },
        ],
      },
      {
        headerFooterPictures: [
          old,
          {
            id: 'hf-picture-0-cheven',
            position: 'CHEVEN',
            widthPt: 18,
            heightPt: 10,
            mediaType: 'image/jpeg',
          },
        ],
      },
      null,
    )
    expect(setup.headerFooterPictures).toEqual([
      {
        id: 'hf-picture-0-cheven',
        position: 'CHEVEN',
        widthPt: 18,
        heightPt: 10,
        mediaType: 'image/jpeg',
      },
      {
        id: 'journal:RFFIRST',
        position: 'RFFIRST',
        widthPt: 42,
        heightPt: 12,
        mediaType: 'image/png',
      },
    ])
  })

  it('restores old bytes and size after a save only for the changed slot', () => {
    const saved = {
      ...shape,
      headerFooterPictures: [
        oldWithoutBytes,
        {
          id: 'hf-picture-0-ch',
          position: 'CH',
          widthPt: 30,
          heightPt: 12,
          mediaType: 'image/png',
        },
      ],
    }
    const before = {
      ...shape,
      headerFooterPictures: [
        old,
        {
          id: 'hf-picture-0-ch',
          position: 'CH',
          widthPt: 30,
          heightPt: 12,
          mediaType: 'image/png',
        },
      ],
    }
    expect(pageSetupPatch(saved, before, new Set(['LH']))).toEqual({
      headerFooterPictureEdits: [
        {
          position: 'LH',
          image: { mediaType: 'image/png', base64: PNG },
          widthPt: 80,
          heightPt: 20,
        },
      ],
    })
    expect(pageSetupPatch(saved, before, new Set())).toEqual({})
  })

  it('rejects oversized image payloads at the public save boundary', () => {
    expect(
      workbookPageSetupStateSchema.safeParse({
        sheetId: 'sheet',
        headerFooterPictureEdits: [{ position: 'BAD', image: null }],
      }).success,
    ).toBe(false)
    expect(
      workbookPageSetupStateSchema.safeParse({
        sheetId: 'sheet',
        headerFooterPictureEdits: [
          {
            position: 'LH',
            widthPt: 20,
            heightPt: 20,
            image: { mediaType: 'text/html', base64: PNG },
          },
        ],
      }).success,
    ).toBe(false)
  })
})
