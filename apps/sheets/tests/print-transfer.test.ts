import { describe, expect, it } from 'vitest'

import { resolvePrintRequest, storePrintHtmlChunk } from '../src/main/print-transfer'
import type { WorkbookExportPdfRequest } from '../src/shared/desktop-api'

const base = {
  fileName: 'a.pdf',
  landscape: false,
  pageSize: 'A4',
  margins: { top: 0, bottom: 0, left: 0, right: 0 },
  scale: 1,
} as const

describe('chunked print HTML', () => {
  it('reassembles chunks in index order regardless of arrival order', () => {
    storePrintHtmlChunk('t1', 1, 'BBB')
    storePrintHtmlChunk('t1', 0, 'AAA')
    storePrintHtmlChunk('t1', 2, 'CCC')
    const request = { ...base, html: '', htmlTransferId: 't1' } as WorkbookExportPdfRequest
    expect(resolvePrintRequest(request).html).toBe('AAABBBCCC')
    // a transfer is consumed once
    expect(() => resolvePrintRequest(request)).toThrow()
  })
  it('refuses an incomplete transfer', () => {
    storePrintHtmlChunk('t2', 1, 'B')
    const request = { ...base, html: '', htmlTransferId: 't2' } as WorkbookExportPdfRequest
    expect(() => resolvePrintRequest(request)).toThrow()
  })
  it('passes a plain request through and rejects an empty one', () => {
    expect(
      resolvePrintRequest({ ...base, html: '<p>x</p>' } as WorkbookExportPdfRequest).html,
    ).toBe('<p>x</p>')
    expect(() => resolvePrintRequest({ ...base, html: '' } as WorkbookExportPdfRequest)).toThrow()
  })
})
