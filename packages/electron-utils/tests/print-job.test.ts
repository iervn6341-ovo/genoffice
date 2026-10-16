import { describe, expect, it, vi } from 'vitest'
import {
  DEFAULT_HTML_PRINT_OPTIONS,
  defaultPaperFor,
  normalizeHtmlPrintOptions,
  parsePrintRanges,
  printEntryFor,
  printLoadedDocument,
  webContentsPrintArgs,
  withPrintPageCss,
  type PrintableContents,
} from '../src/print-job'
import { printJobStringsFor } from '../src/print-job-strings'

describe('printEntryFor', () => {
  it('sends macOS straight to the native panel and everything else to the in-app sheet', () => {
    expect(printEntryFor('MacIntel')).toBe('native-panel')
    expect(printEntryFor('macOS')).toBe('native-panel')
    expect(printEntryFor('Win32')).toBe('in-app')
    expect(printEntryFor('Linux x86_64')).toBe('in-app')
  })
})

describe('parsePrintRanges', () => {
  it('reads singles, ranges and open-ended ranges (1-based)', () => {
    expect(parsePrintRanges('')).toEqual([])
    expect(parsePrintRanges('3')).toEqual([{ from: 3, to: 3 }])
    expect(parsePrintRanges('1-3, 5，8-')).toEqual([
      { from: 1, to: 3 },
      { from: 5, to: 5 },
      { from: 8, to: 100_000 },
    ])
  })

  it('rejects zero, reversed and malformed text', () => {
    for (const bad of ['0', '3-1', 'a', '1--2', '-4', '1;x']) {
      expect(parsePrintRanges(bad), bad).toBeNull()
    }
  })
})

describe('normalizeHtmlPrintOptions', () => {
  it('fills defaults and clamps untrusted input', () => {
    expect(normalizeHtmlPrintOptions(undefined)).toEqual(DEFAULT_HTML_PRINT_OPTIONS)
    expect(
      normalizeHtmlPrintOptions({ paper: 'Foo', scale: 9000, margin: 'wide', landscape: 'yes' }),
    ).toMatchObject({ paper: 'A4', scale: 200, margin: 'wide', landscape: false })
    expect(normalizeHtmlPrintOptions({ scale: 1 }).scale).toBe(10)
    expect(normalizeHtmlPrintOptions({ scale: 'x' }).scale).toBe(100)
  })
})

describe('defaultPaperFor', () => {
  it('uses Letter in Letter regions and A4 elsewhere', () => {
    expect(defaultPaperFor('US')).toBe('Letter')
    expect(defaultPaperFor('us')).toBe('Letter')
    expect(defaultPaperFor('TW')).toBe('A4')
    expect(defaultPaperFor(undefined)).toBe('A4')
  })
})

describe('webContentsPrintArgs', () => {
  it('maps every option onto the Electron print call (pages counted from 0)', () => {
    const args = webContentsPrintArgs(
      {
        paper: 'Letter',
        landscape: true,
        margin: 'none',
        scale: 80,
        headerFooter: true,
        background: false,
        pages: '2-3,7',
      },
      'My doc',
    )
    expect(args).toEqual({
      silent: false,
      printBackground: false,
      landscape: true,
      scaleFactor: 80,
      pageSize: 'Letter',
      pageRanges: [
        { from: 1, to: 2 },
        { from: 6, to: 6 },
      ],
      header: 'My doc',
      footer: ' ',
    })
  })

  it('omits ranges and header/footer when unused, and rejects a malformed range', () => {
    const args = webContentsPrintArgs(DEFAULT_HTML_PRINT_OPTIONS, 't')
    expect(args).not.toHaveProperty('pageRanges')
    expect(args).not.toHaveProperty('header')
    expect(webContentsPrintArgs({ ...DEFAULT_HTML_PRINT_OPTIONS, pages: '9-2' }, 't')).toBeNull()
  })
})

describe('withPrintPageCss', () => {
  it('writes the chosen margin into the document', () => {
    const html = withPrintPageCss('<html><head></head><body>x</body></html>', {
      ...DEFAULT_HTML_PRINT_OPTIONS,
      margin: 'narrow',
    })
    expect(html).toContain('@page { margin: 0.25in; }')
    expect(html.indexOf('@page')).toBeLessThan(html.indexOf('</head>'))
    expect(withPrintPageCss('<p>x</p>', DEFAULT_HTML_PRINT_OPTIONS)).toMatch(/^<style>.*0\.6in/)
  })
})

describe('printLoadedDocument', () => {
  const run = (success: boolean, reason: string) => {
    const contents: PrintableContents = { print: vi.fn((_o, cb) => cb(success, reason)) }
    return printLoadedDocument(contents, DEFAULT_HTML_PRINT_OPTIONS, 't')
  }

  it('maps the print callback to printed / canceled / error', async () => {
    expect(await run(true, '')).toEqual({ status: 'printed' })
    expect(await run(false, 'Print job canceled')).toEqual({ status: 'canceled' })
    expect(await run(false, 'cancelled')).toEqual({ status: 'canceled' })
    expect(await run(false, 'no printer')).toEqual({ status: 'error', message: 'no printer' })
  })

  it('reports a thrown print call and a bad range as errors', async () => {
    const throwing: PrintableContents = {
      print: () => {
        throw new Error('boom')
      },
    }
    expect(await printLoadedDocument(throwing, DEFAULT_HTML_PRINT_OPTIONS, 't')).toEqual({
      status: 'error',
      message: 'boom',
    })
    const never: PrintableContents = { print: vi.fn() }
    const bad = await printLoadedDocument(never, { ...DEFAULT_HTML_PRINT_OPTIONS, pages: 'x' }, 't')
    expect(bad.status).toBe('error')
    expect(never.print).not.toHaveBeenCalled()
  })
})

describe('printJobStringsFor', () => {
  it('keeps one key set and the same placeholders in every language', () => {
    const en = printJobStringsFor('en')
    for (const lang of ['zh', 'zh-TW', 'ja', 'ko', 'fr']) {
      const strings = printJobStringsFor(lang)
      expect(Object.keys(strings).sort()).toEqual(Object.keys(en).sort())
      expect(strings.printSheetFailed).toContain('{error}')
    }
    expect(printJobStringsFor('zh-TW').printSheetTitle).toBe('列印')
    expect(printJobStringsFor('fr')).toEqual(en)
  })
})
