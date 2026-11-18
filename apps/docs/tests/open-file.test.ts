import { describe, expect, it } from 'vitest'
import {
  OPENABLE_DOC_EXTENSIONS,
  docImportKindForPath,
  findDocxPath,
  findOpenableDocPath,
  isOpenableDocPath,
} from '../src/shared/open-file'

describe('findOpenableDocPath', () => {
  it('finds Finder and Explorer document arguments case-insensitively', () => {
    expect(findOpenableDocPath(['/Applications/GenOffice Docs.app', '/tmp/Quarterly Plan.docx'])).toBe(
      '/tmp/Quarterly Plan.docx',
    )
    expect(findOpenableDocPath(['GenOffice Docs.exe', 'C:\\Users\\Me\\REPORT.DOCX'])).toBe(
      'C:\\Users\\Me\\REPORT.DOCX',
    )
    expect(findOpenableDocPath(['GenOffice Docs', '/tmp/legacy.doc'])).toBe('/tmp/legacy.doc')
    expect(findOpenableDocPath(['GenOffice Docs', '/tmp/TEMPLATE.DOT'])).toBe('/tmp/TEMPLATE.DOT')
    expect(findDocxPath(['GenOffice Docs', '/tmp/legacy.doc'])).toBe('/tmp/legacy.doc')
  })

  it('ignores Electron switches and unrelated files', () => {
    expect(findOpenableDocPath(['GenOffice Docs', '--inspect=document.docx', '/tmp/notes.txt'])).toBeNull()
  })
})

describe('OPENABLE_DOC_EXTENSIONS', () => {
  it('covers every extension the argv route and the guard accept', () => {
    expect(OPENABLE_DOC_EXTENSIONS).toEqual(['docx', 'doc', 'dot'])
    for (const extension of OPENABLE_DOC_EXTENSIONS) {
      expect(isOpenableDocPath(`/tmp/Quarterly Plan.${extension.toUpperCase()}`)).toBe(true)
    }
  })

  it('rejects the neighbouring formats Docs does not open', () => {
    for (const path of ['/tmp/notes.txt', '/tmp/legacy.rtf', '/tmp/deck.odp', '/tmp/book.ods']) {
      expect(isOpenableDocPath(path)).toBe(false)
    }
  })
})

describe('docImportKindForPath', () => {
  it('separates the import filter from the native and foreign formats', () => {
    expect(docImportKindForPath('/tmp/legacy.doc')).toBe('word97')
    expect(docImportKindForPath('/tmp/TEMPLATE.DOT')).toBe('word97')
    expect(docImportKindForPath('/tmp/report.docx')).toBeNull()
    expect(docImportKindForPath('/tmp/notes.txt')).toBeNull()
    expect(docImportKindForPath('/tmp/noextension')).toBeNull()
  })
})
