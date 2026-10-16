import { describe, expect, it } from 'vitest'

import {
  hasHeaderFooterPageOffset,
  headerFooterTemplateForPage,
} from '../src/main/pdf-header-offsets'

describe('PDF header/footer page offsets', () => {
  it('resolves renderer-generated positive and negative offsets for the source page', () => {
    const template =
      'Page <span data-genoffice-page-offset="+2"></span> / ' +
      '<span data-genoffice-page-offset="-4"></span> of ' +
      '<span class="totalPages"></span>'
    expect(hasHeaderFooterPageOffset(template)).toBe(true)
    expect(headerFooterTemplateForPage(template, 3, 8)).toBe('Page 5 / -1 of 8')
  })

  it('keeps plain page spans and adjacent text untouched', () => {
    const template =
      '<span class="pageNumber"></span>+2 ' + '<span data-genoffice-page-offset="+2"></span>x3'
    expect(headerFooterTemplateForPage(template, 2, 10)).toBe(
      '<span class="pageNumber"></span>+2 4x3',
    )
    expect(hasHeaderFooterPageOffset('<span class="pageNumber"></span>')).toBe(false)
  })

  it('does not lose precision for a large author-supplied offset', () => {
    expect(
      headerFooterTemplateForPage(
        '<span data-genoffice-page-offset="+9007199254740993"></span>',
        2,
        9,
      ),
    ).toBe('9007199254740995')
  })
})
