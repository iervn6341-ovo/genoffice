/** Chromium substitutes pageNumber but cannot add Excel's &P offset. */
const PAGE_OFFSET = /<span data-genoffice-page-offset="([+-])(\d+)"><\/span>/g

export function hasHeaderFooterPageOffset(template: string | undefined): boolean {
  return template?.includes('data-genoffice-page-offset=') ?? false
}

/** Resolve only renderer-generated offset spans in a single page's template. */
export function headerFooterTemplateForPage(
  template: string | undefined,
  page: number,
  total: number,
): string | undefined {
  return template
    ?.replace(PAGE_OFFSET, (_whole, sign: string, digits: string) => {
      const offset = BigInt(digits)
      return String(BigInt(page) + (sign === '-' ? -offset : offset))
    })
    .replaceAll('<span class="totalPages"></span>', String(total))
}
