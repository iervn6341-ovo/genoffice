export type HeaderFooterVariant = 'odd' | 'first' | 'even'
export type HeaderFooterSection =
  | 'header-left'
  | 'header-center'
  | 'header-right'
  | 'footer-left'
  | 'footer-center'
  | 'footer-right'

export function headerFooterPicturePosition(
  variant: HeaderFooterVariant,
  section: HeaderFooterSection,
): string {
  const [kind, align] = section.split('-')
  const prefix = align === 'left' ? 'L' : align === 'right' ? 'R' : 'C'
  return `${prefix}${kind === 'header' ? 'H' : 'F'}${variant === 'odd' ? '' : variant.toUpperCase()}`
}

export function hasHeaderFooterPictureCode(text: string): boolean {
  for (let index = 0; index < text.length - 1; index += 1) {
    if (text[index] !== '&') continue
    const next = text[index + 1]?.toUpperCase()
    if (next === '"') {
      const end = text.indexOf('"', index + 2)
      if (end < 0) return false
      index = end
      continue
    }
    if (next === '&') {
      index += 1
      continue
    }
    if (next === 'G') return true
  }
  return false
}

export function withoutHeaderFooterPictureCode(text: string): string {
  let result = ''
  for (let index = 0; index < text.length; index += 1) {
    if (text[index] === '&' && index + 1 < text.length) {
      const next = text[index + 1]?.toUpperCase()
      if (next === '"') {
        const end = text.indexOf('"', index + 2)
        if (end < 0) return result + text.slice(index)
        result += text.slice(index, end + 1)
        index = end
        continue
      }
      if (next === '&') {
        result += '&&'
        index += 1
        continue
      }
      if (next === 'G') {
        index += 1
        continue
      }
    }
    result += text[index]
  }
  return result
}

export function withHeaderFooterPictureCode(text: string): string {
  return hasHeaderFooterPictureCode(text) ? text : `${text}&G`
}
