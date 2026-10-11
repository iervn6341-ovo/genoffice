import type { StyleDisplay, StyleInfo, StyleUpsert } from '@genoffice/docx-engine'

/**
 * Word's built-in paragraph styles the Styles gallery offers beyond Normal and Heading 1–3. A
 * document that already defines one (any file Word wrote with it in use) keeps its own definition;
 * for the rest the definition below is added on first use — pending upsert for the save, and the
 * matching display info so the paragraph looks right at once. Values follow Word's default template.
 */
export interface BuiltinParaStyle {
  /** w:styleId */
  styleId: string
  /** the English built-in name Word writes into styles.xml */
  name: string
  upsert: StyleUpsert
}

const para = (
  styleId: string,
  name: string,
  pPr: NonNullable<StyleUpsert['pPr']>,
  rPr: NonNullable<StyleUpsert['rPr']> = {},
): BuiltinParaStyle => ({
  styleId,
  name,
  upsert: {
    styleId,
    type: 'paragraph',
    name,
    basedOn: 'Normal',
    next: 'Normal',
    quickFormat: true,
    pPr,
    rPr,
  },
})

export const BUILTIN_PARA_STYLES: readonly BuiltinParaStyle[] = [
  para('NoSpacing', 'No Spacing', { spaceAfterTwips: 0, lineSpacing: 1 }),
  para(
    'Title',
    'Title',
    { spaceAfterTwips: 0, lineSpacing: 1 },
    { font: 'Calibri Light', sizeHalfPoints: 56 },
  ),
  para('Subtitle', 'Subtitle', { spaceAfterTwips: 160 }, { color: '5A5A5A' }),
  para(
    'Quote',
    'Quote',
    {
      align: 'center',
      spaceBeforeTwips: 200,
      spaceAfterTwips: 160,
      indentLeftTwips: 864,
      indentRightTwips: 864,
    },
    { italic: true, color: '404040' },
  ),
  para(
    'IntenseQuote',
    'Intense Quote',
    {
      align: 'center',
      spaceBeforeTwips: 360,
      spaceAfterTwips: 360,
      indentLeftTwips: 864,
      indentRightTwips: 864,
    },
    { italic: true, color: '4472C4' },
  ),
  para('ListParagraph', 'List Paragraph', { indentLeftTwips: 720 }),
]

export const builtinStyle = (styleId: string): BuiltinParaStyle | undefined =>
  BUILTIN_PARA_STYLES.find((s) => s.styleId === styleId)

/** the display hints the editor's style CSS reads, from the same properties the upsert writes */
export function displayOfUpsert(up: StyleUpsert): StyleDisplay {
  const p = up.pPr ?? {}
  const r = up.rPr ?? {}
  const d: StyleDisplay = {}
  if (r.sizeHalfPoints != null) d.sizeHalfPoints = r.sizeHalfPoints
  if (r.bold !== undefined) d.bold = r.bold
  if (r.italic !== undefined) d.italic = r.italic
  if (r.underline !== undefined) d.underline = r.underline
  if (r.color) d.color = r.color
  if (r.font) d.fontAscii = r.font
  if (p.align) d.align = p.align
  if (p.spaceBeforeTwips != null) d.spaceBeforeTwips = p.spaceBeforeTwips
  if (p.spaceAfterTwips != null) d.spaceAfterTwips = p.spaceAfterTwips
  if (p.lineSpacing != null) {
    d.lineSpacing = p.lineSpacing
    d.lineRule = 'auto'
  }
  if (p.indentLeftTwips != null) d.indentLeftTwips = p.indentLeftTwips
  if (p.indentRightTwips != null) d.indentRightTwips = p.indentRightTwips
  if (p.firstLineTwips != null) d.indentFirstLineTwips = p.firstLineTwips
  return d
}

export function styleInfoOfUpsert(up: StyleUpsert): StyleInfo {
  return {
    styleId: up.styleId,
    name: up.name ?? up.styleId,
    type: up.type ?? 'paragraph',
    ...(up.basedOn ? { basedOn: up.basedOn } : {}),
    ...(up.next ? { next: up.next } : {}),
    qFormat: up.quickFormat ?? true,
    display: displayOfUpsert(up),
  }
}
