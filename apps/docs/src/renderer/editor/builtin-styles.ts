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

const charStyle = (
  styleId: string,
  name: string,
  rPr: NonNullable<StyleUpsert['rPr']>,
): BuiltinParaStyle => ({
  styleId,
  name,
  upsert: {
    styleId,
    type: 'character',
    name,
    quickFormat: true,
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
  // Thesis & academic document styles
  para(
    'Caption',
    'caption',
    { spaceBeforeTwips: 120, spaceAfterTwips: 120 },
    { font: 'Calibri', sizeHalfPoints: 18, italic: true, color: '44546A' },
  ),
  para(
    'Bibliography',
    'bibliography',
    { spaceAfterTwips: 120, indentLeftTwips: 720, firstLineTwips: -720 },
    { font: 'Calibri', sizeHalfPoints: 22 },
  ),
  para(
    'TOCHeading',
    'TOC Heading',
    { spaceBeforeTwips: 240, spaceAfterTwips: 120 },
    { font: 'Calibri Light', sizeHalfPoints: 32, bold: true, color: '2E74B5' },
  ),
  para('TableofFigures', 'table of figures', { spaceAfterTwips: 60 }),
  para('Header', 'header', { spaceAfterTwips: 0 }, { font: 'Calibri', sizeHalfPoints: 18 }),
  para('Footer', 'footer', { spaceAfterTwips: 0 }, { font: 'Calibri', sizeHalfPoints: 18 }),
  para('NormalWeb', 'Normal (Web)', { spaceAfterTwips: 120 }),
  para('TOC1', 'toc 1', { spaceAfterTwips: 100 }),
  para('TOC2', 'toc 2', { indentLeftTwips: 220, spaceAfterTwips: 100 }),
  para('TOC3', 'toc 3', { indentLeftTwips: 440, spaceAfterTwips: 100 }),
  para('TOC4', 'toc 4', { indentLeftTwips: 660, spaceAfterTwips: 100 }),
  para('TOC5', 'toc 5', { indentLeftTwips: 880, spaceAfterTwips: 100 }),
  para('TOC6', 'toc 6', { indentLeftTwips: 1100, spaceAfterTwips: 100 }),
  para('TOC7', 'toc 7', { indentLeftTwips: 1320, spaceAfterTwips: 100 }),
  para('TOC8', 'toc 8', { indentLeftTwips: 1540, spaceAfterTwips: 100 }),
  para('TOC9', 'toc 9', { indentLeftTwips: 1760, spaceAfterTwips: 100 }),
  charStyle('SubtleReference', 'Subtle Reference', { italic: true, color: '595959' }),
  charStyle('IntenseReference', 'Intense Reference', { bold: true, italic: true, color: '4472C4' }),
  charStyle('Hyperlink', 'Hyperlink', { underline: true, color: '0563C1' }),
]

export const builtinStyle = (styleId: string): BuiltinParaStyle | undefined =>
  BUILTIN_PARA_STYLES.find((s) => s.styleId === styleId)

const ZH_TW_STYLE_NAMES: Record<string, string> = {
  Normal: '內文',
  Heading1: '標題 1',
  Heading2: '標題 2',
  Heading3: '標題 3',
  Heading4: '標題 4',
  Heading5: '標題 5',
  Heading6: '標題 6',
  Heading7: '標題 7',
  Heading8: '標題 8',
  Heading9: '標題 9',
  'heading 1': '標題 1',
  'heading 2': '標題 2',
  'heading 3': '標題 3',
  'heading 4': '標題 4',
  'heading 5': '標題 5',
  'heading 6': '標題 6',
  'heading 7': '標題 7',
  'heading 8': '標題 8',
  'heading 9': '標題 9',
  NoSpacing: '無間距',
  'No Spacing': '無間距',
  Title: '書名',
  Subtitle: '副標題',
  Quote: '引文',
  IntenseQuote: '鮮明引文',
  'Intense Quote': '鮮明引文',
  SubtleReference: '區別參考',
  'Subtle Reference': '區別參考',
  IntenseReference: '鮮明參考',
  'Intense Reference': '鮮明參考',
  ListParagraph: '清單段落',
  'List Paragraph': '清單段落',
  Caption: '標號',
  caption: '標號',
  Bibliography: '書目',
  bibliography: '書目',
  TOC1: '目錄 1',
  TOC2: '目錄 2',
  TOC3: '目錄 3',
  TOC4: '目錄 4',
  TOC5: '目錄 5',
  TOC6: '目錄 6',
  TOC7: '目錄 7',
  TOC8: '目錄 8',
  TOC9: '目錄 9',
  'toc 1': '目錄 1',
  'toc 2': '目錄 2',
  'toc 3': '目錄 3',
  'toc 4': '目錄 4',
  'toc 5': '目錄 5',
  'toc 6': '目錄 6',
  'toc 7': '目錄 7',
  'toc 8': '目錄 8',
  'toc 9': '目錄 9',
  TOCHeading: '目錄標題',
  'TOC Heading': '目錄標題',
  TableofFigures: '圖表目錄',
  'table of figures': '圖表目錄',
  NormalWeb: '內文 (Web)',
  'Normal (Web)': '內文 (Web)',
  Header: '頁首',
  header: '頁首',
  Footer: '頁尾',
  footer: '頁尾',
  Hyperlink: '超連結',
}

const ZH_CN_STYLE_NAMES: Record<string, string> = {
  Normal: '正文',
  Heading1: '标题 1',
  Heading2: '标题 2',
  Heading3: '标题 3',
  Heading4: '标题 4',
  Heading5: '标题 5',
  Heading6: '标题 6',
  Heading7: '标题 7',
  Heading8: '标题 8',
  Heading9: '标题 9',
  'heading 1': '标题 1',
  'heading 2': '标题 2',
  'heading 3': '标题 3',
  'heading 4': '标题 4',
  'heading 5': '标题 5',
  'heading 6': '标题 6',
  'heading 7': '标题 7',
  'heading 8': '标题 8',
  'heading 9': '标题 9',
  NoSpacing: '无间隔',
  'No Spacing': '无间隔',
  Title: '标题',
  Subtitle: '副标题',
  Quote: '引用',
  IntenseQuote: '明显引用',
  'Intense Quote': '明显引用',
  SubtleReference: '微弱参考',
  'Subtle Reference': '微弱参考',
  IntenseReference: '明显参考',
  'Intense Reference': '明显参考',
  ListParagraph: '列表段落',
  'List Paragraph': '列表段落',
  Caption: '题注',
  caption: '题注',
  Bibliography: '书目',
  bibliography: '书目',
  TOC1: '目录 1',
  TOC2: '目录 2',
  TOC3: '目录 3',
  TOC4: '目录 4',
  TOC5: '目录 5',
  TOC6: '目录 6',
  TOC7: '目录 7',
  TOC8: '目录 8',
  TOC9: '目录 9',
  'toc 1': '目录 1',
  'toc 2': '目录 2',
  'toc 3': '目录 3',
  'toc 4': '目录 4',
  'toc 5': '目录 5',
  'toc 6': '目录 6',
  'toc 7': '目录 7',
  'toc 8': '目录 8',
  'toc 9': '目录 9',
  TOCHeading: '目录标题',
  'TOC Heading': '目录标题',
  TableofFigures: '图表目录',
  'table of figures': '图表目录',
  NormalWeb: '正文 (Web)',
  'Normal (Web)': '正文 (Web)',
  Header: '页眉',
  header: '页眉',
  Footer: '页脚',
  footer: '页脚',
  Hyperlink: '超链接',
}

/** Localizes standard Word style names according to UI language */
export function localizedStyleName(styleId: string, name: string, lang = 'zh-TW'): string {
  if (lang === 'zh-TW') {
    return ZH_TW_STYLE_NAMES[styleId] ?? ZH_TW_STYLE_NAMES[name] ?? name
  }
  if (lang.startsWith('zh')) {
    return ZH_CN_STYLE_NAMES[styleId] ?? ZH_CN_STYLE_NAMES[name] ?? name
  }
  return name
}

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
