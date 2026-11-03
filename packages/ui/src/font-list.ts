import type { Lang } from '@genoffice/i18n'

/**
 * Common Western / Latin typography candidates widely used across Office,
 * Windows, macOS, and standard document workflows.
 */
export const LATIN_FONT_FAMILIES: readonly string[] = [
  'Aptos',
  'Aptos Display',
  'Arial',
  'Arial Black',
  'Arial Narrow',
  'Bahnschrift',
  'Book Antiqua',
  'Bookman Old Style',
  'Calibri',
  'Calibri Light',
  'Cambria',
  'Cambria Math',
  'Candara',
  'Century Gothic',
  'Comic Sans MS',
  'Consolas',
  'Constantia',
  'Corbel',
  'Courier New',
  'Franklin Gothic Medium',
  'Gabriola',
  'Garamond',
  'Georgia',
  'Gill Sans MT',
  'Helvetica',
  'Helvetica Neue',
  'Impact',
  'Lucida Console',
  'Lucida Sans Unicode',
  'Palatino Linotype',
  'Segoe UI',
  'Segoe UI Semibold',
  'Sitka',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
]

/**
 * Traditional Chinese fonts covering Windows (PMingLiU, JhengHei, DFKai),
 * macOS (PingFang TC, BiauKai, Songti TC), and open-source standards (Source Han / Noto).
 */
export const TRADITIONAL_CHINESE_FONT_FAMILIES: readonly string[] = [
  '微軟正黑體',
  '微軟正黑體 Light',
  '新細明體',
  '細明體',
  '標楷體',
  '蘋方-繁',
  '蘋方-港',
  '黑體-繁',
  '宋體-繁',
  '楷體-繁',
  '圓體-繁',
  '思源黑體',
  '思源宋體',
  'BiauKai',
  'DFKai-SB',
  'Microsoft JhengHei',
  'Microsoft JhengHei Light',
  'MingLiU',
  'Noto Sans TC',
  'Noto Serif TC',
  'PingFang HK',
  'PingFang TC',
  'PMingLiU',
]

/**
 * Simplified Chinese fonts covering standard Office/Windows fonts,
 * GB/T 9704 official document fonts, macOS system fonts, and Noto/Source Han.
 */
export const SIMPLIFIED_CHINESE_FONT_FAMILIES: readonly string[] = [
  '等线',
  '等线 Light',
  '微软雅黑',
  '微软雅黑 Light',
  '宋体',
  '新宋体',
  '黑体',
  '楷体',
  '仿宋',
  '仿宋_GB2312',
  '楷体_GB2312',
  '方正小标宋简体',
  '苹方-简',
  '华文中宋',
  '华文宋体',
  '华文黑体',
  '华文楷体',
  '华文仿宋',
  '华文细黑',
  '华文行楷',
  '华文新魏',
  '思源黑体',
  '思源宋体',
  'DengXian',
  'DengXian Light',
  'FangSong',
  'KaiTi',
  'Microsoft YaHei',
  'Microsoft YaHei Light',
  'Noto Sans SC',
  'Noto Serif SC',
  'NSimSun',
  'PingFang SC',
  'SimHei',
  'SimSun',
]

/**
 * Japanese fonts covering standard Windows, macOS, and modern UI/open-source fonts.
 */
export const JAPANESE_FONT_FAMILIES: readonly string[] = [
  'Yu Gothic',
  'Yu Gothic Medium',
  'Yu Mincho',
  'Meiryo',
  'Meiryo UI',
  'MS Gothic',
  'MS PGothic',
  'MS Mincho',
  'MS PMincho',
  'MS UI Gothic',
  'Hiragino Sans',
  'Hiragino Mincho ProN',
  'BIZ UDPGothic',
  'BIZ UDPMincho',
  'Noto Sans JP',
  'Noto Serif JP',
]

/**
 * Korean fonts covering standard Windows, macOS, and modern Korean typography.
 */
export const KOREAN_FONT_FAMILIES: readonly string[] = [
  'Malgun Gothic',
  'Batang',
  'BatangChe',
  'Gulim',
  'GulimChe',
  'Dotum',
  'DotumChe',
  'Gungsuh',
  'Apple SD Gothic Neo',
  'NanumGothic',
  'NanumMyeongjo',
  'Noto Sans KR',
  'Noto Serif KR',
]

const ALL_CANDIDATE_FONTS: readonly string[] = [
  ...LATIN_FONT_FAMILIES,
  ...TRADITIONAL_CHINESE_FONT_FAMILIES,
  ...SIMPLIFIED_CHINESE_FONT_FAMILIES,
  ...JAPANESE_FONT_FAMILIES,
  ...KOREAN_FONT_FAMILIES,
]

/**
 * Distinct built-in font families sorted alphabetically / by name.
 */
export const BUILTIN_FONT_FAMILIES: readonly string[] = Array.from(
  new Set(ALL_CANDIDATE_FONTS),
).sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base', numeric: true }))

/**
 * Sort font family names alphabetically / by name according to target language locale.
 */
export function sortFontFamilies(families: readonly string[], lang?: Lang): string[] {
  const locale = lang === 'zh' ? 'zh-CN' : lang === 'zh-TW' ? 'zh-TW' : (lang ?? 'en')
  return [...families].sort((a, b) =>
    a.localeCompare(b, locale, { sensitivity: 'base', numeric: true }),
  )
}

/**
 * Return candidate font families for the specified UI language, sorted by name
 * using the locale's natural collation order (e.g. stroke count for zh-TW, Pinyin
 * for zh-CN, alphabetical A-Z for Latin and other scripts).
 *
 * For market-specific locales (e.g. Traditional Chinese 'zh-TW', Japanese 'ja',
 * Korean 'ko', Simplified Chinese 'zh'), only the matching script plus standard
 * Latin typography are offered by default, avoiding cluttering the menu with
 * irrelevant scripts (e.g. no Simplified Chinese fonts under Traditional Chinese).
 */
export function fontFamiliesFor(lang: Lang): readonly string[] {
  let candidates: readonly string[]
  switch (lang) {
    case 'zh-TW':
      candidates = [...TRADITIONAL_CHINESE_FONT_FAMILIES, ...LATIN_FONT_FAMILIES]
      break
    case 'zh':
      candidates = [...SIMPLIFIED_CHINESE_FONT_FAMILIES, ...LATIN_FONT_FAMILIES]
      break
    case 'ja':
      candidates = [...JAPANESE_FONT_FAMILIES, ...LATIN_FONT_FAMILIES]
      break
    case 'ko':
      candidates = [...KOREAN_FONT_FAMILIES, ...LATIN_FONT_FAMILIES]
      break
    default:
      candidates = [
        ...LATIN_FONT_FAMILIES,
        ...TRADITIONAL_CHINESE_FONT_FAMILIES,
        ...SIMPLIFIED_CHINESE_FONT_FAMILIES,
        ...JAPANESE_FONT_FAMILIES,
        ...KOREAN_FONT_FAMILIES,
      ]
      break
  }
  return sortFontFamilies(Array.from(new Set(candidates)), lang)
}

