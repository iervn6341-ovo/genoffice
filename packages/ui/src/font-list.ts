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

export interface SubstituteFontOption {
  readonly name: string
  readonly label: string
  readonly group: 'system' | 'office'
}

/**
 * Curated font candidate list for font substitution settings.
 * Prioritizes macOS/system-native and GenOffice-bundled fonts,
 * followed by standard Microsoft Office fonts.
 */
export const RECOMMENDED_SUBSTITUTE_FONTS: readonly SubstituteFontOption[] = [
  // ── System and bundled fonts (prioritized for macOS, cross-platform and open source) ──
  // Traditional Chinese
  { name: 'PingFang TC', label: 'PingFang TC (蘋方-繁)', group: 'system' },
  { name: 'PingFang HK', label: 'PingFang HK (蘋方-港)', group: 'system' },
  { name: 'Songti TC', label: 'Songti TC (宋體-繁)', group: 'system' },
  { name: 'Heiti TC', label: 'Heiti TC (黑體-繁)', group: 'system' },
  { name: 'Kaiti SC', label: 'Kaiti SC (楷體/華文楷體)', group: 'system' },
  { name: 'Kaiti TC', label: 'Kaiti TC (楷體-繁)', group: 'system' },
  { name: 'Yuanti TC', label: 'Yuanti TC (圓體-繁)', group: 'system' },
  { name: 'BiauKai', label: 'BiauKai (標楷體-系統)', group: 'system' },
  { name: '思源黑體', label: '思源黑體 (Source Han Sans TC)', group: 'system' },
  { name: '思源宋體', label: '思源宋體 (Source Han Serif TC)', group: 'system' },
  { name: 'Noto Sans TC', label: 'Noto Sans TC', group: 'system' },
  { name: 'Noto Serif TC', label: 'Noto Serif TC', group: 'system' },
  { name: 'GenOffice Songti TC', label: 'GenOffice Songti TC (內建開源)', group: 'system' },

  // Simplified Chinese
  { name: 'PingFang SC', label: 'PingFang SC (苹方-简)', group: 'system' },
  { name: 'Songti SC', label: 'Songti SC (宋体-简)', group: 'system' },
  { name: 'Heiti SC', label: 'Heiti SC (黑体-简)', group: 'system' },
  { name: '华文黑体', label: '华文黑体 (STHeiti)', group: 'system' },
  { name: '华文宋体', label: '华文宋体 (STSong)', group: 'system' },
  { name: '华文楷体', label: '华文楷体 (STKaiti)', group: 'system' },
  { name: '华文仿宋', label: '华文仿宋 (STFangsong)', group: 'system' },
  { name: '思源黑体', label: '思源黑体 (Source Han Sans SC)', group: 'system' },
  { name: '思源宋体', label: '思源宋体 (Source Han Serif SC)', group: 'system' },
  { name: '方正小标宋简体', label: '方正小标宋简体', group: 'system' },
  { name: 'Noto Sans SC', label: 'Noto Sans SC', group: 'system' },
  { name: 'Noto Serif SC', label: 'Noto Serif SC', group: 'system' },

  // Japanese
  { name: 'Hiragino Sans', label: 'Hiragino Sans (ヒラギノ角ゴ)', group: 'system' },
  { name: 'Hiragino Mincho ProN', label: 'Hiragino Mincho ProN (ヒラギノ明朝)', group: 'system' },
  { name: 'BIZ UDPGothic', label: 'BIZ UDPGothic', group: 'system' },
  { name: 'BIZ UDPMincho', label: 'BIZ UDPMincho', group: 'system' },
  { name: 'Noto Sans JP', label: 'Noto Sans JP', group: 'system' },
  { name: 'Noto Serif JP', label: 'Noto Serif JP', group: 'system' },

  // Korean
  { name: 'Apple SD Gothic Neo', label: 'Apple SD Gothic Neo', group: 'system' },
  { name: 'NanumGothic', label: 'NanumGothic (나눔고딕)', group: 'system' },
  { name: 'NanumMyeongjo', label: 'NanumMyeongjo (나눔명조)', group: 'system' },
  { name: 'Noto Sans KR', label: 'Noto Sans KR', group: 'system' },
  { name: 'Noto Serif KR', label: 'Noto Serif KR', group: 'system' },

  // Common Latin
  { name: 'Carlito GO', label: 'Carlito GO (相容 Calibri, 內建開源)', group: 'system' },
  { name: 'Caladea', label: 'Caladea (相容 Cambria, 內建開源)', group: 'system' },
  { name: 'Arial', label: 'Arial', group: 'system' },
  { name: 'Times New Roman', label: 'Times New Roman', group: 'system' },
  { name: 'Helvetica Neue', label: 'Helvetica Neue', group: 'system' },

  // ── Microsoft Office fonts ──
  // Traditional Chinese
  { name: '微軟正黑體', label: '微軟正黑體 (Microsoft JhengHei)', group: 'office' },
  { name: '微軟正黑體 Light', label: '微軟正黑體 Light (Microsoft JhengHei Light)', group: 'office' },
  { name: '新細明體', label: '新細明體 (PMingLiU)', group: 'office' },
  { name: '細明體', label: '細明體 (MingLiU)', group: 'office' },
  { name: '標楷體', label: '標楷體 (DFKai-SB)', group: 'office' },

  // Simplified Chinese
  { name: '微软雅黑', label: '微软雅黑 (Microsoft YaHei)', group: 'office' },
  { name: '微软雅黑 Light', label: '微软雅黑 Light (Microsoft YaHei Light)', group: 'office' },
  { name: '宋体', label: '宋体 (SimSun)', group: 'office' },
  { name: '新宋体', label: '新宋体 (NSimSun)', group: 'office' },
  { name: '黑体', label: '黑体 (SimHei)', group: 'office' },
  { name: '楷体', label: '楷体 (KaiTi)', group: 'office' },
  { name: '仿宋', label: '仿宋 (FangSong)', group: 'office' },
  { name: '等线', label: '等线 (DengXian)', group: 'office' },
  { name: '等线 Light', label: '等线 Light (DengXian Light)', group: 'office' },
  { name: '华文中宋', label: '华文中宋 (STZhongsong)', group: 'office' },
  { name: '华文细黑', label: '华文细黑 (STXihei)', group: 'office' },
  { name: '华文行楷', label: '华文行楷 (STXingkai)', group: 'office' },
  { name: '华文新魏', label: '华文新魏 (STXinwei)', group: 'office' },

  // Japanese
  { name: 'Yu Gothic', label: 'Yu Gothic (游ゴシック)', group: 'office' },
  { name: 'Yu Gothic Medium', label: 'Yu Gothic Medium', group: 'office' },
  { name: 'Yu Mincho', label: 'Yu Mincho (游明朝)', group: 'office' },
  { name: 'Meiryo', label: 'Meiryo (メイリオ)', group: 'office' },
  { name: 'Meiryo UI', label: 'Meiryo UI', group: 'office' },
  { name: 'MS Gothic', label: 'MS Gothic (ＭＳ ゴシック)', group: 'office' },
  { name: 'MS PGothic', label: 'MS PGothic (ＭＳ Ｐゴシック)', group: 'office' },
  { name: 'MS Mincho', label: 'MS Mincho (ＭＳ 明朝)', group: 'office' },
  { name: 'MS PMincho', label: 'MS PMincho (ＭＳ Ｐ明朝)', group: 'office' },
  { name: 'MS UI Gothic', label: 'MS UI Gothic', group: 'office' },

  // Korean
  { name: 'Malgun Gothic', label: 'Malgun Gothic (맑은 고딕)', group: 'office' },
  { name: 'Batang', label: 'Batang (바탕)', group: 'office' },
  { name: 'BatangChe', label: 'BatangChe', group: 'office' },
  { name: 'Gulim', label: 'Gulim (굴림)', group: 'office' },
  { name: 'GulimChe', label: 'GulimChe', group: 'office' },
  { name: 'Dotum', label: 'Dotum (돋움)', group: 'office' },
  { name: 'DotumChe', label: 'DotumChe', group: 'office' },
  { name: 'Gungsuh', label: 'Gungsuh (궁서)', group: 'office' },

  // Common Latin
  { name: 'Calibri', label: 'Calibri', group: 'office' },
  { name: 'Calibri Light', label: 'Calibri Light', group: 'office' },
  { name: 'Cambria', label: 'Cambria', group: 'office' },
  { name: 'Cambria Math', label: 'Cambria Math', group: 'office' },
  { name: 'Segoe UI', label: 'Segoe UI', group: 'office' },
  { name: 'Consolas', label: 'Consolas', group: 'office' },
  { name: 'Aptos', label: 'Aptos', group: 'office' },
  { name: 'Aptos Display', label: 'Aptos Display', group: 'office' },
]

