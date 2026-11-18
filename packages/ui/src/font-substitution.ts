/**
 * Font substitution types and default mapping rules for cross-platform
 * and missing font handling in GenOffice.
 */

export interface FontSubstitutionRule {
  /** Original font name as declared in the document, e.g. PMingLiU */
  original: string
  /** Substituted local font name, e.g. 'Songti TC' */
  substitute: string
  /** Whether this rule is a system-provided default */
  isDefault?: boolean
}

export const DEFAULT_FONT_SUBSTITUTIONS: readonly FontSubstitutionRule[] = [
  { original: '新細明體', substitute: 'Songti TC', isDefault: true },
  { original: '細明體', substitute: 'Songti TC', isDefault: true },
  { original: '標楷體', substitute: 'BiauKai', isDefault: true },
  { original: '微軟正黑體', substitute: 'PingFang TC', isDefault: true },
  { original: 'PMingLiU', substitute: 'Songti TC', isDefault: true },
  { original: 'MingLiU', substitute: 'Songti TC', isDefault: true },
  { original: 'DFKai-SB', substitute: 'BiauKai', isDefault: true },
  { original: 'Microsoft JhengHei', substitute: 'PingFang TC', isDefault: true },
  { original: '宋体', substitute: 'Songti SC', isDefault: true },
  { original: '微软雅黑', substitute: 'PingFang SC', isDefault: true },
  { original: '楷体', substitute: 'KaiTi SC', isDefault: true },
  { original: '黑体', substitute: 'Heiti SC', isDefault: true },
  { original: '仿宋', substitute: 'STFangsong', isDefault: true },
  { original: 'SimSun', substitute: 'Songti SC', isDefault: true },
  { original: 'Microsoft YaHei', substitute: 'PingFang SC', isDefault: true },
  { original: 'KaiTi', substitute: 'KaiTi SC', isDefault: true },
  { original: 'Calibri', substitute: 'Carlito GO', isDefault: true },
  { original: 'Cambria', substitute: 'Caladea', isDefault: true },
]

/**
 * Normalizes a list of font substitution rules to ensure valid entries,
 * trimming whitespace and removing empty originals.
 */
export function normalizeFontSubstitutions(
  rules: readonly FontSubstitutionRule[] | unknown,
): FontSubstitutionRule[] {
  if (!Array.isArray(rules)) return [...DEFAULT_FONT_SUBSTITUTIONS]
  const seen = new Set<string>()
  const result: FontSubstitutionRule[] = []
  for (const item of rules) {
    if (!item || typeof item !== 'object') continue
    const r = item as Partial<FontSubstitutionRule>
    const original = typeof r.original === 'string' ? r.original.trim() : ''
    const substitute = typeof r.substitute === 'string' ? r.substitute.trim() : ''
    if (!original || !substitute) continue
    const key = original.toLowerCase()
    if (seen.has(key)) continue
    seen.add(key)
    result.push({
      original,
      substitute,
      ...(r.isDefault ? { isDefault: true } : {}),
    })
  }
  return result
}

/**
 * The same face under another name. Chromium matches a family by the names in the font file, so a
 * face installed as "DFKai-SB" (Office, a hand-installed kaiu.ttf) does NOT answer to its localized name, and
 * the localized Windows name of an English-named font often misses too. Both directions are listed so
 * the availability check, the substitution chain, the missing-font report and the settings pane agree.
 */
const FONT_NAME_EQUIVALENTS: Record<string, readonly string[]> = {
  標楷體: ['DFKai-SB', 'BiauKai', 'BiauKaiTC', 'BiauKaiHK', '標楷體-繁', '標楷體-港澳'],
  'DFKai-SB': ['標楷體', 'BiauKai', 'BiauKaiTC', 'BiauKaiHK', '標楷體-繁', '標楷體-港澳'],
  BiauKai: ['BiauKaiTC', 'BiauKaiHK', '標楷體-繁', '標楷體-港澳', '標楷體', 'DFKai-SB'],
  BiauKaiTC: ['BiauKai', 'BiauKaiHK', '標楷體-繁', '標楷體-港澳', '標楷體', 'DFKai-SB'],
  BiauKaiHK: ['BiauKai', 'BiauKaiTC', '標楷體-繁', '標楷體-港澳', '標楷體', 'DFKai-SB'],
  '標楷體-繁': ['BiauKaiTC', 'BiauKai', '標楷體', 'DFKai-SB'],
  '標楷體-港澳': ['BiauKaiHK', 'BiauKai', '標楷體', 'DFKai-SB'],
  新細明體: ['PMingLiU'],
  PMingLiU: ['新細明體'],
  細明體: ['MingLiU'],
  MingLiU: ['細明體'],
  微軟正黑體: ['Microsoft JhengHei'],
  'Microsoft JhengHei': ['微軟正黑體'],
  '微軟正黑體 Light': ['Microsoft JhengHei Light'],
  'Microsoft JhengHei Light': ['微軟正黑體 Light'],
  微软雅黑: ['Microsoft YaHei'],
  'Microsoft YaHei': ['微软雅黑'],
  '微软雅黑 Light': ['Microsoft YaHei Light'],
  'Microsoft YaHei Light': ['微软雅黑 Light'],
  宋体: ['SimSun'],
  SimSun: ['宋体'],
  新宋体: ['NSimSun'],
  NSimSun: ['新宋体'],
  黑体: ['SimHei'],
  SimHei: ['黑体'],
  楷体: ['KaiTi'],
  KaiTi: ['楷体'],
  仿宋: ['FangSong'],
  FangSong: ['仿宋'],
  等线: ['DengXian'],
  DengXian: ['等线'],
  '等线 Light': ['DengXian Light'],
  'DengXian Light': ['等线 Light'],
  '蘋方-繁': ['PingFang TC'],
  'PingFang TC': ['蘋方-繁'],
  '蘋方-港': ['PingFang HK'],
  'PingFang HK': ['蘋方-港'],
  '苹方-简': ['PingFang SC'],
  'PingFang SC': ['苹方-简'],
  '宋體-繁': ['Songti TC'],
  'Songti TC': ['宋體-繁'],
  '黑體-繁': ['Heiti TC'],
  'Heiti TC': ['黑體-繁'],
  '楷體-繁': ['Kaiti SC', 'STKaiti', 'Kaiti TC', 'BiauKai'],
  'Kaiti SC': ['楷體-繁', 'STKaiti', '楷体', 'KaiTi'],
  STKaiti: ['Kaiti SC', '楷體-繁', '楷体', 'KaiTi'],
  '圓體-繁': ['Yuanti TC', 'Yuanti SC'],
  'Yuanti TC': ['圓體-繁'],
  'Yuanti SC': ['圓體-繁'],
  思源黑體: [
    'Source Han Sans TC',
    'Source Han Sans TW',
    'Source Han Sans HK',
    'Source Han Sans',
    'Noto Sans CJK TC',
    'Noto Sans TC',
  ],
  'Source Han Sans TC': ['思源黑體', 'Noto Sans TC'],
  思源宋體: [
    'Source Han Serif TC',
    'Source Han Serif TW',
    'Source Han Serif HK',
    'Source Han Serif',
    'Noto Serif CJK TC',
    'Noto Serif TC',
  ],
  'Source Han Serif TC': ['思源宋體', 'Noto Serif TC'],
  思源黑体: [
    'Source Han Sans SC',
    'Source Han Sans CN',
    'Source Han Sans',
    'Noto Sans CJK SC',
    'Noto Sans SC',
  ],
  'Source Han Sans SC': ['思源黑体', 'Noto Sans SC'],
  思源宋体: [
    'Source Han Serif SC',
    'Source Han Serif CN',
    'Source Han Serif',
    'Noto Serif CJK SC',
    'Noto Serif SC',
  ],
  'Source Han Serif SC': ['思源宋体', 'Noto Serif SC'],
  华文中宋: ['STZhongsong'],
  STZhongsong: ['华文中宋'],
  华文细黑: ['STXihei'],
  STXihei: ['华文细黑'],
  华文行楷: ['STXingkai'],
  STXingkai: ['华文行楷'],
  华文新魏: ['STXinwei'],
  STXinwei: ['华文新魏'],
  方正小标宋简体: ['FZXiaoBiaoSong-B05S', 'FZXiaoBiaoSong-B05'],
  'FZXiaoBiaoSong-B05S': ['方正小标宋简体', 'FZXiaoBiaoSong-B05'],
  'Songti SC': ['宋体', 'STSong'],
  STSong: ['Songti SC', '宋体'],
  'Heiti SC': ['黑体', 'STHeiti'],
  STHeiti: ['Heiti SC', '黑体'],
  'Yu Gothic': ['游ゴシック'],
  '游ゴシック': ['Yu Gothic'],
  'Yu Mincho': ['游明朝'],
  '游明朝': ['Yu Mincho'],
  Meiryo: ['メイリオ', 'Meiryo UI'],
  'メイリオ': ['Meiryo', 'Meiryo UI'],
  'MS Gothic': ['ＭＳ ゴシック', 'MS PGothic', 'MS UI Gothic'],
  'ＭＳ ゴシック': ['MS Gothic', 'MS PGothic', 'MS UI Gothic'],
  'MS Mincho': ['ＭＳ 明朝', 'MS PMincho'],
  'ＭＳ 明朝': ['MS Mincho', 'MS PMincho'],
  'Malgun Gothic': ['맑은 고딕'],
  '맑은 고딕': ['Malgun Gothic'],
  Batang: ['바탕', 'BatangChe'],
  '바탕': ['Batang', 'BatangChe'],
  Gulim: ['굴림', 'GulimChe', 'Dotum', 'DotumChe'],
  '굴림': ['Gulim', 'GulimChe', 'Dotum', 'DotumChe'],
  'Hiragino Sans': ['ヒラギノ角ゴ', 'Hiragino Kaku Gothic ProN'],
  'Hiragino Mincho ProN': ['ヒラギノ明朝'],
}

const FONT_NAME_EQUIVALENTS_LC = new Map(
  Object.entries(FONT_NAME_EQUIVALENTS).map(([k, v]) => [k.toLowerCase(), v] as const),
)

/** Other names the same face may be installed under (declared name excluded) */
export function fontNameEquivalents(font: string): readonly string[] {
  return FONT_NAME_EQUIVALENTS_LC.get(font.trim().toLowerCase()) ?? []
}

/** Check if a font family belongs to Microsoft Office font catalog */
export function isOfficeFamily(family: string): boolean {
  const norm = family.trim().toLowerCase()
  return (
    norm.includes('微軟') ||
    norm.includes('微软') ||
    norm.includes('細明') ||
    norm.includes('细明') ||
    norm.includes('mingliu') ||
    norm.includes('jhenghei') ||
    norm.includes('yahei') ||
    norm.includes('simsun') ||
    norm.includes('simhei') ||
    norm.includes('dengxian') ||
    norm.includes('等线') ||
    norm.includes('等線') ||
    norm.includes('標楷') ||
    norm.includes('标楷') ||
    norm.includes('dfkai') ||
    norm.includes('calibri') ||
    norm.includes('cambria') ||
    norm.includes('consolas') ||
    norm.includes('candara') ||
    norm.includes('constantia') ||
    norm.includes('corbel') ||
    norm.includes('aptos') ||
    norm.includes('segoe') ||
    norm.includes('meiryo') ||
    norm.includes('msgothic') ||
    norm.includes('ms gothic') ||
    norm.includes('ms pgothic') ||
    norm.includes('ms mincho') ||
    norm.includes('ms pmincho') ||
    norm.includes('ms ui gothic') ||
    norm.includes('yu gothic') ||
    norm.includes('yugoth') ||
    norm.includes('yu mincho') ||
    norm.includes('yumin') ||
    norm.includes('malgun') ||
    norm.includes('batang') ||
    norm.includes('gulim') ||
    norm.includes('dotum') ||
    norm.includes('gungsuh') ||
    norm.includes('zhongsong') ||
    norm.includes('xihei') ||
    norm.includes('xingkai') ||
    norm.includes('xinwei')
  )
}

/** Faces the app ships or aliases itself (fonts.css): they exist inside the editors, not as system fonts */
export function isAppBundledFamily(name: string): boolean {
  return /^(genoffice |carlito|caladea|liberation |noto )|( go)$/i.test(name.trim())
}

/**
 * Whether a family resolves on this machine, measured on a canvas: an unresolvable family inherits the
 * metrics of the generic fallback (document.fonts.check answers true for anything). False without a DOM.
 */
export function isFontInstalled(family: string): boolean {
  if (typeof document === 'undefined') return false
  try {
    const cx = document.createElement('canvas').getContext('2d')
    if (!cx) return false
    const sample = '한글あア中文abcWXYmm123永体宋黑'
    const q = `"${family.replace(/"/g, '')}"`
    const width = (stack: string): number => {
      cx.font = `32px ${stack}`
      return cx.measureText(sample).width
    }
    return (
      width(`${q}, monospace`) !== width('monospace') || width(`${q}, serif`) !== width('serif')
    )
  } catch {
    return false
  }
}

/**
 * A substitution target that nothing on this machine answers to (neither the name nor another name of
 * the same face): the rule cannot do what it says and the document silently renders with whatever
 * generic fallback follows. App-bundled stand-ins are exempt — they exist inside the editors.
 */
export function substituteIsUnavailable(substitute: string): boolean {
  const name = substitute.trim()
  if (!name || isAppBundledFamily(name)) return false
  return ![name, ...fontNameEquivalents(name)].some((n) => isFontInstalled(n))
}
