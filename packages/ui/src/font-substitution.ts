/**
 * Font substitution types and default mapping rules for cross-platform
 * and missing font handling in GenOffice.
 */

export interface FontSubstitutionRule {
  /** Original font name as declared in the document, e.g. '新細明體' */
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
 * face installed as "DFKai-SB" (Office, a hand-installed kaiu.ttf) does NOT answer to "標楷體", and
 * the localized Windows name of an English-named font often misses too. Both directions are listed so
 * the availability check, the substitution chain, the missing-font report and the settings pane agree.
 */
const FONT_NAME_EQUIVALENTS: Record<string, readonly string[]> = {
  標楷體: ['DFKai-SB', 'BiauKai'],
  'DFKai-SB': ['標楷體', 'BiauKai'],
  BiauKai: ['標楷體', 'DFKai-SB'],
  新細明體: ['PMingLiU'],
  PMingLiU: ['新細明體'],
  細明體: ['MingLiU'],
  MingLiU: ['細明體'],
  微軟正黑體: ['Microsoft JhengHei', 'PingFang TC'],
  'Microsoft JhengHei': ['微軟正黑體', 'PingFang TC'],
  '微軟正黑體 Light': ['Microsoft JhengHei Light'],
  'Microsoft JhengHei Light': ['微軟正黑體 Light'],
  微软雅黑: ['Microsoft YaHei', 'PingFang SC'],
  'Microsoft YaHei': ['微软雅黑', 'PingFang SC'],
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
}

const FONT_NAME_EQUIVALENTS_LC = new Map(
  Object.entries(FONT_NAME_EQUIVALENTS).map(([k, v]) => [k.toLowerCase(), v] as const),
)

/** Other names the same face may be installed under (declared name excluded) */
export function fontNameEquivalents(font: string): readonly string[] {
  return FONT_NAME_EQUIVALENTS_LC.get(font.trim().toLowerCase()) ?? []
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
