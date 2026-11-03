import type { Lang } from '@genoffice/i18n'
import { BUILTIN_FONT_FAMILIES, fontFamiliesFor, sortFontFamilies } from '@genoffice/ui'

export { BUILTIN_FONT_FAMILIES, fontFamiliesFor, sortFontFamilies }

const EAST_ASIAN_FONT_RE =
  /[⺀-鿿豈-﫿぀-ヿㇰ-ㇿ가-힯]|sim(sun|hei)|nsimsun|kaiti|fangsong|dengxian|yahei|songti|heiti|xingkai|lisu|youyuan|st(zhongsong|song|kai|fangsong|xihei|hupo|liti|caiyun)|pingfang|hiragino|meiryo|osaka|kozuka|yu (gothic|mincho)|yugoth|ms (ui )?p?(gothic|mincho)|biz ud|malgun|batang|gulim|dotum|gungsuh|m(ye|yu)ngjo|nanum|apple (sd )?gothic|applemyungjo|jhenghei|p?mingliu|biaukai|dfkai|kaiu|source han|noto (sans|serif) (cjk|sc|tc|hk|jp|kr)|wenquanyi/i

/**
 * Which rFonts slot a font-box pick should target: East Asian names go to
 * w:eastAsia, everything else to w:ascii/w:hAnsi — mirroring Word, where
 * picking a Latin font never clobbers the Chinese font and vice versa.
 */
export function isEastAsianFontName(name: string): boolean {
  return EAST_ASIAN_FONT_RE.test(name.normalize('NFKC'))
}

/**
 * docDefaults w:eastAsia font for new blank documents, matching what Word
 * ships per market (zh → SimSun, ja → Yu Mincho, ko → Malgun Gothic,
 * zh-TW → PMingLiU). English and every other language return undefined —
 * like en-US Word, whose theme leaves the East Asian slot empty and lets
 * per-script substitution kick in only when CJK text actually appears.
 * Latin default stays Calibri for every language.
 */
export function defaultEastAsiaFontFor(lang: Lang): string | undefined {
  switch (lang) {
    case 'zh':
      return '宋体'
    case 'ja':
      return 'Yu Mincho'
    case 'ko':
      return 'Malgun Gothic'
    case 'zh-TW':
      return 'PMingLiU'
    default:
      return undefined
  }
}

/**
 * settings.xml w:defaultTabStop (twips) for new blank documents: Chinese Word
 * templates use a two-character grid (probed zh-TW: 480 = 2 × 12pt; zh: 420 =
 * 2 × 10.5pt). Other languages keep Word's 0.5" default (undefined).
 */
export function defaultTabStopFor(lang: Lang): number | undefined {
  switch (lang) {
    case 'zh':
      return 420
    case 'zh-TW':
      return 480
    default:
      return undefined
  }
}
