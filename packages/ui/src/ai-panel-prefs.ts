export const AI_FONT_SIZES = ['default', 'large', 'xlarge', 'custom'] as const
export type AiPanelSide = 'left' | 'right'
/** Ribbon layout: `classic` keeps groups collapsed to dropdown buttons; `auto` expands them while there is room */
export type RibbonFold = 'classic' | 'auto'
/** Document area appearance: follows the app theme, or is pinned light / dark (View ▸ Dark Mode) */
export type ContentTheme = 'auto' | 'light' | 'dark'
export type AiFontSize = (typeof AI_FONT_SIZES)[number]

/** Body text size of `.ai-chat` in every app's stylesheet; the presets scale from it */
export const AI_FONT_BASE_PX = 14
export const AI_CUSTOM_FONT_MIN_PX = 10
export const AI_CUSTOM_FONT_MAX_PX = 32

const PRESET_ZOOM: Record<Exclude<AiFontSize, 'custom'>, number> = {
  default: 1,
  large: 1.15,
  xlarge: 1.3,
}

/** AI panel display preferences, persisted by the shell in app-settings.json */
export interface AiPanelPrefs {
  readonly side: AiPanelSide
  readonly fontSize: AiFontSize
  /** Body text size in px, used only when `fontSize` is `'custom'` */
  readonly customFontSize: number
  readonly spellcheck: boolean
  /** Ribbon overflow style; rides on this record because it reaches every renderer the same way */
  readonly ribbonFold: RibbonFold
  /** Eases notched mouse-wheel scrolling (see smooth-scroll.ts); rides on this record like `ribbonFold` */
  readonly smoothScroll: boolean
  /** GPU-accelerated rendering; read by the shell at startup (a change takes effect after a restart) */
  readonly gpuAcceleration: boolean
  /** Light / dark preview of the document content; shared by Docs, Sheets and Slides and kept across launches */
  readonly contentTheme: ContentTheme
}

export const DEFAULT_AI_PANEL_PREFS: AiPanelPrefs = {
  side: 'left',
  fontSize: 'default',
  customFontSize: AI_FONT_BASE_PX,
  spellcheck: true,
  ribbonFold: 'auto',
  smoothScroll: false,
  gpuAcceleration: true,
  contentTheme: 'auto',
}

export function isAiFontSize(value: unknown): value is AiFontSize {
  return typeof value === 'string' && (AI_FONT_SIZES as readonly string[]).includes(value)
}

export function clampAiCustomFontSize(value: unknown): number | null {
  const n = typeof value === 'string' ? Number(value) : value
  if (typeof n !== 'number' || !Number.isFinite(n)) return null
  return Math.min(AI_CUSTOM_FONT_MAX_PX, Math.max(AI_CUSTOM_FONT_MIN_PX, Math.round(n)))
}

/** Effective body text size in px for the given preferences */
export function aiPanelFontPx(prefs: AiPanelPrefs): number {
  return prefs.fontSize === 'custom'
    ? prefs.customFontSize
    : Math.round(AI_FONT_BASE_PX * PRESET_ZOOM[prefs.fontSize])
}

/** Block zoom applied to the chat log and composer (see ai-panel-prefs.css) */
export function aiPanelZoom(prefs: AiPanelPrefs): number {
  return prefs.fontSize === 'custom'
    ? prefs.customFontSize / AI_FONT_BASE_PX
    : PRESET_ZOOM[prefs.fontSize]
}

export function sameAiPanelPrefs(a: AiPanelPrefs, b: AiPanelPrefs): boolean {
  return (
    a.side === b.side &&
    a.fontSize === b.fontSize &&
    a.customFontSize === b.customFontSize &&
    a.spellcheck === b.spellcheck &&
    a.ribbonFold === b.ribbonFold &&
    a.smoothScroll === b.smoothScroll &&
    a.gpuAcceleration === b.gpuAcceleration &&
    a.contentTheme === b.contentTheme
  )
}

/** Fills in defaults for missing or malformed fields (settings file, IPC payloads) */
export function normalizeAiPanelPrefs(raw: unknown): AiPanelPrefs {
  const obj = raw !== null && typeof raw === 'object' ? (raw as Record<string, unknown>) : {}
  return {
    side: obj.side === 'right' ? 'right' : 'left',
    fontSize: isAiFontSize(obj.fontSize) ? obj.fontSize : DEFAULT_AI_PANEL_PREFS.fontSize,
    customFontSize:
      clampAiCustomFontSize(obj.customFontSize) ?? DEFAULT_AI_PANEL_PREFS.customFontSize,
    spellcheck:
      typeof obj.spellcheck === 'boolean' ? obj.spellcheck : DEFAULT_AI_PANEL_PREFS.spellcheck,
    ribbonFold: obj.ribbonFold === 'classic' ? 'classic' : 'auto',
    smoothScroll:
      typeof obj.smoothScroll === 'boolean'
        ? obj.smoothScroll
        : DEFAULT_AI_PANEL_PREFS.smoothScroll,
    gpuAcceleration:
      typeof obj.gpuAcceleration === 'boolean'
        ? obj.gpuAcceleration
        : DEFAULT_AI_PANEL_PREFS.gpuAcceleration,
    contentTheme:
      obj.contentTheme === 'light' || obj.contentTheme === 'dark' ? obj.contentTheme : 'auto',
  }
}
