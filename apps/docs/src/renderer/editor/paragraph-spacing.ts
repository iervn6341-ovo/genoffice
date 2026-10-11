import type { StyleDisplay, StyleUpsert } from '@genoffice/docx-engine'

/**
 * Design → Paragraph Spacing: Word's presets rewrite the spacing of the document's base style
 * (Normal), so every paragraph that inherits it changes at once. `after` / `before` are twips,
 * `line` a multiple of single spacing.
 */
export interface ParagraphSpacingPreset {
  id: 'default' | 'none' | 'compact' | 'tight' | 'open' | 'relaxed' | 'double'
  before: number
  after: number
  line: number
}

export const PARAGRAPH_SPACING_PRESETS: readonly ParagraphSpacingPreset[] = [
  { id: 'default', before: 0, after: 160, line: 1.08 },
  { id: 'none', before: 0, after: 0, line: 1 },
  { id: 'compact', before: 0, after: 80, line: 1 },
  { id: 'tight', before: 0, after: 120, line: 1 },
  { id: 'open', before: 0, after: 200, line: 1.15 },
  { id: 'relaxed', before: 0, after: 120, line: 1.5 },
  { id: 'double', before: 0, after: 160, line: 2 },
]

export const spacingPreset = (id: string | null | undefined): ParagraphSpacingPreset | undefined =>
  PARAGRAPH_SPACING_PRESETS.find((p) => p.id === id)

/** the style-definition patch that saves a preset into styles.xml */
export function spacingUpsert(preset: ParagraphSpacingPreset): StyleUpsert {
  return {
    styleId: 'Normal',
    pPr: {
      spaceBeforeTwips: preset.before,
      spaceAfterTwips: preset.after,
      lineSpacing: preset.line,
    },
  }
}

export type SpacingFields = Pick<
  StyleDisplay,
  | 'spaceBeforeTwips'
  | 'spaceAfterTwips'
  | 'spaceBeforeAuto'
  | 'spaceAfterAuto'
  | 'lineSpacing'
  | 'lineRule'
  | 'lineRawTwips'
>

const KEYS: Array<keyof SpacingFields> = [
  'spaceBeforeTwips',
  'spaceAfterTwips',
  'spaceBeforeAuto',
  'spaceAfterAuto',
  'lineSpacing',
  'lineRule',
  'lineRawTwips',
]

/** the spacing fields of a style display, to put back when the preset is undone */
export function captureSpacing(display: StyleDisplay | undefined): SpacingFields {
  const out: SpacingFields = {}
  for (const k of KEYS)
    if (display?.[k] !== undefined) (out as Record<string, unknown>)[k] = display[k]
  return out
}

/** write a preset (or the captured original) into a style display in place */
export function writeSpacing(
  display: StyleDisplay,
  fields: SpacingFields | ParagraphSpacingPreset,
): void {
  const d = display as Record<string, unknown>
  for (const k of KEYS) delete d[k]
  if ('id' in fields) {
    display.spaceBeforeTwips = fields.before
    display.spaceAfterTwips = fields.after
    display.lineSpacing = fields.line
    display.lineRule = 'auto'
  } else {
    Object.assign(display, fields)
  }
}
