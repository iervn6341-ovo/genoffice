import type { SelectionFormat } from './selection-format'
import type { SelectionProtection } from './selection-protection'
import { BORDER_CONTROLS, borderActionCommand, type BorderAction } from './border-format'
import type { SelectionBorderControl } from './selection-protection'

export type TriState = '' | 'on' | 'off'

/// Everything the Format Cells dialog edits, prefilled from the selection so
/// it opens showing the current format.
export interface FormatCellsDraft {
  readonly pattern: string
  readonly family: string
  readonly size: string
  readonly bold: TriState
  readonly italic: TriState
  readonly underline: TriState
  /// Empty means a mixed selection; an unchanged draft never writes an XF.
  readonly underlineMode:
    '' | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting'
  readonly fontVerticalAlign: '' | 'baseline' | 'superscript' | 'subscript'
  readonly strike: TriState
  readonly fontColor: string
  readonly border: string
  readonly borderColor: string
  /// OOXML ST_BorderStyle name applied with the preset.
  readonly borderStyle: string
  readonly borderEdges: Readonly<Record<SelectionBorderControl, TriState>>
  readonly fill: string
  /// Fill → Pattern Style: '' for none, else an OOXML pattern type (see fill-pattern.ts).
  readonly fillPattern: string
  /// Fill → Pattern Color, #rrggbb.
  readonly patternColor: string
  readonly noFill: boolean
  readonly hAlign: string
  readonly vAlign: string
  readonly wrapText: TriState
  readonly shrinkToFit: TriState
  readonly mergeCells: TriState
  /// Ribbon rotate preset keys ('' = unchanged, '0' = clear).
  readonly rotation: string
  /// '' = unchanged, otherwise the indent step count as text.
  readonly indent: string
  readonly locked: TriState
  readonly hidden: TriState
}

export function draftFromSelection(
  format: SelectionFormat | null,
  protection?: SelectionProtection | null,
): FormatCellsDraft {
  const toTriState = (flag: boolean | undefined): TriState =>
    flag === undefined ? '' : flag ? 'on' : 'off'
  return {
    pattern: format?.numberFormat ?? '',
    family: format?.fontFamily ?? '',
    size: format?.fontSize != null ? String(format.fontSize) : '',
    bold: toTriState(format?.bold),
    italic: toTriState(format?.italic),
    underline: toTriState(format?.underline),
    underlineMode: protection?.font
      ? (protection.font.underlineStyle ?? '')
      : format
        ? format.underline
          ? 'single'
          : 'none'
        : '',
    fontVerticalAlign: protection?.font?.verticalAlign ?? '',
    strike: toTriState(format?.strike),
    // Native color inputs report lowercase hex; store the same casing so
    // "user picked the value it already had" compares equal.
    fontColor: format?.fontColor?.toLowerCase() ?? '',
    border: '',
    borderColor: protection?.border?.lineColor?.toLowerCase() ?? '#000000',
    borderStyle: protection?.border?.lineStyle ?? 'thin',
    borderEdges: Object.fromEntries(
      BORDER_CONTROLS.map((control) => [
        control,
        protection?.border?.[control] === null || protection?.border?.[control] === undefined
          ? ''
          : protection.border[control]
            ? 'on'
            : 'off',
      ]),
    ) as Record<SelectionBorderControl, TriState>,
    // a pattern fill shows its own background color, not the blended flat color
    fill: (format?.fillPattern?.bg ?? format?.fillColor)?.toLowerCase() ?? '',
    fillPattern: format?.fillPattern?.pattern ?? '',
    patternColor: format?.fillPattern?.fg.toLowerCase() ?? '#000000',
    noFill: false,
    hAlign: protection
      ? (protection.alignment.horizontal ?? '')
      : (format?.horizontalAlignment ?? ''),
    vAlign:
      (protection ? protection.alignment.vertical : format?.verticalAlignment) === 'center'
        ? 'middle'
        : protection
          ? (protection.alignment.vertical ?? '')
          : (format?.verticalAlignment ?? ''),
    wrapText: toTriState(format?.wrap),
    shrinkToFit: toTriState(protection?.alignment.shrinkToFit ?? undefined),
    mergeCells: toTriState(protection?.alignment.merged ?? undefined),
    rotation: rotationPreset(protection ? protection.alignment.rotation : format?.textRotation),
    // 0 echoes as '' so an untouched dialog never emits an indent command.
    indent: format?.indent ? String(format.indent) : '',
    locked: toTriState(protection?.locked ?? undefined),
    hidden: toTriState(protection?.hidden ?? undefined),
  }
}

/** Univer renders rich and rotated cells through a document skeleton, not the plain-cell canvas hook. */
export function accountingUnderlineUnavailable(
  selection: SelectionProtection | null,
  draftRotation: string,
): boolean {
  if (!selection) return true
  return selection.cells.some((cell) => {
    const rotation =
      draftRotation === ''
        ? cell.textRotation
        : draftRotation === 'vertical'
          ? 255
          : Number(draftRotation)
    return cell.richDocument === true || rotation !== 0
  })
}

/** Preserve an imported unsupported mode, but reject a newly created unsupported combination. */
export function accountingUnderlineChangeBlocked(
  initial: FormatCellsDraft,
  edited: FormatCellsDraft,
  selection: SelectionProtection | null,
): boolean {
  const accounting =
    edited.underlineMode === 'singleAccounting' || edited.underlineMode === 'doubleAccounting'
  const changed =
    edited.underlineMode !== initial.underlineMode || edited.rotation !== initial.rotation
  return accounting && changed && accountingUnderlineUnavailable(selection, edited.rotation)
}

/// OOXML rotation value → the ribbon preset key it corresponds to; values
/// between presets echo as unchanged.
function rotationPreset(value: number | null | undefined): string {
  if (value == null) return ''
  return value === 255 ? 'vertical' : String(value)
}

export function ooxmlRotationDegrees(value: string): string {
  if (!value || value === 'vertical') return ''
  const angle = Number(value)
  return Number.isInteger(angle) && angle >= 0 && angle <= 180
    ? String(angle <= 90 ? angle : 90 - angle)
    : ''
}

export function degreesToOoxmlRotation(value: string): string {
  if (value.trim() === '') return ''
  const degrees = Number(value)
  if (!Number.isInteger(degrees) || degrees < -90 || degrees > 90) return ''
  return String(degrees >= 0 ? degrees : 90 - degrees)
}

/// Ribbon command strings for exactly the settings the user changed relative
/// to the prefilled draft — untouched settings never emit.
export function formatCellsCommands(initial: FormatCellsDraft, edited: FormatCellsDraft): string[] {
  const commands: string[] = []
  if (edited.pattern && edited.pattern !== initial.pattern) {
    commands.push(`format:${edited.pattern}`)
  }
  if (edited.family && edited.family !== initial.family) {
    commands.push(`font-family:${edited.family}`)
  }
  if (edited.size && edited.size !== initial.size) commands.push(`font-size:${edited.size}`)
  for (const key of ['bold', 'italic', 'strike'] as const) {
    if (edited[key] && edited[key] !== initial[key]) commands.push(`${key}:${edited[key]}`)
  }
  if (edited.fontColor && edited.fontColor !== initial.fontColor) {
    commands.push(`font-color:${edited.fontColor}`)
  }
  const underlineMode =
    edited.underlineMode && edited.underlineMode !== initial.underlineMode
      ? edited.underlineMode
      : ''
  const fontVerticalAlign =
    edited.fontVerticalAlign && edited.fontVerticalAlign !== initial.fontVerticalAlign
      ? edited.fontVerticalAlign
      : ''
  if (underlineMode || fontVerticalAlign) {
    commands.push(`font-effect:${underlineMode}:${fontVerticalAlign}`)
  }
  const modes: Partial<Record<SelectionBorderControl, boolean>> = {}
  const selectedEdges = edited.border
    ? borderPresetEdges(edited.borderEdges, edited.border)
    : edited.borderEdges
  for (const control of BORDER_CONTROLS) {
    const value = selectedEdges[control]
    if (value && (edited.border || value !== initial.borderEdges[control]))
      modes[control] = value === 'on'
  }
  if (Object.keys(modes).length > 0) {
    commands.push(
      borderActionCommand({
        modes,
        style: edited.borderStyle as BorderAction['style'],
        color: edited.borderColor,
      }),
    )
  }
  if (edited.noFill) {
    if (initial.fill) commands.push('fill:none')
  } else if (edited.fillPattern) {
    if (
      edited.fillPattern !== initial.fillPattern ||
      edited.patternColor !== initial.patternColor ||
      edited.fill !== initial.fill
    ) {
      commands.push(
        `fill-pattern:${edited.fillPattern}:${edited.patternColor}:${edited.fill || '#ffffff'}`,
      )
    }
  } else if (edited.fill && (edited.fill !== initial.fill || initial.fillPattern)) {
    // choosing "no pattern" on a patterned cell rewrites it as the plain background color
    commands.push(`fill:${edited.fill}`)
  }
  if (edited.hAlign && edited.hAlign !== initial.hAlign) {
    commands.push(`align:${edited.hAlign}`)
  }
  if (edited.vAlign && edited.vAlign !== initial.vAlign) {
    commands.push(`valign:${edited.vAlign}`)
  }
  if (edited.wrapText && edited.wrapText !== initial.wrapText) {
    commands.push(`wrap:${edited.wrapText}`)
  }
  if (edited.shrinkToFit && edited.shrinkToFit !== initial.shrinkToFit) {
    commands.push(`shrink:${edited.shrinkToFit}`)
  }
  if (edited.mergeCells && edited.mergeCells !== initial.mergeCells) {
    commands.push(edited.mergeCells === 'on' ? 'merge:merge' : 'merge:unmerge')
  }
  if (edited.rotation && edited.rotation !== initial.rotation) {
    commands.push(`rotate-ooxml:${edited.rotation === 'vertical' ? 255 : edited.rotation}`)
  }
  if (edited.indent && edited.indent !== initial.indent) {
    commands.push(`indent:${edited.indent}`)
  }
  if (
    (edited.locked && edited.locked !== initial.locked) ||
    (edited.hidden && edited.hidden !== initial.hidden)
  ) {
    const locked =
      edited.locked && edited.locked !== initial.locked ? `locked-${edited.locked}` : ''
    const hidden =
      edited.hidden && edited.hidden !== initial.hidden ? `hidden-${edited.hidden}` : ''
    commands.push(`cellprot:${locked}:${hidden}`)
  }
  return commands
}

export function borderPresetEdges(
  current: FormatCellsDraft['borderEdges'],
  preset: string,
): FormatCellsDraft['borderEdges'] {
  const next = { ...current }
  const change = (controls: readonly SelectionBorderControl[], mode: TriState) => {
    for (const control of controls) next[control] = mode
  }
  if (preset === 'none') change(BORDER_CONTROLS, 'off')
  else if (preset === 'all') change(BORDER_CONTROLS.slice(0, 6), 'on')
  else if (preset === 'inside') change(['insideHorizontal', 'insideVertical'], 'on')
  else if (preset === 'outer' || preset === 'thick-outer')
    change(['left', 'top', 'bottom', 'right'], 'on')
  else if (['left', 'top', 'bottom', 'right'].includes(preset))
    change([preset as SelectionBorderControl], 'on')
  return next
}
