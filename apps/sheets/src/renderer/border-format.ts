import type { WorkbookStyleEdit } from '../shared/desktop-api'
import type {
  CellProtectionFlags,
  ProtectionRange,
  SelectionBorderControl,
} from './selection-protection'
import { visibleBorderSeam } from './merged-border-geometry'

export const BORDER_CONTROLS = [
  'left',
  'top',
  'bottom',
  'right',
  'insideHorizontal',
  'insideVertical',
  'diagonalUp',
  'diagonalDown',
] as const satisfies readonly SelectionBorderControl[]

export interface BorderAction {
  readonly modes: Readonly<Partial<Record<SelectionBorderControl, boolean>>>
  readonly style: NonNullable<NonNullable<WorkbookStyleEdit['borderTop']>>['style']
  readonly color: string
}

export function borderActionCommand(action: BorderAction): string {
  const modes = BORDER_CONTROLS.map((control) =>
    action.modes[control] === undefined ? '-' : action.modes[control] ? '1' : '0',
  ).join('')
  return `border-edges:${modes}:${action.style}:${action.color}`
}

export function borderActionOfCommand(command: string): BorderAction | null {
  const match =
    /^border-edges:([01-]{8}):(thin|medium|thick|dashed|dotted|double|hair|dashDot|dashDotDot|mediumDashed|mediumDashDot|mediumDashDotDot|slantDashDot):(#[0-9a-fA-F]{6})$/.exec(
      command,
    )
  if (!match?.[1] || !match[2] || !match[3] || !/[01]/.test(match[1])) return null
  const modes: Partial<Record<SelectionBorderControl, boolean>> = {}
  for (const [index, control] of BORDER_CONTROLS.entries()) {
    const mode = match[1][index]
    if (mode !== '-') modes[control] = mode === '1'
  }
  return {
    modes,
    style: match[2] as BorderAction['style'],
    color: match[3].toUpperCase(),
  }
}

export function borderDeltaForCell(
  action: BorderAction,
  range: ProtectionRange,
  cell: CellProtectionFlags,
  merges: readonly ProtectionRange[] = [],
): WorkbookStyleEdit | null {
  const delta: WorkbookStyleEdit = {}
  const edge = { style: action.style, color: action.color }
  const set = (key: 'borderTop' | 'borderBottom' | 'borderLeft' | 'borderRight', on: boolean) => {
    delta[key] = on ? edge : null
  }
  const { modes } = action
  if (cell.column === range.startColumn && modes.left !== undefined) set('borderLeft', modes.left)
  if (cell.column === range.endColumn && modes.right !== undefined) set('borderRight', modes.right)
  if (cell.row === range.startRow && modes.top !== undefined) set('borderTop', modes.top)
  if (cell.row === range.endRow && modes.bottom !== undefined) set('borderBottom', modes.bottom)
  if (
    cell.row > range.startRow &&
    modes.insideHorizontal !== undefined &&
    visibleBorderSeam(merges, cell.row - 1, cell.column, 'down')
  )
    set('borderTop', modes.insideHorizontal)
  if (
    cell.row < range.endRow &&
    modes.insideHorizontal !== undefined &&
    visibleBorderSeam(merges, cell.row, cell.column, 'down')
  )
    set('borderBottom', modes.insideHorizontal)
  if (
    cell.column > range.startColumn &&
    modes.insideVertical !== undefined &&
    visibleBorderSeam(merges, cell.row, cell.column - 1, 'right')
  )
    set('borderLeft', modes.insideVertical)
  if (
    cell.column < range.endColumn &&
    modes.insideVertical !== undefined &&
    visibleBorderSeam(merges, cell.row, cell.column, 'right')
  )
    set('borderRight', modes.insideVertical)
  if (modes.diagonalUp !== undefined || modes.diagonalDown !== undefined) {
    const up = modes.diagonalUp ?? cell.diagonalUp === true
    const down = modes.diagonalDown ?? cell.diagonalDown === true
    if (modes.diagonalUp !== undefined) delta.diagonalUp = modes.diagonalUp
    if (modes.diagonalDown !== undefined) delta.diagonalDown = modes.diagonalDown
    if (up || down) {
      // An imported diagonal can use theme/auto color that the editor cannot express.
      // Toggling its direction must keep the original diagonal stroke intact.
      if (
        (modes.diagonalUp === true || modes.diagonalDown === true) &&
        cell.diagonalUp !== true &&
        cell.diagonalDown !== true
      )
        delta.borderDiagonal = edge
    } else delta.borderDiagonal = null
  }
  return Object.keys(delta).length === 0 ? null : delta
}
