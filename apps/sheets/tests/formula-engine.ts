/**
 * Headless Univer formula engine for tests: no Electron, no renderer. Write
 * cells (strings starting with "=" are formulas), wait for the calculation to
 * settle, read the raw results back. Pass `install` to register the app's own
 * function patches on the same runtime before any formula runs.
 */
import {
  ICommandService,
  IUniverInstanceService,
  LocaleType,
  LogLevel,
  Univer,
  UniverInstanceType,
} from '@univerjs/core'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'

const NOTIFICATION = 'formula.mutation.set-formula-calculation-notification'
const RESULT = 'formula.mutation.set-formula-calculation-result'
const SUCCESS = 3
const QUIET_MS = 120
const TIMEOUT_MS = 6000

export type CellInput = string | number | boolean | null

export interface EngineCell {
  v?: unknown
  t?: number
  f?: string
}

export type CellResult = number | string | boolean | null

export interface FormulaEngine {
  univer: Univer
  /** write cells ('A1' → value) and wait for the formulas they trigger */
  set(cells: Record<string, CellInput>): Promise<void>
  get(address: string): EngineCell | null
  /** the computed value as a plain JS value (null for a blank result) */
  value(address: string): CellResult
  dispose(): void
}

function parseAddress(address: string): { row: number; column: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(address)
  if (!match) throw new Error(`bad address ${address}`)
  let column = 0
  for (const letter of match[1]!) column = column * 26 + (letter.charCodeAt(0) - 64)
  return { row: Number(match[2]) - 1, column: column - 1 }
}

export async function createFormulaEngine(
  install?: (univer: Univer) => void,
  size = { rows: 400, columns: 30 },
): Promise<FormulaEngine> {
  const univer = new Univer({ logLevel: LogLevel.ERROR, locale: LocaleType.EN_US, locales: {} })
  univer.registerPlugin(UniverFormulaEnginePlugin)
  univer.registerPlugin(UniverSheetsPlugin)
  univer.registerPlugin(UniverSheetsFormulaPlugin)
  install?.(univer)
  univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: 'wb1',
    sheetOrder: ['s1'],
    name: 'wb',
    styles: {},
    sheets: {
      s1: { id: 's1', name: 'Sheet1', rowCount: size.rows, columnCount: size.columns, cellData: {} },
    },
  })
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const instanceService = injector.get(IUniverInstanceService)
  const sheet = () =>
    (
      instanceService.getUnit('wb1') as unknown as {
        getSheetBySheetId(id: string): { getCellRaw(row: number, column: number): EngineCell | null }
      }
    ).getSheetBySheetId('s1')

  let passes = 0
  let lastResultAt = 0
  commandService.onCommandExecuted((command) => {
    if (command.id === RESULT) lastResultAt = Date.now()
    if (
      command.id === NOTIFICATION &&
      (command.params as { functionsExecutedState?: number } | undefined)?.functionsExecutedState ===
        SUCCESS
    ) {
      passes += 1
      lastResultAt = Date.now()
    }
  })

  return {
    univer,
    async set(cells) {
      const value: Record<number, Record<number, unknown>> = {}
      let startRow = Infinity
      let endRow = -1
      let startColumn = Infinity
      let endColumn = -1
      let hasFormula = false
      for (const [address, input] of Object.entries(cells)) {
        const { row, column } = parseAddress(address)
        startRow = Math.min(startRow, row)
        endRow = Math.max(endRow, row)
        startColumn = Math.min(startColumn, column)
        endColumn = Math.max(endColumn, column)
        let cell: unknown
        if (typeof input === 'string' && input.startsWith('=')) {
          cell = { f: input }
          hasFormula = true
        } else if (typeof input === 'string') cell = { v: input, t: 1 }
        else if (typeof input === 'number') cell = { v: input, t: 2 }
        else if (typeof input === 'boolean') cell = { v: input ? 1 : 0, t: 3 }
        else cell = { v: null }
        ;(value[row] ??= {})[column] = cell
      }
      const before = passes
      await commandService.executeCommand('sheet.command.set-range-values', {
        unitId: 'wb1',
        subUnitId: 's1',
        range: { startRow, endRow, startColumn, endColumn },
        value,
      })
      if (!hasFormula) return
      const deadline = Date.now() + TIMEOUT_MS
      while (Date.now() < deadline) {
        await new Promise((resolve) => setTimeout(resolve, 25))
        if (passes > before && Date.now() - lastResultAt >= QUIET_MS) return
      }
    },
    get(address) {
      const { row, column } = parseAddress(address)
      return sheet().getCellRaw(row, column) ?? null
    },
    value(address) {
      const cell = this.get(address)
      if (!cell || cell.v === undefined || cell.v === null) return null
      if (cell.t === 3) return cell.v === 1 || cell.v === true
      return cell.v as number | string
    },
    dispose() {
      univer.dispose()
    },
  }
}
