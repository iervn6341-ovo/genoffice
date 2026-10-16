import { describe, expect, it, vi } from 'vitest'

import {
  caseInsensitiveSheetNameMap,
  installCaseInsensitiveFormulaSheetLookup,
} from '../src/renderer/formula-sheet-lookup'

describe('Excel-compatible formula sheet lookup', () => {
  it('resolves mixed-case names without changing exact keys or the source map', () => {
    const names = { Inputs: 'sheet-inputs', "O'Brien": 'sheet-quoted' }
    const map = caseInsensitiveSheetNameMap({ workbook: names })
    expect(map.workbook?.Inputs).toBe('sheet-inputs')
    expect(map.workbook?.iNpUtS).toBe('sheet-inputs')
    expect(map.workbook?.["o'bRiEn"]).toBe('sheet-quoted')
    expect(Object.keys(map.workbook ?? {})).toEqual(['Inputs', "O'Brien"])
    expect(names).toEqual({ Inputs: 'sheet-inputs', "O'Brien": 'sheet-quoted' })
  })

  it('patches engine and core workbook lookups once while preserving exact-match priority', () => {
    const names = { Inputs: 'sheet-inputs', Data: 'sheet-data' }
    const config = { getSheetNameMap: () => ({ unit: names }) }
    const runtime = { univer: { __getInjector: () => ({ get: () => config }) } }
    const sheets = [{ getName: () => 'Inputs' }, { getName: () => 'Data' }]
    const exact = vi.fn((name: string) => sheets.find((sheet) => sheet.getName() === name))
    const workbook = { getSheetBySheetName: exact, getSheets: () => sheets }

    installCaseInsensitiveFormulaSheetLookup(runtime as never, workbook as never)
    installCaseInsensitiveFormulaSheetLookup(runtime as never, workbook as never)

    expect((config.getSheetNameMap() as Record<string, Record<string, string>>).unit?.iNpUtS).toBe(
      'sheet-inputs',
    )
    expect(workbook.getSheetBySheetName('Inputs')).toBe(sheets[0])
    expect(workbook.getSheetBySheetName('iNpUtS')).toBe(sheets[0])
    expect(workbook.getSheetBySheetName('unknown')).toBeUndefined()
    expect(exact).toHaveBeenCalledTimes(3)
  })
})
