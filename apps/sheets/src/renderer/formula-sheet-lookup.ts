/**
 * Excel sheet names compare without case, but Univer 0.25.1 resolves both
 * formula-reference sheetNameMap keys and Workbook.getSheetBySheetName with
 * exact string equality. Keep the XLSX and model formula source untouched;
 * only those two lookup boundaries need the Excel-compatible fallback.
 */
import { IFormulaCurrentConfigService } from '@univerjs/engine-formula'
import type { Workbook } from '@univerjs/core'

import type { UniverRuntime } from './univer-state'

const patchedRuntimes = new WeakSet<UniverRuntime>()
const patchedWorkbooks = new WeakSet<Workbook>()

function caseInsensitiveNames<T>(names: Record<string, T>): Record<string, T> {
  return new Proxy(names, {
    get(target, key, receiver) {
      if (typeof key !== 'string') return Reflect.get(target, key, receiver)
      if (Object.hasOwn(target, key)) return target[key]
      const match = Object.keys(target).find((name) => name.toLowerCase() === key.toLowerCase())
      return match === undefined ? undefined : target[match]
    },
  })
}

export function caseInsensitiveSheetNameMap<T>(
  units: Record<string, Record<string, T> | null | void>,
): Record<string, Record<string, T> | null | void> {
  return Object.fromEntries(
    Object.entries(units).map(([unitId, names]) => [
      unitId,
      names == null ? names : caseInsensitiveNames(names),
    ]),
  )
}

export function installCaseInsensitiveFormulaSheetLookup(
  runtime: UniverRuntime,
  workbook: Workbook,
): void {
  if (!patchedRuntimes.has(runtime)) {
    const service = runtime.univer.__getInjector().get(IFormulaCurrentConfigService)
    const original = service.getSheetNameMap.bind(service)
    service.getSheetNameMap = () => caseInsensitiveSheetNameMap(original())
    patchedRuntimes.add(runtime)
  }
  if (patchedWorkbooks.has(workbook)) return
  const original = workbook.getSheetBySheetName.bind(workbook)
  workbook.getSheetBySheetName = (name) =>
    original(name) ??
    workbook.getSheets().find((sheet) => sheet.getName().toLowerCase() === name.toLowerCase())
  patchedWorkbooks.add(workbook)
}
