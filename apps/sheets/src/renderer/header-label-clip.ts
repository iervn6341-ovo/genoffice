/**
 * Row / column header labels must stay inside their own header cell. Stock
 * Univer only clips a label when the header has a per-row config, so a row
 * dragged (or set) shorter than the label's font height lets the number spill
 * over its neighbours and the labels of adjacent short rows pile on top of
 * each other. Excel clips the label to the header cell; marking every header
 * as "specially styled" makes Univer take its own clip-to-cell-bounds path.
 */
import { SheetColumnHeaderExtensionRegistry, SheetRowHeaderExtensionRegistry } from '@univerjs/engine-render'

type CfgGetter = (...args: unknown[]) => [unknown, boolean]

interface HeaderExtension {
  getCfgOfCurrentRow?: CfgGetter
  getCfgOfCurrentColumn?: CfgGetter
  __labelClip?: boolean
}

export function forceHeaderLabelClip(
  extension: HeaderExtension,
  method: 'getCfgOfCurrentRow' | 'getCfgOfCurrentColumn',
): boolean {
  const original = extension[method]
  if (typeof original !== 'function' || extension.__labelClip) return false
  extension[method] = function (this: HeaderExtension, ...args: unknown[]) {
    const [cfg] = original.apply(this, args)
    return [cfg, true]
  }
  extension.__labelClip = true
  return true
}

let installed = false

/** Idempotent; the registries hand the same extension objects to every sheet. */
export function installHeaderLabelClip(): void {
  if (installed) return
  installed = true
  for (const extension of SheetRowHeaderExtensionRegistry.getData() as HeaderExtension[]) {
    forceHeaderLabelClip(extension, 'getCfgOfCurrentRow')
  }
  for (const extension of SheetColumnHeaderExtensionRegistry.getData() as HeaderExtension[]) {
    forceHeaderLabelClip(extension, 'getCfgOfCurrentColumn')
  }
}
