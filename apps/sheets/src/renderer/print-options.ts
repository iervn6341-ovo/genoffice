/**
 * Print-only options of the print sheet: Excel's Page Setup / Sheet-tab choices that the workbook
 * journal cannot save to the XLSX yet (custom margins, centering on the page, "cell errors as",
 * black-and-white, draft quality, first page number). They shape this print job only: the layout
 * (print-html.ts) and the preview read the same object, so preview and printout agree, and
 * nothing here is written to the file.
 */
import type { PrintMargins } from './print-settings'

export type PrintErrorsAs = 'displayed' | 'blank' | 'dash' | 'na'

export interface PrintOnlyOptions {
  /// Page margins in inches (replaces the sheet's margins for this job).
  readonly margins?: PrintMargins
  readonly centerHorizontally?: boolean
  readonly centerVertically?: boolean
  /// How error values (#DIV/0!, #N/A, …) print; absent = as displayed.
  readonly errorsAs?: PrintErrorsAs
  /// Text black, fills removed, graphics grayscale.
  readonly blackAndWhite?: boolean
  /// No gridlines and no drawings (charts, pictures, shapes).
  readonly draft?: boolean
  /// Page number the first page carries (&P in header/footer); absent = 1.
  readonly firstPageNumber?: number
}

const ERROR_VALUE = /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|GETTING_DATA)$/

/// The text an error cell prints under "Cell errors as"; other text is returned unchanged.
export function printedErrorText(text: string, mode: PrintErrorsAs | undefined): string {
  if (mode === undefined || mode === 'displayed' || !ERROR_VALUE.test(text)) return text
  return mode === 'blank' ? '' : mode === 'dash' ? '--' : '#N/A'
}

/// Offset added to &P so the first printed page carries `firstPageNumber`.
export function firstPageOffset(options: PrintOnlyOptions | undefined): number {
  const n = options?.firstPageNumber
  return n === undefined || !Number.isFinite(n) ? 0 : Math.trunc(n) - 1
}

/// Top padding (pt, before print scale) that centers a block of `contentPt` on a page of
/// `capacityPt`; 0 when the block fills or exceeds the page (it then flows over several pages).
export function verticalCenterPadding(contentPt: number, capacityPt: number): number {
  if (contentPt <= 0 || contentPt >= capacityPt) return 0
  // a little slack so rounding never pushes the last row onto a new page
  return Math.max(0, Math.floor((capacityPt - contentPt) / 2) - 2)
}

/// True when no option departs from the defaults.
export function isDefaultPrintOptions(options: PrintOnlyOptions | undefined): boolean {
  if (!options) return true
  return (
    options.margins === undefined &&
    !options.centerHorizontally &&
    !options.centerVertically &&
    (options.errorsAs === undefined || options.errorsAs === 'displayed') &&
    !options.blackAndWhite &&
    !options.draft &&
    (options.firstPageNumber === undefined || options.firstPageNumber === 1)
  )
}
