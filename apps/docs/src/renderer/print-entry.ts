/**
 * Where "Print" leads. macOS goes straight to the system print panel (printer, copies, duplex, the
 * PDF menu and its page preview all come from the OS, like Office for Mac); Windows and Linux open
 * the in-app print dialog with its own preview and Word-style options (like Office for Windows).
 */
export type PrintEntry = 'native-panel' | 'in-app'

/** localStorage key that forces the in-app dialog on macOS (automated tests of that dialog) */
export const FORCE_IN_APP_PRINT_KEY = 'genoffice:docs-force-in-app-print'

export function printEntryFor(platform: string, forceInApp = false): PrintEntry {
  if (forceInApp) return 'in-app'
  return platform.toLowerCase().includes('mac') ? 'native-panel' : 'in-app'
}

/** The entry for this window (reads the platform and the test override) */
export function currentPrintEntry(): PrintEntry {
  let force = false
  try {
    force = window.localStorage.getItem(FORCE_IN_APP_PRINT_KEY) === '1'
  } catch {
    /* storage unavailable: no override */
  }
  return printEntryFor(navigator.platform, force)
}
