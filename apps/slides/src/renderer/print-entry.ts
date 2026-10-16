/**
 * Where "Print" leads. macOS goes straight to the system print panel (printer, copies, duplex, the
 * PDF menu and its page preview all come from the OS, like Office for Mac); Windows and Linux open
 * the in-app print sheet with its own preview and printer settings (like Office for Windows).
 */
export type PrintEntry = 'native-panel' | 'in-app'

export function printEntryFor(platform: string): PrintEntry {
  return platform.toLowerCase().includes('mac') ? 'native-panel' : 'in-app'
}
