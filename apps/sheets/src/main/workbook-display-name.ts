import { basename } from 'node:path'

/// The sidecar opens a UUID-named snapshot. Renderer-facing fields must name
/// the user's workbook, including recovery and converted import sessions.
export function workbookDisplayName(
  openedPath: string,
  options: {
    readonly csvSourcePath?: string | undefined
    readonly restoreTarget?: string | undefined
    readonly suggestSaveAs?: string | undefined
  } = {},
): string {
  return basename(
    options.csvSourcePath ?? options.restoreTarget ?? options.suggestSaveAs ?? openedPath,
  )
}
