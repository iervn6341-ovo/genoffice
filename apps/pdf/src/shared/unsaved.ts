const UNSAVED_PDF_PREFIX = 'genoffice-memory-pdf://'

/** An opaque view-local source identifier, never a filesystem path. */
export function unsavedPdfPath(viewId: number, displayName = 'Untitled PDF'): string {
  const stem = displayName
    .replace(/\.pdf$/i, '')
    // eslint-disable-next-line no-control-regex -- control characters cannot appear in a save-dialog filename
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
    .trim()
    .slice(0, 120)
  return `${UNSAVED_PDF_PREFIX}${viewId}/${stem || 'Untitled PDF'}.pdf`
}

export function isUnsavedPdfPath(path: string): boolean {
  return path.startsWith(UNSAVED_PDF_PREFIX)
}
