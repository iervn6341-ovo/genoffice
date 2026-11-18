/**
 * The documents the Docs app opens. Legacy Word .doc/.dot arrives through the
 * import filter, which converts a copy — the same route the shell takes — so
 * the standalone argv route, the external-open guard and the open dialog's
 * picker filter all read this one list instead of drifting apart.
 */

export const OPENABLE_DOC_EXTENSIONS = ['docx', 'doc', 'dot'] as const

const OPENABLE_DOC_PATTERN = new RegExp(`\\.(${OPENABLE_DOC_EXTENSIONS.join('|')})$`, 'i')

export function isOpenableDocPath(path: string): boolean {
  return OPENABLE_DOC_PATTERN.test(path.trim())
}

/** Which import filter converts a non-native openable document. */
export type DocImportKind = 'word97'

/**
 * The import filter an openable document needs: legacy Word 97-2003
 * `.doc`/`.dot` through doc2docx; native `.docx` returns null (no conversion on open).
 */
export function docImportKindForPath(path: string): DocImportKind | null {
  const extension = path
    .trim()
    .toLowerCase()
    .match(/\.([a-z0-9]+)$/)?.[1]
  if (extension === 'doc' || extension === 'dot') return 'word97'
  return null
}

export function findOpenableDocPath(argv: readonly string[]): string | null {
  return (
    argv.find((arg) => {
      const value = arg.trim()
      return value.length > 0 && !value.startsWith('-') && isOpenableDocPath(value)
    }) ?? null
  )
}

export function findDocxPath(argv: readonly string[]): string | null {
  return findOpenableDocPath(argv)
}
