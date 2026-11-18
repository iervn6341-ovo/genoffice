/**
 * UI theme attributes on `<html>`. "gray" is a dark-family theme: it sets
 * `data-theme="dark"` (so every existing dark rule, Univer darkMode and canvas
 * color table keeps working) plus `data-theme-variant="gray"`, which only
 * lifts the chrome tokens in tokens.css to Office's softer dark gray.
 */
export type UiThemeName = 'light' | 'dark' | 'gray' | 'system'

interface UiThemeSource {
  getTheme(): Promise<UiThemeName>
  onThemeChanged(handler: (theme: UiThemeName) => void): () => void
}

/** Subscribe before reading the snapshot; a newer event always wins. Start this
 * before awaiting language, fonts or other bootstrap work, and await ready
 * before first paint without reapplying the possibly stale snapshot later. */
export function syncUiTheme(
  source: UiThemeSource | undefined,
  apply: (theme: UiThemeName) => void = applyUiTheme,
): { ready: Promise<void>; dispose: () => void } {
  let disposed = false
  let receivedEvent = false
  const off = source?.onThemeChanged?.((theme) => {
    if (disposed) return
    receivedEvent = true
    apply(theme)
  })
  const ready = Promise.resolve()
    .then(() => source?.getTheme() ?? 'system')
    .then(
      (theme) => {
        if (!disposed && !receivedEvent) apply(theme)
      },
      () => {
        if (!disposed && !receivedEvent) apply('system')
      },
    )
  return {
    ready,
    dispose: () => {
      disposed = true
      off?.()
    },
  }
}

export function applyUiTheme(
  theme: UiThemeName,
  root: HTMLElement = document.documentElement,
): void {
  if (theme === 'system') {
    root.removeAttribute('data-theme')
    root.removeAttribute('data-theme-variant')
    return
  }
  root.setAttribute('data-theme', theme === 'gray' ? 'dark' : theme)
  if (theme === 'gray') root.setAttribute('data-theme-variant', 'gray')
  else root.removeAttribute('data-theme-variant')
}
