import { useEffect, useState } from 'react'
import { useContentDarkPref } from '@genoffice/ui'
import { BehaviorSubject, type Observable } from 'rxjs'

/**
 * Effective darkness of the app UI theme: `<html data-theme>` (set by main.tsx
 * from the shell's light/dark/system setting); in system mode the OS appearance.
 */
export function uiThemeIsDark(): boolean {
  const attr = document.documentElement.getAttribute('data-theme')
  if (attr === 'dark') return true
  if (attr === 'light') return false
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ?? false
}

/** live uiThemeIsDark(): follows theme broadcasts and OS appearance flips */
export function useUiThemeIsDark(): boolean {
  const [dark, setDark] = useState(uiThemeIsDark)
  useEffect(() => {
    const update = (): void => setDark(uiThemeIsDark())
    // main.tsx's listener (registered at bootstrap) updates data-theme first,
    // so reading the attribute in ours is safe
    const off = window.desktopApi?.onThemeChanged?.(update)
    const mq = window.matchMedia?.('(prefers-color-scheme: dark)')
    mq?.addEventListener('change', update)
    return () => {
      off?.()
      mq?.removeEventListener('change', update)
    }
  }, [])
  return dark
}

/**
 * Sheet-content darkness (display-time remap only, never written to the workbook): follows the UI
 * theme until the View ▸ Dark Mode button pins light or dark; the pinned choice is a saved
 * preference shared with Docs and Slides (Settings can put it back on "follow theme").
 */
export function useContentDark(): readonly [boolean, (dark: boolean) => void] {
  const themeDark = useUiThemeIsDark()
  const [contentDark, setContentDark] = useContentDarkPref(themeDark, saveDisplayPrefs)
  // CSS hook: the canvas-drawn formula bar renders in the content shade, so it
  // is filter-inverted back to the UI shade whenever the two disagree
  useEffect(() => {
    document.documentElement.dataset.contentMismatch =
      contentDark === themeDark ? 'false' : themeDark ? 'ui-dark' : 'ui-light'
  }, [contentDark, themeDark])
  return [contentDark, setContentDark]
}

/** Display preferences go to the shell's shared preferences record */
const saveDisplayPrefs = (patch: Record<string, unknown>): unknown =>
  window.desktopApi.setAiPanelPrefs(patch)

/**
 * Univer's `setDarkMode` is one global flag: the canvas colour inversion reads
 * `themeService.darkMode`, while every DOM piece of Univer UI (name box,
 * formula bar, sheet tabs, menus, dialogs) subscribes to `darkMode$` and sets
 * the `univer-dark` class from it. Re-point `darkMode$` at the *UI* theme so
 * the DOM follows the app theme (light / dark / gray) while the canvas keeps
 * following the content-darkness flag (the engine-render repaint subscription
 * was made on the original observable and still fires on `setDarkMode`).
 * Returns a disposer.
 */
export function decoupleUniverUiTheme(themeService: {
  darkMode$: Observable<boolean>
}): () => void {
  const ui$ = new BehaviorSubject<boolean>(uiThemeIsDark())
  const update = (): void => ui$.next(uiThemeIsDark())
  const off = window.desktopApi?.onThemeChanged?.(update)
  const mq = window.matchMedia?.('(prefers-color-scheme: dark)')
  mq?.addEventListener('change', update)
  themeService.darkMode$ = ui$.asObservable()
  return () => {
    off?.()
    mq?.removeEventListener('change', update)
    ui$.complete()
  }
}
