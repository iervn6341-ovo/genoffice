/**
 * UI theme attributes on `<html>`. "gray" is a dark-family theme: it sets
 * `data-theme="dark"` (so every existing dark rule, Univer darkMode and canvas
 * color table keeps working) plus `data-theme-variant="gray"`, which only
 * lifts the chrome tokens in tokens.css to Office's softer dark gray.
 */
export type UiThemeName = 'light' | 'dark' | 'gray' | 'system'

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
