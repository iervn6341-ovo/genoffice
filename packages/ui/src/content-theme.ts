import { useCallback } from 'react'

import type { AiPanelPrefs, ContentTheme } from './ai-panel-prefs'
import { useAiPanelPrefs } from './ai-panel-prefs-store'

/** Whether the document content is shown dark, given the saved choice and the app theme. */
export function resolveContentDark(contentTheme: ContentTheme, themeDark: boolean): boolean {
  return contentTheme === 'auto' ? themeDark : contentTheme === 'dark'
}

/**
 * Content darkness for an editor's View ▸ Dark Mode: `auto` follows the app theme (light / dark /
 * gray), otherwise the saved choice wins and survives restarts. Pressing the button pins the
 * opposite of what is shown; Settings can put it back on `auto`. `save` is the app's
 * `setAiPanelPrefs` (the preferences record is shared by every renderer).
 */
export function useContentDarkPref(
  themeDark: boolean,
  save: (patch: Partial<AiPanelPrefs>) => unknown,
): readonly [boolean, (dark: boolean) => void] {
  const { contentTheme } = useAiPanelPrefs()
  const dark = resolveContentDark(contentTheme, themeDark)
  const setDark = useCallback(
    (next: boolean) => void save({ contentTheme: next ? 'dark' : 'light' }),
    [save],
  )
  return [dark, setDark]
}
