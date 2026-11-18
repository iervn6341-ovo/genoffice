import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlLang } from '@genoffice/i18n'
import { AppFrame } from './AppFrame'
import { LocaleProvider } from './locale'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/dropdown.css'
import './home.css'
import './tabbar.css'
import { syncUiTheme, installScreenTips } from '@genoffice/ui'

installScreenTips()
const themeReady = syncUiTheme(window.aiOffice).ready

// macOS shell window is created with vibrancy; a transparent body lets the
// editor views' translucent regions (e.g. slides thumbnail pane) show it
const IS_MAC = navigator.platform.toLowerCase().includes('mac')
if (IS_MAC) document.body.classList.add('vib')
// non-mac: the tab strip doubles as the title bar (caption buttons overlay it)
document.body.classList.add(IS_MAC ? 'mac' : 'overlay-title-bar')

// resolve the persisted language, first-run flag, and theme before first paint
// so the UI never flashes (home showing briefly before the onboarding overlay)
void Promise.all([
  window.aiOffice.getLanguage(),
  // if the flag is unreadable, skip onboarding rather than block the home screen
  window.aiOffice.onboardingSeen().catch(() => true),
  themeReady,
]).then(([lang, onboardingSeen]) => {
  document.documentElement.lang = htmlLang(lang)
  createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <LocaleProvider initial={lang}>
        <AppFrame initialOnboardingSeen={onboardingSeen} />
      </LocaleProvider>
    </React.StrictMode>,
  )
})
