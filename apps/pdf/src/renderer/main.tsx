import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@genoffice/i18n'
import App from './App'
import { LocaleProvider } from './i18n/locale'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/color-picker.css'
import '@genoffice/ui/dropdown.css'
import '@genoffice/ui/floating-outline.css'
import '@genoffice/ui/ribbon-collapse.css'
import '@genoffice/ui/markdown.css'
import '@genoffice/ui/ai-panel-prefs.css'
import '@genoffice/ui/ai-scope-quote.css'
import './styles.css'
import { syncUiTheme, applyAiPanelPrefs, installScreenTips } from '@genoffice/ui'

installScreenTips()

const themeReady = syncUiTheme(window.pdfApi).ready

void (async () => {
  const [lang] = await Promise.all([
    window.pdfApi.getLanguage().catch(() => 'zh' as const),
    themeReady,
  ])
  document.documentElement.lang = htmlLang(lang as Lang)
  void window.pdfApi
    ?.getAiPanelPrefs?.()
    .then(applyAiPanelPrefs)
    .catch(() => {})
  window.pdfApi?.onAiPanelPrefsChanged?.(applyAiPanelPrefs)
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
})()
