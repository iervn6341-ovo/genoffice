import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@genoffice/i18n'
import App from './App'
import { LocaleProvider } from './i18n/locale'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/dropdown.css'
import '@genoffice/ui/find-panel.css'
import '@genoffice/ui/floating-outline.css'
import '@genoffice/ui/ribbon-collapse.css'
import '@genoffice/ui/ribbon-fit.css'
import '@genoffice/ui/markdown.css'
import '@genoffice/ui/ai-panel-prefs.css'
import '@genoffice/ui/ai-scope-quote.css'
import '@genoffice/ui/image-viewer.css'
import 'katex/dist/katex.min.css'
import './styles.css'
// compact Microsoft 365 ribbon metrics: after styles.css so they win over its base rules
import '@genoffice/ui/ribbon-density.css'
import {
  syncUiTheme,
  applyAiPanelPrefs,
  installMacLineBoundaryKeys,
  installScreenTips,
} from '@genoffice/ui'

installScreenTips()
installMacLineBoundaryKeys()

const themeReady = syncUiTheme(window.markdownApi).ready

void (async () => {
  const [lang] = await Promise.all([
    window.markdownApi.getLanguage().catch(() => 'zh' as const),
    themeReady,
  ])
  document.documentElement.lang = htmlLang(lang as Lang)
  void window.markdownApi
    ?.getAiPanelPrefs?.()
    .then(applyAiPanelPrefs)
    .catch(() => {})
  window.markdownApi?.onAiPanelPrefsChanged?.(applyAiPanelPrefs)
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
})()
