import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@genoffice/i18n'
import { App } from './App'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/color-picker.css'
import '@genoffice/ui/dropdown.css'
import '@genoffice/ui/floating-outline.css'
import '@genoffice/ui/ribbon-collapse.css'
import '@genoffice/ui/ribbon-fit.css'
import '@genoffice/ui/markdown.css'
import '@genoffice/ui/ai-panel-prefs.css'
import '@genoffice/ui/ai-scope-quote.css'
import '@genoffice/ui/image-viewer.css'
import '@genoffice/ui/printer-settings.css'
import '@genoffice/ui/craft-icon.css'
import '@genoffice/ui/backstage.css'
import '@genoffice/ui/math-tools.css'
import './styles.css'
// compact Microsoft 365 ribbon metrics: after styles.css so they win over its base rules
import '@genoffice/ui/ribbon-density.css'
import './fonts/fonts.css'
import {
  syncUiTheme,
  applyAiPanelPrefs,
  installMacLineBoundaryKeys,
  installScreenTips,
} from '@genoffice/ui'
import { setAltChunkHtmlConverter } from '@genoffice/docx-engine'

installScreenTips()
installMacLineBoundaryKeys()
if (window.desktop?.convertAltChunkHtml) {
  setAltChunkHtmlConverter((html) => window.desktop.convertAltChunkHtml(html))
}

const themeReady = syncUiTheme(window.desktop).ready

async function bootstrap(): Promise<void> {
  let lang: Lang = 'zh'
  try {
    // per-promise catch: standalone runs have no app:get-theme handler, and
    // that rejection must not drop a resolved language
    ;[lang] = await Promise.all([
      window.desktop.getLanguage().catch(() => 'zh' as const),
      themeReady,
    ])
  } catch {
    /* dev renderer without the preload bridge */
  }
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  void window.desktop
    ?.getAiPanelPrefs?.()
    .then(applyAiPanelPrefs)
    .catch(() => {})
  window.desktop?.onAiPanelPrefsChanged?.(applyAiPanelPrefs)
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
}

void bootstrap()
