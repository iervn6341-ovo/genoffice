import { printJobStringsFor } from '@genoffice/electron-utils/print-job-strings'
import { aiStrings } from './strings-ai'
import { appStrings } from './strings-app'

export const strings = {
  zh: { ...printJobStringsFor('zh'), ...appStrings.zh, ...aiStrings.zh },
  en: { ...printJobStringsFor('en'), ...appStrings.en, ...aiStrings.en },
  ja: { ...printJobStringsFor('ja'), ...appStrings.ja, ...aiStrings.ja },
  ko: { ...printJobStringsFor('ko'), ...appStrings.ko, ...aiStrings.ko },
  fr: { ...printJobStringsFor('fr'), ...appStrings.fr, ...aiStrings.fr },
  de: { ...printJobStringsFor('de'), ...appStrings.de, ...aiStrings.de },
  es: { ...printJobStringsFor('es'), ...appStrings.es, ...aiStrings.es },
  th: { ...printJobStringsFor('th'), ...appStrings.th, ...aiStrings.th },
  id: { ...printJobStringsFor('id'), ...appStrings.id, ...aiStrings.id },
  ru: { ...printJobStringsFor('ru'), ...appStrings.ru, ...aiStrings.ru },
  ar: { ...printJobStringsFor('ar'), ...appStrings.ar, ...aiStrings.ar },
  pt: { ...printJobStringsFor('pt'), ...appStrings.pt, ...aiStrings.pt },
  it: { ...printJobStringsFor('it'), ...appStrings.it, ...aiStrings.it },
  pl: { ...printJobStringsFor('pl'), ...appStrings.pl, ...aiStrings.pl },
  cs: { ...printJobStringsFor('cs'), ...appStrings.cs, ...aiStrings.cs },
  nl: { ...printJobStringsFor('nl'), ...appStrings.nl, ...aiStrings.nl },
  ms: { ...printJobStringsFor('ms'), ...appStrings.ms, ...aiStrings.ms },
  he: { ...printJobStringsFor('he'), ...appStrings.he, ...aiStrings.he },
  hi: { ...printJobStringsFor('hi'), ...appStrings.hi, ...aiStrings.hi },
  'zh-TW': { ...printJobStringsFor('zh-TW'), ...appStrings['zh-TW'], ...aiStrings['zh-TW'] },
}
