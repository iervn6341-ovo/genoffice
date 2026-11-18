import type { Lang } from '@genoffice/i18n'

const labels: Record<Lang, readonly [string, string, string, string]> = {
  en: ['Document outline', 'Search outline', 'No results', 'Close (Esc)'],
  zh: ['文档目录', '搜索目录', '无结果', '关闭 (Esc)'],
  'zh-TW': ['文件目錄', '搜尋目錄', '無符合的結果', '關閉 (Esc)'],
  ja: ['文書の目次', '目次を検索', '結果なし', '閉じる (Esc)'],
  ko: ['문서 목차', '목차 검색', '결과 없음', '닫기 (Esc)'],
  fr: ['Plan du document', 'Rechercher dans le plan', 'Aucun résultat', 'Fermer (Esc)'],
  de: ['Dokumentgliederung', 'Gliederung durchsuchen', 'Keine Ergebnisse', 'Schließen (Esc)'],
  es: ['Índice del documento', 'Buscar en el índice', 'Sin resultados', 'Cerrar (Esc)'],
  th: ['เค้าโครงเอกสาร', 'ค้นหาเค้าโครง', 'ไม่พบผลลัพธ์', 'ปิด (Esc)'],
  id: ['Kerangka dokumen', 'Cari kerangka', 'Tidak ada hasil', 'Tutup (Esc)'],
  ru: ['Структура документа', 'Поиск по структуре', 'Нет результатов', 'Закрыть (Esc)'],
  ar: ['مخطط المستند', 'البحث في المخطط', 'لا توجد نتائج', 'إغلاق (Esc)'],
  pt: ['Estrutura do documento', 'Pesquisar na estrutura', 'Sem resultados', 'Fechar (Esc)'],
  it: ['Struttura del documento', 'Cerca nella struttura', 'Nessun risultato', 'Chiudi (Esc)'],
  pl: ['Konspekt dokumentu', 'Szukaj w konspekcie', 'Brak wyników', 'Zamknij (Esc)'],
  cs: ['Osnova dokumentu', 'Hledat v osnově', 'Žádné výsledky', 'Zavřít (Esc)'],
  nl: ['Documentoverzicht', 'Zoeken in overzicht', 'Geen resultaten', 'Sluiten (Esc)'],
  ms: ['Rangka dokumen', 'Cari rangka', 'Tiada hasil', 'Tutup (Esc)'],
  he: ['חלוקת המסמך', 'חיפוש בחלוקה', 'אין תוצאות', 'סגירה (Esc)'],
  hi: ['दस्तावेज़ की रूपरेखा', 'रूपरेखा खोजें', 'कोई परिणाम नहीं', 'बंद करें (Esc)'],
}

export function floatingOutlineLabels(lang: Lang) {
  const [title, searchLabel, emptyLabel, closeLabel] = labels[lang] ?? labels.en
  return { title, searchLabel, emptyLabel, closeLabel }
}
