import type { Lang } from '@genoffice/i18n'

export const tabDragStrings: Record<Lang, { newWindow: string; joinWindow: string }> = {
  zh: { newWindow: '松开以新建窗口', joinWindow: '松开以移入「{name}」' },
  'zh-TW': { newWindow: '放開以建立新視窗', joinWindow: '放開以移入「{name}」' },
  en: { newWindow: 'Release to open a new window', joinWindow: 'Release to move into “{name}”' },
  ja: { newWindow: '離すと新しいウィンドウで開く', joinWindow: '離すと「{name}」に移動' },
  ko: { newWindow: '놓아서 새 창에서 열기', joinWindow: '놓아서 “{name}” 창으로 이동' },
  fr: {
    newWindow: 'Relâchez pour ouvrir une nouvelle fenêtre',
    joinWindow: 'Relâchez pour déplacer vers « {name} »',
  },
  de: {
    newWindow: 'Loslassen, um ein neues Fenster zu öffnen',
    joinWindow: 'Loslassen, um zu „{name}“ zu verschieben',
  },
  es: {
    newWindow: 'Suelta para abrir una ventana nueva',
    joinWindow: 'Suelta para mover a «{name}»',
  },
  th: { newWindow: 'ปล่อยเพื่อเปิดหน้าต่างใหม่', joinWindow: 'ปล่อยเพื่อย้ายไปยัง “{name}”' },
  id: {
    newWindow: 'Lepaskan untuk membuka jendela baru',
    joinWindow: 'Lepaskan untuk memindahkan ke “{name}”',
  },
  ru: {
    newWindow: 'Отпустите, чтобы открыть новое окно',
    joinWindow: 'Отпустите, чтобы переместить в «{name}»',
  },
  ar: { newWindow: 'أفلت لفتح نافذة جديدة', joinWindow: 'أفلت للنقل إلى «{name}»' },
  pt: {
    newWindow: 'Solte para abrir uma nova janela',
    joinWindow: 'Solte para mover para “{name}”',
  },
  it: {
    newWindow: 'Rilascia per aprire una nuova finestra',
    joinWindow: 'Rilascia per spostare in «{name}»',
  },
  pl: { newWindow: 'Puść, aby otworzyć nowe okno', joinWindow: 'Puść, aby przenieść do „{name}”' },
  cs: { newWindow: 'Puštěním otevřete nové okno', joinWindow: 'Puštěním přesunete do „{name}“' },
  nl: {
    newWindow: 'Laat los om een nieuw venster te openen',
    joinWindow: 'Laat los om naar ‘{name}’ te verplaatsen',
  },
  ms: {
    newWindow: 'Lepaskan untuk membuka tetingkap baharu',
    joinWindow: 'Lepaskan untuk memindahkan ke “{name}”',
  },
  he: { newWindow: 'שחרר כדי לפתוח חלון חדש', joinWindow: 'שחרר כדי להעביר אל „{name}”' },
  hi: {
    newWindow: 'नई विंडो खोलने के लिए छोड़ें',
    joinWindow: '“{name}” में ले जाने के लिए छोड़ें',
  },
}
