import { useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { SymbolPicker, type SymbolPickerLabels, BUILTIN_FONT_FAMILIES } from '@genoffice/ui'
import { AUTOCORRECT_DEFAULTS, SPECIAL_CHARACTERS, symbolCodePoint } from '@genoffice/office-math'
import { readUserAutoCorrect, writeUserAutoCorrect } from '../editor/autocorrect'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/** Apply the chosen font to the inserted scalar without changing adjacent text. */
export function insertSymbolAtSelection(
  editor: Editor,
  selection: { from: number; to: number },
  char: string,
  font: string | null,
): void {
  const position = editor.state.doc.resolve(selection.from)
  const inherited = editor.state.storedMarks ?? position.marks()
  const textStyle = inherited.find((mark) => mark.type.name === 'docTextStyle')
  const marks = inherited
    .filter((mark) => mark.type.name !== 'docTextStyle')
    .map((mark) => ({ type: mark.type.name, attrs: mark.attrs }))
  // special characters keep the surrounding font; symbols carry the font chosen in the dialog
  if (font) {
    marks.push({
      type: 'docTextStyle',
      attrs: { ...textStyle?.attrs, font, fontAscii: font },
    })
  } else if (textStyle) {
    marks.push({ type: 'docTextStyle', attrs: textStyle.attrs })
  }
  editor
    .chain()
    .focus()
    .insertContentAt(selection, {
      type: 'text',
      text: char,
      marks,
    })
    .run()
}

/** Insert a chosen Unicode scalar at the document selection captured before modal focus. */
export function SymbolDialog({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const selection = useRef({ from: editor.state.selection.from, to: editor.state.selection.to })
  const labels: SymbolPickerLabels = {
    font: t('appSymbolFont'),
    subset: t('appSymbolSubset'),
    search: t('appSymbolSearch'),
    codePoint: t('appSymbolCodePoint'),
    recent: t('appSymbolRecent'),
    insert: t('appInsert'),
    empty: t('appSymbolNoResults'),
    invalidCodePoint: t('appSymbolInvalidCodePoint'),
    subsetNames: {
      latin: t('appSymbolLatin'),
      greek: t('appSymbolGreek'),
      math: t('appSymbolMath'),
      arrows: t('appSymbolArrows'),
      currency: t('appSymbolCurrency'),
      misc: t('appSymbolMisc'),
    },
  }
  const [tab, setTab] = useState<'symbols' | 'special'>('symbols')
  const [chosen, setChosen] = useState<{ char: string; font: string } | null>(null)
  const [trigger, setTrigger] = useState('')
  const [entries, setEntries] = useState(readUserAutoCorrect)
  const [special, setSpecial] = useState(SPECIAL_CHARACTERS[0]?.char ?? '')
  const current = tab === 'special' ? special : chosen?.char
  const info = current ? SPECIAL_CHARACTERS.find((item) => item.char === current) : undefined
  const addEntry = () => {
    const text = trigger.trim()
    if (!text || !current || text in AUTOCORRECT_DEFAULTS) return
    const next = { ...entries, [text]: current }
    setEntries(next)
    writeUserAutoCorrect(next)
    setTrigger('')
  }
  const removeEntry = (text: string) => {
    const { [text]: _removed, ...rest } = entries
    setEntries(rest)
    writeUserAutoCorrect(rest)
  }
  const insert = (char: string, font: string | null) => {
    if (!editor.isEditable) return
    insertSymbolAtSelection(editor, selection.current, char, font)
    onClose()
  }
  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(event) => event.target === event.currentTarget && onClose()}
    >
      <div className="modal docs-symbol-modal">
        <h2>{t('appAdvancedSymbol')}</h2>
        <div className="docs-symbol-tabs" role="tablist">
          {(
            [
              ['symbols', 'appSymbolTabSymbols'],
              ['special', 'appSymbolTabSpecial'],
            ] as const
          ).map(([id, label]) => (
            <button
              key={id}
              type="button"
              role="tab"
              aria-selected={tab === id}
              className={tab === id ? 'active' : ''}
              onClick={() => setTab(id)}
            >
              {t(label)}
            </button>
          ))}
        </div>
        {tab === 'symbols' ? (
          <SymbolPicker
            onPick={insert}
            onSelect={(char, font) => setChosen({ char, font })}
            labels={labels}
            fontFamilies={['Cambria Math', 'Segoe UI Symbol', ...BUILTIN_FONT_FAMILIES]}
          />
        ) : (
          <div className="docs-special-chars" role="listbox" aria-label={t('appSymbolTabSpecial')}>
            {SPECIAL_CHARACTERS.map((item) => (
              <button
                key={item.id}
                type="button"
                role="option"
                aria-selected={special === item.char}
                className={special === item.char ? 'active' : ''}
                onClick={() => setSpecial(item.char)}
                onDoubleClick={() => insert(item.char, null)}
              >
                <span className="docs-special-glyph">
                  {item.char === '\u00A0' || item.char === '\u00AD' || /\s/.test(item.char)
                    ? '\u2423'
                    : item.char}
                </span>
                <span className="docs-special-name">{item.name}</span>
                <span className="docs-special-code">U+{symbolCodePoint(item.char)}</span>
              </button>
            ))}
            <div className="docs-special-actions">
              <button
                type="button"
                className="btn-primary"
                disabled={!special}
                onClick={() => insert(special, null)}
              >
                {t('appInsert')}
              </button>
            </div>
          </div>
        )}
        {current && (
          <p className="docs-symbol-info" role="status">
            {t('appSymbolInfo')}: {info?.name ?? ''} U+{symbolCodePoint(current)}
          </p>
        )}
        <fieldset className="docs-autocorrect">
          <legend>{t('appSymbolAutoCorrect')}</legend>
          <p className="modal-desc">{t('appSymbolAutoCorrectHint')}</p>
          <div className="docs-autocorrect-row">
            <input
              value={trigger}
              maxLength={24}
              placeholder={t('appSymbolAutoCorrectReplace')}
              aria-label={t('appSymbolAutoCorrectReplace')}
              onChange={(event) => setTrigger(event.target.value)}
              onKeyDown={(event) => event.key === 'Enter' && addEntry()}
            />
            <span className="docs-autocorrect-with">{current ?? ''}</span>
            <button
              type="button"
              className="btn-ghost"
              disabled={!trigger.trim() || !current}
              onClick={addEntry}
            >
              {t('appSymbolAutoCorrectAdd')}
            </button>
          </div>
          {Object.entries(entries).map(([text, value]) => (
            <div key={text} className="docs-autocorrect-entry">
              <code>{text}</code>
              <span>{value}</span>
              <button type="button" className="btn-ghost" onClick={() => removeEntry(text)}>
                {t('appSymbolAutoCorrectRemove')}
              </button>
            </div>
          ))}
        </fieldset>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
        </div>
      </div>
    </div>
  )
}
