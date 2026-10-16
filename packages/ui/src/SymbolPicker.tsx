import { useEffect, useMemo, useState } from 'react'
import {
  SYMBOL_SUBSETS,
  findSymbolByCodePoint,
  symbolCodePoint,
  symbolSearchText,
} from '@genoffice/office-math'

const RECENT_KEY = 'genoffice.symbols.recent'
const DEFAULT_FONTS = ['Cambria Math', 'Segoe UI Symbol', 'Arial']

export interface SymbolPickerLabels {
  font: string
  subset: string
  search: string
  codePoint: string
  recent: string
  insert: string
  empty: string
  invalidCodePoint: string
  subsetNames: Record<string, string>
}

export interface SymbolPickerProps {
  onPick: (char: string, font: string) => void
  labels: SymbolPickerLabels
  fontFamilies?: readonly string[]
  /** Fires when the highlighted symbol changes (grid, recent strip or code point box). */
  onSelect?: (char: string, font: string) => void
}

function readRecent(): string[] {
  try {
    const value = JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]')
    return Array.isArray(value)
      ? value.filter((char): char is string => typeof char === 'string').slice(0, 12)
      : []
  } catch {
    return []
  }
}

function normalizedHex(raw: string): string {
  return raw.trim().replace(/^U\+/i, '').toUpperCase()
}

/** Search and insert Unicode symbols using a chosen document font. */
export function SymbolPicker({ onPick, labels, fontFamilies, onSelect }: SymbolPickerProps) {
  const fonts = fontFamilies?.length ? [...new Set(fontFamilies)] : DEFAULT_FONTS
  const [font, setFont] = useState(fonts[0] ?? 'Cambria Math')
  const [subset, setSubset] = useState<string>(SYMBOL_SUBSETS[0]?.id ?? '')
  const [search, setSearch] = useState('')
  const [codePointInput, setCodePointInput] = useState('')
  const [selected, setSelected] = useState<string>(SYMBOL_SUBSETS[0]?.symbols[0] ?? '')
  const [recent, setRecent] = useState(readRecent)
  const subsetInfo = SYMBOL_SUBSETS.find((item) => item.id === subset)
  const filtered = useMemo(() => {
    const query = search.trim().toLocaleLowerCase()
    const hex = normalizedHex(query)
    const source = query
      ? SYMBOL_SUBSETS.flatMap((item) => item.symbols)
      : (subsetInfo?.symbols ?? [])
    return [...new Set<string>(source)].filter(
      (char) =>
        !query ||
        char.toLocaleLowerCase().includes(query) ||
        symbolSearchText(char).toLocaleLowerCase().includes(query) ||
        symbolCodePoint(char).toLocaleLowerCase().includes(hex.toLocaleLowerCase()) ||
        SYMBOL_SUBSETS.some(
          (item) =>
            (item.symbols as readonly string[]).includes(char) &&
            (labels.subsetNames[item.labelKey] ?? labels.subsetNames[item.id] ?? item.id)
              .toLocaleLowerCase()
              .includes(query),
        ),
    )
  }, [search, subsetInfo, labels.subsetNames])
  const inputHex = normalizedHex(codePointInput)
  const fromCodePoint = inputHex ? findSymbolByCodePoint(inputHex) : null
  const codePointError = Boolean(codePointInput.trim() && !fromCodePoint)
  const chosen = codePointInput.trim() ? (fromCodePoint ?? '') : selected
  useEffect(() => {
    if (chosen) onSelect?.(chosen, font)
    // onSelect identity is the caller's; only the choice itself should re-fire
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chosen, font])

  const pick = () => {
    if (!chosen) return
    const next = [chosen, ...recent.filter((char) => char !== chosen)].slice(0, 12)
    setRecent(next)
    try {
      localStorage.setItem(RECENT_KEY, JSON.stringify(next))
    } catch {
      /* storage is optional */
    }
    onPick(chosen, font)
  }

  return (
    <div className="go-symbol-picker">
      <div className="go-symbol-filters">
        <label>
          {labels.font}
          <select
            aria-label={labels.font}
            value={font}
            onChange={(event) => setFont(event.target.value)}
          >
            {fonts.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        </label>
        <label>
          {labels.subset}
          <select
            aria-label={labels.subset}
            value={subset}
            onChange={(event) => setSubset(event.target.value)}
          >
            {SYMBOL_SUBSETS.map((item) => (
              <option key={item.id} value={item.id}>
                {labels.subsetNames[item.labelKey] ?? labels.subsetNames[item.id] ?? item.id}
              </option>
            ))}
          </select>
        </label>
        <label>
          {labels.search}
          <input
            type="search"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            aria-label={labels.search}
          />
        </label>
      </div>
      <div className="go-symbol-grid" role="listbox" aria-label={labels.subset}>
        {filtered.map((char) => (
          <button
            key={char}
            type="button"
            role="option"
            aria-selected={selected === char && !codePointInput.trim()}
            className={selected === char && !codePointInput.trim() ? 'active' : ''}
            title={`U+${symbolCodePoint(char)}`}
            style={{ fontFamily: font }}
            onClick={() => {
              setSelected(char)
              setCodePointInput('')
            }}
          >
            {char}
          </button>
        ))}
        {filtered.length === 0 && <span className="go-symbol-empty">{labels.empty}</span>}
      </div>
      {recent.length > 0 && (
        <div className="go-symbol-recent">
          <span>{labels.recent}</span>
          <div>
            {recent.map((char) => (
              <button
                key={char}
                type="button"
                title={`U+${symbolCodePoint(char)}`}
                style={{ fontFamily: font }}
                onClick={() => {
                  setSelected(char)
                  setCodePointInput('')
                }}
              >
                {char}
              </button>
            ))}
          </div>
        </div>
      )}
      <div className="go-symbol-bottom">
        <label>
          {labels.codePoint}
          <input
            value={codePointInput || (selected ? `U+${symbolCodePoint(selected)}` : '')}
            onChange={(event) => setCodePointInput(event.target.value)}
            aria-label={labels.codePoint}
            spellCheck={false}
          />
        </label>
        {codePointError && <span className="go-symbol-error">{labels.invalidCodePoint}</span>}
        <span className="go-symbol-chosen" style={{ fontFamily: font }}>
          {chosen}
        </span>
        <button type="button" className="go-symbol-insert" disabled={!chosen} onClick={pick}>
          {labels.insert}
        </button>
      </div>
    </div>
  )
}
