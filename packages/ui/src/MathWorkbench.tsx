import { useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent } from 'react'
import {
  EQUATION_STRUCTURES,
  SYMBOL_SUBSETS,
  latexToOmml,
  ommlToLatex,
  ommlToMathML,
  ommlToUnicodeMath,
  symbolSearchText,
  unicodeMathToOmml,
} from '@genoffice/office-math'

const SELECTION_MARK = '{{selection}}'
const MATH_SYMBOL_SUBSETS = SYMBOL_SUBSETS.filter(
  (item) => item.id === 'math' || item.id === 'greek' || item.id === 'arrows',
)

function handleTabKeys(event: KeyboardEvent<HTMLButtonElement>): void {
  const tabs = Array.from(
    event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]') ?? [],
  )
  const current = tabs.indexOf(event.currentTarget)
  const step = getComputedStyle(event.currentTarget).direction === 'rtl' ? -1 : 1
  let next: number
  if (event.key === 'ArrowRight') next = (current + step + tabs.length) % tabs.length
  else if (event.key === 'ArrowLeft') next = (current - step + tabs.length) % tabs.length
  else if (event.key === 'Home') next = 0
  else if (event.key === 'End') next = tabs.length - 1
  else return
  event.preventDefault()
  tabs[next]?.focus()
  tabs[next]?.click()
}

export interface MathWorkbenchLabels {
  input: string
  structures: string
  professional: string
  linear: string
  preview: string
  empty: string
  categoryNames: Record<string, string>
  templateNames?: Record<string, string>
  symbols?: string
  symbolSearch?: string
  symbolSubsetNames?: Record<string, string>
  /** when set, the source box can switch between LaTeX and UnicodeMath (Word's linear format) */
  inputFormats?: { latex: string; unicodeMath: string }
}

export interface MathWorkbenchProps {
  latex: string
  onLatexChange: (latex: string) => void
  labels: MathWorkbenchLabels
  autoFocus?: boolean
}

export interface MathTemplateInsertion {
  latex: string
  selectionStart: number
  selectionEnd: number
}

/** Replace the selected source inside a catalog template and keep the next slot editable. */
export function insertMathTemplate(
  source: string,
  start: number,
  end: number,
  template: string,
): MathTemplateInsertion {
  const from = Math.max(0, Math.min(start, source.length))
  const to = Math.max(from, Math.min(end, source.length))
  const selected = source.slice(from, to)
  const mark = template.indexOf(SELECTION_MARK)
  const replacement = mark >= 0 ? template.replace(SELECTION_MARK, selected || 'x') : template
  const latex = source.slice(0, from) + replacement + source.slice(to)
  const cursor = mark >= 0 && !selected ? from + mark : from + replacement.length
  return {
    latex,
    selectionStart: cursor,
    selectionEnd: mark >= 0 && !selected ? cursor + 1 : cursor,
  }
}

/** UnicodeMath text of a LaTeX source, or null when it holds structures the linear form cannot express. */
export function latexToLinear(latex: string): string | null {
  try {
    return ommlToUnicodeMath(`<m:oMath>${latexToOmml(latex)}</m:oMath>`)
  } catch {
    return null
  }
}

/** LaTeX source of UnicodeMath text, or an error message. */
export function linearToLatex(text: string): { latex: string } | { error: string } {
  try {
    const latex = ommlToLatex(`<m:oMath>${unicodeMathToOmml(text)}</m:oMath>`)
    return latex === null
      ? { error: 'This linear formula uses a structure the source editor cannot express' }
      : { latex }
  } catch (error) {
    return { error: error instanceof Error ? error.message : String(error) }
  }
}

export function mathPreview(latex: string): { mathml: string } | { error: string } | null {
  if (!latex.trim()) return null
  try {
    const omml = latexToOmml(latex)
    const mathml = ommlToMathML(`<m:oMath>${omml}</m:oMath>`)
    return mathml ? { mathml } : { error: 'MathML preview unavailable' }
  } catch (error) {
    return { error: error instanceof Error ? error.message : String(error) }
  }
}

/** Shared source editor and structure gallery; the host owns insertion and dialog chrome. */
export function MathWorkbench({ latex, onLatexChange, labels, autoFocus }: MathWorkbenchProps) {
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const pendingSelection = useRef<{ start: number; end: number } | null>(null)
  const [category, setCategory] = useState<string>(EQUATION_STRUCTURES[0]?.id ?? '')
  const [symbolSubset, setSymbolSubset] = useState<string>('math')
  const [symbolSearch, setSymbolSearch] = useState('')
  const [mode, setMode] = useState<'professional' | 'linear'>('professional')
  const [inputFormat, setInputFormat] = useState<'latex' | 'unicode'>('latex')
  const [linearText, setLinearText] = useState('')
  const [linearError, setLinearError] = useState<string | null>(null)
  const preview = useMemo(() => mathPreview(latex), [latex])
  useLayoutEffect(() => {
    const selection = pendingSelection.current
    if (!selection || !inputRef.current) return
    inputRef.current.focus()
    inputRef.current.setSelectionRange(selection.start, selection.end)
    pendingSelection.current = null
  }, [latex])

  const group = EQUATION_STRUCTURES.find((item) => item.id === category)
  const filteredSymbols = useMemo(() => {
    const query = symbolSearch.trim().toLocaleLowerCase()
    const subsets = query
      ? MATH_SYMBOL_SUBSETS
      : MATH_SYMBOL_SUBSETS.filter((item) => item.id === symbolSubset)
    return [...new Set<string>(subsets.flatMap((item) => item.symbols))].filter(
      (char) => !query || symbolSearchText(char).toLocaleLowerCase().includes(query),
    )
  }, [symbolSubset, symbolSearch])
  const switchFormat = (next: 'latex' | 'unicode') => {
    if (next === inputFormat) return
    if (next === 'unicode') {
      setLinearText(latexToLinear(latex) ?? latex)
      setLinearError(null)
    }
    setInputFormat(next)
  }
  const editLinear = (text: string) => {
    setLinearText(text)
    if (text.trim() === '') {
      setLinearError(null)
      onLatexChange('')
      return
    }
    const result = linearToLatex(text)
    if ('latex' in result) {
      setLinearError(null)
      onLatexChange(result.latex)
    } else setLinearError(result.error)
  }
  const insert = (template: string) => {
    if (inputFormat === 'unicode') {
      // structures and symbols are authored in LaTeX; keep the linear box in step
      const next = insertMathTemplate(latex, latex.length, latex.length, template)
      onLatexChange(next.latex)
      setLinearText(latexToLinear(next.latex) ?? next.latex)
      setLinearError(null)
      return
    }
    const input = inputRef.current
    const start = input?.selectionStart ?? latex.length
    const end = input?.selectionEnd ?? start
    const next = insertMathTemplate(latex, start, end, template)
    pendingSelection.current = { start: next.selectionStart, end: next.selectionEnd }
    onLatexChange(next.latex)
  }

  return (
    <div className="go-math-workbench">
      <label className="go-math-label">
        {labels.input}
        {labels.inputFormats && (
          <span className="go-math-view-switch" role="group" aria-label={labels.input}>
            <button
              type="button"
              className={inputFormat === 'latex' ? 'active' : ''}
              aria-pressed={inputFormat === 'latex'}
              onClick={() => switchFormat('latex')}
            >
              {labels.inputFormats.latex}
            </button>
            <button
              type="button"
              className={inputFormat === 'unicode' ? 'active' : ''}
              aria-pressed={inputFormat === 'unicode'}
              onClick={() => switchFormat('unicode')}
            >
              {labels.inputFormats.unicodeMath}
            </button>
          </span>
        )}
        {inputFormat === 'unicode' && labels.inputFormats ? (
          <textarea
            className="go-math-input"
            value={linearText}
            onChange={(event) => editLinear(event.target.value)}
            spellCheck={false}
            aria-label={labels.inputFormats.unicodeMath}
            aria-invalid={linearError !== null}
          />
        ) : (
          <textarea
            ref={inputRef}
            className="go-math-input"
            value={latex}
            onChange={(event) => onLatexChange(event.target.value)}
            autoFocus={autoFocus}
            spellCheck={false}
            aria-label={labels.input}
          />
        )}
        {linearError && inputFormat === 'unicode' && (
          <span className="go-math-error">{linearError}</span>
        )}
      </label>
      <div className="go-math-structures" aria-label={labels.structures}>
        <div className="go-math-categories" role="tablist" aria-label={labels.structures}>
          {EQUATION_STRUCTURES.map((item) => (
            <button
              key={item.id}
              type="button"
              role="tab"
              aria-selected={category === item.id}
              tabIndex={category === item.id ? 0 : -1}
              className={category === item.id ? 'active' : ''}
              onMouseDown={(event) => event.preventDefault()}
              onKeyDown={handleTabKeys}
              onClick={() => setCategory(item.id)}
            >
              {labels.categoryNames[item.labelKey] ?? labels.categoryNames[item.id] ?? item.id}
            </button>
          ))}
          <button
            type="button"
            role="tab"
            aria-selected={category === 'symbols'}
            tabIndex={category === 'symbols' ? 0 : -1}
            className={category === 'symbols' ? 'active' : ''}
            onMouseDown={(event) => event.preventDefault()}
            onKeyDown={handleTabKeys}
            onClick={() => setCategory('symbols')}
          >
            {labels.symbols ?? 'Symbols'}
          </button>
        </div>
        {category === 'symbols' ? (
          <div className="go-math-symbols" role="tabpanel">
            <div className="go-math-symbol-tools">
              <div
                className="go-math-symbol-subsets"
                role="tablist"
                aria-label={labels.symbols ?? 'Symbols'}
              >
                {MATH_SYMBOL_SUBSETS.map((item) => (
                  <button
                    key={item.id}
                    type="button"
                    role="tab"
                    aria-selected={symbolSubset === item.id}
                    tabIndex={symbolSubset === item.id ? 0 : -1}
                    className={symbolSubset === item.id ? 'active' : ''}
                    onMouseDown={(event) => event.preventDefault()}
                    onKeyDown={handleTabKeys}
                    onClick={() => {
                      setSymbolSubset(item.id)
                      setSymbolSearch('')
                    }}
                  >
                    {labels.symbolSubsetNames?.[item.labelKey] ??
                      labels.symbolSubsetNames?.[item.id] ??
                      item.id}
                  </button>
                ))}
              </div>
              <input
                type="search"
                value={symbolSearch}
                onChange={(event) => setSymbolSearch(event.target.value)}
                placeholder={labels.symbolSearch ?? labels.symbols ?? 'Symbols'}
                aria-label={labels.symbolSearch ?? labels.symbols ?? 'Symbols'}
              />
            </div>
            <div className="go-math-symbol-grid" aria-label={labels.symbols ?? 'Symbols'}>
              {filteredSymbols.map((char) => (
                <button
                  key={char}
                  type="button"
                  title={symbolSearchText(char)}
                  aria-label={symbolSearchText(char)}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => insert(char)}
                >
                  {char}
                </button>
              ))}
              {filteredSymbols.length === 0 && (
                <span className="go-math-empty">{labels.empty}</span>
              )}
            </div>
          </div>
        ) : (
          <div className="go-math-templates" role="tabpanel">
            {group?.templates.map((template) => {
              const sample = template.latex.replace(SELECTION_MARK, 'x')
              const result = mathPreview(sample)
              if (!result || 'error' in result) return null
              return (
                <button
                  key={template.id}
                  type="button"
                  className="go-math-template"
                  title={labels.templateNames?.[template.id] ?? sample}
                  aria-label={labels.templateNames?.[template.id] ?? sample}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => insert(template.latex)}
                >
                  <span dangerouslySetInnerHTML={{ __html: result.mathml }} />
                </button>
              )
            })}
          </div>
        )}
      </div>
      <div className="go-math-preview-heading">
        <span>{labels.preview}</span>
        <div className="go-math-view-switch" role="group" aria-label={labels.preview}>
          <button
            type="button"
            className={mode === 'professional' ? 'active' : ''}
            aria-pressed={mode === 'professional'}
            onClick={() => setMode('professional')}
          >
            {labels.professional}
          </button>
          <button
            type="button"
            className={mode === 'linear' ? 'active' : ''}
            aria-pressed={mode === 'linear'}
            onClick={() => setMode('linear')}
          >
            {labels.linear}
          </button>
        </div>
      </div>
      <div className="go-math-preview" aria-live="polite">
        {preview === null && <span className="go-math-empty">{labels.empty}</span>}
        {preview && 'error' in preview && <span className="go-math-error">{preview.error}</span>}
        {preview && 'mathml' in preview && mode === 'professional' && (
          <span dangerouslySetInnerHTML={{ __html: preview.mathml }} />
        )}
        {preview && 'mathml' in preview && mode === 'linear' && (
          <code className="go-math-linear">{latexToLinear(latex) ?? latex}</code>
        )}
      </div>
    </div>
  )
}
