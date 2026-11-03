import { useEffect, useMemo, useState } from 'react'
import {
  BUILTIN_FONT_FAMILIES,
  DEFAULT_FONT_SUBSTITUTIONS,
  substituteIsUnavailable,
  type FontSubstitutionRule,
} from '@genoffice/ui'
import type { TFunc } from './locale'

/**
 * Font substitution settings pane allowing users to view, search, add,
 * customize, and reset font replacement mappings across documents.
 */
export function FontSubstitutionPane({ t }: { t: TFunc }) {
  const [rules, setRules] = useState<FontSubstitutionRule[]>([])
  const [loading, setLoading] = useState(true)
  const [query, setQuery] = useState('')
  const [isAdding, setIsAdding] = useState(false)
  const [newOriginal, setNewOriginal] = useState('')
  const [newSubstitute, setNewSubstitute] = useState('Songti TC')
  const [savedFeedback, setSavedFeedback] = useState(false)

  useEffect(() => {
    let canceled = false
    void window.aiOffice.getFontSubstitutions?.().then((loaded) => {
      if (!canceled && loaded) {
        setRules(loaded)
        setLoading(false)
      }
    })
    return () => {
      canceled = true
    }
  }, [])

  const saveRules = (updated: FontSubstitutionRule[]) => {
    setRules(updated)
    void window.aiOffice.setFontSubstitutions?.(updated)
    setSavedFeedback(true)
    setTimeout(() => setSavedFeedback(false), 2000)
  }

  const handleReset = () => {
    saveRules([...DEFAULT_FONT_SUBSTITUTIONS])
  }

  const handleAdd = () => {
    const orig = newOriginal.trim()
    const sub = newSubstitute.trim()
    if (!orig || !sub) return
    const filtered = rules.filter((r) => r.original.toLowerCase() !== orig.toLowerCase())
    const next = [{ original: orig, substitute: sub, isDefault: false }, ...filtered]
    saveRules(next)
    setNewOriginal('')
    setIsAdding(false)
  }

  const handleDelete = (original: string) => {
    const next = rules.filter((r) => r.original.toLowerCase() !== original.toLowerCase())
    saveRules(next)
  }

  const handleUpdateSubstitute = (original: string, newSub: string) => {
    const next = rules.map((r) =>
      r.original.toLowerCase() === original.toLowerCase()
        ? { ...r, substitute: newSub, isDefault: false }
        : r,
    )
    saveRules(next)
  }

  const filteredRules = useMemo(() => {
    const q = query.trim().toLowerCase()
    if (!q) return rules
    return rules.filter(
      (r) => r.original.toLowerCase().includes(q) || r.substitute.toLowerCase().includes(q),
    )
  }, [rules, query])

  // a rule whose target nothing on this machine answers to cannot do what it says: say so instead of
  // letting the document fall back to a generic face silently
  const unavailable = useMemo(
    () =>
      new Set(rules.filter((r) => substituteIsUnavailable(r.substitute)).map((r) => r.original)),
    [rules],
  )

  return (
    <div className="font-sub-pane">
      <h3 className="set-pane-title">{t('setSecFonts')}</h3>
      <div className="set-field-desc font-sub-desc">{t('fontSubDesc')}</div>

      <div className="font-sub-toolbar">
        <div className="font-sub-search-wrap">
          <svg
            className="font-sub-search-icon"
            width="14"
            height="14"
            viewBox="0 0 16 16"
            fill="none"
            aria-hidden="true"
          >
            <circle cx="7" cy="7" r="5" stroke="currentColor" strokeWidth="1.3" />
            <path
              d="M10.8 10.8L14.2 14.2"
              stroke="currentColor"
              strokeWidth="1.3"
              strokeLinecap="round"
            />
          </svg>
          <input
            type="search"
            className="set-input font-sub-search"
            placeholder={t('fontSubSearch')}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <div className="font-sub-actions">
          <button
            type="button"
            className="set-btn primary font-sub-btn"
            onClick={() => setIsAdding(!isAdding)}
          >
            + {t('fontSubAdd')}
          </button>
          <button
            type="button"
            className="set-btn font-sub-btn"
            onClick={handleReset}
            title={t('fontSubReset')}
          >
            {t('fontSubReset')}
          </button>
          {savedFeedback && (
            <span className="font-sub-saved" aria-live="polite">
              ✓
            </span>
          )}
        </div>
      </div>

      {isAdding && (
        <div className="font-sub-add-card">
          <div className="font-sub-add-fields">
            <div className="font-sub-add-col">
              <label className="font-sub-label">{t('fontSubOriginal')}</label>
              <input
                type="text"
                className="set-input font-sub-input"
                placeholder={t('fontSubOriginalPlaceholder')}
                value={newOriginal}
                onChange={(e) => setNewOriginal(e.target.value)}
                autoFocus
              />
            </div>
            <div className="font-sub-add-col">
              <label className="font-sub-label">{t('fontSubTarget')}</label>
              <div className="font-sub-target-inputs">
                <input
                  type="text"
                  className="set-input font-sub-input"
                  placeholder={t('fontSubTargetPlaceholder')}
                  value={newSubstitute}
                  onChange={(e) => setNewSubstitute(e.target.value)}
                />
                <select
                  className="set-input font-sub-quick-select"
                  value=""
                  onChange={(e) => {
                    if (e.target.value) setNewSubstitute(e.target.value)
                  }}
                  aria-label={t('fontSubTarget')}
                >
                  <option value="" disabled>
                    ▾ 候選字型
                  </option>
                  {BUILTIN_FONT_FAMILIES.map((font) => (
                    <option key={font} value={font}>
                      {font}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>
          <div className="font-sub-add-buttons">
            <button
              type="button"
              className="set-btn primary"
              disabled={!newOriginal.trim() || !newSubstitute.trim()}
              onClick={handleAdd}
            >
              {t('fontSubSave')}
            </button>
            <button type="button" className="set-btn" onClick={() => setIsAdding(false)}>
              {t('cancel')}
            </button>
          </div>
        </div>
      )}

      {loading ? (
        <div className="font-sub-empty">...</div>
      ) : filteredRules.length === 0 ? (
        <div className="font-sub-empty">{t('fontSubEmpty')}</div>
      ) : (
        <div className="font-sub-table-wrap">
          <table className="font-sub-table">
            <thead>
              <tr>
                <th className="font-sub-th">{t('fontSubOriginal')}</th>
                <th className="font-sub-th">{t('fontSubTarget')}</th>
                <th className="font-sub-th font-sub-col-type">{t('fontSubType')}</th>
                <th className="font-sub-th font-sub-col-actions">{t('moreActions')}</th>
              </tr>
            </thead>
            <tbody>
              {filteredRules.map((rule) => (
                <tr key={rule.original} className="font-sub-tr">
                  <td className="font-sub-td">
                    <span className="font-sub-orig-badge" title={rule.original}>
                      {rule.original}
                    </span>
                  </td>
                  <td className="font-sub-td">
                    <div className="font-sub-target-cell">
                      <input
                        type="text"
                        className="set-input font-sub-cell-input"
                        value={rule.substitute}
                        onChange={(e) => handleUpdateSubstitute(rule.original, e.target.value)}
                      />
                      <select
                        className="set-input font-sub-cell-select"
                        value=""
                        onChange={(e) => {
                          if (e.target.value) {
                            handleUpdateSubstitute(rule.original, e.target.value)
                          }
                        }}
                        aria-label={t('fontSubTarget')}
                      >
                        <option value="" disabled>
                          ▾
                        </option>
                        {BUILTIN_FONT_FAMILIES.map((font) => (
                          <option key={font} value={font}>
                            {font}
                          </option>
                        ))}
                      </select>
                    </div>
                    {unavailable.has(rule.original) && (
                      <div className="font-sub-warn" role="note">
                        ⚠ {t('fontSubSubstituteMissing')}
                      </div>
                    )}
                  </td>
                  <td className="font-sub-td font-sub-col-type">
                    <span
                      className={`font-sub-badge ${rule.isDefault ? 'badge-default' : 'badge-custom'}`}
                    >
                      {rule.isDefault ? t('fontSubTypeDefault') : t('fontSubTypeCustom')}
                    </span>
                  </td>
                  <td className="font-sub-td font-sub-col-actions">
                    <button
                      type="button"
                      className="set-btn danger font-sub-action-btn"
                      onClick={() => handleDelete(rule.original)}
                      title={t('fontSubDelete')}
                    >
                      {t('fontSubDelete')}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
