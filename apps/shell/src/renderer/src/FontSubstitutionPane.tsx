import { useEffect, useMemo, useState } from 'react'
import {
  DEFAULT_FONT_SUBSTITUTIONS,
  RECOMMENDED_SUBSTITUTE_FONTS,
  isOfficeFamily,
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
  const [newSubstitute, setNewSubstitute] = useState('PingFang TC')
  const [isNewCustom, setIsNewCustom] = useState(false)
  const [customInputRows, setCustomInputRows] = useState<Set<string>>(new Set())
  const [acquiredFonts, setAcquiredFonts] = useState<Set<string>>(new Set())
  const [acquiringFont, setAcquiringFont] = useState<string | null>(null)
  const [savedFeedback, setSavedFeedback] = useState(false)
  const [toastMessage, setToastMessage] = useState<string | null>(null)
  const [noOfficeDialogOpen, setNoOfficeDialogOpen] = useState(false)

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
    const sub = (newSubstitute || systemFonts[0]?.name || 'Carlito GO').trim()
    if (!orig || !sub) return
    const filtered = rules.filter((r) => r.original.toLowerCase() !== orig.toLowerCase())
    const next = [{ original: orig, substitute: sub, isDefault: false }, ...filtered]
    saveRules(next)
    setNewOriginal('')
    setNewSubstitute('')
    setIsAdding(false)
    setIsNewCustom(false)
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

  const handleAcquire = async (fontFamily: string) => {
    setAcquiringFont(fontFamily)
    try {
      const res = await window.aiOffice.acquireOfficeFont(fontFamily)
      if (res.success) {
        setAcquiredFonts((prev) => {
          const next = new Set(prev)
          next.add(fontFamily.toLowerCase())
          return next
        })
        setToastMessage(`✓ ${t('fontSubAcquireSuccess')}`)
        setTimeout(() => setToastMessage(null), 3500)
      } else if (res.reason === 'no-office') {
        setNoOfficeDialogOpen(true)
      } else {
        setToastMessage(`⚠ ${res.message}`)
        setTimeout(() => setToastMessage(null), 4000)
      }
    } catch (err) {
      setToastMessage(`⚠ ${String(err)}`)
      setTimeout(() => setToastMessage(null), 4000)
    } finally {
      setAcquiringFont(null)
    }
  }

  const filteredRules = useMemo(() => {
    const q = query.trim().toLowerCase()
    if (!q) return rules
    return rules.filter(
      (r) => r.original.toLowerCase().includes(q) || r.substitute.toLowerCase().includes(q),
    )
  }, [rules, query])

  // Only surface system & bundled fonts that actually resolve on this machine,
  // preventing misleading warnings when users pick a font from the "System" group.
  const systemFonts = useMemo(
    () =>
      RECOMMENDED_SUBSTITUTE_FONTS.filter(
        (f) => f.group === 'system' && !substituteIsUnavailable(f.name),
      ),
    [acquiredFonts],
  )

  const officeFonts = useMemo(
    () => RECOMMENDED_SUBSTITUTE_FONTS.filter((f) => f.group === 'office'),
    [],
  )

  const isKnownFont = (val: string) => {
    const target = val.trim().toLowerCase()
    return (
      systemFonts.some((f) => f.name.toLowerCase() === target) ||
      officeFonts.some((f) => f.name.toLowerCase() === target)
    )
  }

  // Mark rules whose substitute target is unavailable on this machine,
  // excluding fonts that have been acquired in this session
  const unavailable = useMemo(
    () =>
      new Set(
        rules
          .filter(
            (r) =>
              !acquiredFonts.has(r.substitute.toLowerCase()) &&
              substituteIsUnavailable(r.substitute),
          )
          .map((r) => r.original),
      ),
    [rules, acquiredFonts],
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
            onClick={() => {
              setIsAdding(!isAdding)
              setIsNewCustom(false)
            }}
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
                {isNewCustom ? (
                  <div className="font-sub-custom-input-wrap">
                    <input
                      type="text"
                      className="set-input font-sub-input"
                      placeholder={t('fontSubTargetPlaceholder')}
                      value={newSubstitute}
                      onChange={(e) => setNewSubstitute(e.target.value)}
                      autoFocus
                    />
                    <button
                      type="button"
                      className="set-btn font-sub-back-btn"
                      onClick={() => setIsNewCustom(false)}
                      title={t('fontSubBackToSelect')}
                    >
                      ↩
                    </button>
                  </div>
                ) : (
                  <select
                    className="set-input font-sub-select font-sub-input"
                    value={newSubstitute || systemFonts[0]?.name || 'Carlito GO'}
                    onChange={(e) => {
                      const val = e.target.value
                      if (val === '__custom__') {
                        setIsNewCustom(true)
                      } else {
                        setNewSubstitute(val)
                      }
                    }}
                    aria-label={t('fontSubTarget')}
                  >
                    <optgroup label={t('fontSubGroupSystem')}>
                      {systemFonts.map((font) => (
                        <option key={font.name} value={font.name}>
                          {font.label}
                        </option>
                      ))}
                    </optgroup>
                    <optgroup label={t('fontSubGroupOffice')}>
                      {officeFonts.map((font) => (
                        <option key={font.name} value={font.name}>
                          {font.label}
                        </option>
                      ))}
                    </optgroup>
                    {!isKnownFont(newSubstitute || systemFonts[0]?.name || 'Carlito GO') &&
                      (newSubstitute || systemFonts[0]?.name || 'Carlito GO').trim() && (
                        <optgroup label={t('fontSubGroupCustom')}>
                          <option value={newSubstitute}>{newSubstitute}</option>
                        </optgroup>
                      )}
                    <option value="__custom__">{t('fontSubCustomOption')}</option>
                  </select>
                )}
              </div>
            </div>
          </div>
          <div className="font-sub-add-buttons">
            <button
              type="button"
              className="set-btn primary"
              disabled={
                !newOriginal.trim() ||
                !(newSubstitute || systemFonts[0]?.name || 'Carlito GO').trim()
              }
              onClick={handleAdd}
            >
              {t('fontSubSave')}
            </button>
            <button
              type="button"
              className="set-btn"
              onClick={() => {
                setIsAdding(false)
                setIsNewCustom(false)
              }}
            >
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
              {filteredRules.map((rule) => {
                const isCustomEditing = customInputRows.has(rule.original)
                const isMissing =
                  unavailable.has(rule.original) &&
                  !acquiredFonts.has(rule.substitute.toLowerCase())
                const canAcquire = isOfficeFamily(rule.substitute) && isMissing

                return (
                  <tr key={rule.original} className="font-sub-tr">
                    <td className="font-sub-td">
                      <span className="font-sub-orig-badge" title={rule.original}>
                        {rule.original}
                      </span>
                    </td>
                    <td className="font-sub-td">
                      <div className="font-sub-target-cell">
                        {isCustomEditing ? (
                          <div className="font-sub-custom-input-wrap">
                            <input
                              type="text"
                              className="set-input font-sub-cell-input"
                              value={rule.substitute}
                              onChange={(e) =>
                                handleUpdateSubstitute(rule.original, e.target.value)
                              }
                              placeholder={t('fontSubTargetPlaceholder')}
                              autoFocus
                            />
                            <button
                              type="button"
                              className="set-btn font-sub-back-btn"
                              onClick={() => {
                                setCustomInputRows((prev) => {
                                  const next = new Set(prev)
                                  next.delete(rule.original)
                                  return next
                                })
                              }}
                              title={t('fontSubBackToSelect')}
                            >
                              ↩
                            </button>
                          </div>
                        ) : (
                          <select
                            className="set-input font-sub-select font-sub-cell-select"
                            value={rule.substitute}
                            onChange={(e) => {
                              const val = e.target.value
                              if (val === '__custom__') {
                                setCustomInputRows((prev) => new Set(prev).add(rule.original))
                              } else {
                                handleUpdateSubstitute(rule.original, val)
                              }
                            }}
                            aria-label={t('fontSubTarget')}
                          >
                            <optgroup label={t('fontSubGroupSystem')}>
                              {systemFonts.map((font) => (
                                <option key={font.name} value={font.name}>
                                  {font.label}
                                </option>
                              ))}
                            </optgroup>
                            <optgroup label={t('fontSubGroupOffice')}>
                              {officeFonts.map((font) => (
                                <option key={font.name} value={font.name}>
                                  {font.label}
                                </option>
                              ))}
                            </optgroup>
                            {!isKnownFont(rule.substitute) && rule.substitute.trim() && (
                              <optgroup label={t('fontSubGroupCustom')}>
                                <option value={rule.substitute}>{rule.substitute}</option>
                              </optgroup>
                            )}
                            <option value="__custom__">{t('fontSubCustomOption')}</option>
                          </select>
                        )}
                        {canAcquire && (
                          <button
                            type="button"
                            className="set-btn font-sub-acquire-btn"
                            disabled={acquiringFont === rule.substitute}
                            onClick={() => handleAcquire(rule.substitute)}
                            title={t('fontSubAcquire')}
                          >
                            {acquiringFont === rule.substitute
                              ? t('fontSubAcquiring')
                              : t('fontSubAcquire')}
                          </button>
                        )}
                      </div>
                      {isMissing && (
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
                )
              })}
            </tbody>
          </table>
        </div>
      )}

      {noOfficeDialogOpen && (
        <div
          className="font-sub-modal-backdrop"
          onClick={() => setNoOfficeDialogOpen(false)}
        >
          <div
            className="font-sub-modal-card"
            role="dialog"
            aria-modal="true"
            aria-labelledby="font-sub-no-office-title"
            onClick={(e) => e.stopPropagation()}
          >
            <h4 id="font-sub-no-office-title" className="font-sub-modal-title">
              {t('fontSubNoOfficeTitle')}
            </h4>
            <p className="font-sub-modal-desc">{t('fontSubNoOfficeDesc')}</p>
            <div className="font-sub-modal-actions">
              <button
                type="button"
                className="set-btn primary"
                onClick={async () => {
                  await window.aiOffice.openFontsFolder()
                  setNoOfficeDialogOpen(false)
                }}
              >
                {t('fontSubOpenFontsFolder')}
              </button>
              <button
                type="button"
                className="set-btn"
                onClick={() => setNoOfficeDialogOpen(false)}
              >
                {t('cancel')}
              </button>
            </div>
          </div>
        </div>
      )}

      {toastMessage && (
        <div className="font-sub-toast" role="status">
          {toastMessage}
        </div>
      )}
    </div>
  )
}
