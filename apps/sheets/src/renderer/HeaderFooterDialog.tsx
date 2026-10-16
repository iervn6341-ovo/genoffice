import { useEffect, useRef, useState } from 'react'
import { isMetafileMime, metafileToDataUrl } from '@genoffice/docx-engine/metafile'

import { useI18n, type StringKey } from './i18n/locale'
import type { HeaderFooterParts, HeaderFooterPictureEdit } from './edit-journal'
import type { HeaderFooterPictureSlot } from './print-settings'
import {
  headerFooterPicturePosition,
  withHeaderFooterPictureCode,
  withoutHeaderFooterPictureCode,
} from './header-footer-picture-codes'
import { useDialogEscape } from './use-dialog-escape'

interface HeaderFooterPair {
  readonly header: HeaderFooterParts | null
  readonly footer: HeaderFooterParts | null
}

export interface HeaderFooterResult extends HeaderFooterPair {
  readonly sheetId: string | null
  readonly firstPage: HeaderFooterPair | null
  readonly evenPages: HeaderFooterPair | null
  readonly firstPageStored: HeaderFooterPair
  readonly evenPagesStored: HeaderFooterPair
  readonly pictureEdits: readonly HeaderFooterPictureEdit[]
  readonly loadedPictures: readonly LoadedHeaderFooterPicture[]
}

export interface LoadedHeaderFooterPicture {
  readonly position: string
  readonly widthPt: number
  readonly heightPt: number
  readonly mediaType: NonNullable<NonNullable<HeaderFooterPictureEdit['image']>>['mediaType']
  readonly base64: string
}

type Variant = 'odd' | 'first' | 'even'
type SectionKey =
  | 'header-left'
  | 'header-center'
  | 'header-right'
  | 'footer-left'
  | 'footer-center'
  | 'footer-right'

const SECTION_KEYS: readonly SectionKey[] = [
  'header-left',
  'header-center',
  'header-right',
  'footer-left',
  'footer-center',
  'footer-right',
]
const SECTION_LABELS: Record<SectionKey, StringKey> = {
  'header-left': 'dlgHfHeaderLeft',
  'header-center': 'dlgHfHeaderCenter',
  'header-right': 'dlgHfHeaderRight',
  'footer-left': 'dlgHfFooterLeft',
  'footer-center': 'dlgHfFooterCenter',
  'footer-right': 'dlgHfFooterRight',
}
const FIELD_CODES: readonly { readonly code: string; readonly labelKey: StringKey }[] = [
  { code: '&P', labelKey: 'dlgHfPageNumber' },
  { code: '&N', labelKey: 'dlgHfPageCount' },
  { code: '&D', labelKey: 'dlgHfDate' },
  { code: '&T', labelKey: 'dlgHfTime' },
  { code: '&F', labelKey: 'dlgHfFileName' },
  { code: '&Z', labelKey: 'dlgHfFilePath' },
  { code: '&A', labelKey: 'dlgHfSheetName' },
  { code: '&B', labelKey: 'dlgHfBold' },
  { code: '&I', labelKey: 'dlgHfItalic' },
  { code: '&U', labelKey: 'dlgHfUnderline' },
]

function isHeaderFooterImageMime(value: string): value is LoadedHeaderFooterPicture['mediaType'] {
  return [
    'image/png',
    'image/jpeg',
    'image/gif',
    'image/x-emf',
    'image/x-wmf',
    'image/bmp',
    'image/webp',
    'image/svg+xml',
    'image/x-emz',
    'image/x-wmz',
  ].includes(value)
}

function toParts(left: string, center: string, right: string): HeaderFooterParts | null {
  if (left === '' && center === '' && right === '') return null
  return {
    ...(left === '' ? {} : { left }),
    ...(center === '' ? {} : { center }),
    ...(right === '' ? {} : { right }),
  }
}

function initialValues(pair: HeaderFooterPair | null): Record<SectionKey, string> {
  return {
    'header-left': pair?.header?.left ?? '',
    'header-center': pair?.header?.center ?? '',
    'header-right': pair?.header?.right ?? '',
    'footer-left': pair?.footer?.left ?? '',
    'footer-center': pair?.footer?.center ?? '',
    'footer-right': pair?.footer?.right ?? '',
  }
}

function toPair(values: Record<SectionKey, string>): HeaderFooterPair {
  return {
    header: toParts(values['header-left'], values['header-center'], values['header-right']),
    footer: toParts(values['footer-left'], values['footer-center'], values['footer-right']),
  }
}

export function HeaderFooterDialog({
  initialHeader,
  initialFooter,
  initialFirstPage,
  initialEvenPages,
  firstPageEnabled,
  evenPagesEnabled,
  sessionId,
  sheetId,
  initialPictures,
  pendingPictureEdits,
  onApply,
  onClose,
}: {
  readonly initialHeader: HeaderFooterParts | null
  readonly initialFooter: HeaderFooterParts | null
  readonly initialFirstPage: HeaderFooterPair | null
  readonly initialEvenPages: HeaderFooterPair | null
  readonly firstPageEnabled: boolean
  readonly evenPagesEnabled: boolean
  readonly sessionId: string | null
  readonly sheetId: string | null
  readonly initialPictures: readonly HeaderFooterPictureSlot[]
  readonly pendingPictureEdits: readonly HeaderFooterPictureEdit[]
  readonly onApply: (result: HeaderFooterResult) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  useDialogEscape(onClose)
  const { t } = useI18n()
  const translateRef = useRef(t)
  translateRef.current = t
  const [variant, setVariant] = useState<Variant>('odd')
  const [enabled, setEnabled] = useState({
    first: firstPageEnabled,
    even: evenPagesEnabled,
  })
  const [values, setValues] = useState<Record<Variant, Record<SectionKey, string>>>(() => ({
    odd: initialValues({ header: initialHeader, footer: initialFooter }),
    first: initialValues(initialFirstPage),
    even: initialValues(initialEvenPages),
  }))
  const [error, setError] = useState<string | null>(null)
  const [pictures, setPictures] = useState<Record<string, LoadedHeaderFooterPicture>>({})
  const [previews, setPreviews] = useState<Record<string, string>>({})
  const [pictureEdits, setPictureEdits] = useState<Record<string, HeaderFooterPictureEdit>>({})
  const [pictureLoading, setPictureLoading] = useState(initialPictures.length > 0)
  const loadedPictures = useRef<LoadedHeaderFooterPicture[]>([])
  const initialSheetId = useRef(sheetId)
  const focusedKey = useRef<SectionKey>('header-center')
  const inputs = useRef(new Map<SectionKey, HTMLInputElement>())

  const initialPicturesRef = useRef(initialPictures)
  const pendingPictureEditsRef = useRef(pendingPictureEdits)
  useEffect(() => {
    let active = true
    void (async () => {
      try {
        const entries = await Promise.all(
          initialPicturesRef.current.map(async (slot) => {
            const pending = slot.id.startsWith('journal:')
              ? pendingPictureEditsRef.current.find((edit) => edit.position === slot.position)
                  ?.image
              : undefined
            const media =
              pending && pending !== null
                ? pending
                : await window.desktopApi.readWorkbookMedia({
                    sessionId: sessionId!,
                    visualId: slot.id,
                  })
            if (!isHeaderFooterImageMime(media.mediaType))
              throw new Error(translateRef.current('dlgHfPictureInvalid'))
            const entry: LoadedHeaderFooterPicture = {
              position: slot.position,
              widthPt: slot.widthPt,
              heightPt: slot.heightPt,
              mediaType: media.mediaType,
              base64: media.base64,
            }
            return entry
          }),
        )
        if (!active) return
        loadedPictures.current = entries
        setPictures(Object.fromEntries(entries.map((entry) => [entry.position, entry])))
        const urls = await Promise.all(
          entries.map(async (entry) => {
            try {
              const url = isMetafileMime(entry.mediaType)
                ? await metafileToDataUrl(
                    Uint8Array.from(atob(entry.base64), (character) => character.charCodeAt(0)),
                    entry.mediaType,
                  )
                : `data:${entry.mediaType};base64,${entry.base64}`
              return [entry.position, url ?? ''] as const
            } catch {
              return [entry.position, ''] as const
            }
          }),
        )
        if (active) setPreviews(Object.fromEntries(urls))
      } catch (reason: unknown) {
        if (active) setError(reason instanceof Error ? reason.message : String(reason))
      } finally {
        if (active) setPictureLoading(false)
      }
    })()
    return () => {
      active = false
    }
  }, [sessionId])

  const updatePicture = (
    position: string,
    entry: LoadedHeaderFooterPicture | null,
    replaceImage: boolean,
  ): void => {
    setPictures((prior) => {
      const next = { ...prior }
      if (entry) next[position] = entry
      else delete next[position]
      return next
    })
    setPictureEdits((prior) => ({
      ...prior,
      [position]:
        entry === null
          ? { position, image: null }
          : {
              position,
              ...(replaceImage
                ? { image: { mediaType: entry.mediaType, base64: entry.base64 } }
                : {}),
              widthPt: entry.widthPt,
              heightPt: entry.heightPt,
            },
    }))
  }

  const choosePicture = async (key: SectionKey, file: File | undefined): Promise<void> => {
    if (!file) return
    if (
      !['image/png', 'image/jpeg', 'image/gif'].includes(file.type) ||
      file.size > 2 * 1024 * 1024
    ) {
      setError(t('dlgHfPictureInvalid'))
      return
    }
    try {
      const url = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader()
        reader.onload = () => resolve(String(reader.result))
        reader.onerror = () => reject(reader.error)
        reader.readAsDataURL(file)
      })
      const natural = await new Promise<{ width: number; height: number }>((resolve, reject) => {
        const image = new Image()
        image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight })
        image.onerror = () => reject(new Error(t('dlgHfPictureInvalid')))
        image.src = url
      })
      const position = headerFooterPicturePosition(variant, key)
      const base64 = url.slice(url.indexOf(',') + 1)
      const scale = Math.min(1, 2_000 / Math.max(natural.width * 0.75, natural.height * 0.75))
      const entry = {
        position,
        mediaType: file.type as LoadedHeaderFooterPicture['mediaType'],
        base64,
        widthPt: Math.max(1, Math.round(natural.width * 0.75 * scale * 100) / 100),
        heightPt: Math.max(1, Math.round(natural.height * 0.75 * scale * 100) / 100),
      }
      updatePicture(position, entry, true)
      setPreviews((prior) => ({ ...prior, [position]: url }))
      setValues((prior) => ({
        ...prior,
        [variant]: { ...prior[variant], [key]: withHeaderFooterPictureCode(prior[variant][key]) },
      }))
      setError(null)
    } catch (reason: unknown) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }
  }

  const insertCode = (code: string): void => {
    const key = focusedKey.current
    const input = inputs.current.get(key)
    const value = values[variant][key]
    const start = input?.selectionStart ?? value.length
    const end = input?.selectionEnd ?? value.length
    setValues((prev) => ({
      ...prev,
      [variant]: { ...prev[variant], [key]: value.slice(0, start) + code + value.slice(end) },
    }))
    requestAnimationFrame(() => {
      input?.focus()
      input?.setSelectionRange(start + code.length, start + code.length)
    })
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog header-footer-dialog"
        role="dialog"
        aria-label={t('dlgHfTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgHfTitle')}</header>
        <div className="dialog-tabs" role="tablist">
          {(['odd', 'first', 'even'] as const).map((tab) => (
            <button
              key={tab}
              type="button"
              role="tab"
              aria-selected={variant === tab}
              className={variant === tab ? 'active' : ''}
              onClick={() => setVariant(tab)}
            >
              {t(
                tab === 'odd'
                  ? 'dlgHfOddPages'
                  : tab === 'first'
                    ? 'dlgHfFirstPage'
                    : 'dlgHfEvenPages',
              )}
            </button>
          ))}
        </div>
        {variant !== 'odd' && (
          <label className="dialog-note">
            <input
              type="checkbox"
              checked={enabled[variant]}
              onChange={(event) =>
                setEnabled((prev) => ({ ...prev, [variant]: event.target.checked }))
              }
            />
            {t(variant === 'first' ? 'dlgHfDifferentFirst' : 'dlgHfDifferentEven')}
          </label>
        )}
        <div className="dialog-grid" style={{ gridTemplateColumns: '1fr 1fr 1fr' }}>
          {SECTION_KEYS.map((key) => (
            <div key={key}>
              <label>
                {t(SECTION_LABELS[key])}
                <input
                  value={values[variant][key]}
                  disabled={variant !== 'odd' && !enabled[variant]}
                  ref={(element) => {
                    if (element) inputs.current.set(key, element)
                    else inputs.current.delete(key)
                  }}
                  onFocus={() => {
                    focusedKey.current = key
                  }}
                  onChange={(event) => {
                    const next = event.target.value
                    setValues((prev) => ({
                      ...prev,
                      [variant]: { ...prev[variant], [key]: next },
                    }))
                  }}
                />
              </label>
              {(() => {
                const position = headerFooterPicturePosition(variant, key)
                const picture = pictures[position]
                const disabled = pictureLoading || (variant !== 'odd' && !enabled[variant])
                return (
                  <div className="hf-picture-controls">
                    {picture && previews[position] && (
                      <img src={previews[position]} alt={t('dlgHfPicture')} />
                    )}
                    <label className="secondary hf-picture-pick">
                      {picture ? t('dlgHfPictureReplace') : t('dlgHfPictureInsert')}
                      <input
                        type="file"
                        accept="image/png,image/jpeg,image/gif"
                        disabled={disabled}
                        onChange={(event) => {
                          void choosePicture(key, event.target.files?.[0])
                          event.target.value = ''
                        }}
                      />
                    </label>
                    {picture && (
                      <>
                        <label>
                          {t('dlgHfPictureWidth')}
                          <input
                            type="number"
                            min="1"
                            max="2000"
                            step="0.1"
                            value={picture.widthPt}
                            disabled={disabled}
                            onChange={(event) => {
                              const widthPt = Number(event.target.value)
                              if (widthPt > 0 && widthPt <= 2_000)
                                updatePicture(
                                  position,
                                  { ...picture, widthPt },
                                  (pictureEdits[position]?.image !== undefined &&
                                    pictureEdits[position]?.image !== null) ||
                                    initialPicturesRef.current.some(
                                      (slot) =>
                                        slot.position === position &&
                                        slot.id.startsWith('journal:'),
                                    ),
                                )
                            }}
                          />
                        </label>
                        <label>
                          {t('dlgHfPictureHeight')}
                          <input
                            type="number"
                            min="1"
                            max="2000"
                            step="0.1"
                            value={picture.heightPt}
                            disabled={disabled}
                            onChange={(event) => {
                              const heightPt = Number(event.target.value)
                              if (heightPt > 0 && heightPt <= 2_000)
                                updatePicture(
                                  position,
                                  { ...picture, heightPt },
                                  (pictureEdits[position]?.image !== undefined &&
                                    pictureEdits[position]?.image !== null) ||
                                    initialPicturesRef.current.some(
                                      (slot) =>
                                        slot.position === position &&
                                        slot.id.startsWith('journal:'),
                                    ),
                                )
                            }}
                          />
                        </label>
                        <button
                          type="button"
                          className="secondary"
                          disabled={disabled}
                          onClick={() => {
                            updatePicture(position, null, false)
                            setValues((prior) => ({
                              ...prior,
                              [variant]: {
                                ...prior[variant],
                                [key]: withoutHeaderFooterPictureCode(prior[variant][key]),
                              },
                            }))
                          }}
                        >
                          {t('dlgHfPictureRemove')}
                        </button>
                      </>
                    )}
                  </div>
                )
              })()}
            </div>
          ))}
          <div className="dialog-span" style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
            {FIELD_CODES.map((field) => (
              <button
                key={field.code}
                className="secondary"
                disabled={variant !== 'odd' && !enabled[variant]}
                data-tip={t('dlgHfInsertCode', { code: field.code })}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => insertCode(field.code)}
              >
                {t(field.labelKey)} {field.code}
              </button>
            ))}
            <label>
              {t('dlgHfFontSize')}
              <select
                aria-label={t('dlgHfFontSize')}
                value=""
                disabled={variant !== 'odd' && !enabled[variant]}
                onChange={(event) => insertCode(event.target.value)}
              >
                <option value="">{t('dlgHfFontSize')}</option>
                {[8, 9, 10, 11, 12, 14, 16, 18, 20, 24].map((size) => (
                  <option key={size} value={`&${size}`}>
                    {size} pt
                  </option>
                ))}
              </select>
            </label>
            <label>
              {t('dlgHfFontFamily')}
              <select
                aria-label={t('dlgHfFontFamily')}
                value=""
                disabled={variant !== 'odd' && !enabled[variant]}
                onChange={(event) => insertCode(event.target.value)}
              >
                <option value="">{t('dlgHfFontFamily')}</option>
                {['Arial', 'Calibri', 'Times New Roman'].map((family) => (
                  <option key={family} value={`&"${family},Regular"`}>
                    {family}
                  </option>
                ))}
              </select>
            </label>
          </div>
          <p className="dialog-note dialog-span">
            {pictureLoading ? t('dlgHfPictureLoading') : t('dlgHfNote')}
          </p>
        </div>
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            disabled={pictureLoading || error !== null}
            onClick={() => {
              const odd = toPair(values.odd)
              const failure = onApply({
                sheetId: initialSheetId.current,
                ...odd,
                firstPage: enabled.first ? toPair(values.first) : null,
                evenPages: enabled.even ? toPair(values.even) : null,
                firstPageStored: toPair(values.first),
                evenPagesStored: toPair(values.even),
                pictureEdits: Object.values(pictureEdits),
                loadedPictures: loadedPictures.current,
              })
              setError(failure)
              if (failure === null) onClose()
            }}
          >
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
