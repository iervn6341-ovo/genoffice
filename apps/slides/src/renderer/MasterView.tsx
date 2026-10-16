/**
 * Slide master editing view (modeled on PowerPoint View → Slide Master).
 * Full-screen overlay: [master, ...layouts] thumbnails on the left, editable canvas for the
 * current part on the right. All edits go through slides:master-* IPC; App only mounts it and
 * refreshes all slides after exit.
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { RenderFill, RenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import { SlideCanvas } from './SlideCanvas'
import { SlideThumb } from './SlideThumb'
import { TextEditOverlay } from './TextEditOverlay'
import { armColorInput } from './color-input'
import { ContextMenu, type CtxItem } from './components/ContextMenu'
import type { EditParagraph, MasterPartItem } from '../shared/ipc'
import { t } from './i18n/locale'

interface Props {
  initialItems: MasterPartItem[]
  initialPartPath: string | null
  onClose: () => void
}

export function MasterView({ initialItems, initialPartPath, onClose }: Props) {
  const [items, setItems] = useState(initialItems)
  const [sel, setSel] = useState(
    Math.max(
      0,
      initialItems.findIndex((x) => x.partPath === initialPartPath),
    ),
  )
  const [handoutPreview, setHandoutPreview] = useState<'2' | '3' | '4' | '6' | '9' | 'outline'>('6')
  const [selectedIds, setSelectedIds] = useState<string[]>([])
  const [editing, setEditing] = useState<{
    sourceId: string
    caret?: { x: number; y: number }
  } | null>(null)
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; targetId: string | null } | null>(
    null,
  )
  const [images, setImages] = useState<Map<string, HTMLImageElement>>(new Map())
  const colorInputRef = useRef<HTMLInputElement>(null)
  const colorTargetRef = useRef<{
    kind: 'fill' | 'stroke' | 'font' | 'background'
    sourceId?: string
  } | null>(null)

  const slide = items[sel]?.slide
  const selectedNode = slide?.nodes.find((n) => n.sourceId === selectedIds[0])
  const selectedTextNode =
    selectedNode && (selectedNode.type === 'shape' || selectedNode.type === 'text')
      ? selectedNode
      : null
  const firstRun = selectedTextNode?.text?.lines.flatMap((line) => line.runs)[0]
  const selectedStyle = items[sel]?.elementStyles?.[selectedIds[0] ?? '']

  const replaceResult = useCallback(
    (result: { items: MasterPartItem[]; selectedPartPath: string } | null) => {
      if (!result) return
      setItems(result.items)
      setSel(
        Math.max(
          0,
          result.items.findIndex((x) => x.partPath === result.selectedPartPath),
        ),
      )
      setSelectedIds([])
      setEditing(null)
    },
    [],
  )

  const applyCurrent = useCallback(
    (updated: RenderSlide | null) => {
      if (updated)
        setItems((prev) => prev.map((it, i) => (i === sel ? { ...it, slide: updated } : it)))
    },
    [sel],
  )

  // Lazy image loading (picture elements + picture fills + background images; including the decoration layer)
  useEffect(() => {
    const urls = new Set<string>()
    const addFill = (fill: RenderFill | undefined) => {
      if (fill && fill.kind === 'image' && fill.dataUrl) urls.add(fill.dataUrl)
    }
    const walk = (nodes: readonly RenderNode[]) => {
      for (const n of nodes) {
        if (n.type === 'picture' && n.dataUrl) urls.add(n.dataUrl)
        if ((n.type === 'shape' || n.type === 'text') && n.fill) addFill(n.fill)
        if (n.type === 'group' && Array.isArray(n.children)) walk(n.children)
      }
    }
    for (const it of items) {
      addFill(it.slide.background)
      walk(it.slide.nodes)
    }
    urls.forEach((u) => {
      if (images.has(u)) return
      const img = new Image()
      img.onload = () => setImages((prev) => new Map(prev).set(u, img))
      img.src = u
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [items])

  const selectPart = useCallback(
    async (i: number) => {
      if (i === sel) return
      const it = items[i]
      if (!it) return
      setEditing(null)
      setSelectedIds([])
      const rs = await window.slidesApi.masterOpen(it.partPath)
      if (rs) {
        setItems((prev) => prev.map((x, k) => (k === i ? { ...x, slide: rs } : x)))
        setSel(i)
      }
    },
    [sel, items],
  )

  const onTransform = useCallback(
    async (
      sourceId: string,
      box: { x: number; y: number; w: number; h: number; rotationDeg: number },
      preview?: boolean,
    ) => {
      applyCurrent(
        await window.slidesApi.masterEditTransform({
          sourceId,
          xPx: box.x,
          yPx: box.y,
          wPx: box.w,
          hPx: box.h,
          rotationDeg: box.rotationDeg,
          fitWidthPx: 1280,
          ...(preview ? { preview: true } : {}),
        }),
      )
    },
    [applyCurrent],
  )

  const commitEdit = useCallback(
    async (paragraphs: EditParagraph[]) => {
      if (!editing) return
      applyCurrent(
        await window.slidesApi.masterEditText({ sourceId: editing.sourceId, paragraphs }),
      )
      setEditing(null)
      setSelectedIds([editing.sourceId])
    },
    [editing, applyCurrent],
  )

  const deleteSelected = useCallback(async () => {
    for (const id of selectedIds) {
      applyCurrent(await window.slidesApi.masterDeleteElement({ sourceId: id }))
    }
    setSelectedIds([])
  }, [selectedIds, applyCurrent])

  const pickColor = useCallback(
    (kind: 'fill' | 'stroke' | 'font' | 'background', sourceId?: string) => {
      colorTargetRef.current = { kind, sourceId }
      if (colorInputRef.current) {
        armColorInput(colorInputRef.current)
        colorInputRef.current.click()
      }
    },
    [],
  )

  const onColorPicked = useCallback(
    async (hex: string) => {
      const target = colorTargetRef.current
      if (!target) return
      colorTargetRef.current = null
      if (target.kind === 'background')
        applyCurrent(await window.slidesApi.masterEditBackground(hex))
      else if (target.sourceId)
        applyCurrent(
          target.kind === 'fill'
            ? await window.slidesApi.masterEditFill({ sourceId: target.sourceId, fill: hex })
            : target.kind === 'font'
              ? await window.slidesApi.masterEditFont({
                  sourceId: target.sourceId,
                  font: { color: hex },
                })
              : await window.slidesApi.masterEditStroke({
                  sourceId: target.sourceId,
                  stroke: { color: hex, widthPt: 1 },
                }),
        )
    },
    [applyCurrent],
  )

  // Delete/Backspace deletes selected elements (not intercepted while editing)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (editing || !selectedIds.length) return
      const el = e.target as HTMLElement | null
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable))
        return
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault()
        void deleteSelected()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [editing, selectedIds, deleteSelected])

  const editNode = useMemo(() => {
    if (!editing || !slide) return null
    return (
      (slide.nodes.find((n) => n.sourceId === editing.sourceId) as ShapeRenderNode | undefined) ??
      null
    )
  }, [editing, slide])

  const ctxItems = useMemo<Array<CtxItem | null>>(() => {
    if (!ctxMenu?.targetId) return []
    const id = ctxMenu.targetId
    const node = slide?.nodes.find((n) => n.sourceId === id)
    const textual = node && (node.type === 'text' || node.type === 'shape')
    return [
      ...(textual && !items[sel]?.nativeFieldIds?.includes(id)
        ? [{ label: t('appCtxEditText'), onClick: () => setEditing({ sourceId: id }) } as CtxItem]
        : []),
      ...(textual
        ? [
            { label: t('appMasterFill'), onClick: () => pickColor('fill', id) },
            {
              label: t('appMasterFillNone'),
              onClick: () =>
                void window.slidesApi
                  .masterEditFill({ sourceId: id, fill: 'none' })
                  .then(applyCurrent),
            },
            { label: t('appMasterStroke'), onClick: () => pickColor('stroke', id) },
            {
              label: t('appMasterStrokeNone'),
              onClick: () =>
                void window.slidesApi
                  .masterEditStroke({ sourceId: id, stroke: null })
                  .then(applyCurrent),
            },
          ]
        : []),
      null,
      { label: t('appCtxDelete'), hint: '⌫', danger: true, onClick: () => void deleteSelected() },
    ]
  }, [ctxMenu, slide, items, sel, pickColor, applyCurrent, deleteSelected])

  if (!slide) return null

  const current = items[sel]!
  const notesOrHandout = current.kind === 'notesMaster' || current.kind === 'handoutMaster'
  const title =
    current.kind === 'notesMaster'
      ? t('appNotesMasterTitle')
      : current.kind === 'handoutMaster'
        ? t('appHandoutMasterTitle')
        : t('appMasterTitle')
  const formatSelected = async (
    font: Parameters<typeof window.slidesApi.masterEditFont>[0]['font'],
  ) => {
    const id = selectedIds[0]
    if (!id) return
    const updated = await window.slidesApi.masterEditFont({ sourceId: id, font })
    applyCurrent(updated)
    if (updated)
      setItems((prev) =>
        prev.map((it, i) =>
          i === sel
            ? {
                ...it,
                elementStyles: {
                  ...it.elementStyles,
                  [id]: { ...it.elementStyles?.[id], ...font },
                },
              }
            : it,
        ),
      )
  }
  const paragraphSelected = async (
    format: Parameters<typeof window.slidesApi.masterEditParagraph>[0]['format'],
  ) => {
    const id = selectedIds[0]
    if (!id) return
    const updated = await window.slidesApi.masterEditParagraph({ sourceId: id, format })
    applyCurrent(updated)
    if (updated)
      setItems((prev) =>
        prev.map((it, i) =>
          i === sel
            ? {
                ...it,
                elementStyles: {
                  ...it.elementStyles,
                  [id]: { ...it.elementStyles?.[id], ...format },
                },
              }
            : it,
        ),
      )
  }

  return (
    <div className="master-view">
      <div className="master-bar">
        <span className="master-bar-title">{title}</span>
        <button
          className="master-tool"
          onClick={() => void window.slidesApi.masterUndo().then(replaceResult)}
          data-testid="master-undo"
        >
          {t('ribbonUndo')}
        </button>
        <button
          className="master-tool"
          onClick={() => void window.slidesApi.masterRedo().then(replaceResult)}
          data-testid="master-redo"
        >
          {t('ribbonRedo')}
        </button>
        <button className="master-close-btn" onClick={onClose}>
          {t('appMasterClose')}
        </button>
      </div>
      <div className="master-format-bar">
        <input
          aria-label={t('appMasterFontFamily')}
          data-testid="master-font-family"
          key={`font-${selectedIds[0]}-${selectedStyle?.fontFamily}`}
          defaultValue={
            selectedStyle?.fontFamily ?? firstRun?.srcFontFamily ?? firstRun?.fontFamily ?? ''
          }
          disabled={!selectedTextNode}
          onKeyDown={(e) => {
            if (e.key === 'Enter') e.currentTarget.blur()
          }}
          onBlur={(e) => {
            const value = e.currentTarget.value.trim()
            if (value && value !== selectedStyle?.fontFamily)
              void formatSelected({ fontFamily: value })
          }}
        />
        <input
          aria-label={t('ribbonFontSizeTip')}
          data-testid="master-font-size"
          type="number"
          min="1"
          max="400"
          step="1"
          key={`size-${selectedIds[0]}-${selectedStyle?.fontSizePt}`}
          defaultValue={
            selectedStyle?.fontSizePt ??
            (firstRun ? Math.round((firstRun.fontSizePx / slide.scale) * 0.75) : 18)
          }
          disabled={!selectedTextNode}
          onKeyDown={(e) => {
            if (e.key === 'Enter') e.currentTarget.blur()
          }}
          onBlur={(e) => {
            const value = Number(e.currentTarget.value)
            if (value > 0 && value <= 400 && value !== selectedStyle?.fontSizePt)
              void formatSelected({ fontSizePt: value })
          }}
        />
        <button
          disabled={!selectedTextNode}
          aria-label={t('ribbonBold')}
          aria-pressed={!!(selectedStyle?.bold ?? firstRun?.bold)}
          onClick={() => void formatSelected({ bold: !(selectedStyle?.bold ?? firstRun?.bold) })}
        >
          <b>B</b>
        </button>
        <button
          disabled={!selectedTextNode}
          aria-label={t('ribbonItalic')}
          aria-pressed={!!(selectedStyle?.italic ?? firstRun?.italic)}
          onClick={() =>
            void formatSelected({ italic: !(selectedStyle?.italic ?? firstRun?.italic) })
          }
        >
          <i>I</i>
        </button>
        <button
          disabled={!selectedTextNode}
          aria-label={t('ribbonUnderline')}
          aria-pressed={!!(selectedStyle?.underline ?? firstRun?.underline)}
          onClick={() =>
            void formatSelected({ underline: !(selectedStyle?.underline ?? firstRun?.underline) })
          }
        >
          <u>U</u>
        </button>
        <button
          disabled={!selectedTextNode}
          aria-label={t('ribbonFontColor')}
          onClick={() => pickColor('font', selectedIds[0])}
        >
          A ▾
        </button>
        <select
          aria-label={t('ribbonAlignText')}
          data-testid="master-paragraph-align"
          disabled={!selectedTextNode}
          onChange={(e) =>
            void paragraphSelected({
              align: e.currentTarget.value as 'left' | 'center' | 'right' | 'justify',
            })
          }
          value={selectedStyle?.align ?? 'left'}
        >
          <option value="left">{t('ribbonAlignLeft')}</option>
          <option value="center">{t('ribbonAlignCenter')}</option>
          <option value="right">{t('ribbonAlignRight')}</option>
          <option value="justify">{t('ribbonAlignJustify')}</option>
        </select>
        <input
          aria-label={t('ribbonLineSpacing')}
          data-testid="master-line-spacing"
          type="number"
          min="50"
          max="400"
          step="5"
          key={`line-${selectedIds[0]}-${selectedStyle?.lineSpacingPct}`}
          defaultValue={selectedStyle?.lineSpacingPct ?? 100}
          disabled={!selectedTextNode}
          onKeyDown={(e) => {
            if (e.key === 'Enter') e.currentTarget.blur()
          }}
          onBlur={(e) => {
            const value = Number(e.currentTarget.value)
            if (value >= 50 && value <= 400 && value !== selectedStyle?.lineSpacingPct)
              void paragraphSelected({ lineSpacingPct: value })
          }}
        />
        <button onClick={() => pickColor('background')} data-testid="master-background">
          {t('appMasterBackground')}
        </button>
        {notesOrHandout && (
          <select
            aria-label={t('appPrintOrientation')}
            data-testid="master-page-orientation"
            value={current.pageOrientation ?? 'portrait'}
            onChange={(e) =>
              void window.slidesApi
                .masterSetPageOrientation(e.currentTarget.value as 'portrait' | 'landscape')
                .then(replaceResult)
            }
          >
            <option value="portrait">{t('appPrintPortrait')}</option>
            <option value="landscape">{t('appPrintLandscape')}</option>
          </select>
        )}
        {current.kind === 'handoutMaster' && (
          <select
            aria-label={t('appMasterPreviewOnly')}
            data-testid="handout-master-preview-layout"
            value={handoutPreview}
            onChange={(e) => setHandoutPreview(e.currentTarget.value as typeof handoutPreview)}
          >
            <option value="2">2 {t('appPrintSlidesLabel')}</option>
            <option value="3">3 {t('appPrintSlidesLabel')}</option>
            <option value="4">4 {t('appPrintSlidesLabel')}</option>
            <option value="6">6 {t('appPrintSlidesLabel')}</option>
            <option value="9">9 {t('appPrintSlidesLabel')}</option>
            <option value="outline">{t('appMasterOutline')}</option>
          </select>
        )}
        {notesOrHandout &&
          current.headerFooterFlags &&
          (['hdr', 'ftr', 'dt', 'sldNum'] as const).map((kind) => (
            <label key={kind} className="master-flag">
              <input
                type="checkbox"
                data-testid={`master-flag-${kind}`}
                checked={current.headerFooterFlags![kind]}
                onChange={(e) =>
                  void window.slidesApi
                    .masterSetHeaderFooterFlag(kind, e.currentTarget.checked)
                    .then(replaceResult)
                }
              />
              {t(
                kind === 'hdr'
                  ? 'ribbonDlgHeader'
                  : kind === 'ftr'
                    ? 'ribbonDlgFooter'
                    : kind === 'dt'
                      ? 'ribbonDlgDateTime'
                      : 'ribbonDlgPageNum',
              )}
            </label>
          ))}
        {current.kind === 'notesMaster' &&
          current.notesPlaceholders &&
          (['sldImg', 'body'] as const).map((kind) => (
            <label key={kind} className="master-flag">
              <input
                type="checkbox"
                data-testid={`master-placeholder-${kind}`}
                checked={current.notesPlaceholders![kind]}
                onChange={(e) =>
                  void window.slidesApi
                    .masterSetNotesPlaceholder(kind, e.currentTarget.checked)
                    .then(replaceResult)
                }
              />
              {t(kind === 'sldImg' ? 'appMasterSlideImage' : 'appMasterNotesBody')}
            </label>
          ))}
      </div>
      <div className="master-body">
        <div className="slide-list master-list">
          {items.map((it, i) => (
            <div
              key={it.partPath}
              className={`thumb master-thumb ${i === sel ? 'active' : ''} ${it.kind === 'layout' ? 'master-thumb-layout' : ''}`}
              onClick={() => void selectPart(i)}
              data-tip={
                it.kind === 'notesMaster'
                  ? t('appNotesMasterTitle')
                  : it.kind === 'handoutMaster'
                    ? t('appHandoutMasterTitle')
                    : it.kind === 'master'
                      ? t('appMasterThumbMaster')
                      : it.name
              }
            >
              <SlideThumb
                slide={it.slide}
                images={images}
                width={it.kind === 'layout' ? 108 : 126}
              />
              <span className="master-thumb-label">
                {it.kind === 'notesMaster'
                  ? t('appNotesMasterTitle')
                  : it.kind === 'handoutMaster'
                    ? t('appHandoutMasterTitle')
                    : it.kind === 'master'
                      ? t('appMasterThumbMaster')
                      : it.name || t('appMasterThumbLayout', { n: i })}
              </span>
            </div>
          ))}
        </div>
        <div className="master-stage">
          <div
            className="stage-rel"
            style={{ position: 'relative', width: slide.widthPx, height: slide.heightPx }}
          >
            <SlideCanvas
              slide={slide}
              selectedIds={selectedIds}
              onSelect={(id, additive) =>
                setSelectedIds((prev) => {
                  if (!id) return []
                  if (additive)
                    return prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
                  return [id]
                })
              }
              onEditText={(sourceId, caret) => {
                if (!current.nativeFieldIds?.includes(sourceId)) setEditing({ sourceId, caret })
              }}
              onTransform={onTransform}
              onEditTableCell={() => {}}
              onTableColResize={() => {}}
              onContextMenu={(sourceId, x, y) => {
                if (editing) return
                if (sourceId)
                  setSelectedIds((prev) => (prev.includes(sourceId) ? prev : [sourceId]))
                setCtxMenu({ x, y, targetId: sourceId })
              }}
              onMarqueeSelect={setSelectedIds}
              images={images}
              editingText={editing ? { sourceId: editing.sourceId } : null}
            />
            {current.kind === 'handoutMaster' && (
              <div
                className={`master-handout-preview master-handout-${handoutPreview}`}
                aria-hidden="true"
              >
                {handoutPreview === 'outline'
                  ? Array.from({ length: 8 }, (_, i) => (
                      <span key={i} className="master-outline-line" />
                    ))
                  : Array.from({ length: Number(handoutPreview) }, (_, i) => (
                      <span key={i} className="master-slide-slot" />
                    ))}
              </div>
            )}
            {editing && editNode && (
              <TextEditOverlay
                node={editNode}
                scale={slide.scale}
                caretPoint={editing.caret}
                onCommit={commitEdit}
                onCancel={() => setEditing(null)}
              />
            )}
          </div>
        </div>
      </div>
      <input
        ref={colorInputRef}
        type="color"
        style={{ position: 'fixed', left: -9999, top: -9999 }}
        onChange={(e) => void onColorPicked(e.target.value.toUpperCase())}
      />
      {ctxMenu && (
        <ContextMenu
          x={ctxMenu.x}
          y={ctxMenu.y}
          items={ctxItems}
          onClose={() => setCtxMenu(null)}
        />
      )}
    </div>
  )
}
