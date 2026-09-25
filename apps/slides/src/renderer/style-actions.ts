/**
 * Element/style actions extracted from App.tsx: fonts, paragraph
 * format, fill/stroke, background, theme, table style, and chart edits.
 * Functions read the latest App state through ActionCtx.
 */
import type { RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import type { TextCaseMode } from '@genoffice/pptx-engine/text-case'
import type {
  EditBackgroundOp,
  EditChartOp,
  EditStrokeOp,
  EditTableStyleOp,
  GradientFillSpec,
} from '../shared/ipc'
import type { ActionCtx } from './action-context'
import { FIT_WIDTH } from './app-constants'
import {
  applySelectionFontFamily,
  applySelectionParagraphFormat,
  changeSelectionCase,
  resizeSelectionFont,
  restoreEditSelection,
  setSelectionFontSizePt,
  setSelectionHighlight,
  setSelectionLetterSpacingPt,
} from './TextEditOverlay'
import type { FormatCmd } from './components/Ribbon'
import { ELEMENT_FORMAT_CMDS } from './components/ribbon-shared'
import type { SlideThemePreset } from './themes'
import { t } from './i18n/locale'

/** Text is being typed — in a slide box, a table cell or the notes pane: commands act on the live selection */
function typing(ctx: ActionCtx): boolean {
  return !!(ctx.editing || ctx.editingCell || ctx.editingNotes)
}

export function onFormat(ctx: ActionCtx, cmd: FormatCmd): void {
  if (typing(ctx) || !ELEMENT_FORMAT_CMDS.has(cmd)) {
    if (cmd === 'fontSizeUp') resizeSelectionFont(1)
    else if (cmd === 'fontSizeDown') resizeSelectionFont(-1)
    else document.execCommand(cmd)
    return
  }
  if (!ctx.selectedIds.length) return
  let patch: { fontSizeStep?: 1 | -1; baseline?: number }
  if (cmd === 'fontSizeUp' || cmd === 'fontSizeDown') {
    patch = { fontSizeStep: cmd === 'fontSizeUp' ? 1 : -1 }
  } else {
    // Toggle like Bold: every run already raised (lowered) → back to normal, else apply to all
    const sign = cmd === 'superscript' ? 1 : -1
    let allOn = true
    for (const id of ctx.selectedIds) {
      const node = ctx.findNodeCtx(id)?.node
      const text =
        node && (node.type === 'text' || node.type === 'shape')
          ? (node as ShapeRenderNode).text
          : undefined
      const runs =
        text?.lines.flatMap((l) => l.runs).filter((r) => !r.isBullet && r.text.trim()) ?? []
      if (!runs.length || runs.some((r) => Math.sign(r.baselinePct ?? 0) !== sign)) allOn = false
    }
    patch = { baseline: allOn ? 0 : cmd === 'superscript' ? 30 : -25 }
  }
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      ...patch,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

// While editing, change the caret selection; with only an element selected, change it wholesale (applies to all the element's text runs)
export function onFontFamily(ctx: ActionCtx, family: string): void {
  if (typing(ctx)) {
    applySelectionFontFamily(family)
    return
  }
  if (!ctx.selectedIds.length) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      fontFamily: family,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

export function onFontSize(ctx: ActionCtx, pt: number): void {
  if (typing(ctx)) {
    setSelectionFontSizePt(pt)
    return
  }
  if (!ctx.selectedIds.length) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      fontSizePt: pt,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

// While editing, change the caret's paragraph; with only an element selected, use the element-level paragraph format op (all paragraphs)
export function onAlign(ctx: ActionCtx, align: 'left' | 'center' | 'right' | 'justify'): void {
  if (typing(ctx)) {
    document.execCommand(
      align === 'left'
        ? 'justifyLeft'
        : align === 'center'
          ? 'justifyCenter'
          : align === 'right'
            ? 'justifyRight'
            : 'justifyFull',
    )
    return
  }
  if (!ctx.selectedIds.length) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementParagraphFormat({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      align,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

// Toggling B/I/U/strikethrough on a selected element without editing: if all runs have it on, turn off, else turn all on
export function onTextToggle(
  ctx: ActionCtx,
  kind: 'bold' | 'italic' | 'underline' | 'strike',
): void {
  if (!ctx.selectedIds.length) return
  let allOn = true
  for (const id of ctx.selectedIds) {
    const node = ctx.findNodeCtx(id)?.node
    const text =
      node && (node.type === 'text' || node.type === 'shape')
        ? (node as ShapeRenderNode).text
        : undefined
    const runs =
      text?.lines.flatMap((l) => l.runs).filter((r) => !r.isBullet && r.text.trim()) ?? []
    if (!runs.length || runs.some((r) => !r[kind])) allOn = false
  }
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      [kind]: !allOn,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

// Change font color directly on a selected element (editing mode goes through execCommand on the selection)
export function onElementTextColor(ctx: ActionCtx, hex: string): void {
  if (!ctx.selectedIds.length) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      color: hex,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

export interface ParagraphFormatPatch {
  bullet?: 'char' | 'number' | 'blip' | 'none'
  bulletChar?: string
  bulletFont?: string
  numType?: string
  startAt?: number
  bulletImage?: { base64: string; ext: string }
  bulletHangEmu?: number
  bulletSizePct?: number
  bulletColor?: string
  lineSpacingPct?: number
  spaceBeforePt?: number
  spaceAfterPt?: number
  rtl?: boolean
  indentDelta?: 1 | -1
}

/** Patch keys the editing-mode selection path can express; anything else stays element-level */
const SELECTION_PATCH_KEYS = new Set([
  'bullet',
  'bulletChar',
  'bulletFont',
  'numType',
  'startAt',
  'bulletImage',
  'bulletHangEmu',
  'bulletSizePct',
  'bulletColor',
  'indentDelta',
  'lineSpacingPct',
  'spaceBeforePt',
  'spaceAfterPt',
  'rtl',
])

// While editing, bullets/numbering/line spacing/paragraph spacing apply to the paragraphs covered
// by the caret/selection (PowerPoint semantics, committed with the edit); with an element selected
// they apply element-wide. Clicking the same bullet kind again = turn off (toggle semantics);
// editing mode judges by the paragraph div's marks, element mode by the render tree's bullet glyphs
export function onParagraphFormat(ctx: ActionCtx, patch: ParagraphFormatPatch): void {
  // Cell editing commits regenerate whole paragraphs from the overlay DOM, which round-trips
  // the rtl mark; the other selection keys keep their historical element-wide semantics there
  const selectable = ctx.editing
    ? Object.keys(patch).every((k) => SELECTION_PATCH_KEYS.has(k))
    : ctx.editingCell != null && Object.keys(patch).every((k) => k === 'rtl')
  if (selectable) {
    const active = document.activeElement
    if (!(active instanceof HTMLElement && active.isContentEditable)) restoreEditSelection()
    if (applySelectionParagraphFormat(patch)) return
  }
  // Typing in a text box may leave selectedIds empty: fall back to the box being edited
  const targetIds = ctx.selectedIds.length
    ? ctx.selectedIds
    : ctx.editing
      ? [ctx.editing.sourceId]
      : []
  if (!targetIds.length) return
  // Picking an explicit glyph / scheme / picture always applies (no toggle-off)
  if (
    patch.bullet &&
    patch.bullet !== 'none' &&
    !patch.bulletChar &&
    !patch.numType &&
    !patch.bulletImage &&
    targetIds.length === 1
  ) {
    const node = ctx.findNodeCtx(targetIds[0]!)?.node
    const text =
      node && (node.type === 'text' || node.type === 'shape')
        ? (node as ShapeRenderNode).text
        : undefined
    const bulletRun = text?.lines.flatMap((l) => l.runs).find((r) => r.isBullet)
    const cur = bulletRun
      ? bulletRun.numType
        ? 'number'
        : bulletRun.image
          ? 'blip'
          : 'char'
      : null
    if (cur === patch.bullet) patch = { ...patch, bullet: 'none' }
  }
  const groupId = ctx.groupIdOf(targetIds[0]!)
  void window.slidesApi
    .setElementParagraphFormat({
      slideIndex: ctx.current,
      sourceIds: targetIds,
      ...patch,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

export async function onFill(
  ctx: ActionCtx,
  sourceId: string,
  fill: string | GradientFillSpec,
): Promise<void> {
  const groupId = ctx.groupIdOf(sourceId)
  const updated = await window.slidesApi.editFill({
    slideIndex: ctx.current,
    sourceId,
    fill,
    ...(groupId ? { groupId } : {}),
  })
  if (updated) ctx.applySlide(ctx.current, updated)
}

export async function onStroke(
  ctx: ActionCtx,
  sourceId: string,
  stroke: EditStrokeOp['stroke'],
): Promise<void> {
  const groupId = ctx.groupIdOf(sourceId)
  const updated = await window.slidesApi.editStroke({
    slideIndex: ctx.current,
    sourceId,
    stroke,
    ...(groupId ? { groupId } : {}),
  })
  if (updated) ctx.applySlide(ctx.current, updated)
}

/** Omit that distributes over union members (plain Omit collapses the EditBackgroundOp union). */
type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never

export async function onBackground(
  ctx: ActionCtx,
  op: DistributiveOmit<EditBackgroundOp, 'fitWidthPx'>,
): Promise<void> {
  if (!ctx.slide) return
  const r = await window.slidesApi.editBackground({
    ...op,
    fitWidthPx: FIT_WIDTH,
  } as EditBackgroundOp)
  if (r) {
    ctx.setSlides(r)
    ctx.setDirty(true)
    ctx.setStatus(op.slideIndex === -1 ? t('appStatusBgAppliedAll') : '')
  }
}

// Apply theme: main process rewrites theme*.xml + per-page backgrounds then reparses, sending back the whole RenderSlide set.
// All element ids change (save→reopen); selection/edit state is cleared too.
export async function applyThemePreset(ctx: ActionCtx, preset: SlideThemePreset): Promise<void> {
  if (!ctx.slide) return
  const r = await window.slidesApi.applyTheme({
    name: preset.name,
    colors: preset.colors,
    ...(preset.majorFont ? { majorFont: preset.majorFont } : {}),
    ...(preset.minorFont ? { minorFont: preset.minorFont } : {}),
    fitWidthPx: FIT_WIDTH,
  })
  if (r && !Array.isArray(r)) {
    ctx.setStatus(t('appStatusThemeApplyFailed', { error: r.error }))
    return
  }
  if (r) {
    ctx.setSlides(r)
    ctx.setSelectedIds([])
    ctx.setEditing(null)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusThemeApplied', { name: preset.name }))
  }
}

/** Table style operations (delegated to IPC) */
export async function onEditTableStyle(
  ctx: ActionCtx,
  op: Omit<EditTableStyleOp, 'slideIndex' | 'sourceId'>,
): Promise<void> {
  if (!ctx.selectedNode || ctx.selectedNode.type !== 'table') return
  const oldId = ctx.selectedNode.sourceId
  const result = await window.slidesApi.editTableStyle({
    ...op,
    slideIndex: ctx.current,
    sourceId: oldId,
  })
  if (result) {
    ctx.applySlide(ctx.current, result.slide)
    // Element ids change after reparse: re-select so the "Table Design" tab doesn't jump away
    if (result.sourceId) {
      ctx.setSelectedIds([result.sourceId])
      if (ctx.editingCell?.sourceId === oldId)
        ctx.setEditingCell({ ...ctx.editingCell, sourceId: result.sourceId })
    }
  }
}

/** Chart edit operations (delegated to IPC) */
export async function onEditChart(
  ctx: ActionCtx,
  op: Omit<EditChartOp, 'slideIndex' | 'sourceId'>,
): Promise<void> {
  if (!ctx.selectedNode || ctx.selectedNode.type !== 'chart') return
  const result = await window.slidesApi.editChart({
    ...op,
    slideIndex: ctx.current,
    sourceId: ctx.selectedNode.sourceId,
  })
  if (result) {
    ctx.applySlide(ctx.current, result.slide)
    // Element ids change after reparse: re-select so the "Chart Design" tab doesn't jump away
    if (result.sourceId) ctx.setSelectedIds([result.sourceId])
  }
}

/** Open the chart data edit dialog */
export async function openChartDataDialog(ctx: ActionCtx): Promise<void> {
  if (!ctx.selectedNode || ctx.selectedNode.type !== 'chart') return
  const data = await window.slidesApi.getChartData(ctx.current, ctx.selectedNode.sourceId)
  if (data) {
    ctx.setChartDataDialogInit(data)
    ctx.setChartDataDialogOpen(true)
  }
}

/** Boxes a text-box command applies to: the selection, or the box being typed in */
function textTargets(ctx: ActionCtx): string[] {
  if (ctx.selectedIds.length) return ctx.selectedIds
  return ctx.editing ? [ctx.editing.sourceId] : []
}

// Home → Font extras (Character Spacing / Text Highlight / Change Case): while editing they act on
// the selection (the word at the caret when nothing is selected), otherwise on every run of the
// selected boxes
export function onFontExtra(
  ctx: ActionCtx,
  patch: { letterSpacingPt?: number; highlight?: string | null; textCase?: TextCaseMode },
): void {
  if (typing(ctx)) {
    const active = document.activeElement
    if (!(active instanceof HTMLElement && active.isContentEditable)) restoreEditSelection()
    if (patch.textCase) changeSelectionCase(patch.textCase)
    if (patch.letterSpacingPt !== undefined) setSelectionLetterSpacingPt(patch.letterSpacingPt)
    if (patch.highlight !== undefined) setSelectionHighlight(patch.highlight)
    return
  }
  if (!ctx.selectedIds.length) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  void window.slidesApi
    .setElementFont({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
      ...patch,
      ...(groupId ? { groupId } : {}),
    })
    .then((r) => r && ctx.applySlide(ctx.current, r))
}

// Home → Paragraph → Align Text / Text Direction: shape-level body properties, so they apply to
// the whole box even while typing in it (the editor follows the re-rendered node)
export async function onTextAnchor(
  ctx: ActionCtx,
  anchor: 'top' | 'middle' | 'bottom',
): Promise<void> {
  let slide: RenderSlide | null = null
  for (const sourceId of textTargets(ctx)) {
    slide =
      (await window.slidesApi.setTextAnchor({ slideIndex: ctx.current, sourceId, anchor })) ?? slide
  }
  if (slide) ctx.applySlide(ctx.current, slide)
}

export async function onTextDirection(
  ctx: ActionCtx,
  vert: 'horz' | 'vert' | 'vert270' | 'wordArtVert',
): Promise<void> {
  let slide: RenderSlide | null = null
  for (const sourceId of textTargets(ctx)) {
    slide =
      (await window.slidesApi.setTextBodyProps({
        slideIndex: ctx.current,
        sourceId,
        props: { vert },
      })) ?? slide
  }
  if (slide) ctx.applySlide(ctx.current, slide)
}
