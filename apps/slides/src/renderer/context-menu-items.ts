/**
 * Context-menu item builder extracted from App.tsx. Builds the
 * thumbnail/section/canvas/element menus from the latest App state (ActionCtx);
 * item callbacks dispatch into the extracted action modules.
 */
import type { PictureRenderNode, TableRenderNode } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import type { CtxItem } from './components/ContextMenu'
import { isEditableText } from './konva-adapter'
import * as clipboardActions from './clipboard-actions'
import * as slideActions from './slide-actions'
import * as showActions from './show-actions'
import * as arrangeActions from './arrange-actions'
import * as insertActions from './insert-actions'
import * as pictureEditActions from './picture-edit-actions'
import * as styleActions from './style-actions'
import * as tableActions from './table-actions'
import * as tableLayoutActions from './table-layout-actions'
import * as tableSelectActions from './table-select-actions'
import { cellAddrsInRect } from './table-selection'
import { getLang, t } from './i18n/locale'
import { tableRibbonLabels } from './i18n/table-ribbon'
import { TABLE_SHADING_COLORS } from './components/table-shading-colors'

export function buildCtxItems(ctx: ActionCtx): Array<CtxItem | null> {
  const { ctxMenu, slides, sections, selectedIds, slide, current } = ctx
  if (!ctxMenu) return []
  if (ctxMenu.kind === 'thumb') {
    const i = ctxMenu.index
    return [
      { label: t('appCtxNewSlide'), onClick: () => void slideActions.addSlide(ctx) },
      {
        label: t('appCtxDuplicateSlide'),
        onClick: () => void slideActions.duplicateSlideAt(ctx, i),
      },
      {
        label: slides[i]?.hidden ? t('appCtxUnhideSlide') : t('appCtxHideSlide'),
        onClick: () => void showActions.toggleHidden(ctx, i),
      },
      null,
      {
        label: t('appCtxCutSlide'),
        disabled: slides.length <= 1,
        onClick: () => void slideActions.cutSlideAt(ctx, i),
      },
      { label: t('appCtxCopySlide'), onClick: () => void clipboardActions.copySlideAt(ctx, i) },
      {
        label: t('appCtxPasteSlide'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, i),
      },
      {
        label: t('appCtxPasteSlideKeepSource'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, i, 'source'),
      },
      {
        label: t('appCtxPasteSlideAsPicture'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, i, 'picture'),
      },
      null,
      { label: t('appCtxAddSectionBefore'), onClick: () => void slideActions.addSectionAt(ctx, i) },
      null,
      {
        label: t('appCtxChangeBgImage'),
        onClick: () => {
          ctx.setCurrent(i)
          void styleActions.onBackground(ctx, {
            kind: 'image',
            mode: 'stretch',
            pick: true,
            slideIndex: i,
          })
        },
      },
      {
        label: t('appCtxFormatBackground'),
        onClick: () => {
          ctx.setCurrent(i)
          ctx.openBgFormat()
        },
      },
      null,
      {
        label: t('appCtxDeleteSlide'),
        danger: true,
        disabled: slides.length <= 1,
        onClick: () => void slideActions.deleteSlideAt(ctx, i),
      },
    ]
  }
  if (ctxMenu.kind === 'section') {
    const sid = ctxMenu.sectionId
    const si = sections.findIndex((s) => s.id === sid)
    const sec = sections[si]
    return [
      {
        label: t('appCtxRenameSection'),
        onClick: () => ctx.setRenamingSec({ id: sid, value: sec?.name ?? '' }),
      },
      {
        label: t('appCtxRemoveSection'),
        danger: true,
        onClick: () => void slideActions.removeSectionAt(ctx, sid),
      },
      null,
      {
        label: t('appCtxMoveSectionUp'),
        disabled: si <= 0,
        onClick: () => void slideActions.moveSectionDir(ctx, sid, 'up'),
      },
      {
        label: t('appCtxMoveSectionDown'),
        disabled: si < 0 || si >= sections.length - 1,
        onClick: () => void slideActions.moveSectionDir(ctx, sid, 'down'),
      },
    ]
  }
  if (ctxMenu.kind === 'canvas') {
    return [
      {
        label: t('appCtxPaste'),
        hint: '⌘V',
        disabled: !ctx.hasClipboard,
        onClick: () => void clipboardActions.pasteClipboard(ctx),
      },
      null,
      { label: t('appCtxNewSlide'), onClick: () => void slideActions.addSlide(ctx) },
      null,
      { label: t('appCtxFormatBackground'), onClick: () => ctx.openBgFormat() },
    ]
  }
  const node = slide?.nodes.find((n) => n.sourceId === ctxMenu.targetId)
  const single = selectedIds.length === 1
  const cell = ctxMenu.cell
  // Group: multi-select and all are text/shape/picture
  const canGroup =
    selectedIds.length >= 2 &&
    selectedIds.every((id) => {
      const n = slide?.nodes.find((nn) => nn.sourceId === id)
      return n && (n.type === 'text' || n.type === 'shape' || n.type === 'picture')
    })
  // Ungroup: single selection and it's a group
  const canUngroup = single && node?.type === 'group'
  // Table commands act on the selected cell block (right-clicking inside it keeps the block,
  // outside it targets just that cell)
  const ctxTable = cell && node?.type === 'table' ? (node as TableRenderNode) : null
  const range = ctxTable ? tableSelectActions.commandRange(ctx, ctxTable, cell) : null
  const multiCell = !!range && (range.r1 > range.r0 || range.c1 > range.c0)
  const structure = (kind: tableSelectActions.StructureKind) =>
    ctxTable && range
      ? void tableSelectActions.structureForRange(ctx, ctxTable, range, kind)
      : undefined
  const anchorOp = (anchor: 'top' | 'middle' | 'bottom') => {
    if (!ctxTable || !range) return
    const cells = cellAddrsInRect(ctxTable, range)
    void tableLayoutActions.runOps(
      ctx,
      cells.map((c) => ({
        op: 'setTableCellAnchor',
        target: { slide: current, el: ctxMenu.targetId },
        row: c.row,
        col: c.col,
        anchor,
      })),
    )
  }
  const tableItems: Array<CtxItem | null> =
    cell && ctxTable && range
      ? [
          { label: t('appCtxInsertRowAbove'), onClick: () => structure('row-above') },
          { label: t('appCtxInsertRowBelow'), onClick: () => structure('row-below') },
          { label: t('appCtxInsertColLeft'), onClick: () => structure('col-left') },
          { label: t('appCtxInsertColRight'), onClick: () => structure('col-right') },
          { label: t('appCtxDeleteRow'), danger: true, onClick: () => structure('del-row') },
          { label: t('appCtxDeleteCol'), danger: true, onClick: () => structure('del-col') },
          null,
          ...(multiCell
            ? [
                {
                  label: tableRibbonLabels(getLang()).mergeCells,
                  onClick: () => void tableSelectActions.mergeSelectedCells(ctx),
                } as CtxItem,
              ]
            : [
                {
                  label: t('appCtxMergeRight'),
                  onClick: () =>
                    void tableActions.tableMergeOp(
                      ctx,
                      ctxMenu.targetId,
                      'merge-right',
                      cell.row,
                      cell.col,
                    ),
                } as CtxItem,
                {
                  label: t('appCtxMergeDown'),
                  onClick: () =>
                    void tableActions.tableMergeOp(
                      ctx,
                      ctxMenu.targetId,
                      'merge-down',
                      cell.row,
                      cell.col,
                    ),
                } as CtxItem,
              ]),
          { label: t('appCtxCellAnchorTop'), onClick: () => anchorOp('top') },
          { label: t('appCtxCellAnchorMiddle'), onClick: () => anchorOp('middle') },
          { label: t('appCtxCellAnchorBottom'), onClick: () => anchorOp('bottom') },
          {
            label: t('appCtxSplitCell'),
            disabled: !tableSelectActions.rangeHasMerge(ctxTable, range),
            onClick: () =>
              multiCell
                ? void tableSelectActions.splitSelectedCells(ctx)
                : void tableActions.tableMergeOp(
                    ctx,
                    ctxMenu.targetId,
                    'split',
                    cell.row,
                    cell.col,
                  ),
          },
          null,
          {
            label: t('appCtxCellShading'),
            swatches: [...TABLE_SHADING_COLORS, 'none'],
            onSwatch: (c) =>
              void window.slidesApi
                .editTableStyle({
                  slideIndex: current,
                  sourceId: ctxMenu.targetId,
                  shadingColor: c,
                  cells: cellAddrsInRect(ctxTable, range),
                })
                .then((r) => {
                  if (!r) return
                  ctx.applySlide(current, r.slide)
                  if (r.sourceId) {
                    ctx.setSelectedIds([r.sourceId])
                    if (ctx.cellSel) ctx.setCellSel({ ...ctx.cellSel, sourceId: r.sourceId })
                  }
                }),
          },
          null,
        ]
      : []
  return [
    ...tableItems,
    { label: t('appCtxCut'), hint: '⌘X', onClick: () => void clipboardActions.cutSelected(ctx) },
    { label: t('appCtxCopy'), hint: '⌘C', onClick: () => void clipboardActions.copySelected(ctx) },
    {
      label: t('appCtxPaste'),
      hint: '⌘V',
      disabled: !ctx.hasClipboard,
      onClick: () => void clipboardActions.pasteClipboard(ctx),
    },
    null,
    {
      label: t('appCtxBringToFront'),
      onClick: () => void arrangeActions.reorderSelected(ctx, ctxMenu.targetId, 'front'),
    },
    {
      label: t('appCtxSendToBack'),
      onClick: () => void arrangeActions.reorderSelected(ctx, ctxMenu.targetId, 'back'),
    },
    {
      label: t('appCtxBringForward'),
      onClick: () => void arrangeActions.reorderSelected(ctx, ctxMenu.targetId, 'forward'),
    },
    {
      label: t('appCtxSendBackward'),
      onClick: () => void arrangeActions.reorderSelected(ctx, ctxMenu.targetId, 'backward'),
    },
    null,
    // Rotate/flip group (PowerPoint parity) for shapes, pictures and groups; connectors
    // are endpoint-based and keep their existing menu
    ...(node &&
    (node.type === 'shape' || node.type === 'picture' || node.type === 'group') &&
    !(node as { line?: unknown }).line
      ? [
          {
            label: t('appCtxRotateLeft90'),
            onClick: () => void arrangeActions.rotateSelected(ctx, -90),
          } as CtxItem,
          {
            label: t('appCtxRotateRight90'),
            onClick: () => void arrangeActions.rotateSelected(ctx, 90),
          } as CtxItem,
          null,
          {
            label: t('appCtxFlipH'),
            onClick: () => void arrangeActions.flipSelected(ctx, 'h'),
          } as CtxItem,
          {
            label: t('appCtxFlipV'),
            onClick: () => void arrangeActions.flipSelected(ctx, 'v'),
          } as CtxItem,
          null,
        ]
      : []),
    ...(node && isEditableText(node)
      ? [{ label: t('appCtxEditText'), onClick: () => ctx.startEdit(ctxMenu.targetId) } as CtxItem]
      : []),
    // Connectors are endpoint-based (p:cxnSp): swapping their prstGeom would
    // orphan the connection metadata, so they keep their existing menu.
    ...(single && node?.type === 'shape' && !(node as { line?: unknown }).line
      ? [
          {
            label: t('appCtxChangeShape'),
            onClick: () => ctx.openChangeShape(ctxMenu.targetId, ctxMenu.x, ctxMenu.y),
          } as CtxItem,
        ]
      : []),
    ...(single
      ? [
          {
            label: t('appCtxHyperlink'),
            onClick: () => void insertActions.openLinkDialog(ctx),
          } as CtxItem,
        ]
      : []),
    ...(node && node.type === 'picture'
      ? [
          {
            label: t('appCtxCropPicture'),
            onClick: () => pictureEditActions.startCrop(ctx),
          } as CtxItem,
          ...((node as PictureRenderNode).srcRect
            ? [
                {
                  label: t('appCtxResetCrop'),
                  onClick: () => void pictureEditActions.resetCrop(ctx),
                } as CtxItem,
              ]
            : []),
          ...(!(node as PictureRenderNode).media
            ? [
                {
                  label: t('appCtxRemoveBackground'),
                  onClick: () => pictureEditActions.startCutout(ctx),
                } as CtxItem,
              ]
            : []),
        ]
      : []),
    ...(canGroup
      ? [
          {
            label: t('appCtxGroup'),
            hint: '⌘G',
            onClick: () => void arrangeActions.groupSelected(ctx),
          } as CtxItem,
        ]
      : []),
    ...(canUngroup
      ? [
          {
            label: t('appCtxUngroup'),
            hint: '⌘⇧G',
            onClick: () => void arrangeActions.ungroupSelected(ctx),
          } as CtxItem,
        ]
      : []),
    ...(selectedIds.length >= 1
      ? [
          null,
          {
            label: t('appCtxAlignLeft'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'left'),
          } as CtxItem,
          {
            label: t('appCtxAlignCenterH'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'center-h'),
          } as CtxItem,
          {
            label: t('appCtxAlignRight'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'right'),
          } as CtxItem,
          {
            label: t('appCtxAlignTop'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'top'),
          } as CtxItem,
          {
            label: t('appCtxAlignCenterV'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'center-v'),
          } as CtxItem,
          {
            label: t('appCtxAlignBottom'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'bottom'),
          } as CtxItem,
        ]
      : []),
    ...(selectedIds.length >= 3
      ? [
          {
            label: t('appCtxDistributeH'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'distribute-h'),
          } as CtxItem,
          {
            label: t('appCtxDistributeV'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'distribute-v'),
          } as CtxItem,
        ]
      : []),
    null,
    {
      // findNodeCtx also resolves children of the group being edited (node is top-level only)
      label: t('paneFormatTitleTyped', {
        type: nodeTypeName(node ?? ctx.findNodeCtx(ctxMenu.targetId)?.node),
      }),
      onClick: ctx.openFormat,
    },
    null,
    {
      label: t('appCtxDelete'),
      hint: '⌫',
      danger: true,
      onClick: () => void clipboardActions.deleteSelected(ctx),
    },
  ]
}

/** Format-pane type name (same mapping as FormatPane's title) for the context-menu label */
function nodeTypeName(node: { type: string } | undefined): string {
  switch (node?.type) {
    case 'picture':
      return t('paneFormatPicture')
    case 'group':
      return t('paneFormatGroup')
    case 'text':
      return t('paneFormatTextBox')
    case 'table':
      return t('ribbonGroupTable')
    case 'chart':
      return t('ribbonChart')
    default:
      return t('paneFormatShape')
  }
}
