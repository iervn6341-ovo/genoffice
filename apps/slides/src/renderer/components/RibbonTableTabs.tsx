/**
 * The two contextual tabs of a selected table, laid out like PowerPoint's:
 *   Table Design  — style options (6 toggles) · style gallery · shading / borders · pen
 *   Table Layout  — delete / insert · merge · cell size · alignment · table size
 */
import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { ColorPicker, CraftIcon, useDismissablePopover, type CraftSet } from '@genoffice/ui'
import { DECK_ICONS } from '@genoffice/ui/craft-icons/deck'
import type { StyleGalleryEntry, StylePreviewCell } from '@genoffice/pptx-engine'
import type { EditTableStyleOp } from '../../shared/ipc'
import { armColorInput } from '../color-input'
import { useI18n } from '../i18n/locale'
import { localizedTableStyleName, tableRibbonLabels } from '../i18n/table-ribbon'
import { Group } from './ribbon-shared'

const DECK: CraftSet = { grid: 24, icons: DECK_ICONS }
const BIG = 28
const SMALL = 16

export type TableFlagKey = 'firstRow' | 'lastRow' | 'firstCol' | 'lastCol' | 'bandRow' | 'bandCol'

export interface TableTabInfo {
  rows: number
  cols: number
  rowHeightsPx: number[]
  colWidthsPx: number[]
  /** layout px per centimetre (96 dpi × the slide's zoom) */
  pxPerCm: number
  styleId?: string
  flags: Record<TableFlagKey, boolean> & { rtl: boolean }
  /** the cell being edited; structure / merge / alignment act on it */
  activeCell: { row: number; col: number } | null
}

export interface TableTabActions {
  editStyle: (op: Omit<EditTableStyleOp, 'slideIndex' | 'sourceId'>) => void
  structure: (
    kind: 'insert-row' | 'delete-row' | 'insert-col' | 'delete-col',
    index: number,
    before?: boolean,
  ) => void
  merge: (kind: 'merge-right' | 'merge-down' | 'split', row: number, col: number) => void
  setRowHeightPx: (row: number, hPx: number) => void
  setColWidthPx: (col: number, wPx: number) => void
  distribute: (axis: 'rows' | 'cols') => void
  alignH: (align: 'left' | 'center' | 'right') => void
  alignV: (anchor: 'top' | 'middle' | 'bottom') => void
  /** scale every row (h) or column (w) so the table reaches this size */
  resizeTable: (axis: 'h' | 'w', px: number) => void
  deleteTable: () => void
  toggleRtl: (rtl: boolean) => void
}

// ── shared bits ─────────────────────────────────────────────────────────

/** A trigger + popover pair that closes on an outside press or Escape. */
function Pop({
  trigger,
  children,
}: {
  trigger: (open: boolean, toggle: () => void) => ReactNode
  children: (close: () => void) => ReactNode
}) {
  const [open, setOpen] = useState(false)
  const wrap = useRef<HTMLDivElement>(null)
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrap.current] })
  return (
    <div className="rb-drop-wrap" ref={wrap} onMouseDown={(e) => e.stopPropagation()}>
      {trigger(open, () => setOpen((v) => !v))}
      {open && <div className="rb-drop rb-tbl-drop">{children(() => setOpen(false))}</div>}
    </div>
  )
}

function Caret() {
  return (
    <svg className="rb-tbl-caret" width="8" height="8" viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M5.5 9.25 12 15.75l6.5-6.5"
        fill="none"
        stroke="currentColor"
        strokeWidth="3"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function BigBtn({
  icon,
  label,
  onClick,
  disabled,
  caret,
  active,
}: {
  icon: string
  label: string
  onClick: () => void
  disabled?: boolean
  caret?: boolean
  active?: boolean
}) {
  return (
    <button
      type="button"
      className={`rb-big${active ? ' active' : ''}`}
      data-tip={label}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      <span className="rb-big-icon">
        <CraftIcon set={DECK} name={icon} size={BIG} muted={disabled} />
      </span>
      <span className="rb-tbl-biglabel">
        {label}
        {caret && <Caret />}
      </span>
    </button>
  )
}

function SmallBtn({
  icon,
  label,
  onClick,
  disabled,
  active,
  iconOnly,
}: {
  icon: string
  label: string
  onClick: () => void
  disabled?: boolean
  active?: boolean
  /** label only as the tooltip (keeps the tab inside a narrow window) */
  iconOnly?: boolean
}) {
  return (
    <button
      type="button"
      className={`rb-tbl-small${active ? ' active' : ''}`}
      data-tip={label}
      aria-label={label}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      <CraftIcon set={DECK} name={icon} size={SMALL} muted={disabled} />
      {!iconOnly && <span>{label}</span>}
    </button>
  )
}

// ── gallery thumbnails ──────────────────────────────────────────────────

const THUMB_W = 52
const THUMB_H = 38

function StyleThumb({ cells }: { cells: StylePreviewCell[][] }) {
  const rows = cells.length
  const cols = cells[0]?.length ?? 0
  const cw = THUMB_W / cols
  const ch = THUMB_H / rows
  return (
    <svg
      width={THUMB_W}
      height={THUMB_H}
      viewBox={`0 0 ${THUMB_W} ${THUMB_H}`}
      className="rb-tbl-thumb-svg"
      aria-hidden="true"
    >
      {cells.flatMap((row, r) =>
        row.map((c, ci) => {
          const x = ci * cw
          const y = r * ch
          return (
            <g key={`${r}-${ci}`}>
              <rect x={x} y={y} width={cw} height={ch} fill={c.fill ?? 'none'} />
              {/* the text stand-in: a short bar in the part's text color */}
              <rect
                x={x + cw * 0.22}
                y={y + ch * 0.4}
                width={cw * 0.56}
                height={Math.max(1, ch * 0.2)}
                rx={0.5}
                fill={c.text ?? '#000000'}
                opacity={c.fill || c.text ? 0.55 : 0.3}
              />
              {c.t && <line x1={x} x2={x + cw} y1={y} y2={y} stroke={c.t} strokeWidth="0.8" />}
              {c.b && (
                <line x1={x} x2={x + cw} y1={y + ch} y2={y + ch} stroke={c.b} strokeWidth="0.8" />
              )}
              {c.l && <line x1={x} x2={x} y1={y} y2={y + ch} stroke={c.l} strokeWidth="0.8" />}
              {c.r && (
                <line x1={x + cw} x2={x + cw} y1={y} y2={y + ch} stroke={c.r} strokeWidth="0.8" />
              )}
            </g>
          )
        }),
      )}
    </svg>
  )
}

const VISIBLE = 7

function StyleGallery({
  info,
  actions,
  disabled,
}: {
  info: TableTabInfo
  actions: TableTabActions
  disabled: boolean
}) {
  const { lang } = useI18n()
  const l = tableRibbonLabels(lang)
  const [gallery, setGallery] = useState<StyleGalleryEntry[]>([])
  const [offset, setOffset] = useState(0)
  const f = info.flags
  const key = `${f.firstRow}${f.lastRow}${f.firstCol}${f.lastCol}${f.bandRow}${f.bandCol}`
  useEffect(() => {
    let stale = false
    void window.slidesApi
      .getTableStyleGallery?.({
        firstRow: f.firstRow,
        lastRow: f.lastRow,
        firstCol: f.firstCol,
        lastCol: f.lastCol,
        bandRow: f.bandRow,
        bandCol: f.bandCol,
      })
      .then((g) => {
        if (!stale && g) setGallery(g)
      })
    return () => {
      stale = true
    }
    // the flags are folded into `key`
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key])

  // PowerPoint's gallery order: the styles of the deck's accent first (Medium Style 2 …)
  const ordered = useMemo(() => {
    const rank = (g: StyleGalleryEntry): number =>
      g.family === 'medium2' ? 0 : g.family === 'themed1' || g.family === 'themed2' ? 1 : 2
    // within a family the accent variants lead (Medium Style 2 - Accent 1 is the gallery's first tile)
    const plain = (g: StyleGalleryEntry): number => (g.accent ? 0 : 1)
    return [...gallery].sort((a, b) => rank(a) - rank(b) || plain(a) - plain(b))
  }, [gallery])
  const selected = (info.styleId ?? '').toUpperCase()
  const apply = (g: StyleGalleryEntry) => actions.editStyle({ styleId: g.id })
  const thumb = (g: StyleGalleryEntry) => {
    const name = localizedTableStyleName(l, g)
    return (
      <button
        key={g.id}
        type="button"
        className={`rb-tbl-thumb${g.id.toUpperCase() === selected ? ' selected' : ''}`}
        data-tip={name}
        aria-label={name}
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => apply(g)}
      >
        <StyleThumb cells={g.cells} />
      </button>
    )
  }
  const max = Math.max(0, ordered.length - VISIBLE)
  const groups: Array<{ title: string; items: StyleGalleryEntry[] }> = [
    { title: l.famMedium, items: ordered.filter((g) => g.family.startsWith('medium')) },
    { title: l.famLight, items: ordered.filter((g) => g.family.startsWith('light')) },
    { title: l.famDark, items: ordered.filter((g) => g.family.startsWith('dark')) },
    { title: l.noStyleNoGrid, items: ordered.filter((g) => g.family.startsWith('themed')) },
  ]
  return (
    <div className="rb-tbl-gallery">
      <button
        type="button"
        className="rb-tbl-arrow"
        aria-label="‹"
        disabled={offset <= 0}
        onClick={() => setOffset((o) => Math.max(0, o - VISIBLE))}
      >
        ‹
      </button>
      <div className="rb-tbl-strip">{ordered.slice(offset, offset + VISIBLE).map(thumb)}</div>
      <button
        type="button"
        className="rb-tbl-arrow"
        aria-label="›"
        disabled={offset >= max}
        onClick={() => setOffset((o) => Math.min(max, o + VISIBLE))}
      >
        ›
      </button>
      <Pop
        trigger={(open, toggle) => (
          <button
            type="button"
            className={`rb-tbl-arrow${open ? ' active' : ''}`}
            data-tip={l.moreStyles}
            aria-label={l.moreStyles}
            disabled={disabled}
            onClick={toggle}
          >
            <Caret />
          </button>
        )}
      >
        {(close) => (
          <div className="rb-tbl-allstyles" onClick={close}>
            {groups
              .filter((g) => g.items.length > 0)
              .map((g) => (
                <section key={g.title}>
                  <div className="rb-drop-title">{g.title}</div>
                  <div className="rb-tbl-allgrid">{g.items.map(thumb)}</div>
                </section>
              ))}
          </div>
        )}
      </Pop>
    </div>
  )
}

// ── Table Design ────────────────────────────────────────────────────────

const DASHES = ['solid', 'dash', 'sysDot', 'dashDot', 'lgDash'] as const
const WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]

function DashPreview({ dash }: { dash: string }) {
  const arr =
    dash === 'dash'
      ? '6 3'
      : dash === 'sysDot'
        ? '1.5 2.5'
        : dash === 'dashDot'
          ? '6 3 1.5 3'
          : dash === 'lgDash'
            ? '9 3'
            : undefined
  return (
    <svg width="64" height="8" viewBox="0 0 64 8" aria-hidden="true">
      <line
        x1="2"
        x2="62"
        y1="4"
        y2="4"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeDasharray={arr}
      />
    </svg>
  )
}

export function TableDesignTab({
  info,
  actions,
}: {
  info: TableTabInfo
  actions: TableTabActions
}) {
  const { t, lang } = useI18n()
  const l = tableRibbonLabels(lang)
  const [pen, setPen] = useState<{ dash: string; weight: number; color: string }>({
    dash: 'solid',
    weight: 1,
    color: '#000000',
  })
  const f = info.flags
  const check = (key: TableFlagKey, label: string) => (
    <label className="rb-tbl-check">
      <input
        type="checkbox"
        checked={f[key]}
        onChange={(e) => actions.editStyle({ [key]: e.target.checked })}
      />
      <span>{label}</span>
    </label>
  )
  const cells = info.activeCell ? [info.activeCell] : undefined
  const border = (
    borderMode: NonNullable<EditTableStyleOp['borderMode']>,
    label: string,
    close: () => void,
  ) => (
    <button
      key={borderMode}
      type="button"
      className="rb-tbl-menuitem"
      onClick={() => {
        close()
        actions.editStyle({
          borderMode,
          borderColor: pen.color,
          borderWidthPt: pen.weight,
          borderDash: pen.dash,
          ...(cells ? { cells } : {}),
        })
      }}
    >
      {label}
    </button>
  )
  return (
    <>
      <Group label={t('ribbonGroupTableOptions')}>
        <div className="rb-tbl-options">
          {check('firstRow', t('ribbonHeaderRow'))}
          {check('firstCol', l.firstColumn)}
          {check('lastRow', l.totalRow)}
          {check('lastCol', l.lastColumn)}
          {check('bandRow', t('ribbonBandedRows'))}
          {check('bandCol', l.bandedColumns)}
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupTableStyles')}>
        <StyleGallery info={info} actions={actions} disabled={false} />
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupShading')}>
        <div className="rb-tbl-stack">
          <Pop
            trigger={(open, toggle) => (
              <button
                type="button"
                className={`rb-tbl-wide${open ? ' active' : ''}`}
                data-tip={info.activeCell ? t('ribbonGroupShadingCell') : t('ribbonGroupShading')}
                onMouseDown={(e) => e.preventDefault()}
                onClick={toggle}
              >
                <CraftIcon set={DECK} name="shading" size={SMALL} />
                <span>{t('ribbonGroupShading')}</span>
                <Caret />
              </button>
            )}
          >
            {(close) => (
              <ColorPicker
                className="rb-tbl-colorpop"
                value="#ffffff"
                strings={{
                  themeColors: t('ribbonThemeColorsSection'),
                  standardColors: t('ribbonStandardColors'),
                  moreColors: t('ribbonMoreColors'),
                }}
                onPick={(hex) => {
                  close()
                  actions.editStyle({
                    shadingColor: hex ?? 'none',
                    ...(cells ? { cells } : {}),
                  })
                }}
                moreInputProps={{
                  onPointerDown: (e) => armColorInput(e.currentTarget),
                  onChange: (e) =>
                    actions.editStyle({
                      shadingColor: e.currentTarget.value,
                      ...(cells ? { cells } : {}),
                    }),
                }}
              />
            )}
          </Pop>
          <Pop
            trigger={(open, toggle) => (
              <button
                type="button"
                className={`rb-tbl-wide${open ? ' active' : ''}`}
                data-tip={t('ribbonGroupBorders')}
                onMouseDown={(e) => e.preventDefault()}
                onClick={toggle}
              >
                <CraftIcon set={DECK} name="borders" size={SMALL} />
                <span>{t('ribbonGroupBorders')}</span>
                <Caret />
              </button>
            )}
          >
            {(close) => (
              <div className="rb-tbl-menu">
                {border('bottom', l.bBottom, close)}
                {border('top', l.bTop, close)}
                {border('left', l.bLeft, close)}
                {border('right', l.bRight, close)}
                <div className="row-menu-divider" />
                {border('none', l.bNone, close)}
                {border('all', l.bAll, close)}
                {border('outside', l.bOutside, close)}
                {border('inside', l.bInside, close)}
                {border('insideH', l.bInsideH, close)}
                {border('insideV', l.bInsideV, close)}
                <div className="row-menu-divider" />
                {border('diagDown', l.bDiagDown, close)}
                {border('diagUp', l.bDiagUp, close)}
              </div>
            )}
          </Pop>
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={l.grpDrawBorders}>
        <div className="rb-tbl-stack">
          <Pop
            trigger={(open, toggle) => (
              <button
                type="button"
                className={`rb-tbl-wide${open ? ' active' : ''}`}
                data-tip={l.penStyle}
                onClick={toggle}
              >
                <DashPreview dash={pen.dash} />
                <Caret />
              </button>
            )}
          >
            {(close) => (
              <div className="rb-tbl-menu">
                {DASHES.map((d) => (
                  <button
                    key={d}
                    type="button"
                    className={`rb-tbl-menuitem${pen.dash === d ? ' active' : ''}`}
                    onClick={() => {
                      close()
                      setPen((p) => ({ ...p, dash: d }))
                    }}
                  >
                    <DashPreview dash={d} />
                  </button>
                ))}
              </div>
            )}
          </Pop>
          <Pop
            trigger={(open, toggle) => (
              <button
                type="button"
                className={`rb-tbl-wide${open ? ' active' : ''}`}
                data-tip={l.penWeight}
                onClick={toggle}
              >
                <span>{pen.weight} pt</span>
                <Caret />
              </button>
            )}
          >
            {(close) => (
              <div className="rb-tbl-menu">
                {WEIGHTS.map((w) => (
                  <button
                    key={w}
                    type="button"
                    className={`rb-tbl-menuitem${pen.weight === w ? ' active' : ''}`}
                    onClick={() => {
                      close()
                      setPen((p) => ({ ...p, weight: w }))
                    }}
                  >
                    {w} pt
                  </button>
                ))}
              </div>
            )}
          </Pop>
          <Pop
            trigger={(open, toggle) => (
              <button
                type="button"
                className={`rb-tbl-wide${open ? ' active' : ''}`}
                data-tip={l.penColor}
                onClick={toggle}
              >
                <span className="rb-tbl-pencolor" style={{ background: pen.color }} />
                <span>{l.penColor}</span>
                <Caret />
              </button>
            )}
          >
            {(close) => (
              <ColorPicker
                className="rb-tbl-colorpop"
                value={pen.color}
                strings={{
                  themeColors: t('ribbonThemeColorsSection'),
                  standardColors: t('ribbonStandardColors'),
                  moreColors: t('ribbonMoreColors'),
                }}
                onPick={(hex) => {
                  close()
                  if (hex) setPen((p) => ({ ...p, color: hex }))
                }}
                moreInputProps={{
                  onPointerDown: (e) => armColorInput(e.currentTarget),
                  onChange: (e) => setPen((p) => ({ ...p, color: e.currentTarget.value })),
                }}
              />
            )}
          </Pop>
        </div>
      </Group>
    </>
  )
}

// ── Table Layout ────────────────────────────────────────────────────────

/** A numeric field in centimetres that commits on Enter / blur (a keystroke must not resize the table). */
function CmField({
  label,
  valuePx,
  pxPerCm,
  onCommit,
  disabled,
}: {
  label: string
  valuePx: number
  pxPerCm: number
  onCommit: (px: number) => void
  disabled?: boolean
}) {
  const shown = (valuePx / pxPerCm).toFixed(2)
  const [draft, setDraft] = useState<string | null>(null)
  const commit = () => {
    if (draft === null) return
    const cm = Number.parseFloat(draft.replace(',', '.'))
    setDraft(null)
    if (Number.isFinite(cm) && cm > 0 && Math.abs(cm - Number(shown)) > 0.004) {
      onCommit(cm * pxPerCm)
    }
  }
  return (
    <label className="rb-tbl-field">
      <span>{label}</span>
      <input
        type="text"
        inputMode="decimal"
        disabled={disabled}
        value={draft ?? shown}
        onChange={(e) => setDraft(e.target.value)}
        onFocus={(e) => e.currentTarget.select()}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            commit()
            e.currentTarget.blur()
          } else if (e.key === 'Escape') {
            setDraft(null)
            e.currentTarget.blur()
          }
        }}
      />
      <em>cm</em>
    </label>
  )
}

export function TableLayoutTab({
  info,
  actions,
}: {
  info: TableTabInfo
  actions: TableTabActions
}) {
  const { t, lang } = useI18n()
  const l = tableRibbonLabels(lang)
  const cell = info.activeCell
  // no cell cursor: rows/columns go after the last / before the first, the way a whole-table selection behaves
  const row = cell?.row ?? 0
  const col = cell?.col ?? 0
  const lastRow = info.rows - 1
  const lastCol = info.cols - 1
  const sumPx = (a: number[]) => a.reduce((x, y) => x + y, 0)
  return (
    <>
      <Group label={l.grpRowsCols}>
        <Pop
          trigger={(open, toggle) => (
            <BigBtn icon="delete-table" label={l.del} caret active={open} onClick={toggle} />
          )}
        >
          {(close) => (
            <div className="rb-tbl-menu">
              <button
                type="button"
                className="rb-tbl-menuitem"
                disabled={info.cols < 2}
                onClick={() => {
                  close()
                  actions.structure('delete-col', col)
                }}
              >
                {t('appCtxDeleteCol')}
              </button>
              <button
                type="button"
                className="rb-tbl-menuitem"
                disabled={info.rows < 2}
                onClick={() => {
                  close()
                  actions.structure('delete-row', row)
                }}
              >
                {t('appCtxDeleteRow')}
              </button>
              <div className="row-menu-divider" />
              <button
                type="button"
                className="rb-tbl-menuitem danger"
                onClick={() => {
                  close()
                  actions.deleteTable()
                }}
              >
                {l.deleteTable}
              </button>
            </div>
          )}
        </Pop>
        <div className="rb-tbl-grid2">
          <SmallBtn
            icon="insert-above"
            iconOnly
            label={t('appCtxInsertRowAbove')}
            onClick={() => actions.structure('insert-row', row, true)}
          />
          <SmallBtn
            icon="insert-below"
            iconOnly
            label={t('appCtxInsertRowBelow')}
            onClick={() => actions.structure('insert-row', cell ? row : lastRow)}
          />
          <SmallBtn
            icon="insert-left"
            iconOnly
            label={t('appCtxInsertColLeft')}
            onClick={() => actions.structure('insert-col', col, true)}
          />
          <SmallBtn
            icon="insert-right"
            iconOnly
            label={t('appCtxInsertColRight')}
            onClick={() => actions.structure('insert-col', cell ? col : lastCol)}
          />
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={l.grpMerge}>
        <Pop
          trigger={(open, toggle) => (
            <BigBtn
              icon="merge-cells"
              label={l.mergeCells}
              caret
              active={open}
              disabled={!cell}
              onClick={toggle}
            />
          )}
        >
          {(close) => (
            <div className="rb-tbl-menu">
              <button
                type="button"
                className="rb-tbl-menuitem"
                onClick={() => {
                  close()
                  if (cell) actions.merge('merge-right', cell.row, cell.col)
                }}
              >
                {t('appCtxMergeRight')}
              </button>
              <button
                type="button"
                className="rb-tbl-menuitem"
                onClick={() => {
                  close()
                  if (cell) actions.merge('merge-down', cell.row, cell.col)
                }}
              >
                {t('appCtxMergeDown')}
              </button>
            </div>
          )}
        </Pop>
        <BigBtn
          icon="split-cells"
          label={t('appCtxSplitCell')}
          disabled={!cell}
          onClick={() => cell && actions.merge('split', cell.row, cell.col)}
        />
      </Group>
      <div className="ribbon-sep" />
      <Group label={l.grpCellSize}>
        <div className="rb-tbl-fields">
          <CmField
            label={l.rowHeight}
            valuePx={info.rowHeightsPx[row] ?? 0}
            pxPerCm={info.pxPerCm}
            onCommit={(px) => actions.setRowHeightPx(row, px)}
          />
          <CmField
            label={l.colWidth}
            valuePx={info.colWidthsPx[col] ?? 0}
            pxPerCm={info.pxPerCm}
            onCommit={(px) => actions.setColWidthPx(col, px)}
          />
        </div>
        <div className="rb-tbl-stack">
          <SmallBtn
            icon="distribute-v"
            label={l.distRows}
            onClick={() => actions.distribute('rows')}
          />
          <SmallBtn
            icon="distribute-h"
            label={l.distCols}
            onClick={() => actions.distribute('cols')}
          />
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={l.grpAlignment}>
        <div className="rb-tbl-alignstack">
          <div className="rb-tbl-aligngrid">
            <button
              type="button"
              className="rb-icon"
              data-tip={t('ribbonAlignLeft')}
              aria-label={t('ribbonAlignLeft')}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => actions.alignH('left')}
            >
              <CraftIcon set={DECK} name="align-left" size={SMALL} />
            </button>
            <button
              type="button"
              className="rb-icon"
              data-tip={t('ribbonAlignCenter')}
              aria-label={t('ribbonAlignCenter')}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => actions.alignH('center')}
            >
              <CraftIcon set={DECK} name="align-center" size={SMALL} />
            </button>
            <button
              type="button"
              className="rb-icon"
              data-tip={t('ribbonAlignRight')}
              aria-label={t('ribbonAlignRight')}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => actions.alignH('right')}
            >
              <CraftIcon set={DECK} name="align-right" size={SMALL} />
            </button>
            <button
              type="button"
              className="rb-icon"
              data-tip={t('appCtxCellAnchorTop')}
              aria-label={t('appCtxCellAnchorTop')}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => actions.alignV('top')}
            >
              <CraftIcon set={DECK} name="align-obj-top" size={SMALL} />
            </button>
            <button
              type="button"
              className="rb-icon"
              data-tip={t('appCtxCellAnchorMiddle')}
              aria-label={t('appCtxCellAnchorMiddle')}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => actions.alignV('middle')}
            >
              <CraftIcon set={DECK} name="align-obj-middle" size={SMALL} />
            </button>
            <button
              type="button"
              className="rb-icon"
              data-tip={t('appCtxCellAnchorBottom')}
              aria-label={t('appCtxCellAnchorBottom')}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => actions.alignV('bottom')}
            >
              <CraftIcon set={DECK} name="align-obj-bottom" size={SMALL} />
            </button>
          </div>
          <button
            type="button"
            className={`rb-icon${info.flags.rtl ? ' active' : ''}`}
            data-tip={t('ribbonTableRtl')}
            aria-label={t('ribbonTableRtl')}
            aria-pressed={info.flags.rtl}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => actions.toggleRtl(!info.flags.rtl)}
          >
            <CraftIcon set={DECK} name="text-direction" size={SMALL} />
          </button>
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={l.grpTableSize}>
        <div className="rb-tbl-fields">
          <CmField
            label={l.tableHeight}
            valuePx={sumPx(info.rowHeightsPx)}
            pxPerCm={info.pxPerCm}
            onCommit={(px) => actions.resizeTable('h', px)}
          />
          <CmField
            label={l.tableWidth}
            valuePx={sumPx(info.colWidthsPx)}
            pxPerCm={info.pxPerCm}
            onCommit={(px) => actions.resizeTable('w', px)}
          />
        </div>
      </Group>
    </>
  )
}
