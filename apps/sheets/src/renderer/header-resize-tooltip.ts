/**
 * Excel-style resize tooltip for sheet row and column resizing.
 * Displays character width / row height in points and pixels in real time:
 * e.g. "Width: 14.00 (89 pixels)" or "寬度: 14.00 (89 像素)".
 */
import { HeaderResizeRenderController } from '@univerjs/sheets-ui'
import { getWorkbookMdw } from './app-constants'
import { getLang } from './i18n/locale'

export function pixelsToExcelColWidth(pixels: number, mdw = getWorkbookMdw()): number {
  if (pixels <= 0) return 0
  const padding = 5
  if (pixels <= mdw + padding) {
    return Math.round((pixels / (mdw + padding)) * 100) / 100
  }
  return Math.round(((pixels - padding) / mdw) * 100) / 100
}

export function pixelsToExcelRowHeight(pixels: number): number {
  if (pixels <= 0) return 0
  const pt = pixels * 0.75
  return Math.round(pt * 100) / 100
}

export function formatResizeTooltip(
  kind: 'column' | 'row',
  pixels: number,
  lang: string = getLang(),
  mdw = getWorkbookMdw(),
): string {
  const safePx = Math.max(0, Math.round(pixels))
  const pxLabel = safePx === 1 ? 'pixel' : 'pixels'

  if (kind === 'column') {
    const chars = pixelsToExcelColWidth(safePx, mdw).toFixed(2)
    if (lang === 'zh-TW') {
      return `寬度: ${chars} (${safePx} 像素)`
    }
    if (lang === 'zh' || lang === 'zh-CN') {
      return `宽度: ${chars} (${safePx} 像素)`
    }
    if (lang === 'ja') {
      return `幅: ${chars} (${safePx} ピクセル)`
    }
    return `Width: ${chars} (${safePx} ${pxLabel})`
  }

  const pt = pixelsToExcelRowHeight(safePx).toFixed(2)
  if (lang === 'zh-TW') {
    return `高度: ${pt} (${safePx} 像素)`
  }
  if (lang === 'zh' || lang === 'zh-CN') {
    return `高度: ${pt} (${safePx} 像素)`
  }
  if (lang === 'ja') {
    return `高さ: ${pt} (${safePx} ピクセル)`
  }
  return `Height: ${pt} (${safePx} ${pxLabel})`
}

let tooltipEl: HTMLElement | null = null

function getOrCreateTooltipElement(): HTMLElement {
  if (!tooltipEl || !document.body.contains(tooltipEl)) {
    tooltipEl = document.createElement('div')
    tooltipEl.id = 'sheets-header-resize-tooltip'
    tooltipEl.className = 'sheets-header-resize-tooltip'
    tooltipEl.setAttribute('role', 'tooltip')
    tooltipEl.setAttribute('aria-hidden', 'true')
    document.body.appendChild(tooltipEl)
  }
  return tooltipEl
}

export interface ShowResizeTooltipOptions {
  kind: 'column' | 'row'
  pixels: number
  clientX: number
  clientY: number
  canvasRect?: DOMRect | null
  rowHeaderWidth?: number
  lang?: string
  mdw?: number
}

export function showResizeTooltip(opts: ShowResizeTooltipOptions): void {
  if (typeof document === 'undefined') return
  const el = getOrCreateTooltipElement()
  const text = formatResizeTooltip(opts.kind, opts.pixels, opts.lang, opts.mdw)
  el.textContent = text
  el.style.display = 'block'

  const { kind, clientX, clientY, canvasRect, rowHeaderWidth = 46 } = opts

  if (kind === 'column') {
    // Horizontally centered at the column divider line;
    // Vertically positioned just above the column header bar.
    const top = canvasRect && canvasRect.top >= 26
      ? Math.max(6, canvasRect.top - 4)
      : Math.max(6, clientY - 30)

    el.style.left = `${Math.round(clientX)}px`
    el.style.top = `${Math.round(top)}px`
    el.style.transform = 'translate(-50%, -100%)'
  } else {
    // Vertically centered at the row divider line;
    // Horizontally positioned just to the right of the row header strip.
    const left = canvasRect
      ? Math.round(canvasRect.left + rowHeaderWidth + 8)
      : Math.round(clientX + 16)

    el.style.left = `${left}px`
    el.style.top = `${Math.round(clientY)}px`
    el.style.transform = 'translateY(-50%)'
  }
}

export function hideResizeTooltip(): void {
  if (tooltipEl) {
    tooltipEl.style.display = 'none'
  }
}

let installed = false

export function installHeaderResizeTooltip(): void {
  if (installed) return
  installed = true

  const proto = HeaderResizeRenderController.prototype as unknown as {
    _initialHoverResize: (initialType?: number) => void
  }
  if (!proto || typeof proto._initialHoverResize !== 'function') return

  const origInitialHoverResize = proto._initialHoverResize

  proto._initialHoverResize = function (initialType = 0) {
    origInitialHoverResize.call(this, initialType)

    const controller = this as any
    const scene = controller._context?.scene
    if (!scene) return

    const eventBindingObject = initialType === 0 ? controller._rowResizeRect : controller._columnResizeRect
    if (!eventBindingObject) return

    controller.disposeWithMe(
      eventBindingObject.onPointerDown$.subscribeEvent((downEvt: any) => {
        const skeleton = controller._sheetSkeletonManagerService?.getCurrentParam()?.skeleton
        if (!skeleton) return

        const isCol = initialType === 1
        const col = controller._currentColumn
        const row = controller._currentRow
        const cell = skeleton.getNoMergeCellWithCoordByIndex(row, col)
        if (!cell) return

        const startClientX = Number(downEvt.clientX ?? (downEvt.inputEvent as MouseEvent)?.clientX ?? 0)
        const startClientY = Number(downEvt.clientY ?? (downEvt.inputEvent as MouseEvent)?.clientY ?? 0)
        let hasMoved = false

        let moveSub: any = null
        let upSub: any = null

        const cleanup = () => {
          hideResizeTooltip()
          if (moveSub) {
            moveSub.unsubscribe()
            moveSub = null
          }
          if (upSub) {
            upSub.unsubscribe()
            upSub = null
          }
          window.removeEventListener('pointerup', cleanup)
          window.removeEventListener('keydown', onKeyDown)
        }

        const onKeyDown = (e: KeyboardEvent) => {
          if (e.key === 'Escape') cleanup()
        }

        window.addEventListener('pointerup', cleanup, { once: true })
        window.addEventListener('keydown', onKeyDown, { once: true })

        moveSub = scene.onPointerMove$.subscribeEvent((moveEvt: any) => {
          const helper = controller._resizeHelperShape
          if (!helper) return

          const clientX = Number(moveEvt.clientX ?? (moveEvt.inputEvent as MouseEvent)?.clientX ?? startClientX)
          const clientY = Number(moveEvt.clientY ?? (moveEvt.inputEvent as MouseEvent)?.clientY ?? startClientY)

          if (!hasMoved) {
            if (Math.abs(clientX - startClientX) >= 2 || Math.abs(clientY - startClientY) >= 2) {
              hasMoved = true
            }
          }

          let newPx: number
          if (isCol) {
            const helperLeft = (helper as { left?: number }).left
            if (typeof helperLeft !== 'number') return
            newPx = Math.max(0, Math.round(helperLeft - cell.startX))
          } else {
            const helperTop = (helper as { top?: number }).top
            if (typeof helperTop !== 'number') return
            newPx = Math.max(0, Math.round(helperTop - cell.startY))
          }

          const canvas = scene.getEngine()?.getCanvasElement()
          const canvasRect = canvas ? canvas.getBoundingClientRect() : null

          showResizeTooltip({
            kind: isCol ? 'column' : 'row',
            pixels: newPx,
            clientX,
            clientY,
            canvasRect,
            rowHeaderWidth: skeleton.rowHeaderWidthAndMarginLeft || 46,
          })
        })

        upSub = scene.onPointerUp$.subscribeEvent(() => {
          cleanup()
        })
      }),
    )
  }
}
