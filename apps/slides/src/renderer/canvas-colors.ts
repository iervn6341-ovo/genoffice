/**
 * Colors of canvas-drawn editing chrome (overlays painted on the slide, not app UI). They sit on the
 * authored slide, so they pick a light or dark variant by what is under them — the same verdict the
 * selection frame uses (selectionChromeColor) — instead of following the UI theme.
 */
export interface CellSelColors {
  /** outline of the selected block */
  stroke: string
  /** translucent wash over the selected cells */
  fill: string
  /** the press-to-select row / column handles outside the table */
  gutter: string
}

const ON_LIGHT: CellSelColors = {
  stroke: '#232425',
  fill: 'rgba(35, 36, 37, 0.16)',
  gutter: 'rgba(35, 36, 37, 0.55)',
}

const ON_DARK: CellSelColors = {
  stroke: '#ffffff',
  fill: 'rgba(255, 255, 255, 0.26)',
  gutter: 'rgba(255, 255, 255, 0.7)',
}

/** `chrome` = the frame color selectionChromeColor() chose for this spot (white on dark backgrounds) */
export function tableCellSelColors(chrome: string): CellSelColors {
  return chrome.toLowerCase() === '#ffffff' ? ON_DARK : ON_LIGHT
}
