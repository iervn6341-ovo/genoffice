import { BooleanNumber, BorderStyleTypes } from '@univerjs/core'

/**
 * Home → Format as Table. Univer paints every table with its own lavender theme, so the style the
 * user picked (Light / Medium / Dark, which the save writes as `tableStyleInfo`) was never drawn.
 * This maps Excel's built-in style names to a Univer table theme in the Office accent colors:
 * header band, first-row stripe and (Dark) second-row stripe, which is what the gallery's
 * thumbnails show. Excel draws its own exact style from the saved name on open.
 */
export interface TableThemeItem {
  bg?: { rgb: string }
  cl?: { rgb: string }
  bl?: BooleanNumber
  bd?: { t?: BorderEdge; b?: BorderEdge }
}

interface BorderEdge {
  s: BorderStyleTypes
  cl: { rgb: string }
}

export interface TableThemeJson {
  name: string
  headerRowStyle?: TableThemeItem
  firstRowStyle?: TableThemeItem
  secondRowStyle?: TableThemeItem
}

/** Office theme accents 1–6 */
const ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
const BLACK = '#000000'
const WHITE = '#FFFFFF'

const channel = (hex: string, at: number) => parseInt(hex.slice(1 + at * 2, 3 + at * 2), 16)
const toHex = (r: number, g: number, b: number) =>
  `#${[r, g, b].map((v) => Math.round(v).toString(16).padStart(2, '0')).join('')}`.toUpperCase()
/** "Lighter 80%" — mixed towards white */
const tint = (hex: string, amount: number) =>
  toHex(
    ...([0, 1, 2].map((i) => channel(hex, i) + (255 - channel(hex, i)) * amount) as [
      number,
      number,
      number,
    ]),
  )
/** "Darker 50%" — mixed towards black */
const shade = (hex: string, amount: number) =>
  toHex(...([0, 1, 2].map((i) => channel(hex, i) * (1 - amount)) as [number, number, number]))

const rgb = (hex: string) => ({ rgb: hex })
const edge = (hex: string): BorderEdge => ({ s: BorderStyleTypes.THIN, cl: rgb(hex) })

/** the color a style number stands for: 1 is black, 2–7 are accents 1–6 (Light 9–14 too) */
function colorOf(family: string, n: number): string | null {
  if (family === 'Light') {
    if (n === 1 || n === 8) return BLACK
    if (n >= 2 && n <= 7) return ACCENTS[n - 2] ?? null
    if (n >= 9 && n <= 14) return ACCENTS[n - 9] ?? null
    return null
  }
  if (n === 1) return BLACK
  if (n >= 2 && n <= 7) return ACCENTS[n - 2] ?? null
  return null
}

/**
 * The theme for a built-in style name, or null for one this draws no look for (the table then keeps
 * Univer's default). `uniqueName` must differ per application — Univer keys themes by name.
 */
export function tableThemeFor(
  style: string | undefined,
  uniqueName: string,
): TableThemeJson | null {
  const match = /^TableStyle(Light|Medium|Dark)(\d{1,2})$/.exec(style ?? '')
  if (!match) return null
  const family = match[1]!
  const n = Number(match[2])
  const color = colorOf(family, n)
  if (!color) return null
  const black = color === BLACK
  if (family === 'Light') {
    if (n >= 8) {
      // solid header, hairlines on the banded rows
      return {
        name: uniqueName,
        headerRowStyle: { bg: rgb(color), cl: rgb(WHITE), bl: BooleanNumber.TRUE },
        firstRowStyle: { bd: { t: edge(color), b: edge(color) } },
      }
    }
    // outline styles: bold header over a rule, tinted stripe
    return {
      name: uniqueName,
      headerRowStyle: { bl: BooleanNumber.TRUE, bd: { b: edge(color) } },
      firstRowStyle: { bg: rgb(black ? '#D9D9D9' : tint(color, 0.8)) },
    }
  }
  if (family === 'Medium') {
    return {
      name: uniqueName,
      headerRowStyle: { bg: rgb(color), cl: rgb(WHITE), bl: BooleanNumber.TRUE },
      firstRowStyle: { bg: rgb(black ? '#D9D9D9' : tint(color, 0.8)) },
    }
  }
  // Dark: a deep header over two tinted stripes
  return {
    name: uniqueName,
    headerRowStyle: {
      bg: rgb(black ? BLACK : shade(color, 0.5)),
      cl: rgb(WHITE),
      bl: BooleanNumber.TRUE,
    },
    firstRowStyle: { bg: rgb(black ? '#A6A6A6' : tint(color, 0.6)) },
    secondRowStyle: { bg: rgb(black ? '#D9D9D9' : tint(color, 0.8)) },
  }
}

/** Univer's facade call that puts a theme on an existing table */
export interface TableThemeSheet {
  addTableTheme(tableId: string, theme: TableThemeJson): Promise<boolean> | boolean
}
