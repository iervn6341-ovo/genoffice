import type { FontMetricsProvider, RunStyle } from './metrics'

interface MathNode {
  name: string
  children: MathNode[]
  text?: string
  bevelled?: boolean
}

function decodeMathText(text: string): string {
  return text
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, decimal: string) => String.fromCodePoint(Number(decimal)))
    .replace(
      /&(?:amp|lt|gt|quot|apos);/g,
      (entity) =>
        ({ '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" })[entity] ??
        entity,
    )
}

/** Parse the MathML subset produced by office-math without requiring browser DOM in layout tests. */
function mathTree(mathML: string): MathNode {
  const root: MathNode = { name: 'root', children: [] }
  const stack = [root]
  for (const part of mathML.match(/<[^>]+>|[^<]+/g) ?? []) {
    if (!part.startsWith('<')) {
      if (part.trim())
        stack.at(-1)!.children.push({ name: '#text', children: [], text: decodeMathText(part) })
      continue
    }
    const match = part.match(/^<(\/?)\s*([\w:-]+)/)
    if (!match) continue
    if (match[1] === '/') {
      if (stack.length > 1 && stack.at(-1)!.name === match[2]) stack.pop()
      continue
    }
    const node: MathNode = {
      name: match[2]!,
      children: [],
      bevelled: /\bbevelled="true"/.test(part),
    }
    stack.at(-1)!.children.push(node)
    if (!/\/\s*>$/.test(part)) stack.push(node)
  }
  return root
}

function treeWidth(node: MathNode, metrics: FontMetricsProvider, style: RunStyle): number {
  const em = style.fontSizePx
  if (node.name === '#text') {
    const text = node.text!.replace(/\u2061/g, '')
    return metrics.measure(text, style)
  }
  const parts = node.children.map((child) => treeWidth(child, metrics, style))
  const sum = parts.reduce((total, width) => total + width, 0)
  switch (node.name) {
    case 'mfrac':
      return node.bevelled ? sum * 0.8 + em * 0.45 : Math.max(...parts, 0) + em * 0.42
    case 'msqrt':
      return sum + em * 0.75
    case 'mroot':
      return (parts[0] ?? 0) + em * 0.75 + (parts[1] ?? 0) * 0.45
    case 'msup':
    case 'msub':
      return (parts[0] ?? 0) + (parts[1] ?? 0) * 0.7
    case 'msubsup':
      return (parts[0] ?? 0) + Math.max(parts[1] ?? 0, parts[2] ?? 0) * 0.7
    case 'mmultiscripts':
      return (parts[0] ?? 0) + Math.max(...parts.slice(1), 0) * 0.7
    case 'mover':
    case 'munder':
    case 'munderover':
      return Math.max(...parts, 0) + em * 0.08
    case 'mtable': {
      const rows = node.children.filter((child) => child.name === 'mtr')
      const columns: number[] = []
      for (const row of rows) {
        row.children
          .filter((cell) => cell.name === 'mtd')
          .forEach((cell, index) => {
            columns[index] = Math.max(columns[index] ?? 0, treeWidth(cell, metrics, style))
          })
      }
      return (
        columns.reduce((total, width) => total + width, 0) +
        Math.max(0, columns.length - 1) * em * 0.45
      )
    }
    case 'mprescripts':
      return 0
    default:
      return sum
  }
}

/** Approximate rendered structure, never the linear fallback that spells out stacked fractions. */
export function mathWidthPx(mathML: string, metrics: FontMetricsProvider, style: RunStyle): number {
  const structural = treeWidth(mathTree(mathML), metrics, style)
  // Cambria Math and the layout font can differ slightly; leave a small safety margin.
  return Math.max(style.fontSizePx * 0.6, structural * 1.12 + style.fontSizePx * 0.12)
}
