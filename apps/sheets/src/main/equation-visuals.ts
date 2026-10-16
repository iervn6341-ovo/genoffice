/// Rehydrates editable equations from the drawing part after the sidecar has
/// assigned anchor indices. The sidecar's generic shape reader intentionally
/// does not interpret Office 2010 DrawingML extensions.

import { createHash } from 'node:crypto'
import { latexToOmml, ommlToLatex } from '@genoffice/office-math'
import { canEditEquationAnchor } from '@genoffice/xlsx-gateway/gateway/xlsx-equation-edit'

export interface DrawingEquation {
  readonly latex: string | null
  readonly omml: string
}

const ANCHOR_PATTERN =
  /<([A-Za-z_][\w.-]*:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/\1\2>/g

export function equationAtDrawingIndex(
  drawingXml: string,
  drawingIndex: number,
): DrawingEquation | null {
  const anchor = [...drawingXml.matchAll(ANCHOR_PATTERN)][drawingIndex]?.[0]
  if (!anchor || !/<a14:m\b/.test(anchor)) return null
  const choices = [
    ...anchor.matchAll(/<(?:[\w.-]+:)?Choice\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?Choice>/g),
  ]
  const content = choices.find((match) => /<a14:m\b/.test(match[1] ?? ''))?.[1] ?? anchor
  const omml = /<([\w.-]+:)?oMath\b[^>]*>([\s\S]*?)<\/\1oMath>/.exec(content)?.[2]
  if (!omml || omml.length > 131_072) return null
  const metadata =
    /<(?:[\w.-]+:)?cNvPr\b[^>]*\bdescr="GenOfficeEquationV2:([A-Za-z0-9+/=]+):([a-f0-9]{64})"/.exec(
      content,
    )
  const encodedSource = metadata?.[1]
  const digest = metadata?.[2]
  const matchesStoredSource = digest === createHash('sha256').update(omml).digest('hex')
  const storedLatex =
    encodedSource && encodedSource.length <= 24_000
      ? Buffer.from(encodedSource, 'base64').toString('utf8')
      : null
  let storedSourceMatchesMath = false
  if (matchesStoredSource && storedLatex?.trim()) {
    try {
      storedSourceMatchesMath = latexToOmml(storedLatex) === omml
    } catch {
      // Imported metadata is only an editing hint; native OMML is authoritative.
    }
  }
  const latex = storedSourceMatchesMath ? storedLatex : ommlToLatex(`<m:oMath>${omml}</m:oMath>`)
  return { latex: canEditEquationAnchor(anchor, latex) ? latex : null, omml }
}
