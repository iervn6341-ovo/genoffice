import type { TextBody } from '@genoffice/pptx-engine'

/** Only metrics and flow affect autofit. Paint-only edits must not resize an
 * imported box whose stored height differs from our measured text height. */
export function textLayoutKey(body: TextBody | undefined): string | null {
  if (!body) return null
  const { paragraphs, ...frame } = body
  return JSON.stringify([
    frame,
    paragraphs.map((p) => [
      p.align ?? 'left',
      p.rtl ?? false,
      p.level ?? 0,
      p.lineHeight,
      p.lineExact,
      p.spaceBefore,
      p.spaceAfter,
      p.spaceBeforePct,
      p.spaceAfterPct,
      p.marL,
      p.marR,
      p.indent,
      p.tabStops,
      p.defTabSz,
      p.bullet && [
        p.bullet.type,
        p.bullet.char,
        p.bullet.mediaRef,
        p.bullet.font,
        p.bullet.sizePct,
        p.bullet.sizePt,
        p.bullet.numType,
        p.bullet.startAt,
      ],
      p.defRPr && [
        p.defRPr.fontSize,
        !!p.defRPr.bold,
        !!p.defRPr.italic,
        p.defRPr.cap,
        p.defRPr.latinFont,
        p.defRPr.eaFont,
        p.defRPr.csFont,
      ],
      p.runs.map((r) => [
        // Stored field caches (such as ‹#›) differ from evaluated editor text.
        // Formatting the evaluated value must not count as a content edit.
        r.field ? null : r.text,
        r.rawXml,
        r.field,
        r.fontSize,
        r.fontFamily,
        r.latinFamily,
        r.fontScriptHint,
        r.latinFont,
        r.eaFont,
        r.csFont,
        !!r.bold,
        !!r.italic,
        r.cap,
        r.baseline ?? 0,
        r.letterSpacing ?? 0,
        r.kern,
        !!r.paraMark,
      ]),
    ]),
  ])
}
