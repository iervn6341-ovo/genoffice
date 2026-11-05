/** Remove native PAGE fields, preserving other fields and surrounding paragraph/table XML. */
export function removePageFields(xml: string): string {
  let next = xml.replace(/<w:fldSimple\b[^>]*>[\s\S]*?<\/w:fldSimple>/g, (field) =>
    /\bw:instr=["']\s*PAGE(?:\s|["'])/i.test(field) ? '' : field,
  )
  const ranges: { start: number; end: number; replacement: string }[] = []
  const stack: { start: number; instruction: string; result: boolean; prefix: string }[] = []
  for (const run of next.matchAll(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)) {
    const text = run[0],
      at = run.index!
    const type = text.match(/\bw:fldCharType=["'](begin|separate|end)["']/)?.[1]
    const tag = text.match(/<w:fldChar\b[^>]*(?:\/>|>[\s\S]*?<\/w:fldChar>)/)
    const visible = (fragment: string) =>
      /<w:(?:t|delText|drawing|pict|object|br|tab)\b/.test(fragment)
    if (type === 'begin') {
      const before = tag ? text.slice(0, tag.index) : ''
      stack.push({
        start: at,
        instruction: '',
        result: false,
        prefix: visible(before) ? before + '</w:r>' : '',
      })
    } else if (type === 'separate' && stack.length) stack[stack.length - 1]!.result = true
    else if (type === 'end') {
      const field = stack.pop()
      if (field && /^\s*PAGE(?:\s|$)/i.test(field.instruction)) {
        const after = tag ? text.slice(tag.index! + tag[0].length) : ''
        const open = text.match(/^<w:r(?:\s[^>]*)?>/)?.[0] ?? '<w:r>'
        const rPr = text.match(/<w:rPr\b[^>]*\/>|<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/)?.[0] ?? ''
        const suffix = visible(after) ? open + rPr + after : ''
        ranges.push({
          start: field.start,
          end: at + text.length,
          replacement: field.prefix + suffix,
        })
      }
    } else if (stack.length && !stack[stack.length - 1]!.result) {
      stack[stack.length - 1]!.instruction += [
        ...text.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g),
      ]
        .map((m) => m[1])
        .join('')
    }
  }
  // Outer fields subsume any nested PAGE fields, so apply disjoint ranges backwards.
  const disjoint = ranges.filter(
    (r) => !ranges.some((other) => other !== r && other.start <= r.start && other.end >= r.end),
  )
  for (const r of disjoint.sort((a, b) => b.start - a.start))
    next = next.slice(0, r.start) + r.replacement + next.slice(r.end)
  return next.replace(/<w:pgNum\b[^>]*\/>/g, '')
}
