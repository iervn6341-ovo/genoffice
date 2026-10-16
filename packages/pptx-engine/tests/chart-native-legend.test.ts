import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import {
  addChart,
  createBlankPptx,
  openPptx,
  savePptx,
  setChartLegendPosition,
  setChartAreaFormat,
  canSetChartQuickLayout,
  setChartQuickLayout,
} from '../src/index'
import type { ChartElement } from '../src/types'

async function importedChart() {
  const created = await openPptx(await createBlankPptx())
  addChart(created, 0, {
    kind: 'bar',
    categories: ['A', 'B'],
    series: [{ name: 'Revenue', values: [4, 7] }],
    offset: { x: 914400, y: 914400, cx: 4572000, cy: 2743200 },
    legendPos: 'b',
  })
  const zip = await JSZip.loadAsync(await savePptx(created))
  const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  zip.file('ppt/slides/slide1.xml', slideXml.replace(' descr="aislides-chart"', ''))
  const chartXml = await zip.file('ppt/charts/chart1.xml')!.async('string')
  zip.file(
    'ppt/charts/chart1.xml',
    chartXml
      .replace(
        '<c:legend><c:legendPos val="b"/>',
        '<c:legend><c:legendPos val="b"/><c:legendEntry><c:idx val="0"/><c:txPr data-preserve="1"/></c:legendEntry>',
      )
      .replace('</c:chartSpace>', '<c:extLst data-preserve="chart-space"/></c:chartSpace>'),
  )
  const opened = await openPptx(await zip.generateAsync({ type: 'uint8array' }))
  const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
  return { opened, chart }
}

describe('native chart legend edit', () => {
  it('maps PowerPoint Mac column Layouts 1 and 3 to native label and legend fields', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    const opened = await openPptx(source)
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    const otherParts = new Map([...opened.archive.entries].filter(([name]) => name !== path))
    expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-1')).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'right-values')).toBe(true)
    expect(opened.archive.readText(path)).toContain('<c:showVal val="1"/>')
    expect(setChartQuickLayout(opened, 0, chart.id, 'mac-1')).toBe(true)
    const right = opened.archive.readText(path)!
    expect(right).toContain('<c:legendPos val="r"/>')
    expect(right).toContain('<c:showVal val="0"/>')
    expect(right).toContain('<c:gapWidth val="150"/>')
    expect(right).not.toContain('<c:overlap')
    expect(right).toContain('<c:majorGridlines>')
    expect(right).toContain('<c:externalData r:id="rId3"')
    expect(chart.chart.gapWidthPct).toBe(150)
    expect(setChartQuickLayout(opened, 0, chart.id, 'mac-3')).toBe(true)
    const bottom = opened.archive.readText(path)!
    expect(bottom).toContain('<c:legendPos val="b"/>')
    expect(bottom).toContain('<c:showVal val="0"/>')
    expect(bottom).toContain('<c:gapWidth val="75"/>')
    expect(bottom).toContain('<c:overlap val="-25"/>')
    for (const tag of ['c:ser', 'c:extLst', 'c:catAx', 'c:valAx']) {
      const fragments = (xml: string) =>
        [...xml.matchAll(new RegExp(`<${tag}\\b[\\s\\S]*?<\\/${tag}>`, 'g'))].map(
          (match) => match[0],
        )
      expect(fragments(bottom)).toEqual(fragments(original))
    }
    expect(chart.chart.dataLabels).toBe(false)
    for (const [name, content] of otherParts)
      expect(opened.archive.entries.get(name)).toEqual(content)
    const reopened = await openPptx(await savePptx(opened))
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restored.chart.legendPos).toBe('b')
    expect(restored.chart.dataLabels).toBeUndefined()
    expect(restored.chart.gapWidthPct).toBe(75)
    expect(restored.chart.overlapPct).toBe(-25)
    expect(setChartQuickLayout(reopened, 0, restored.id, 'mac-1')).toBe(true)
    expect(reopened.archive.readText(path)).toContain('<c:legendPos val="r"/>')
  })

  it('maps Mac Layout 11 title deletion and both tick marks while keeping chart dependencies', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    const native11 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-11-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    const native1 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-01-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    expect(
      native1
        .replace(/<c:title\b[\s\S]*?<\/c:title>/, '')
        .replace('<c:autoTitleDeleted val="0"/>', '<c:autoTitleDeleted val="1"/>')
        .replaceAll('<c:majorTickMark val="none"/>', '<c:majorTickMark val="out"/>'),
    ).toBe(native11)
    const opened = await openPptx(source)
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    const dependencies = new Map(
      [...opened.archive.entries].filter(([name]) =>
        /^(?:ppt\/charts\/(?:_rels\/|style|colors)|ppt\/embeddings\/)/.test(name),
      ),
    )
    expect(native11).not.toMatch(/<c:title\b/)
    expect(native11).toContain('<c:autoTitleDeleted val="1"/>')
    expect((native11.match(/<c:majorTickMark val="out"\/>/g) ?? []).length).toBe(2)
    expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-11')).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'mac-11')).toBe(true)
    const updated = opened.archive.readText(path)!
    expect(updated).not.toMatch(/<c:title\b/)
    expect(updated).toContain('<c:autoTitleDeleted val="1"/>')
    expect((updated.match(/<c:majorTickMark val="out"\/>/g) ?? []).length).toBe(2)
    expect(updated).toContain('<c:legendPos val="r"/>')
    expect(updated).toContain('<c:gapWidth val="150"/>')
    expect(updated).not.toMatch(/<c:overlap\b/)
    for (const name of ['c:ser', 'c:extLst', 'mc:AlternateContent']) {
      const fragments = (xml: string) =>
        [...xml.matchAll(new RegExp(`<${name}\\b[\\s\\S]*?<\\/${name}>`, 'g'))].map(
          (match) => match[0],
        )
      expect(fragments(updated)).toEqual(fragments(original))
    }
    for (const [name, content] of dependencies)
      expect(opened.archive.entries.get(name)).toEqual(content)
    const reopened = await openPptx(await savePptx(opened))
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restored.chart.title).toBeUndefined()
    expect(restored.chart.gapWidthPct).toBe(150)
    expect(restored.chart.legendPos).toBe('r')
    expect(canSetChartQuickLayout(reopened, 0, restored.id, 'mac-11')).toBe(true)
    expect(setChartQuickLayout(reopened, 0, restored.id, 'mac-11')).toBe(true)
    expect(reopened.archive.readText(path)).toBe(updated)
  })

  it('adds Mac Layout 9 native automatic axis titles without rebuilding imported data', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    const native9 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-09-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    const opened = await openPptx(source)
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    const dependencies = new Map(
      [...opened.archive.entries].filter(([name]) =>
        /^(?:ppt\/charts\/(?:_rels\/|style|colors)|ppt\/embeddings\/)/.test(name),
      ),
    )
    expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-9')).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'mac-9')).toBe(true)
    const updated = opened.archive.readText(path)!
    const axisTitle = (xml: string, axis: 'catAx' | 'valAx') =>
      new RegExp(`<c:${axis}>[\\s\\S]*?(<c:title>[\\s\\S]*?<\\/c:title>)`).exec(xml)?.[1]
    expect(axisTitle(updated, 'catAx')).toBe(axisTitle(native9, 'catAx'))
    expect(axisTitle(updated, 'valAx')).toBe(axisTitle(native9, 'valAx'))
    expect(updated).toContain('<c:legendPos val="r"/>')
    expect(updated).toContain('<c:gapWidth val="150"/>')
    expect(updated).not.toMatch(/<c:overlap\b/)
    expect((updated.match(/<c:majorTickMark val="out"\/>/g) ?? []).length).toBe(1)
    for (const name of ['c:ser', 'c:extLst', 'mc:AlternateContent']) {
      const fragments = (xml: string) =>
        [...xml.matchAll(new RegExp(`<${name}\\b[\\s\\S]*?<\\/${name}>`, 'g'))].map(
          (match) => match[0],
        )
      expect(fragments(updated)).toEqual(fragments(original))
    }
    for (const [name, content] of dependencies)
      expect(opened.archive.entries.get(name)).toEqual(content)
    const reopened = await openPptx(await savePptx(opened))
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restored.chart.catAxis?.title).toBe('Axis Title')
    expect(restored.chart.valAxis?.title).toBe('Axis Title')
    expect(restored.chart.gapWidthPct).toBe(150)
    expect(canSetChartQuickLayout(reopened, 0, restored.id, 'mac-9')).toBe(true)
    expect(setChartQuickLayout(reopened, 0, restored.id, 'mac-9')).toBe(true)
    expect(reopened.archive.readText(path)).toBe(updated)
  })

  it('adds Mac Layout 10 labels only to the third Office series and keeps chart dependencies', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    const native10 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-10-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    const opened = await openPptx(source)
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    const dependencies = new Map(
      [...opened.archive.entries].filter(([name]) =>
        /^(?:ppt\/charts\/(?:_rels\/|style|colors)|ppt\/embeddings\/)/.test(name),
      ),
    )
    const series = (xml: string) => [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((m) => m[0])
    const label = (xml: string) => /<c:dLbls>[\s\S]*?<\/c:dLbls>/.exec(xml)?.[0]
    expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-10')).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'mac-10')).toBe(true)
    const updated = opened.archive.readText(path)!
    const beforeSeries = series(original)
    const afterSeries = series(updated)
    const nativeSeries = series(native10)
    expect(afterSeries).toHaveLength(3)
    expect(afterSeries.slice(0, 2)).toEqual(beforeSeries.slice(0, 2))
    expect(afterSeries[2]!.replace(label(afterSeries[2]!)!, '')).toBe(beforeSeries[2])
    expect(label(afterSeries[2]!)).toBe(label(nativeSeries[2]!))
    expect(updated).toContain('<c:gapWidth val="75"/>')
    expect(updated).toContain('<c:overlap val="40"/>')
    expect(updated).toContain('<c:legendPos val="r"/>')
    expect(updated).toContain('<c:externalData r:id="rId3"')
    expect(chart.chart.series.map((item) => item.dataLabels)).toEqual([false, false, true])
    for (const [name, content] of dependencies)
      expect(opened.archive.entries.get(name)).toEqual(content)
    const reopened = await openPptx(await savePptx(opened))
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restored.chart.series.map((item) => item.dataLabels)).toEqual([false, false, true])
    expect(restored.chart.gapWidthPct).toBe(75)
    expect(restored.chart.overlapPct).toBe(40)
    expect(canSetChartQuickLayout(reopened, 0, restored.id, 'mac-10')).toBe(true)
    expect(setChartQuickLayout(reopened, 0, restored.id, 'mac-10')).toBe(true)
    expect(reopened.archive.readText(path)).toBe(updated)
  })

  it('maps Mac Layout 7 title, axis titles, minor gridlines, and gap without changing chart dependencies', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    const native7 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-07-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    const native9 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-09-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    const minor = (xml: string) => /<c:minorGridlines>[\s\S]*?<\/c:minorGridlines>/.exec(xml)?.[0]
    expect(
      native9
        .replace(/<c:title>[\s\S]*?<\/c:title>/, '')
        .replace('<c:autoTitleDeleted val="0"/>', '<c:autoTitleDeleted val="1"/>')
        .replace('<c:gapWidth val="150"/>', '<c:gapWidth val="300"/>')
        .replace('</c:majorGridlines>', `</c:majorGridlines>${minor(native7)}`),
    ).toBe(native7)
    const opened = await openPptx(source)
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    const dependencies = new Map(
      [...opened.archive.entries].filter(([name]) =>
        /^(?:ppt\/charts\/(?:_rels\/|style|colors)|ppt\/embeddings\/)/.test(name),
      ),
    )
    const axisTitle = (xml: string, axis: 'catAx' | 'valAx') =>
      new RegExp(`<c:${axis}>[\\s\\S]*?(<c:title>[\\s\\S]*?<\\/c:title>)`).exec(xml)?.[1]
    expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-7')).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'mac-7')).toBe(true)
    const updated = opened.archive.readText(path)!
    expect(updated).not.toMatch(/<c:chart>\s*<c:title>/)
    expect(updated).toContain('<c:autoTitleDeleted val="1"/>')
    expect(updated).toContain('<c:gapWidth val="300"/>')
    expect(updated).not.toMatch(/<c:overlap\b/)
    expect(updated).toContain('<c:legendPos val="r"/>')
    expect(minor(updated)).toBe(minor(native7))
    expect(axisTitle(updated, 'catAx')).toBe(axisTitle(native7, 'catAx'))
    expect(axisTitle(updated, 'valAx')).toBe(axisTitle(native7, 'valAx'))
    for (const tag of ['c:ser', 'c:extLst', 'mc:AlternateContent']) {
      const fragments = (xml: string) =>
        [...xml.matchAll(new RegExp(`<${tag}\\b[\\s\\S]*?<\\/${tag}>`, 'g'))].map((m) => m[0])
      expect(fragments(updated)).toEqual(fragments(original))
    }
    for (const [name, content] of dependencies)
      expect(opened.archive.entries.get(name)).toEqual(content)
    expect(chart.chart.title).toBeUndefined()
    expect(chart.chart.catAxis?.title).toBe('Axis Title')
    expect(chart.chart.valAxis?.title).toBe('Axis Title')
    expect(chart.chart.valAxis?.minorGridColor).toBeTruthy()
    const reopened = await openPptx(await savePptx(opened))
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restored.chart.valAxis?.minorGridColor).toBe(chart.chart.valAxis?.minorGridColor)
    expect(restored.chart.gapWidthPct).toBe(300)
    expect(canSetChartQuickLayout(reopened, 0, restored.id, 'mac-7')).toBe(true)
    expect(setChartQuickLayout(reopened, 0, restored.id, 'mac-7')).toBe(true)
    expect(reopened.archive.readText(path)).toBe(updated)
  })

  it('fails closed for imported minor-grid extensions or linked titles in Layout 7', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    for (const mutate of [
      (xml: string) =>
        xml.replace(
          '</c:majorGridlines>',
          '</c:majorGridlines><c:minorGridlines><c:extLst data-preserve="1"/></c:minorGridlines>',
        ),
      (xml: string) => xml.replace('</c:title>', '<c:extLst data-preserve="1"/></c:title>'),
      (xml: string) => xml.replace('</c:plotArea>', '<c:dTable/></c:plotArea>'),
    ]) {
      const opened = await openPptx(source)
      const chart = opened.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      const path = 'ppt/charts/chart1.xml'
      const altered = mutate(opened.archive.readText(path)!)
      opened.archive.entries.set(path, Buffer.from(altered, 'utf8'))
      expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-7')).toBe(false)
      expect(setChartQuickLayout(opened, 0, chart.id, 'mac-7')).toBe(false)
      expect(opened.archive.readText(path)).toBe(altered)
    }
  })

  it('maps Mac Layout 8 to native no-legend and no-gridline XML without touching chart data', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    const native9 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-09-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    const native8 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-08-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    expect(
      native9
        .replace('<c:gapWidth val="150"/>', '<c:gapWidth val="0"/>')
        .replace(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/, '')
        .replace(/<c:legend>[\s\S]*?<\/c:legend>/, ''),
    ).toBe(native8)
    const opened = await openPptx(source)
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    const dependencies = new Map(
      [...opened.archive.entries].filter(([name]) =>
        /^(?:ppt\/charts\/(?:_rels\/|style|colors)|ppt\/embeddings\/)/.test(name),
      ),
    )
    expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-8')).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'mac-8')).toBe(true)
    const updated = opened.archive.readText(path)!
    expect(updated).toContain('<c:gapWidth val="0"/>')
    expect(updated).not.toMatch(/<c:majorGridlines\b|<c:legend\b|<c:overlap\b/)
    expect(updated).toContain('<c:externalData r:id="rId3"')
    for (const tag of ['c:ser', 'c:extLst', 'mc:AlternateContent']) {
      const fragments = (xml: string) =>
        [...xml.matchAll(new RegExp(`<${tag}\\b[\\s\\S]*?<\\/${tag}>`, 'g'))].map((m) => m[0])
      expect(fragments(updated)).toEqual(fragments(original))
    }
    for (const [name, content] of dependencies)
      expect(opened.archive.entries.get(name)).toEqual(content)
    expect(chart.chart.valAxis?.gridColor).toBeUndefined()
    expect(chart.chart.legendPos).toBeUndefined()
    expect(chart.chart.catAxis?.title).toBe('Axis Title')
    expect(chart.chart.valAxis?.title).toBe('Axis Title')
    const reopened = await openPptx(await savePptx(opened))
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restored.chart.valAxis?.gridColor).toBeUndefined()
    expect(restored.chart.legendPos).toBeUndefined()
    expect(restored.chart.gapWidthPct).toBe(0)
    expect(canSetChartQuickLayout(reopened, 0, restored.id, 'mac-8')).toBe(true)
    expect(setChartQuickLayout(reopened, 0, restored.id, 'mac-8')).toBe(true)
    expect(reopened.archive.readText(path)).toBe(updated)
  })

  it('refuses Layout 8 when removal would discard unknown gridline or legend XML', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    for (const mutate of [
      (xml: string) =>
        xml.replace('</c:majorGridlines>', '<c:extLst data-preserve="1"/></c:majorGridlines>'),
      (xml: string) => xml.replace('</c:legend>', '<c:extLst data-preserve="1"/></c:legend>'),
      (xml: string) => xml.replace('</c:legend>', '<c:extLst r:id="rId88"/></c:legend>'),
    ]) {
      const opened = await openPptx(source)
      const chart = opened.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      const path = 'ppt/charts/chart1.xml'
      const altered = mutate(opened.archive.readText(path)!)
      opened.archive.entries.set(path, Buffer.from(altered, 'utf8'))
      expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-8')).toBe(false)
      expect(setChartQuickLayout(opened, 0, chart.id, 'mac-8')).toBe(false)
      expect(opened.archive.readText(path)).toBe(altered)
    }
  })

  it('maps Mac Layout 4 to three native value-label series without rewriting imported data', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    const native4 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-04-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    const opened = await openPptx(source)
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    const dependencies = new Map(
      [...opened.archive.entries].filter(([name]) =>
        /^(?:ppt\/charts\/(?:_rels\/|style|colors)|ppt\/embeddings\/)/.test(name),
      ),
    )
    const series = (xml: string) => [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((m) => m[0])
    const label = (xml: string) => /<c:dLbls>[\s\S]*?<\/c:dLbls>/.exec(xml)?.[0]
    expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-4')).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'mac-4')).toBe(true)
    const updated = opened.archive.readText(path)!
    const beforeSeries = series(original)
    const afterSeries = series(updated)
    const nativeSeries = series(native4)
    expect(afterSeries).toHaveLength(3)
    for (let i = 0; i < 3; i++) {
      expect(label(afterSeries[i]!)).toBe(label(nativeSeries[i]!))
      expect(afterSeries[i]!.replace(label(afterSeries[i]!)!, '')).toBe(beforeSeries[i])
    }
    expect(updated).not.toMatch(/<c:chart>\s*<c:title>/)
    expect(updated).toContain('<c:autoTitleDeleted val="1"/>')
    expect(updated).toContain('<c:gapWidth val="75"/>')
    expect(updated).toContain('<c:legendPos val="b"/>')
    expect(updated).not.toMatch(/<c:majorGridlines\b|<c:overlap\b/)
    expect(updated).toContain('<c:externalData r:id="rId3"')
    for (const [name, content] of dependencies)
      expect(opened.archive.entries.get(name)).toEqual(content)
    expect(chart.chart.series.map((item) => item.dataLabels)).toEqual([true, true, true])
    expect(chart.chart.valAxis?.gridColor).toBeUndefined()
    const reopened = await openPptx(await savePptx(opened))
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restored.chart.series.map((item) => item.dataLabels)).toEqual([true, true, true])
    expect(restored.chart.title).toBeUndefined()
    expect(restored.chart.gapWidthPct).toBe(75)
    expect(canSetChartQuickLayout(reopened, 0, restored.id, 'mac-4')).toBe(true)
    expect(setChartQuickLayout(reopened, 0, restored.id, 'mac-4')).toBe(true)
    expect(reopened.archive.readText(path)).toBe(updated)
  })

  it('maps Mac Layout 2 to the native top legend, hidden value axis, and three series labels', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    const native2 = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-02-chart.xml',
        import.meta.url,
      ),
      'utf8',
    )
    const opened = await openPptx(source)
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    const dependencies = new Map(
      [...opened.archive.entries].filter(([name]) =>
        /^(?:ppt\/charts\/(?:_rels\/|style|colors)|ppt\/embeddings\/)/.test(name),
      ),
    )
    const series = (xml: string) => [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((m) => m[0])
    const label = (xml: string) => /<c:dLbls>[\s\S]*?<\/c:dLbls>/.exec(xml)?.[0]
    const valueAxis = (xml: string) => /<c:valAx>[\s\S]*?<\/c:valAx>/.exec(xml)?.[0]
    expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-2')).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'mac-2')).toBe(true)
    const updated = opened.archive.readText(path)!
    expect(series(updated).map(label)).toEqual(series(native2).map(label))
    expect(series(updated).map((fragment) => fragment.replace(label(fragment)!, ''))).toEqual(
      series(original),
    )
    expect(updated).toContain('<c:gapWidth val="150"/>')
    expect(updated).toContain('<c:overlap val="-25"/>')
    expect(updated).toContain('<c:legendPos val="t"/>')
    expect(valueAxis(updated)).toContain('<c:delete val="1"/>')
    expect(valueAxis(updated)).not.toContain('<c:majorGridlines>')
    expect(
      valueAxis(updated)
        ?.replace('<c:delete val="1"/>', '')
        .replace(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/, ''),
    ).toBe(
      valueAxis(original)
        ?.replace('<c:delete val="0"/>', '')
        .replace(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/, ''),
    )
    expect(updated).toContain('<c:externalData r:id="rId3"')
    for (const [name, content] of dependencies)
      expect(opened.archive.entries.get(name)).toEqual(content)
    expect(chart.chart.valAxis?.hidden).toBe(true)
    expect(chart.chart.series.map((item) => item.dataLabels)).toEqual([true, true, true])
    const reopened = await openPptx(await savePptx(opened))
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restored.chart.valAxis?.hidden).toBe(true)
    expect(restored.chart.gapWidthPct).toBe(150)
    expect(restored.chart.overlapPct).toBe(-25)
    expect(canSetChartQuickLayout(reopened, 0, restored.id, 'mac-2')).toBe(true)
    expect(setChartQuickLayout(reopened, 0, restored.id, 'mac-2')).toBe(true)
    expect(reopened.archive.readText(path)).toBe(updated)
  })

  it('fails closed for Layout 2 when labels, gridlines, or axis visibility have unknown XML', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    for (const mutate of [
      (xml: string) =>
        xml.replace(
          '</c:ser>',
          '<c:dLbls><c:dLblPos val="outEnd"/><c:showVal val="1"/></c:dLbls></c:ser>',
        ),
      (xml: string) =>
        xml.replace('</c:majorGridlines>', '<c:extLst data-preserve="1"/></c:majorGridlines>'),
      (xml: string) =>
        xml.replace(
          /(<c:valAx>[\s\S]*?)<c:delete val="0"\/>/,
          '$1<c:delete val="0" data-preserve="1"/>',
        ),
    ]) {
      const opened = await openPptx(source)
      const chart = opened.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      const path = 'ppt/charts/chart1.xml'
      const altered = mutate(opened.archive.readText(path)!)
      expect(altered).not.toBe(opened.archive.readText(path))
      opened.archive.entries.set(path, Buffer.from(altered, 'utf8'))
      expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-2')).toBe(false)
      expect(setChartQuickLayout(opened, 0, chart.id, 'mac-2')).toBe(false)
      expect(opened.archive.readText(path)).toBe(altered)
    }
  })

  it('fails closed for incompatible series labels, unknown gridlines, or linked title in Layout 4', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    for (const mutate of [
      (xml: string) =>
        xml.replace(
          '</c:ser>',
          '<c:dLbls><c:dLblPos val="inEnd"/><c:showVal val="1"/></c:dLbls></c:ser>',
        ),
      (xml: string) =>
        xml.replace('</c:majorGridlines>', '<c:extLst data-preserve="1"/></c:majorGridlines>'),
      (xml: string) => xml.replace('</c:title>', '<c:extLst r:id="rId88"/></c:title>'),
    ]) {
      const opened = await openPptx(source)
      const chart = opened.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      const path = 'ppt/charts/chart1.xml'
      const altered = mutate(opened.archive.readText(path)!)
      opened.archive.entries.set(path, Buffer.from(altered, 'utf8'))
      expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-4')).toBe(false)
      expect(setChartQuickLayout(opened, 0, chart.id, 'mac-4')).toBe(false)
      expect(opened.archive.readText(path)).toBe(altered)
    }
  })

  it('refuses Mac Layout 10 when third-series labels would be overwritten', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    for (const mutate of [
      (xml: string) => xml.replace('</c:ser>', '<c:dLbls><c:showVal val="1"/></c:dLbls></c:ser>'),
      (xml: string) =>
        xml.replace(
          /(<c:ser>[\s\S]*?<\/c:ser>[\s\S]*?<c:ser>[\s\S]*?<\/c:ser>[\s\S]*?<c:ser>)/,
          '$1<c:dLbls><c:showVal val="1"/></c:dLbls>',
        ),
      (xml: string) => xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/, ''),
    ]) {
      const opened = await openPptx(source)
      const chart = opened.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      const path = 'ppt/charts/chart1.xml'
      const altered = mutate(opened.archive.readText(path)!)
      opened.archive.entries.set(path, Buffer.from(altered, 'utf8'))
      expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-10')).toBe(false)
      expect(setChartQuickLayout(opened, 0, chart.id, 'mac-10')).toBe(false)
      expect(opened.archive.readText(path)).toBe(altered)
    }
  })

  it('fails closed for a linked or extension-bearing title and unsupported axis ticks in Layout 11', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    for (const change of [
      (xml: string) => xml.replace('</c:title>', '<c:extLst/></c:title>'),
      (xml: string) =>
        xml.replace(
          '</c:title>',
          '<c:tx><c:strRef><c:f>Sheet1!A1</c:f></c:strRef></c:tx></c:title>',
        ),
      (xml: string) =>
        xml.replace(
          '<c:majorTickMark val="none"/>',
          '<c:majorTickMark val="none" data-preserve="x"/>',
        ),
      (xml: string) => xml.replace('<c:majorTickMark val="none"/>', ''),
    ]) {
      const opened = await openPptx(source)
      const chart = opened.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      const path = 'ppt/charts/chart1.xml'
      const altered = change(opened.archive.readText(path)!)
      opened.archive.entries.set(path, Buffer.from(altered, 'utf8'))
      expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-11')).toBe(false)
      expect(setChartQuickLayout(opened, 0, chart.id, 'mac-11')).toBe(false)
      expect(opened.archive.readText(path)).toBe(altered)
    }
  })

  it('fails closed when a Mac column layout would need title, gridline, or label reconstruction', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    for (const change of [
      (xml: string) => xml.replace(/<c:title>[\s\S]*?<\/c:title>/, ''),
      (xml: string) => xml.replace(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/, ''),
      (xml: string) => xml.replace('<c:grouping val="clustered"/>', '<c:grouping val="stacked"/>'),
      (xml: string) => xml.replace('</c:plotArea>', '<c:dTable/></c:plotArea>'),
      (xml: string) => xml.replace('<c:showSerName val="0"/>', '<c:showSerName val="1"/>'),
      (xml: string) => xml.replace('</c:majorGridlines>', '</c:majorGridlines><c:minorGridlines/>'),
      (xml: string) =>
        xml.replace('<c:overlap val="-27"/>', '<c:overlap val="-27" data-preserve="x"/>'),
    ]) {
      const opened = await openPptx(source)
      const chart = opened.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      const path = 'ppt/charts/chart1.xml'
      const altered = change(opened.archive.readText(path)!)
      opened.archive.entries.set(path, Buffer.from(altered, 'utf8'))
      expect(canSetChartQuickLayout(opened, 0, chart.id, 'mac-1')).toBe(false)
      expect(setChartQuickLayout(opened, 0, chart.id, 'mac-1')).toBe(false)
      expect(opened.archive.readText(path)).toBe(altered)
    }
  })

  it('edits the PowerPoint Mac chart without rewriting its workbook, style, colors, or extensions', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
        import.meta.url,
      ),
    )
    const sourceZip = await JSZip.loadAsync(source)
    const opened = await openPptx(source)
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const chartPath = 'ppt/charts/chart1.xml'
    const original = await sourceZip.file(chartPath)!.async('string')
    const preservedPaths = [
      'ppt/charts/_rels/chart1.xml.rels',
      'ppt/charts/style1.xml',
      'ppt/charts/colors1.xml',
      'ppt/embeddings/Microsoft_Excel____.xlsx',
    ]
    expect(chart).toBeTruthy()
    expect(chart.chart.legendPos).toBe('b')
    expect(chart.chart.dataLabels).toBeUndefined()
    expect(canSetChartQuickLayout(opened, 0, chart.id)).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'right-values')).toBe(true)
    const saved = await savePptx(opened)
    const savedZip = await JSZip.loadAsync(saved)
    const updated = await savedZip.file(chartPath)!.async('string')
    const preservedFragments = (xml: string, name: string) =>
      [...xml.matchAll(new RegExp(`<${name}\\b[\\s\\S]*?<\\/${name}>`, 'g'))].map(
        (match) => match[0],
      )
    expect(updated).toContain('<c:legendPos val="r"/>')
    expect(updated).toMatch(/<c:dLbls>[\s\S]*?<c:showVal val="1"\/>[\s\S]*?<\/c:dLbls>/)
    expect(updated).toContain('<c:extLst>')
    expect(updated).toContain('<c:externalData r:id="rId3"')
    for (const name of ['c:ser', 'c:extLst', 'c:catAx', 'c:valAx', 'mc:AlternateContent']) {
      expect(preservedFragments(updated, name)).toEqual(preservedFragments(original, name))
    }
    for (const path of preservedPaths) {
      expect(await savedZip.file(path)!.async('uint8array')).toEqual(
        await sourceZip.file(path)!.async('uint8array'),
      )
    }
    const reopened = await openPptx(saved)
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restored.chart.legendPos).toBe('r')
    expect(restored.chart.dataLabels).toBe(true)
    expect(setChartQuickLayout(reopened, 0, restored.id, 'values-only')).toBe(true)
    const secondZip = await JSZip.loadAsync(await savePptx(reopened))
    const second = await secondZip.file(chartPath)!.async('string')
    expect(second).not.toMatch(/<c:legend\b/)
    expect(second).toContain('<c:showVal val="1"/>')
    expect(second).toContain('<c:extLst>')
    for (const name of ['c:ser', 'c:extLst', 'c:catAx', 'c:valAx', 'mc:AlternateContent']) {
      expect(preservedFragments(second, name)).toEqual(preservedFragments(original, name))
    }
    for (const path of preservedPaths) {
      expect(await secondZip.file(path)!.async('uint8array')).toEqual(
        await sourceZip.file(path)!.async('uint8array'),
      )
    }
    expect(original).not.toBe(updated)
  })
  it('applies value-label Quick Layouts by patching only native legend and plot labels', async () => {
    const { opened, chart } = await importedChart()
    const before = opened.archive.readText('ppt/charts/chart1.xml')!
    const unrelated = new Map(
      [...opened.archive.entries].filter(([path]) => path !== 'ppt/charts/chart1.xml'),
    )
    expect(canSetChartQuickLayout(opened, 0, chart.id)).toBe(true)
    expect(setChartQuickLayout(opened, 0, chart.id, 'right-values')).toBe(true)
    const first = opened.archive.readText('ppt/charts/chart1.xml')!
    expect(first).toContain('<c:legendPos val="r"/>')
    expect(first).toMatch(/<c:barChart>[\s\S]*?<c:dLbls><c:showVal val="1"\/><\/c:dLbls>/)
    expect(first).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(first).toContain('data-preserve="chart-space"')
    expect(
      first
        .replace('<c:dLbls><c:showVal val="1"/></c:dLbls>', '')
        .replace('<c:legendPos val="r"/>', '<c:legendPos val="b"/>'),
    ).toBe(before)
    expect(setChartQuickLayout(opened, 0, chart.id, 'values-only')).toBe(true)
    const bytes = await savePptx(opened)
    const zip = await JSZip.loadAsync(bytes)
    const after = await zip.file('ppt/charts/chart1.xml')!.async('string')
    expect(after).not.toContain('<c:legend>')
    expect(after).toContain('<c:showVal val="1"/>')
    for (const [path, content] of unrelated)
      expect(opened.archive.entries.get(path)).toEqual(content)
    const reopened = await openPptx(bytes)
    const restored = reopened.deck.slides[0]!.elements.find(
      (item) => item.type === 'chart',
    ) as ChartElement
    expect(restored.chart.dataLabels).toBe(true)
    expect(restored.chart.legendPos).toBeUndefined()
    expect(setChartQuickLayout(reopened, 0, restored.id, 'bottom-values')).toBe(true)
    expect(reopened.archive.readText('ppt/charts/chart1.xml')).toContain('<c:legendPos val="b"/>')
  })

  it('fails closed for series-local labels without changing bytes', async () => {
    const { opened, chart } = await importedChart()
    const original = opened.archive.readText('ppt/charts/chart1.xml')!
    const withSeriesLabels = original.replace(
      '</c:ser>',
      '<c:dLbls><c:showVal val="0"/></c:dLbls></c:ser>',
    )
    opened.archive.entries.set('ppt/charts/chart1.xml', Buffer.from(withSeriesLabels, 'utf8'))
    expect(canSetChartQuickLayout(opened, 0, chart.id)).toBe(false)
    expect(setChartQuickLayout(opened, 0, chart.id, 'right-values')).toBe(false)
    expect(opened.archive.readText('ppt/charts/chart1.xml')).toBe(withSeriesLabels)
  })

  it('updates an existing plot label flag while keeping label formatting and extensions', async () => {
    const { opened, chart } = await importedChart()
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    const labels =
      '<c:dLbls><c:txPr data-preserve="label-font"/><c:showVal val="0"/><c:extLst data-preserve="label-extension"/></c:dLbls>'
    opened.archive.entries.set(
      path,
      Buffer.from(original.replace('</c:ser><c:axId', `</c:ser>${labels}<c:axId`), 'utf8'),
    )
    expect(setChartQuickLayout(opened, 0, chart.id, 'bottom-values')).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain(labels.replace('val="0"', 'val="1"'))
    expect(xml).toContain('data-preserve="chart-space"')
  })

  it('keeps an embedded workbook relationship and unknown chart extensions byte-for-byte', async () => {
    const { opened, chart } = await importedChart()
    const chartPath = 'ppt/charts/chart1.xml'
    const relsPath = 'ppt/charts/_rels/chart1.xml.rels'
    const workbookPath = 'ppt/embeddings/QuickLayout.xlsx'
    const workbook = new JSZip()
    workbook.file(
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    )
    workbook.file(
      'xl/workbook.xml',
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
    )
    const workbookBytes = await workbook.generateAsync({ type: 'uint8array' })
    const rels =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/QuickLayout.xlsx"/></Relationships>'
    opened.archive.entries.set(workbookPath, workbookBytes)
    opened.archive.entries.set(relsPath, Buffer.from(rels, 'utf8'))
    const contentTypes = opened.archive.readText('[Content_Types].xml')!
    opened.archive.entries.set(
      '[Content_Types].xml',
      Buffer.from(
        contentTypes.replace(
          '</Types>',
          '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>',
        ),
        'utf8',
      ),
    )
    opened.archive.entries.set(
      chartPath,
      Buffer.from(
        opened.archive
          .readText(chartPath)!
          .replace(
            '<c:extLst data-preserve="chart-space"/>',
            '<c:externalData r:id="rIdWorkbook"><c:autoUpdate val="0"/></c:externalData><c:extLst data-preserve="chart-space"/>',
          ),
        'utf8',
      ),
    )
    const beforeParts = new Map([...opened.archive.entries].filter(([path]) => path !== chartPath))
    expect(setChartQuickLayout(opened, 0, chart.id, 'right-values')).toBe(true)
    const saved = await JSZip.loadAsync(await savePptx(opened))
    expect(await saved.file(workbookPath)?.async('uint8array')).toEqual(workbookBytes)
    expect(await saved.file(relsPath)?.async('string')).toBe(rels)
    const xml = await saved.file(chartPath)!.async('string')
    expect(xml).toContain('<c:externalData r:id="rIdWorkbook">')
    expect(xml).toContain('data-preserve="chart-space"')
    for (const [path, content] of beforeParts)
      expect(opened.archive.entries.get(path)).toEqual(content)
  })

  it.each(['bar', 'line', 'pie'] as const)(
    'writes value labels for a single %s plot',
    async (kind) => {
      const opened = await openPptx(await createBlankPptx())
      addChart(opened, 0, {
        kind,
        categories: ['A', 'B'],
        series: [{ name: 'Values', values: [2, 5] }],
        offset: { x: 914400, y: 914400, cx: 4572000, cy: 2743200 },
        legendPos: 'b',
      })
      const chart = opened.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      expect(canSetChartQuickLayout(opened, 0, chart.id)).toBe(true)
      expect(setChartQuickLayout(opened, 0, chart.id, 'values-only')).toBe(true)
      const reloaded = await openPptx(await savePptx(opened))
      const nativeXml = reloaded.archive.readText('ppt/charts/chart1.xml')!
      expect(nativeXml).toContain('<c:dLbls><c:showVal val="1"/></c:dLbls>')
      expect(nativeXml).not.toContain('<c:legend>')
      const model = reloaded.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      expect(model.chart.dataLabels).toBe(true)
    },
  )

  it('moves imported legend without rebuilding series, data cache, styled entries, or extensions', async () => {
    const { opened, chart } = await importedChart()
    expect(chart.descr).not.toBe('aislides-chart')
    const before = opened.archive.readText('ppt/charts/chart1.xml')!
    expect(setChartLegendPosition(opened, 0, chart.id, 'r')).toBe(true)
    const zip = await JSZip.loadAsync(await savePptx(opened))
    const after = await zip.file('ppt/charts/chart1.xml')!.async('string')
    expect(after).toBe(before.replace('<c:legendPos val="b"/>', '<c:legendPos val="r"/>'))
    expect(after).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(after).toContain('data-preserve="chart-space"')
    const reopened = await openPptx(await savePptx(opened))
    const refreshed = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(refreshed.chart.legendPos).toBe('r')
    expect(refreshed.descr).not.toBe('aislides-chart')
  })

  it('removes only the native legend for None and keeps every other package part intact', async () => {
    const { opened, chart } = await importedChart()
    const chartPath = 'ppt/charts/chart1.xml'
    // A synthetic external workbook relation exercises the imported-chart
    // package boundary without presenting this fixture as Office-origin.
    const book = new JSZip()
    book.file(
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    )
    book.file(
      '_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    )
    book.file(
      'xl/workbook.xml',
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    book.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    )
    book.file(
      'xl/worksheets/sheet1.xml',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>4</v></c></row></sheetData></worksheet>',
    )
    const workbookBytes = await book.generateAsync({ type: 'uint8array' })
    const workbookPath = 'ppt/embeddings/Workbook.xlsx'
    const chartRelsPath = 'ppt/charts/_rels/chart1.xml.rels'
    const chartRels =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Workbook.xlsx"/></Relationships>'
    opened.archive.entries.set(workbookPath, workbookBytes)
    opened.archive.entries.set(chartRelsPath, Buffer.from(chartRels, 'utf8'))
    const contentTypes = opened.archive.readText('[Content_Types].xml')!
    opened.archive.entries.set(
      '[Content_Types].xml',
      Buffer.from(
        contentTypes.replace(
          '</Types>',
          '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>',
        ),
        'utf8',
      ),
    )
    opened.archive.entries.set(
      chartPath,
      Buffer.from(
        opened.archive
          .readText(chartPath)!
          .replace(
            '<c:extLst data-preserve="chart-space"/>',
            '<c:externalData r:id="rIdWorkbook"><c:autoUpdate val="0"/></c:externalData><c:extLst data-preserve="chart-space"/>',
          ),
        'utf8',
      ),
    )
    const before = opened.archive.readText(chartPath)!
    const legend = /<c:legend>[\s\S]*?<\/c:legend>/.exec(before)![0]
    const otherParts = new Map([...opened.archive.entries].filter(([path]) => path !== chartPath))

    expect(setChartLegendPosition(opened, 0, chart.id, 'none')).toBe(true)
    const out = await savePptx(opened)
    const zip = await JSZip.loadAsync(out)
    const after = await zip.file(chartPath)!.async('string')
    expect(after).toBe(before.replace(legend, ''))
    expect(after).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(after).toContain('data-preserve="chart-space"')
    expect(after).toContain('<c:externalData r:id="rIdWorkbook">')
    expect(await zip.file(workbookPath)?.async('uint8array')).toEqual(workbookBytes)
    expect(await zip.file(chartRelsPath)?.async('string')).toBe(chartRels)
    for (const [path, bytes] of otherParts) {
      expect(opened.archive.entries.get(path)).toEqual(bytes)
    }

    const reopened = await openPptx(out)
    const refreshed = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(refreshed.chart.legendPos).toBeUndefined()
    expect(setChartLegendPosition(reopened, 0, refreshed.id, 'r')).toBe(true)
    const restored = await openPptx(await savePptx(reopened))
    const restoredChart = restored.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(restoredChart.chart.legendPos).toBe('r')
    expect(restored.archive.readText(chartPath)).toContain('<c:legendPos val="r"/>')
    expect(restored.archive.readText(chartPath)).toContain('data-preserve="chart-space"')
  })

  it('changes chart-area fill and outline without touching imported line dash, formula, or extensions', async () => {
    const { opened, chart } = await importedChart()
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    opened.archive.entries.set(
      path,
      Buffer.from(
        original.replace(
          '</c:chart>',
          '</c:chart><c:spPr><a:gradFill data-original="gradient"/><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="dash"/></a:ln><a:effectLst data-preserve="shadow"/></c:spPr>',
        ),
        'utf8',
      ),
    )
    expect(
      setChartAreaFormat(opened, 0, chart.id, {
        fill: '#D9EAF7',
        stroke: { color: '#4472C4', widthEmu: 38100 },
      }),
    ).toBe(true)
    const out = await savePptx(opened)
    const zip = await JSZip.loadAsync(out)
    const xml = await zip.file(path)!.async('string')
    expect(xml).toContain('<a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>')
    expect(xml).toContain(
      '<a:ln w="38100"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:prstDash val="dash"/></a:ln>',
    )
    expect(xml).not.toContain('data-original="gradient"')
    expect(xml).toContain('<a:effectLst data-preserve="shadow"/>')
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(xml).toContain('data-preserve="chart-space"')
    const reopened = await openPptx(out)
    const refreshed = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(refreshed.chart.bgFill).toMatchObject({ type: 'solid', color: '#D9EAF7' })
    expect(refreshed.chart.border?.color).toBe('#4472C4')
    expect(refreshed.chart.border?.widthEmu).toBe(38100)
  })

  it('formats plot-area shape properties without changing series or chart-area formatting', async () => {
    const { opened, chart } = await importedChart()
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    opened.archive.entries.set(
      path,
      Buffer.from(
        original.replace(
          '</c:plotArea>',
          '<c:spPr><a:gradFill data-original="plot-gradient"/><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="dash"/></a:ln><a:effectLst data-preserve="plot-shadow"/></c:spPr><c:extLst data-preserve="plot-extension"/></c:plotArea>',
        ),
        'utf8',
      ),
    )
    expect(
      setChartAreaFormat(opened, 0, chart.id, {
        area: 'plot',
        fill: '#D9EAD3',
        stroke: { color: '#C65911', widthEmu: 19050 },
      }),
    ).toBe(true)
    const out = await savePptx(opened)
    const zip = await JSZip.loadAsync(out)
    const xml = await zip.file(path)!.async('string')
    expect(xml).toContain(
      '<c:spPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="C65911"/></a:solidFill><a:prstDash val="dash"/></a:ln><a:effectLst data-preserve="plot-shadow"/></c:spPr><c:extLst data-preserve="plot-extension"/>',
    )
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(xml).toContain('data-preserve="chart-space"')
    const reopened = await openPptx(out)
    const refreshed = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(refreshed.chart.plotFill).toMatchObject({ type: 'solid', color: '#D9EAD3' })
    expect(refreshed.chart.plotBorder).toMatchObject({ color: '#C65911', widthEmu: 19050 })
    expect(refreshed.chart.bgFill).toBeUndefined()
  })
})
