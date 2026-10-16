import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { materializeSlide, openPptx, type ChartElement } from '@genoffice/pptx-engine'
import { runTxn } from '../src/index'

const officeFixture = new URL(
  '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
  import.meta.url,
)

describe('PowerPoint Mac Quick Layout operation', () => {
  it('validates and applies Layouts 1 and 3 through the runtime op payload', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    const original = opened.archive.readText(path)!
    expect(
      runTxn(opened, {
        dryRun: true,
        ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-1' }],
      }).failures,
    ).toBeUndefined()
    expect(opened.archive.readText(path)).toBe(original)
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-1' }] }).applied,
    ).toBe(true)
    expect(opened.archive.readText(path)).toContain('<c:legendPos val="r"/>')
    expect(
      (opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement).chart
        .gapWidthPct,
    ).toBe(150)
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.gapWidthPct).toBe(150)
    target.el = reparsed.id
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-3' }] }).applied,
    ).toBe(true)
    expect(opened.archive.readText(path)).toContain('<c:legendPos val="b"/>')
    expect(opened.archive.readText(path)).toContain('<c:gapWidth val="75"/>')
    expect(opened.archive.readText(path)).toContain('<c:overlap val="-25"/>')
    expect(opened.archive.readText(path)).toContain('<c:externalData r:id="rId3"')
  })

  it('rejects Layout 1 when the native major gridlines are absent without changing the chart', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const path = 'ppt/charts/chart1.xml'
    const altered = opened.archive
      .readText(path)!
      .replace(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/, '')
    opened.archive.entries.set(path, Buffer.from(altered, 'utf8'))
    const result = runTxn(opened, {
      ops: [{ op: 'setChartQuickLayout', target: { slide: 0, el: chart.id }, layout: 'mac-1' }],
    })
    expect(result.applied).toBe(false)
    expect(result.failures?.[0]?.error).toContain('native layout requirements')
    expect(opened.archive.readText(path)).toBe(altered)
  })

  it('applies Layout 11 through the validated transaction and survives slide materialization', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    expect(
      runTxn(opened, {
        dryRun: true,
        ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-11' }],
      }).failures,
    ).toBeUndefined()
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-11' }] }).applied,
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('<c:autoTitleDeleted val="1"/>')
    expect(xml).not.toMatch(/<c:title\b/)
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.title).toBeUndefined()
    expect(reparsed.chart.gapWidthPct).toBe(150)
    expect(reparsed.chart.legendPos).toBe('r')
    expect(
      runTxn(opened, {
        ops: [
          { op: 'setChartQuickLayout', target: { slide: 0, el: reparsed.id }, layout: 'mac-11' },
        ],
      }).applied,
    ).toBe(true)
    expect(opened.archive.readText(path)).toBe(xml)
  })

  it('applies Layout 9 through the validated transaction with automatic axis titles', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    expect(
      runTxn(opened, {
        dryRun: true,
        ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-9' }],
      }).failures,
    ).toBeUndefined()
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-9' }] }).applied,
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('<c:legendPos val="r"/>')
    expect((xml.match(/<c:title>/g) ?? []).length).toBe(3)
    expect((xml.match(/<c:majorTickMark val="out"\/>/g) ?? []).length).toBe(1)
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.catAxis?.title).toBe('Axis Title')
    expect(reparsed.chart.valAxis?.title).toBe('Axis Title')
    expect(reparsed.chart.gapWidthPct).toBe(150)
  })

  it('validates and materializes Mac Layout 10 series-local labels', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    expect(
      runTxn(opened, {
        dryRun: true,
        ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-10' }],
      }).failures,
    ).toBeUndefined()
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-10' }] }).applied,
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('<c:overlap val="40"/>')
    expect(xml).toContain('<c:gapWidth val="75"/>')
    expect((xml.match(/<c:dLblPos val="inEnd"\/>/g) ?? []).length).toBe(1)
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.series.map((item) => item.dataLabels)).toEqual([false, false, true])
  })

  it('validates and materializes Mac Layout 7 minor gridlines and automatic axis titles', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    expect(
      runTxn(opened, {
        dryRun: true,
        ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-7' }],
      }).failures,
    ).toBeUndefined()
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-7' }] }).applied,
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('<c:gapWidth val="300"/>')
    expect(xml).toContain('<c:minorGridlines>')
    expect(xml).toContain('<c:autoTitleDeleted val="1"/>')
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.title).toBeUndefined()
    expect(reparsed.chart.catAxis?.title).toBe('Axis Title')
    expect(reparsed.chart.valAxis?.title).toBe('Axis Title')
    expect(reparsed.chart.valAxis?.minorGridColor).toBeTruthy()
  })

  it('validates and materializes Mac Layout 8 with no major gridlines or legend', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    expect(
      runTxn(opened, {
        dryRun: true,
        ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-8' }],
      }).failures,
    ).toBeUndefined()
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-8' }] }).applied,
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('<c:gapWidth val="0"/>')
    expect(xml).not.toMatch(/<c:majorGridlines\b|<c:legend\b/)
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.legendPos).toBeUndefined()
    expect(reparsed.chart.valAxis?.gridColor).toBeUndefined()
    expect(reparsed.chart.catAxis?.title).toBe('Axis Title')
    expect(reparsed.chart.valAxis?.title).toBe('Axis Title')
  })

  it('validates and materializes Mac Layout 4 labels on every series', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    expect(
      runTxn(opened, {
        dryRun: true,
        ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-4' }],
      }).failures,
    ).toBeUndefined()
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-4' }] }).applied,
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('<c:gapWidth val="75"/>')
    expect((xml.match(/<c:showVal val="1"\/>/g) ?? []).length).toBe(4)
    expect(xml).not.toMatch(/<c:majorGridlines\b/)
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.series.map((item) => item.dataLabels)).toEqual([true, true, true])
    expect(reparsed.chart.title).toBeUndefined()
    expect(reparsed.chart.valAxis?.gridColor).toBeUndefined()
  })

  it('validates and materializes Mac Layout 2 with a hidden native value axis', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    expect(
      runTxn(opened, {
        dryRun: true,
        ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-2' }],
      }).failures,
    ).toBeUndefined()
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-2' }] }).applied,
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('<c:gapWidth val="150"/>')
    expect(xml).toContain('<c:overlap val="-25"/>')
    expect(xml).toContain('<c:legendPos val="t"/>')
    expect(xml).toMatch(/<c:valAx>[\s\S]*?<c:delete val="1"\/>[\s\S]*?<\/c:valAx>/)
    expect((xml.match(/<c:showVal val="1"\/>/g) ?? []).length).toBe(4)
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.valAxis?.hidden).toBe(true)
    expect(reparsed.chart.series.map((item) => item.dataLabels)).toEqual([true, true, true])
  })

  const macDir = new URL(
    '../../../docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/',
    import.meta.url,
  )
  const normalizeLang = (xml: string) => xml.replace(/lang="[^"]*"/g, 'lang="x"')
  const piece = (xml: string, tag: string) =>
    normalizeLang(new RegExp(`<${tag}\\b[\\s\\S]*?</${tag}>`).exec(xml)?.[0] ?? '')

  it('applies Mac Layout 5 (value-axis title, data table, no legend) matching the native sample', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    const result = runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-5' }] })
    expect(result.applied).toBe(true)
    const xml = opened.archive.readText(path)!
    const native = readFileSync(new URL('layout-05-chart.xml', macDir), 'utf8')
    expect(piece(xml, 'c:dTable')).toBe(piece(native, 'c:dTable'))
    const valAx = (x: string) => /<c:valAx>[\s\S]*?<\/c:valAx>/.exec(x)![0]
    expect(piece(valAx(xml), 'c:title')).toBe(piece(valAx(native), 'c:title'))
    expect(xml).not.toMatch(/<c:legend\b/)
    expect(xml).toContain('<c:externalData r:id="rId3"')
    const catAx = /<c:catAx>[\s\S]*?<\/c:catAx>/.exec(xml)![0]
    expect(catAx).not.toContain('<c:title')
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.dataTable).toBe(true)
    expect(reparsed.chart.valAxis?.title).toBe('Axis Title')
    expect(reparsed.chart.legendPos).toBeUndefined()
    // a second application is refused instead of duplicating the data table
    target.el = reparsed.id
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-5' }] }).applied,
    ).toBe(false)
  })

  it('applies Mac Layout 6 (last-point value + series name labels, value-axis title, no legend)', async () => {
    const opened = await openPptx(readFileSync(officeFixture))
    const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    const target = { slide: 0, el: chart.id }
    const path = 'ppt/charts/chart1.xml'
    expect(
      runTxn(opened, { ops: [{ op: 'setChartQuickLayout', target, layout: 'mac-6' }] }).applied,
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    const native = readFileSync(new URL('layout-06-chart.xml', macDir), 'utf8')
    const labels = (x: string) =>
      [...x.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((m) => {
        const dl = /<c:dLbls>[\s\S]*?<\/c:dLbls>/.exec(m[0])![0]
        // the native sample also carries an empty c15 ext and a c16 uniqueId inside c:dLbl
        return normalizeLang(
          dl.replace(/<c:dLbl>[\s\S]*?<\/c:dLbl>/, (d) =>
            d.replace(/<c:extLst>[\s\S]*<\/c:extLst>/, ''),
          ),
        )
      })
    expect(labels(xml)).toEqual(labels(native))
    expect((xml.match(/<c:idx val="3"\/><c:showLegendKey/g) ?? []).length).toBe(3)
    expect(xml).not.toMatch(/<c:legend\b|<c:dTable\b/)
    const valAx = (x: string) => /<c:valAx>[\s\S]*?<\/c:valAx>/.exec(x)![0]
    expect(piece(valAx(xml), 'c:title')).toBe(piece(valAx(native), 'c:title'))
    const reparsed = materializeSlide(opened, 0)!.elements.find(
      (el) => el.type === 'chart',
    ) as ChartElement
    expect(reparsed.chart.pointValueLabels).toBe(true)
    expect(reparsed.chart.series.map((item) => item.dataLabels)).toEqual([false, false, false])
  })

  it('rejects Layouts 5 and 6 when a series already has labels', async () => {
    for (const layout of ['mac-5', 'mac-6']) {
      const opened = await openPptx(readFileSync(officeFixture))
      const chart = opened.deck.slides[0]!.elements.find(
        (el) => el.type === 'chart',
      ) as ChartElement
      const path = 'ppt/charts/chart1.xml'
      const altered = opened.archive
        .readText(path)!
        .replace(/(<c:ser>[\s\S]*?)(<c:cat>)/, '$1<c:dLbls><c:showVal val="1"/></c:dLbls>$2')
      opened.archive.entries.set(path, Buffer.from(altered, 'utf8'))
      const result = runTxn(opened, {
        ops: [{ op: 'setChartQuickLayout', target: { slide: 0, el: chart.id }, layout }],
      })
      expect(result.applied).toBe(false)
      expect(opened.archive.readText(path)).toBe(altered)
    }
  })
})
