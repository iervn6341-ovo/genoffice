import { describe, expect, it } from 'vitest'
import {
  addChart,
  createBlankPptx,
  editChartElement,
  openPptx,
  setChartAxisVisibility,
} from '../src/index'
import type { ChartElement } from '../src/types'

async function withChart(kind: 'bar' | 'pie' = 'bar') {
  const opened = await openPptx(await createBlankPptx())
  addChart(opened, 0, {
    kind,
    categories: ['A', 'B'],
    series: [{ name: 'Revenue', values: [4, 7] }],
    offset: { x: 914400, y: 914400, cx: 4572000, cy: 2743200 },
    legendPos: 'b',
  })
  const chart = opened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
  return { opened, chart }
}

const del = (xml: string, tag: 'catAx' | 'valAx') =>
  new RegExp(`<c:${tag}>.*?<c:delete val="(\\d)"/>`, 's').exec(xml)?.[1]

describe('chart axis visibility', () => {
  it('hides and re-shows only the requested axis', async () => {
    const { opened, chart } = await withChart()
    const path = 'ppt/charts/chart1.xml'
    expect(setChartAxisVisibility(opened, 0, chart.id, { val: false })).toBe(true)
    let xml = opened.archive.readText(path)!
    expect(del(xml, 'valAx')).toBe('1')
    expect(del(xml, 'catAx')).toBe('0')
    expect(setChartAxisVisibility(opened, 0, chart.id, { cat: false, val: true })).toBe(true)
    xml = opened.archive.readText(path)!
    expect(del(xml, 'valAx')).toBe('0')
    expect(del(xml, 'catAx')).toBe('1')
  })

  it('keeps a hidden axis hidden across a data rebuild', async () => {
    const { opened, chart } = await withChart()
    setChartAxisVisibility(opened, 0, chart.id, { val: false })
    // re-parse so the model sees the hidden axis, as the app does after each edit
    const reopened = await openPptx(await (await import('../src/index')).savePptx(opened))
    const again = reopened.deck.slides[0]!.elements.find((el) => el.type === 'chart') as ChartElement
    expect(again.chart.valAxis?.hidden).toBe(true)
    expect(
      editChartElement(reopened, 0, again.id, { title: 'Renamed' }),
    ).toBe(true)
    expect(del(reopened.archive.readText('ppt/charts/chart1.xml')!, 'valAx')).toBe('1')
  })

  it('rejects a chart without axes', async () => {
    const { opened, chart } = await withChart('pie')
    expect(setChartAxisVisibility(opened, 0, chart.id, { val: false })).toBe(false)
  })
})
