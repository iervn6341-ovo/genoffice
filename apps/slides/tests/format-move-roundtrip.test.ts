import { expect, it } from 'vitest'
import { mkdir, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import {
  addElement,
  createBlankPptx,
  groupElements,
  openPptx,
  savePptx,
} from '@genoffice/pptx-engine'
import { buildRenderSlide } from '@genoffice/pptx-render'
import { runTxn } from '@genoffice/pptx-ops'
import { buildBrushOps, extractBrushFormat } from '../src/renderer/format-brush'

const EMU = 9525
it('paints matching group children through native IDs and preserves the target group text and layout', async () => {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  slide.elements = []
  const shapes = [0, 1, 2, 3].map((i) =>
    addElement(slide, {
      kind: 'rect',
      offset: { x: i * 120 * EMU, y: 80 * EMU, cx: 100 * EMU, cy: 70 * EMU },
      fillColor: ['#ff0000', '#00ff00', '#eeeeee', '#eeeeee'][i],
      paragraphs: [{ runs: [{ text: `Object ${i}`, fontSize: 18 + i, bold: i === 0 }] }],
    }),
  )
  expect(
    groupElements(
      opened,
      0,
      shapes.slice(0, 2).map((s) => s.id),
    ),
  ).toBeTruthy()
  expect(
    groupElements(
      opened,
      0,
      opened.deck.slides[0]!.elements.filter((el) => el.type !== 'group').map((el) => el.id),
    ),
  ).toBeTruthy()
  const render = buildRenderSlide(opened.deck.slides[0]!, opened.deck.size, { fitWidthPx: 960 })
  const [sourceNode, targetNode] = render.nodes.filter((n) => n.type === 'group')
  if (!sourceNode || !targetNode) throw new Error('missing test groups')
  const beforeBox = structuredClone(targetNode.box)
  const result = runTxn(opened, {
    ops: buildBrushOps(extractBrushFormat(sourceNode, render.scale)!, targetNode, 0),
  })
  expect(result.failures).toBeUndefined()
  const reopened = await openPptx(await savePptx(opened))
  const painted = buildRenderSlide(reopened.deck.slides[0]!, reopened.deck.size, {
    fitWidthPx: 960,
  })
    .nodes.filter((n) => n.type === 'group')
    .at(-1)!
  expect(painted.box).toEqual(beforeBox)
  if (painted.type !== 'group') throw new Error('group was flattened')
  expect(painted.children.map((c) => 'fill' in c && c.fill)).toEqual([
    { kind: 'solid', color: '#FF0000' },
    { kind: 'solid', color: '#00FF00' },
  ])
  const nativeGroup = reopened.deck.slides[0]!.elements.filter((el) => el.type === 'group').at(-1)!
  expect(
    nativeGroup.children.map(
      (child) =>
        (child.type === 'shape' || child.type === 'text') &&
        child.text?.paragraphs.flatMap((p) => p.runs.map((r) => r.text)).join(''),
    ),
  ).toEqual(['Object 2', 'Object 3'])
})

it('round-trips painted native fields and moved, rotated shapes without changing content or relative spacing', async () => {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  // A small, controlled deck for native mouse-drag and brush verification.
  slide.elements = []
  const shapes = [
    { x: 90, y: 70, text: 'FORMAT SOURCE', color: '#ffcc00' },
    { x: 390, y: 70, text: 'Target with field ', color: '#dddddd' },
    { x: 690, y: 70, text: 'Second target', color: '#dddddd' },
    { x: 90, y: 320, text: 'MOVE A', color: '#ff9999' },
    { x: 390, y: 320, text: 'MOVE B', color: '#99ccff' },
    { x: 690, y: 320, text: 'MOVE C', color: '#99dd99' },
  ].map((s, i) =>
    addElement(slide, {
      kind: 'rect',
      offset: { x: s.x * EMU, y: s.y * EMU, cx: 200 * EMU, cy: 90 * EMU },
      paragraphs: [
        {
          runs:
            i === 1
              ? [
                  { text: s.text, fontSize: 20, fontFamily: 'Arial' },
                  { text: '1', field: 'slidenum' },
                  { text: '\nsecond line' },
                ]
              : [{ text: s.text, fontSize: 20, fontFamily: 'Arial' }],
        },
      ],
      fillColor: s.color,
    }),
  )
  const tx = runTxn(opened, {
    ops: [
      {
        op: 'setFill',
        target: { slide: 0, el: shapes[0]!.id },
        fill: {
          stops: [
            { pos: 0, color: '#ffcc00' },
            { pos: 1, color: '#ff5555' },
          ],
          angle: 45 * 60000,
        },
      },
      {
        op: 'setStroke',
        target: { slide: 0, el: shapes[0]!.id },
        stroke: { color: '#003399', widthEmu: 3 * 12700, dash: 'dash', cap: 'rnd' },
      },
      {
        op: 'setFont',
        target: { slide: 0, el: shapes[0]!.id },
        font: { fontSizePt: 24, bold: true, strike: true, color: '#001166', letterSpacingPt: 1 },
      },
      {
        op: 'setParagraphFormat',
        target: { slide: 0, el: shapes[0]!.id },
        format: { align: 'center' },
      },
      {
        op: 'setTransform',
        target: { slide: 0, el: shapes[4]!.id },
        rotDeg: 25,
        box: { ...shapes[4]!.transform.offset },
      },
    ],
  })
  expect(tx.failures).toBeUndefined()
  const initialBytes = await savePptx(opened)
  const initialReopen = await openPptx(initialBytes)
  expect(initialReopen.archive.readText('ppt/slides/slide1.xml')).toContain('type="slidenum"')
  if (process.env.GENOFFICE_WRITE_MOVE_QA === '1') {
    const dir = resolve(import.meta.dirname, '../../../docs/qa/slides-move-format-2026-11-05')
    await mkdir(dir, { recursive: true })
    await writeFile(resolve(dir, 'move-format-input.pptx'), await savePptx(opened))
  }
  const rendered = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960 })
  const source = rendered.nodes.find((n) => n.sourceId === shapes[0]!.id)!
  const target = rendered.nodes.find((n) => n.sourceId === shapes[1]!.id)!
  const sourceFormat = extractBrushFormat(source, rendered.scale)!
  const beforeText = structuredClone(
    shapes[1]!.text!.paragraphs.map((p) => p.runs.map((r) => ({ text: r.text, field: r.field }))),
  )
  const painted = runTxn(opened, { ops: buildBrushOps(sourceFormat, target, 0) })
  expect(painted.failures).toBeUndefined()
  expect(
    shapes[1]!.text!.paragraphs.map((p) => p.runs.map((r) => ({ text: r.text, field: r.field }))),
  ).toEqual(beforeText)
  const originals = shapes.slice(3).map((s) => ({ ...s.transform.offset, rot: s.transform.rot }))
  expect(
    runTxn(opened, {
      ops: shapes.slice(3).map((s) => ({
        op: 'setTransform',
        target: { slide: 0, el: s.id },
        rotDeg: s.transform.rot / 60000,
        box: {
          ...s.transform.offset,
          x: s.transform.offset.x + 35 * EMU,
          y: s.transform.offset.y + 40 * EMU,
        },
      })),
    }).applied,
  ).toBe(true)
  const reopened = await openPptx(await savePptx(opened))
  const final = reopened.deck.slides[0]!.elements
  for (let i = 0; i < 3; i++) {
    const tr = final[i + 3]!.transform
    const { rot, ...originalBox } = originals[i]!
    expect(tr.offset).toEqual({
      ...originalBox,
      x: originals[i]!.x + 35 * EMU,
      y: originals[i]!.y + 40 * EMU,
    })
    expect(tr.rot).toBe(rot)
  }
  const xml = reopened.archive.readText('ppt/slides/slide1.xml')!
  expect(xml).toContain('type="slidenum"')
  const paintedTarget = final[1]!
  expect(paintedTarget.type === 'shape' || paintedTarget.type === 'text').toBe(true)
  if (paintedTarget.type === 'shape' || paintedTarget.type === 'text') {
    expect(paintedTarget.fill?.type).toBe('gradient')
    expect(paintedTarget.text?.paragraphs[0]?.runs[0]).toMatchObject({
      fontSize: 24,
      bold: true,
      strike: true,
    })
    expect(paintedTarget.anchor.originalXml).toContain('val="dash"')
  }
})
