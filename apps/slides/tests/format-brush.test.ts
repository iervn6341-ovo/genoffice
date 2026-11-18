import { expect, it } from 'vitest'
import type { RenderNode, ShapeRenderNode } from '@genoffice/pptx-render'
import { buildBrushOps, extractBrushFormat, findBrushTarget } from '../src/renderer/format-brush'

const node = (id = 'shape') =>
  ({
    id,
    sourceId: id,
    type: 'shape',
    box: { x: 10, y: 20, w: 100, h: 80, rotationDeg: 30 },
    fill: {
      kind: 'gradient',
      stops: [
        { pos: 0, color: '#ff000080' },
        { pos: 0.4, color: '#00ff00' },
        { pos: 1, color: '#0000ff' },
      ],
      angleDeg: 45,
    },
    stroke: {
      color: '#112233',
      widthPt: 3,
      dashPreset: 'dash',
      cap: 'round',
      join: 'bevel',
      compound: 'dbl',
    },
    text: {
      fontScale: 0.8,
      lines: [
        {
          paraStart: true,
          align: 'right',
          runs: [
            {
              text: 'Field',
              fontFamily: 'Substitution',
              srcFontFamily: 'Original Font',
              fontSizePx: 16,
              bold: false,
              italic: false,
              underline: false,
              color: '#abcdef',
              strike: true,
              baselinePct: -25,
              letterSpacingPx: 2,
              highlight: '#123456',
            },
          ],
        },
      ],
    },
  }) as unknown as ShapeRenderNode

it('copies gradients, full outline settings and actual document font size rather than fitted display size', () => {
  const brush = extractBrushFormat(node(), 0.5)!
  const fill = node().fill
  expect(brush.fill).toMatchObject({ gradient: { stops: fill.kind === 'gradient' && fill.stops } })
  expect(brush.stroke).toMatchObject({
    dash: 'dash',
    cap: 'rnd',
    join: 'bevel',
    compound: 'dbl',
    widthPt: 3,
  })
  expect(brush.run).toMatchObject({
    fontFamily: 'Original Font',
    fontSize: 30,
    strike: true,
    baseline: -25,
    letterSpacingPt: 3,
  })
  const ops = buildBrushOps(brush, node('target'), 2)
  expect(ops.map((o) => o.op)).toEqual([
    'setFill',
    'setStroke',
    'setFont',
    'setParagraphFormat',
    'setTextAnchor',
    'setTextBodyProps',
    'setEffects',
  ])
  expect(ops[2]!.font).toMatchObject({ fontSizePt: 30, strike: true })
  expect(ops.some((o) => o.op === 'setText')).toBe(false)
})

it('copies explicit no-outline and normal font values so old target formatting is cleared', () => {
  const source = node()
  source.stroke = undefined
  source.text!.lines[0]!.runs[0]!.strike = undefined
  source.text!.lines[0]!.runs[0]!.highlight = undefined
  source.text!.lines[0]!.runs[0]!.baselinePct = undefined
  const brush = extractBrushFormat(source)!
  expect(brush.stroke).toBeNull()
  expect(brush.run).toMatchObject({ strike: false, highlight: null, baseline: 0 })
})

it('finds selected group children and builds native child-targeted operations for a whole group', () => {
  const group = {
    type: 'group',
    sourceId: 'g',
    children: [node('child-a'), node('child-b')],
  } as unknown as RenderNode
  expect(findBrushTarget([group], 'child-b')?.groupId).toBe('g')
  const brush = extractBrushFormat(group)!
  const ops = buildBrushOps(brush, group, 0)
  expect(ops).toHaveLength(8)
  expect(ops.every((op) => op.group === 'g')).toBe(true)
  expect(ops.map((op) => op.target?.el)).toContain('child-b')
})
