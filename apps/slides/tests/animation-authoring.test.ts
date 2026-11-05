import { describe, expect, it } from 'vitest'
import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'
import type { AnimationItem } from '../src/shared/ipc'
import {
  animationMarkers,
  animationStepNumbers,
  animationTargetSelection,
} from '../src/renderer/animation-authoring'
import { editedAnimation, patchAnimation } from '../src/renderer/animation-actions'
import type { ActionCtx } from '../src/renderer/action-context'
import {
  addElement,
  createBlankPptx,
  elementSpid,
  getSlideAnimations,
  openPptx,
  savePptx,
} from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'

const item = (sourceId: string, trigger: AnimationItem['trigger'] = 'onClick'): AnimationItem => ({
  sourceId,
  targetName: sourceId,
  trigger,
  effect: 'fade',
  durationMs: 500,
  delayMs: 0,
})

it('keeps each after-animation choice on the first Appear effect through renderer write/read and save/reopen', async () => {
  let opened = await openPptx(await createBlankPptx())
  const id = addElement(opened.deck.slides[0]!, {
    kind: 'roundRect',
    offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
    fillColor: '#FFFFFF',
  }).id
  expect(
    runTxn(opened, {
      ops: [
        {
          op: 'setAnimations',
          target: { slide: 0 },
          items: [{ ...item(id), effect: 'appear', durationMs: 0 }, item(id)],
        },
      ],
    }).applied,
  ).toBe(true)
  const read = (): AnimationItem[] => {
    const page = opened.deck.slides[0]!
    return getSlideAnimations(page, true).map((a) => ({
      sourceId: page.elements.find((el) => elementSpid(el) === a.spid)!.id,
      targetName: 'Card',
      effect: a.effect,
      trigger: a.trigger,
      durationMs: a.durationMs,
      delayMs: a.delayMs,
      nativeId: a.nativeId,
      ...(a.after ? { after: a.after } : {}),
    }))
  }
  const ctx = {
    current: 0,
    path: '/animation.pptx',
    animations: read(),
    setAnimations: (items: AnimationItem[]) => {
      ctx.animations = items
    },
    setDirty: () => {},
    setStatus: () => {},
  } as unknown as ActionCtx
  Object.assign(window, {
    slidesApi: {
      getAnimations: async () => read(),
      setAnimations: async ({ items }: { items: AnimationItem[] }) =>
        runTxn(opened, {
          ops: [{ op: 'setAnimations', target: { slide: 0 }, items }],
        }).applied,
    },
  })
  for (const after of ['hideAfter', 'hideOnNextClick', null] as const) {
    await patchAnimation(ctx, 0, { after })
    expect(ctx.animations[0]!.after).toBe(after ?? undefined)
    expect(read()[0]).toMatchObject({ effect: 'appear', durationMs: 0 })
    opened = await openPptx(await savePptx(opened))
    expect(read()[0]!.after).toBe(after ?? undefined)
    expect(read()[1]).toMatchObject({ effect: 'fade', durationMs: 500 })
    expect(read()[1]!.after).toBeUndefined()
  }
})
const node = (sourceId: string, box = { x: 200, y: 100, w: 80, h: 60 }): RenderNode =>
  ({ id: sourceId, sourceId, type: 'shape', box }) as RenderNode
const slide = (nodes: RenderNode[]): RenderSlide => ({
  nodes,
  widthPx: 960,
  heightPx: 540,
  scale: 1,
  background: { kind: 'solid', color: '#ffffff' },
})

describe('animation order badges', () => {
  it('shares click order across automatic, with-previous and after-previous effects', () => {
    expect(
      animationStepNumbers([
        item('auto', 'withPrev'),
        item('a'),
        item('b', 'withPrev'),
        item('c', 'afterPrev'),
        item('d'),
      ]),
    ).toEqual([0, 1, 1, 1, 2])
  })
  it('keeps stacked effects separately clickable and avoids overlapping markers at any zoom', () => {
    const s = slide([node('a'), node('b', { x: 205, y: 100, w: 80, h: 60 })])
    const items = [item('a'), item('a'), item('b'), item('missing')]
    for (const zoom of [0.5, 1, 2]) {
      const markers = animationMarkers(s, items, zoom)
      expect(markers.map((m) => m.index)).toEqual([0, 1, 2])
      expect(markers[1]!.y - markers[0]!.y).toBeGreaterThanOrEqual(20 / zoom)
      expect(new Set(markers.map((m) => m.y)).size).toBe(3)
      expect(s.nodes[0]!.box).toEqual({ x: 200, y: 100, w: 80, h: 60 })
    }
  })
  it('positions grouped targets after parent rotation without applying child scaling twice', () => {
    const child = node('child', { x: 10, y: 20, w: 20, h: 10 })
    const group = {
      ...node('group', { x: 100, y: 100, w: 100, h: 100 }),
      type: 'group',
      box: { x: 100, y: 100, w: 100, h: 100, rotationDeg: 90 },
      children: [child],
      childScaleX: 2,
      childScaleY: 2,
    } as RenderNode
    const s = slide([group])
    const marker = animationMarkers(s, [item('child')], 1)[0]!
    expect(marker.x).toBeCloseTo(170 - 29)
    expect(marker.y).toBeCloseTo(110)
    expect(animationTargetSelection(s, 'child')).toEqual({ selectedId: 'child', groupId: 'group' })
  })
})

it('replaces only the selected stacked effect and persists effect options/timing through native save/reopen', async () => {
  const opened = await openPptx(await createBlankPptx())
  const id = addElement(opened.deck.slides[0]!, {
    kind: 'roundRect',
    offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
    fillColor: '#FFFFFF',
  }).id
  const initial = [item(id), { ...item(id), effect: 'disappear' as const, durationMs: 0 }]
  expect(
    runTxn(opened, { ops: [{ op: 'setAnimations', target: { slide: 0 }, items: initial }] })
      .applied,
  ).toBe(true)
  const native = getSlideAnimations(opened.deck.slides[0]!, true)
  const items = initial.map((a, i) => ({ ...a, nativeId: native[i]!.nativeId }))
  const changed = editedAnimation(items[0]!, { effect: 'flyIn' })
  expect(changed.nativeId).toBeUndefined()
  const updated = editedAnimation(changed, {
    direction: 'left',
    after: 'hideOnNextClick',
    trigger: 'afterPrev',
    durationMs: 1250,
    delayMs: 250,
  })
  expect(
    runTxn(opened, {
      ops: [{ op: 'setAnimations', target: { slide: 0 }, items: [updated, items[1]!] }],
    }).applied,
  ).toBe(true)
  const back = await openPptx(await savePptx(opened))
  const saved = getSlideAnimations(back.deck.slides[0]!, true)
  expect(saved).toHaveLength(2)
  expect(saved[0]).toMatchObject({
    effect: 'flyIn',
    direction: 'left',
    after: 'hideOnNextClick',
    trigger: 'afterPrev',
    durationMs: 1250,
    delayMs: 250,
  })
  expect(saved[1]).toMatchObject({
    effect: 'disappear',
    durationMs: 0,
    trigger: 'onClick',
    nativeId: native[1]!.nativeId,
  })
})

it('intentional replacement of an approximated imported preset clears native semantics and stale options', () => {
  const imported = {
    ...item('a'),
    customPreset: true,
    nativeId: 42,
    direction: 'left' as const,
    motionPath: 'M 0 0 L 1 1',
  }
  const changed = editedAnimation(imported, { effect: 'fade' })
  expect(changed).not.toHaveProperty('customPreset')
  expect(changed).not.toHaveProperty('nativeId')
  expect(changed).not.toHaveProperty('direction')
  expect(changed).not.toHaveProperty('motionPath')
  expect(editedAnimation(changed, { after: null })).not.toHaveProperty('after')
})

it('reads back successive single-effect edits while preserving other native effects', async () => {
  const opened = await openPptx(await createBlankPptx())
  const id = addElement(opened.deck.slides[0]!, {
    kind: 'roundRect',
    offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
    fillColor: '#FFFFFF',
  }).id
  const read = (): AnimationItem[] =>
    getSlideAnimations(opened.deck.slides[0]!, true).map((a) => ({
      sourceId: id,
      targetName: 'Card',
      effect: a.effect,
      trigger: a.trigger,
      durationMs: a.durationMs,
      delayMs: a.delayMs,
      ...(a.direction ? { direction: a.direction } : {}),
      ...(a.after ? { after: a.after } : {}),
      ...(a.nativeId ? { nativeId: a.nativeId } : {}),
    }))
  const write = (items: AnimationItem[]) =>
    expect(
      runTxn(opened, {
        ops: [{ op: 'setAnimations', target: { slide: 0 }, items }],
      }).applied,
    ).toBe(true)
  write([item(id), item(id), item(id)])
  const untouched = read().filter((_, i) => i !== 1)
  for (const patch of [
    { effect: 'flyIn' as const },
    { direction: 'left' as const },
    { durationMs: 1250 },
    { delayMs: 250 },
    { after: 'hideOnNextClick' as const },
  ]) {
    const before = read()
    write(before.map((a, i) => (i === 1 ? editedAnimation(a, patch) : a)))
    expect(read()[1]).toMatchObject(patch)
    expect(read().filter((_, i) => i !== 1)).toEqual(untouched)
  }
  const saved = await openPptx(await savePptx(opened))
  expect(getSlideAnimations(saved.deck.slides[0]!)[1]).toMatchObject({
    effect: 'flyIn',
    direction: 'left',
    durationMs: 1250,
    delayMs: 250,
    after: 'hideOnNextClick',
  })
})
