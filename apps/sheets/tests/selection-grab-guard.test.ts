import { describe, expect, it } from 'vitest'
import {
  GRAB_DWELL_MS,
  installSelectionGrabGuard,
  type GrabControlLike,
} from '../src/renderer/selection-grab-guard'

type Handler = (event: unknown) => void

function fakeTarget() {
  const handlers = new Map<string, Handler[]>()
  return {
    addEventListener: (type: string, fn: Handler) => {
      handlers.set(type, [...(handlers.get(type) ?? []), fn])
    },
    removeEventListener: (type: string, fn: Handler) => {
      handlers.set(
        type,
        (handlers.get(type) ?? []).filter((h) => h !== fn),
      )
    },
    fire: (type: string, event: unknown = {}) => {
      for (const fn of handlers.get(type) ?? []) fn(event)
    },
    count: (type: string) => (handlers.get(type) ?? []).length,
  }
}

const control = (): GrabControlLike => ({
  fillControl: { evented: true },
  leftControl: { evented: true },
  rightControl: { evented: true },
  topControl: { evented: true },
  bottomControl: { evented: true },
})
const eventedOf = (c: GrabControlLike) => [
  c.fillControl.evented,
  c.leftControl.evented,
  c.rightControl.evented,
  c.topControl.evented,
  c.bottomControl.evented,
]
const at = (cursor: string, button = 0) => ({ button, target: { style: { cursor } } })

function setup() {
  const target = fakeTarget()
  const c = control()
  let clock = 1000
  const guard = installSelectionGrabGuard({
    getControls: () => [c],
    target: target as never,
    now: () => clock,
  })
  return { target, c, guard, advance: (ms: number) => (clock += ms) }
}

describe('selection grab guard', () => {
  it('a quick press on the fill handle is an ordinary selection start', () => {
    const { target, c, advance } = setup()
    target.fire('pointermove', at('crosshair'))
    advance(40)
    target.fire('pointerdown', at('crosshair'))
    expect(eventedOf(c)).toEqual([false, false, false, false, false])
    target.fire('pointerup')
    expect(eventedOf(c)).toEqual([true, true, true, true, true])
  })

  it('a quick press on the frame border does not move the cells', () => {
    const { target, c, advance } = setup()
    target.fire('pointermove', at('move'))
    advance(GRAB_DWELL_MS - 1)
    target.fire('pointerdown', at('move'))
    expect(c.leftControl.evented).toBe(false)
  })

  it('a press after resting on the handle grabs it', () => {
    const { target, c, advance } = setup()
    target.fire('pointermove', at('crosshair'))
    advance(GRAB_DWELL_MS + 10)
    target.fire('pointermove', at('crosshair'))
    target.fire('pointerdown', at('crosshair'))
    expect(eventedOf(c)).toEqual([true, true, true, true, true])
  })

  it('leaving the handle restarts the rest time', () => {
    const { target, c, advance } = setup()
    target.fire('pointermove', at('crosshair'))
    advance(500)
    target.fire('pointermove', at('cell'))
    advance(10)
    target.fire('pointermove', at('crosshair'))
    advance(10)
    target.fire('pointerdown', at('crosshair'))
    expect(c.fillControl.evented).toBe(false)
  })

  it('a press without any hover is never a grab (touch, scripted click)', () => {
    const { target, c } = setup()
    target.fire('pointerdown', at('crosshair'))
    expect(c.fillControl.evented).toBe(false)
  })

  it('presses on plain cells and right-clicks keep working and are restored', () => {
    const { target, c } = setup()
    target.fire('pointerdown', at('cell'))
    expect(c.fillControl.evented).toBe(false)
    target.fire('pointerup')
    target.fire('pointerdown', at('crosshair', 2))
    expect(eventedOf(c)).toEqual([true, true, true, true, true])
  })

  it('does not enable a handle that was off before the press', () => {
    const { target, c } = setup()
    c.fillControl.evented = false
    target.fire('pointerdown', at('cell'))
    target.fire('pointerup')
    expect(c.fillControl.evented).toBe(false)
    expect(c.leftControl.evented).toBe(true)
  })

  it('dispose removes the listeners and restores the handles', () => {
    const { target, c, guard } = setup()
    target.fire('pointerdown', at('cell'))
    guard.dispose()
    expect(eventedOf(c)).toEqual([true, true, true, true, true])
    expect(target.count('pointerdown')).toBe(0)
    expect(target.count('pointermove')).toBe(0)
  })

  it('an element blur in the middle of a press does not re-arm the handles', () => {
    const { target, c } = setup()
    target.fire('pointerdown', at('cell'))
    target.fire('blur', { target: { tagName: 'INPUT' } })
    expect(eventedOf(c)).toEqual([false, false, false, false, false])
  })

  it('the window losing focus does end the press', () => {
    const { target, c } = setup()
    target.fire('pointerdown', at('cell'))
    target.fire('blur', { target: target })
    expect(eventedOf(c)).toEqual([true, true, true, true, true])
  })

  it('only touches the controls it is given (formula reference boxes stay grabbable)', () => {
    const target = fakeTarget()
    const selection = control()
    const refBox = control()
    installSelectionGrabGuard({
      getControls: () => [selection],
      target: target as never,
      now: () => 0,
    })
    target.fire('pointerdown', at('cell'))
    expect(eventedOf(selection)).toEqual([false, false, false, false, false])
    expect(eventedOf(refBox)).toEqual([true, true, true, true, true])
  })
})
