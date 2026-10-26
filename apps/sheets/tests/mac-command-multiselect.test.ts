// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest'

import {
  installMacCommandMultiSelect,
  shouldPresentAsCtrl,
} from '../src/renderer/mac-command-multiselect'

afterEach(() => {
  document.body.innerHTML = ''
})

function grid() {
  document.body.innerHTML =
    '<div id="univer-container"><canvas id="c"></canvas></div><canvas id="other"></canvas><input id="in" />'
  return document.getElementById('c')!
}

function ev(type: string, init: MouseEventInit): MouseEvent {
  return new MouseEvent(type, { bubbles: true, cancelable: true, ...init })
}

describe('⌘+click adds to the selection on a Mac', () => {
  it('presents a ⌘ press on the grid canvas to Univer as Ctrl', () => {
    const canvas = grid()
    const off = installMacCommandMultiSelect(true)
    let seen: boolean | null = null
    canvas.addEventListener('pointerdown', (e) => (seen = (e as MouseEvent).ctrlKey))
    canvas.dispatchEvent(ev('pointerdown', { metaKey: true }))
    expect(seen).toBe(true)
    off()
  })

  it('also covers the move and up of the same gesture', () => {
    const canvas = grid()
    const off = installMacCommandMultiSelect(true)
    const seen: boolean[] = []
    for (const type of ['pointermove', 'pointerup'])
      canvas.addEventListener(type, (e) => seen.push((e as MouseEvent).ctrlKey))
    canvas.dispatchEvent(ev('pointermove', { metaKey: true }))
    canvas.dispatchEvent(ev('pointerup', { metaKey: true }))
    expect(seen).toEqual([true, true])
    off()
  })

  it('leaves plain presses, Ctrl presses and other surfaces alone', () => {
    const canvas = grid()
    const off = installMacCommandMultiSelect(true)
    const seen: Array<[string, boolean]> = []
    const record = (name: string) => (e: Event) => seen.push([name, (e as MouseEvent).ctrlKey])
    canvas.addEventListener('pointerdown', record('grid'))
    document.getElementById('other')!.addEventListener('pointerdown', record('other-canvas'))
    document.getElementById('in')!.addEventListener('pointerdown', record('input'))
    canvas.dispatchEvent(ev('pointerdown', {}))
    canvas.dispatchEvent(ev('pointerdown', { metaKey: true, altKey: true }))
    document.getElementById('other')!.dispatchEvent(ev('pointerdown', { metaKey: true }))
    document.getElementById('in')!.dispatchEvent(ev('pointerdown', { metaKey: true }))
    expect(seen).toEqual([
      ['grid', false],
      ['grid', false],
      ['other-canvas', false],
      ['input', false],
    ])
    off()
  })

  it('does nothing off a Mac, and stops after uninstall', () => {
    const canvas = grid()
    expect(
      shouldPresentAsCtrl({ metaKey: true, ctrlKey: false, altKey: false, target: canvas }, false),
    ).toBe(false)
    const off = installMacCommandMultiSelect(true)
    off()
    let seen: boolean | null = null
    canvas.addEventListener('pointerdown', (e) => (seen = (e as MouseEvent).ctrlKey))
    canvas.dispatchEvent(ev('pointerdown', { metaKey: true }))
    expect(seen).toBe(false)
  })
})
