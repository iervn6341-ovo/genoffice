// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest'

import {
  EXCLUSIVE_HIDDEN_ATTR,
  installContextSubmenuExclusive,
} from '../src/renderer/context-submenu-exclusive'

const frames = (n: number) =>
  new Promise<void>((resolve) => {
    const step = (left: number) =>
      left === 0 ? resolve() : requestAnimationFrame(() => step(left - 1))
    step(n)
  })

function mountMenu(labels: string[]): HTMLButtonElement[] {
  const popup = document.createElement('div')
  popup.className = 'univer-popup'
  const buttons = labels.map((label) => {
    const b = document.createElement('button')
    b.textContent = label
    popup.appendChild(b)
    return b
  })
  document.body.appendChild(popup)
  return buttons
}

function openSubmenu(name: string): HTMLElement {
  const el = document.createElement('div')
  el.setAttribute('data-u-context-menu-submenu', 'true')
  el.dataset.name = name
  el.style.visibility = 'visible'
  document.body.appendChild(el)
  return el
}

const hover = (el: Element) => el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
const hidden = (el: Element) => el.hasAttribute(EXCLUSIVE_HIDDEN_ATTR)

let disposers: Array<() => void> = []
afterEach(() => {
  for (const d of disposers) d()
  disposers = []
  document.body.innerHTML = ''
})

describe('installContextSubmenuExclusive', () => {
  it('hides the previous submenu when another submenu opens', async () => {
    disposers.push(installContextSubmenuExclusive().dispose)
    const [protect, sort] = mountMenu(['protect', 'sort'])
    hover(protect!)
    const a = openSubmenu('protect')
    await frames(3)
    expect(hidden(a)).toBe(false)

    hover(sort!)
    const b = openSubmenu('sort') // Univer mounts the new one in the same batch
    await frames(3)
    expect(hidden(a)).toBe(true)
    expect(hidden(b)).toBe(false)
  })

  it('hides every submenu when moving to an item without one', async () => {
    disposers.push(installContextSubmenuExclusive().dispose)
    const [protect, plain] = mountMenu(['protect', 'plain'])
    hover(protect!)
    const a = openSubmenu('protect')
    await frames(3)
    hover(plain!)
    await frames(3)
    expect(hidden(a)).toBe(true)
  })

  it('keeps the submenu while the pointer moves into it', async () => {
    disposers.push(installContextSubmenuExclusive().dispose)
    const [protect] = mountMenu(['protect'])
    hover(protect!)
    const a = openSubmenu('protect')
    await frames(3)
    const inner = document.createElement('button')
    a.appendChild(inner)
    hover(inner)
    await frames(3)
    expect(hidden(a)).toBe(false)
  })

  it('shows a hidden submenu again when its owner is re-hovered', async () => {
    disposers.push(installContextSubmenuExclusive().dispose)
    const [protect, sort] = mountMenu(['protect', 'sort'])
    hover(protect!)
    const a = openSubmenu('protect')
    await frames(3)
    hover(sort!)
    const b = openSubmenu('sort')
    await frames(3)
    expect(hidden(a)).toBe(true)

    hover(protect!)
    a.style.visibility = 'hidden' // Univer resets the position flag...
    await frames(1)
    a.style.visibility = 'visible' // ...and re-measures
    await frames(3)
    expect(hidden(a)).toBe(false)
    expect(hidden(b)).toBe(true)
  })

  it('removes its markers on dispose', async () => {
    const fix = installContextSubmenuExclusive()
    const [protect, plain] = mountMenu(['protect', 'plain'])
    hover(protect!)
    const a = openSubmenu('protect')
    await frames(3)
    hover(plain!)
    await frames(3)
    expect(hidden(a)).toBe(true)
    fix.dispose()
    expect(hidden(a)).toBe(false)
  })

  it('uses the React owner: a stale submenu that updates late stays hidden', async () => {
    disposers.push(installContextSubmenuExclusive().dispose)
    const [freeze, protect] = mountMenu(['freeze', 'protect'])
    const own = (submenu: HTMLElement, item: Element) => {
      const host = { tag: 5, stateNode: item.parentElement ? item : item, return: null }
      ;(submenu as unknown as Record<string, unknown>)['__reactFiber$x'] = {
        tag: 5,
        stateNode: submenu,
        return: { tag: 4, stateNode: document.body, return: host },
      }
    }
    hover(freeze!)
    const a = openSubmenu('freeze')
    own(a, freeze!)
    await frames(3)
    hover(protect!)
    const b = openSubmenu('protect')
    own(b, protect!)
    // the stale submenu's own position pass lands after the hover moved on
    a.style.top = '10px'
    await frames(4)
    expect(hidden(a)).toBe(true)
    expect(hidden(b)).toBe(false)
  })
})
