/**
 * Only one context-menu submenu may be open at a time (Excel behaviour).
 *
 * Univer keeps the open flag of each submenu item in that item's own React
 * state and closes it 500 ms after the pointer leaves. Sliding down the menu
 * from "Protect range" to "Sort" therefore leaves both submenus on screen for
 * that delay, stacked on top of each other.
 *
 * The pointer entering a different top-level item makes every submenu that
 * Univer did not just (re)open stale. Those are hidden at once with a marker
 * attribute (styles.css hides it); Univer's own timer still unmounts them. A
 * stale submenu whose owner is hovered again is touched by Univer (mount or
 * style update), which lifts the marker.
 */
const SUBMENU_SELECTOR = '[data-u-context-menu-submenu]'
const ITEM_SELECTOR = 'button, [aria-disabled]'
interface FiberLike {
  tag: number
  stateNode: unknown
  return: FiberLike | null
}

/**
 * The menu item a portaled submenu belongs to. Univer renders the submenu
 * through createPortal as a child of the item's root element, so the React
 * fiber chain still leads back to that element even though the DOM node lives
 * elsewhere. Null when React internals are not reachable.
 */
function ownerOf(submenu: HTMLElement): Element | null {
  const key = Object.keys(submenu).find((k) => k.startsWith('__reactFiber$'))
  if (!key) return null
  let fiber = (submenu as unknown as Record<string, FiberLike>)[key]?.return ?? null
  while (fiber) {
    // 5 = HostComponent; skip the portal (tag 4) and the submenu's own div
    if (fiber.tag === 5 && fiber.stateNode instanceof Element && fiber.stateNode !== submenu) {
      return fiber.stateNode
    }
    fiber = fiber.return
  }
  return null
}

export const EXCLUSIVE_HIDDEN_ATTR = 'data-exclusive-hidden'

export function installContextSubmenuExclusive(root: HTMLElement = document.body): {
  dispose(): void
} {
  let lastItem: Element | null = null
  let touched = new Set<Element>()
  let pending = false
  let timer = 0
  let frame = 0

  const submenusIn = (node: Node): HTMLElement[] => {
    if (!(node instanceof HTMLElement)) return []
    const found: HTMLElement[] = []
    if (node.matches(SUBMENU_SELECTOR)) found.push(node)
    found.push(...node.querySelectorAll<HTMLElement>(SUBMENU_SELECTOR))
    return found
  }

  // keep a submenu only if it belongs to the hovered item (exact when React
  // internals are readable, else: Univer touched it during this hover)
  const decide = (el: HTMLElement) => {
    if (!lastItem) return
    const owner = ownerOf(el)
    const keep = owner ? owner.contains(lastItem) : touched.has(el)
    if (keep) el.removeAttribute(EXCLUSIVE_HIDDEN_ATTR)
    else el.setAttribute(EXCLUSIVE_HIDDEN_ATTR, 'true')
  }

  const touch = (el: HTMLElement) => {
    if (pending) touched.add(el)
    // a hidden submenu Univer updates with no hover to judge by is reopening
    if (!lastItem) el.removeAttribute(EXCLUSIVE_HIDDEN_ATTR)
    else decide(el)
  }

  const observer = new MutationObserver((records) => {
    for (const record of records) {
      if (record.type === 'attributes') {
        if (record.target instanceof HTMLElement && record.target.matches(SUBMENU_SELECTOR)) {
          touch(record.target)
        }
      } else {
        for (const added of record.addedNodes) for (const el of submenusIn(added)) touch(el)
      }
    }
  })
  observer.observe(root, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ['style'],
  })

  const evaluate = () => {
    for (const submenu of root.querySelectorAll<HTMLElement>(SUBMENU_SELECTOR)) decide(submenu)
  }

  const settle = () => {
    pending = false
    touched = new Set()
  }

  const onOver = (event: MouseEvent) => {
    const target = event.target
    if (!(target instanceof Element)) return
    if (target.closest(SUBMENU_SELECTOR)) {
      // moving into a submenu: re-entering its owner must count as a change
      lastItem = null
      return
    }
    const item = target.closest(ITEM_SELECTOR) ?? target.closest('.univer-popup')
    if (!item || item === lastItem) return
    // only react inside an open context menu (it renders submenus beside it)
    if (!target.closest('.univer-popup')) return
    lastItem = item
    touched = new Set()
    pending = true
    window.clearTimeout(timer)
    cancelAnimationFrame(frame)
    // Univer opens the hovered item's submenu in the same React batch; judge
    // once that has landed, then once more after its position pass
    evaluate()
    timer = window.setTimeout(evaluate, 0)
    frame = requestAnimationFrame(() => {
      frame = requestAnimationFrame(() => {
        evaluate()
        settle()
      })
    })
  }

  document.addEventListener('mouseover', onOver, true)
  return {
    dispose() {
      document.removeEventListener('mouseover', onOver, true)
      observer.disconnect()
      window.clearTimeout(timer)
      cancelAnimationFrame(frame)
      for (const el of root.querySelectorAll(`[${EXCLUSIVE_HIDDEN_ATTR}]`)) {
        el.removeAttribute(EXCLUSIVE_HIDDEN_ATTR)
      }
    },
  }
}
