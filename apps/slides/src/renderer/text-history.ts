/**
 * Undo / redo for the text editing overlay.
 *
 * The browser's own undo (`document.execCommand('undo')`) only remembers typing it performed itself.
 * The overlay also edits the DOM directly — Shift+Tab / Backspace change a paragraph's level and
 * margin, the ribbon marks bullets and spacing, Shift+Enter inserts a "\n" text node — and none of
 * that is on the native stack, so ⌘Z skipped those steps or jumped over several at once.
 *
 * This keeps its own stack of whole-editor snapshots (innerHTML + selection). A MutationObserver
 * sees every change, whoever made it; a burst of changes settles into one step after a short idle,
 * and `flush()` closes the step early at discrete keys (Enter, Tab, Backspace, shortcuts).
 */
interface Point {
  path: number[]
  offset: number
}

interface Snapshot {
  html: string
  sel: { anchor: Point; focus: Point } | null
}

export interface TextHistory {
  /** Closes the step in progress (call before a discrete key acts). */
  flush(): void
  /** True when handled (also at the oldest step, where there is nothing to take back). */
  undo(): boolean
  redo(): boolean
  dispose(): void
}

/** Typing settles into one step after this much quiet. */
const IDLE_MS = 500

const pathOf = (root: Node, node: Node): number[] | null => {
  const path: number[] = []
  let cur: Node | null = node
  while (cur && cur !== root) {
    const parent: ParentNode | null = cur.parentNode
    if (!parent) return null
    path.unshift(Array.prototype.indexOf.call(parent.childNodes, cur))
    cur = parent
  }
  return cur === root ? path : null
}

const nodeAt = (root: Node, path: number[]): Node => {
  let cur: Node = root
  for (const index of path) {
    const next = cur.childNodes[index]
    if (!next) break
    cur = next
  }
  return cur
}

const clampOffset = (node: Node, offset: number): number =>
  Math.min(
    offset,
    node.nodeType === Node.TEXT_NODE ? (node.textContent?.length ?? 0) : node.childNodes.length,
  )

function pointOf(root: Node, node: Node, offset: number): Point | null {
  const path = pathOf(root, node)
  return path ? { path, offset } : null
}

function takeSnapshot(root: HTMLElement): Snapshot {
  const sel = window.getSelection()
  let selection: Snapshot['sel'] = null
  if (sel && sel.rangeCount && sel.anchorNode && sel.focusNode) {
    const anchor = pointOf(root, sel.anchorNode, sel.anchorOffset)
    const focus = pointOf(root, sel.focusNode, sel.focusOffset)
    if (anchor && focus) selection = { anchor, focus }
  }
  return { html: root.innerHTML, sel: selection }
}

export function createTextHistory(root: HTMLElement): TextHistory {
  const stack: Snapshot[] = [takeSnapshot(root)]
  let index = 0
  let timer: ReturnType<typeof setTimeout> | null = null
  let restoring = false

  const settle = () => {
    if (timer) {
      clearTimeout(timer)
      timer = null
    }
    const snap = takeSnapshot(root)
    if (snap.html === stack[index]!.html) {
      // markup unchanged (a caret move): only keep the newest caret for the step we are on
      stack[index] = snap
      return
    }
    stack.length = index + 1 // a new edit drops the redo tail
    stack.push(snap)
    index += 1
  }

  const observer = new MutationObserver(() => {
    if (restoring) return
    if (timer) clearTimeout(timer)
    timer = setTimeout(settle, IDLE_MS)
  })
  observer.observe(root, {
    subtree: true,
    childList: true,
    attributes: true,
    characterData: true,
  })

  // changes whose observer callback has not run yet count as pending too (a key can land between)
  const hasPending = () => observer.takeRecords().length > 0 || timer !== null

  const restore = (snap: Snapshot) => {
    restoring = true
    try {
      root.innerHTML = snap.html
      observer.takeRecords()
      const sel = window.getSelection()
      if (sel && snap.sel) {
        const anchor = nodeAt(root, snap.sel.anchor.path)
        const focus = nodeAt(root, snap.sel.focus.path)
        sel.setBaseAndExtent(
          anchor,
          clampOffset(anchor, snap.sel.anchor.offset),
          focus,
          clampOffset(focus, snap.sel.focus.offset),
        )
      }
    } finally {
      restoring = false
    }
  }

  return {
    flush() {
      if (hasPending()) settle()
    },
    undo() {
      if (hasPending()) settle()
      if (index > 0) {
        index -= 1
        restore(stack[index]!)
      }
      return true
    },
    redo() {
      if (hasPending()) settle()
      if (index < stack.length - 1) {
        index += 1
        restore(stack[index]!)
      }
      return true
    },
    dispose() {
      observer.disconnect()
      if (timer) clearTimeout(timer)
      timer = null
    },
  }
}

/** The history of the editor that is open now (there is at most one). */
let active: TextHistory | null = null

export function setActiveTextHistory(history: TextHistory | null): void {
  active = history
}

/** ⌘Z while editing text; false when no overlay is open (the caller falls back). */
export function undoTextEdit(): boolean {
  return active ? active.undo() : false
}

export function redoTextEdit(): boolean {
  return active ? active.redo() : false
}
