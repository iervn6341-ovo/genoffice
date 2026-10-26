import { describe, expect, it } from 'vitest'
import { undoShortcutAction, type UndoShortcutEvent } from '../src/renderer/undo-shortcut'

const target = (inside: { container?: boolean; field?: boolean }) => ({
  closest: (selector: string) => {
    if (selector === '#univer-container') return inside.container ? {} : null
    if (selector.includes('input, textarea, select') || selector === 'input, textarea, select') {
      return inside.field ? {} : null
    }
    return null
  },
})
const key = (over: Partial<UndoShortcutEvent> & { key: string }): UndoShortcutEvent => ({
  metaKey: true,
  ctrlKey: false,
  altKey: false,
  shiftKey: false,
  defaultPrevented: false,
  isComposing: false,
  target: target({}),
  ...over,
})

describe('undo shortcut outside the grid', () => {
  it('undoes and redoes from the ribbon / page body', () => {
    expect(undoShortcutAction(key({ key: 'z' }), false)).toBe('undo')
    expect(undoShortcutAction(key({ key: 'z', metaKey: false, ctrlKey: true }), false)).toBe('undo')
    expect(undoShortcutAction(key({ key: 'Z', shiftKey: true }), false)).toBe('redo')
    expect(undoShortcutAction(key({ key: 'y', metaKey: false, ctrlKey: true }), false)).toBe('redo')
  })
  it('leaves keys inside the Univer container to Univer', () => {
    expect(undoShortcutAction(key({ key: 'z', target: target({ container: true }) }), false)).toBeNull()
  })
  it('keeps native undo in text fields and while editing a cell', () => {
    expect(undoShortcutAction(key({ key: 'z', target: target({ field: true }) }), false)).toBeNull()
    expect(undoShortcutAction(key({ key: 'z' }), true)).toBeNull()
  })
  it('ignores handled, composing, alt and plain keys', () => {
    expect(undoShortcutAction(key({ key: 'z', defaultPrevented: true }), false)).toBeNull()
    expect(undoShortcutAction(key({ key: 'z', isComposing: true }), false)).toBeNull()
    expect(undoShortcutAction(key({ key: 'z', altKey: true }), false)).toBeNull()
    expect(undoShortcutAction(key({ key: 'z', metaKey: false }), false)).toBeNull()
    expect(undoShortcutAction(key({ key: 'a' }), false)).toBeNull()
  })
})
