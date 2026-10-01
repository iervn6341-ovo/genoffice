import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createFontSizeStepper } from '../src/renderer/components/font-size-step'

describe('font size frame updates', () => {
  let editor: Editor
  let frames: Map<number, FrameRequestCallback>
  let serial: number
  let stepper: ReturnType<typeof createFontSizeStepper>
  const paint = () => {
    const callbacks = [...frames.values()]
    frames.clear()
    callbacks.forEach((callback) => callback(0))
  }
  const size = () => editor.getAttributes('docTextStyle').sizeHalfPoints / 2
  beforeEach(() => {
    frames = new Map()
    serial = 0
    vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
      frames.set(++serial, callback)
      return serial
    })
    vi.stubGlobal('cancelAnimationFrame', (id: number) => frames.delete(id))
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Example',
                marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 22 } }],
              },
            ],
          },
        ],
      },
    })
    editor.commands.selectAll()
    stepper = createFontSizeStepper()
  })
  afterEach(() => {
    stepper.cancel()
    editor.destroy()
    vi.unstubAllGlobals()
  })
  it('combines clicks in one frame and uses current editor marks instead of stale toolbar props', () => {
    const update = vi.fn()
    editor.on('update', update)
    for (let i = 0; i < 6; i++) stepper.step(editor, 36, (base) => base + 1)
    expect(size()).toBe(11)
    paint()
    expect(size()).toBe(17)
    expect(update).toHaveBeenCalledTimes(1)
    stepper.step(editor, 36, (base) => base - 1)
    paint()
    expect(size()).toBe(16)
    expect(update).toHaveBeenCalledTimes(2)
  })
  it('does not wait for the end of a click burst to repaint', () => {
    for (let value = 12; value <= 18; value++) {
      stepper.step(editor, 11, (base) => base + 1)
      paint()
      expect(size()).toBe(value)
    }
  })
  it('discards queued work when selection moves', () => {
    stepper.step(editor, 11, (base) => base + 1)
    editor.commands.setTextSelection(2)
    paint()
    expect(size()).toBe(11)
  })
  it('does not overwrite a size entered directly before the frame runs', () => {
    stepper.step(editor, 11, (base) => base + 1)
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 48 })
    stepper.step(editor, 11, (base) => base + 1)
    paint()
    expect(size()).toBe(25)
  })
  it('invalidates caret work if another stored format changes', () => {
    editor.commands.setTextSelection(2)
    stepper.step(editor, 11, (base) => base + 1)
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 40 })
    paint()
    expect(size()).toBe(20)
  })
  it('cancels on unmount and never writes to a read-only editor', () => {
    stepper.step(editor, 11, (base) => base + 1)
    stepper.cancel()
    paint()
    expect(size()).toBe(11)
    stepper.step(editor, 11, (base) => base + 1)
    editor.setEditable(false)
    paint()
    expect(size()).toBe(11)
  })
  it('keeps one frame burst undoable', () => {
    for (let i = 0; i < 3; i++) stepper.step(editor, 11, (base) => base + 1)
    paint()
    expect(size()).toBe(14)
    editor.commands.undo()
    expect(size()).toBe(11)
    editor.commands.redo()
    expect(size()).toBe(14)
  })
})
