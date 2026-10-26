import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createFontSizeStepper } from '../src/renderer/components/font-size-step'
import {
  growFontSize,
  shrinkFontSize,
  stepFontSizeLadder,
} from '../src/renderer/components/font-size-ladder'

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

describe('Word grow/shrink font ladder', () => {
  const up = (from: number, n: number) => {
    const out = [from]
    for (let i = 0; i < n; i++) out.push(growFontSize(out[out.length - 1]))
    return out
  }
  it('grows along the measured Word steps, then ten at a time past 72', () => {
    expect(up(5, 28)).toEqual([
      5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 80, 90,
      100, 110, 120, 130, 140, 150,
    ])
  })
  it('shrinks back down and keeps going below 5 to 1', () => {
    const down = [150]
    for (let i = 0; i < 40; i++) down.push(shrinkFontSize(down[down.length - 1]))
    expect(down.slice(0, 8)).toEqual([150, 140, 130, 120, 110, 100, 90, 80])
    expect(down.slice(8, 12)).toEqual([72, 48, 36, 28])
    const at5 = down.indexOf(5)
    expect(down.slice(at5, at5 + 6)).toEqual([5, 4, 3, 2, 1, 1])
  })
  it('snaps off-ladder sizes to the next rung in the stepping direction', () => {
    expect(growFontSize(11.5)).toBe(12)
    expect(shrinkFontSize(11.5)).toBe(11)
    expect(growFontSize(75)).toBe(80)
    expect(shrinkFontSize(75)).toBe(72)
    expect(growFontSize(3.5)).toBe(4)
    expect(shrinkFontSize(3.5)).toBe(3)
    expect(growFontSize(4)).toBe(5)
  })
  it('clamps at Word limits', () => {
    expect(shrinkFontSize(1)).toBe(1)
    expect(growFontSize(1638)).toBe(1638)
    expect(growFontSize(1635)).toBe(1638)
    expect(stepFontSizeLadder(11, 1)).toBe(12)
    expect(stepFontSizeLadder(11, -1)).toBe(10.5)
  })
})

describe('mixed font sizes step per run', () => {
  let editor: Editor
  let frames: Map<number, FrameRequestCallback>
  let serial: number
  let stepper: ReturnType<typeof createFontSizeStepper>
  const paint = () => {
    const callbacks = [...frames.values()]
    frames.clear()
    callbacks.forEach((callback) => callback(0))
  }
  const run = (text: string, half: number) => ({
    type: 'text',
    text,
    marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: half, color: 'FF0000' } }],
  })
  const sizes = () => {
    const out: Array<[string, number]> = []
    editor.state.doc.descendants((n) => {
      if (n.isText) {
        const m = n.marks.find((x) => x.type.name === 'docTextStyle')
        out.push([n.text!, (m?.attrs.sizeHalfPoints as number) / 2])
      }
    })
    return out
  }
  beforeEach(() => {
    frames = new Map()
    serial = 0
    vi.stubGlobal('requestAnimationFrame', (cb: FrameRequestCallback) => {
      frames.set(++serial, cb)
      return serial
    })
    vi.stubGlobal('cancelAnimationFrame', (id: number) => frames.delete(id))
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [run('AAAA', 20), run('BBBB', 40)] },
          { type: 'docParagraph', content: [run('CCCC', 144)] },
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
  it('walks each run one rung from its own size (10+20 -> 10.5+22, 72 -> 80)', () => {
    stepper.step(editor, 10, (b) => growFontSize(b))
    paint()
    expect(sizes()).toEqual([
      ['AAAA', 10.5],
      ['BBBB', 22],
      ['CCCC', 80],
    ])
    stepper.step(editor, 10, (b) => shrinkFontSize(b))
    paint()
    expect(sizes()).toEqual([
      ['AAAA', 10],
      ['BBBB', 20],
      ['CCCC', 72],
    ])
  })
  it('replays several clicks of one frame per run, in one transaction and one undo step', () => {
    const update = vi.fn()
    editor.on('update', update)
    for (let i = 0; i < 3; i++) stepper.step(editor, 10, (b) => growFontSize(b))
    paint()
    expect(sizes()).toEqual([
      ['AAAA', 12],
      ['BBBB', 26],
      ['CCCC', 100],
    ])
    expect(update).toHaveBeenCalledTimes(1)
    editor.commands.undo()
    expect(sizes()).toEqual([
      ['AAAA', 10],
      ['BBBB', 20],
      ['CCCC', 72],
    ])
  })
  it('only touches the selected part of a run and keeps its other formatting', () => {
    editor.commands.setTextSelection({ from: 3, to: 5 })
    stepper.step(editor, 10, (b) => growFontSize(b))
    paint()
    expect(sizes()).toEqual([
      ['AA', 10],
      ['AA', 10.5],
      ['BBBB', 20],
      ['CCCC', 72],
    ])
    expect(editor.state.doc.firstChild!.child(1).marks[0].attrs.color).toBe('FF0000')
  })
})
