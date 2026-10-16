// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest'
import { insertEquation, insertSymbol } from '../src/renderer/insert-actions'
import { saveEditSelection } from '../src/renderer/TextEditOverlay'
import type { ActionCtx } from '../src/renderer/action-context'

function context() {
  const slide = { widthPx: 960, heightPx: 540, scale: 1, nodes: [] }
  const ctx = {
    slide,
    current: 0,
    setEqDialogOpen: vi.fn(),
    setSlides: vi.fn(),
    setDirty: vi.fn(),
    setSelectedIds: vi.fn(),
    setStatus: vi.fn(),
    applySlide: vi.fn(),
  } as unknown as ActionCtx
  return { ctx, slide }
}

describe('Slides equation and symbol actions', () => {
  it('sends native equation ops through the journaled IPC and preserves the edit target', async () => {
    const { ctx, slide } = context()
    const applyTxn = vi.fn(async (_req: unknown) => ({
      applied: true,
      records: [{ op: 'insertEquation', created: ['math1'] }],
      slides: [slide],
    }))
    ;(window as any).slidesApi = { applyTxn }
    await insertEquation(ctx, '\\frac{a}{b}')
    expect(applyTxn).toHaveBeenCalledOnce()
    const insert = applyTxn.mock.calls[0]![0] as any
    expect(insert.ops[0]).toMatchObject({
      op: 'insertEquation',
      target: { slide: 0 },
      latex: '\\frac{a}{b}',
    })
    expect(insert.ops[0].box.cx).toBe(420 * 9525)
    expect(ctx.setSelectedIds as any).toHaveBeenCalledWith(['math1'])
    await insertEquation(ctx, '\\sqrt{x}', 'math1')
    expect((applyTxn.mock.calls[1]![0] as any).ops[0]).toEqual({
      op: 'replaceEquation',
      target: { slide: 0, el: 'math1' },
      latex: '\\sqrt{x}',
    })
  })

  it('creates an ordinary text box for a symbol when no text caret is active', async () => {
    const { ctx, slide } = context()
    const addElement = vi.fn(async () => ({ sourceId: 'symbol1', slide }))
    ;(window as any).slidesApi = { addElement }
    await insertSymbol(ctx, '∑', 'Cambria Math')
    expect(addElement).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'textbox',
        paragraphs: [{ runs: [{ text: '∑', fontFamily: 'Cambria Math', fontSize: 28 }] }],
      }),
    )
    expect(ctx.setSelectedIds as any).toHaveBeenCalledWith(['symbol1'])
  })

  it('inserts the chosen symbol and font at the saved text selection', async () => {
    const { ctx } = context()
    const editor = document.createElement('div')
    editor.contentEditable = 'true'
    Object.defineProperty(editor, 'isContentEditable', { value: true })
    editor.tabIndex = 0
    editor.textContent = 'AB'
    document.body.appendChild(editor)
    editor.focus()
    const range = document.createRange()
    range.setStart(editor.firstChild!, 1)
    range.collapse(true)
    window.getSelection()!.removeAllRanges()
    window.getSelection()!.addRange(range)
    saveEditSelection()
    const execCommand = vi.fn((_command: string, _showUI: boolean, _value: string) => true)
    Object.defineProperty(document, 'execCommand', { value: execCommand, configurable: true })
    const addElement = vi.fn()
    ;(window as any).slidesApi = { addElement }
    await insertSymbol(ctx, '∑', 'Cambria Math')
    expect(execCommand).toHaveBeenCalledWith('insertHTML', false, expect.stringContaining('∑'))
    expect(execCommand.mock.calls[0]![2]).toContain('font-family: Cambria Math')
    expect(addElement).not.toHaveBeenCalled()
    editor.remove()
  })

  it('keeps arbitrary installed font families but rejects style-attribute injection', async () => {
    const { ctx } = context()
    const editor = document.createElement('div')
    editor.contentEditable = 'true'
    Object.defineProperty(editor, 'isContentEditable', { value: true })
    editor.tabIndex = 0
    editor.textContent = 'AB'
    document.body.appendChild(editor)
    editor.focus()
    const range = document.createRange()
    range.setStart(editor.firstChild!, 1)
    range.collapse(true)
    window.getSelection()!.removeAllRanges()
    window.getSelection()!.addRange(range)
    saveEditSelection()
    const execCommand = vi.fn((_command: string, _showUI: boolean, _value: string) => true)
    Object.defineProperty(document, 'execCommand', { value: execCommand, configurable: true })
    await insertSymbol(ctx, '§', 'Helvetica Neue')
    expect(execCommand.mock.calls[0]![2]).toContain('font-family: Helvetica Neue')
    await insertSymbol(ctx, '§', 'x; background:url(evil)')
    expect(execCommand.mock.calls[1]![2]).toContain('font-family: Cambria Math')
    editor.remove()
  })
})
