import { beforeEach, describe, expect, it, vi } from 'vitest'
import { buildCtxItems } from '../src/renderer/context-menu-items'
import * as arrangeActions from '../src/renderer/arrange-actions'
import { t } from '../src/renderer/i18n/locale'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/konva-adapter', () => ({ isEditableText: () => false }))
vi.mock('../src/renderer/clipboard-actions', () => ({}))
vi.mock('../src/renderer/slide-actions', () => ({}))
vi.mock('../src/renderer/show-actions', () => ({}))
vi.mock('../src/renderer/insert-actions', () => ({}))
vi.mock('../src/renderer/picture-edit-actions', () => ({}))
vi.mock('../src/renderer/table-actions', () => ({}))
vi.mock('../src/renderer/style-actions', () => ({}))
vi.mock('../src/renderer/arrange-actions', () => ({ reorderSelected: vi.fn() }))

const shape = (id: string) => ({ type: 'shape', sourceId: id, box: { x: 0, y: 0, w: 10, h: 10 } })

function elementCtx(selectedIds: string[], targetId: string) {
  return {
    ctxMenu: { kind: 'element', x: 0, y: 0, targetId },
    slides: [{}],
    sections: [],
    selectedIds,
    slide: { nodes: [shape('a'), shape('b'), shape('c')] },
    current: 0,
    hasClipboard: false,
    canPasteSlide: false,
  } as unknown as ActionCtx
}

describe('element context menu: arrange order with a multi-selection', () => {
  beforeEach(() => vi.clearAllMocks())

  for (const [key, dir] of [
    ['appCtxBringToFront', 'front'],
    ['appCtxSendToBack', 'back'],
    ['appCtxBringForward', 'forward'],
    ['appCtxSendBackward', 'backward'],
  ] as const) {
    it(`${key} stays enabled when several elements are selected (PowerPoint reorders them together)`, () => {
      const ctx = elementCtx(['a', 'b'], 'b')
      const item = buildCtxItems(ctx).find((i) => i?.label === t(key))
      expect(item).toBeTruthy()
      expect(item?.disabled).toBeFalsy()
      item?.onClick?.()
      expect(arrangeActions.reorderSelected).toHaveBeenCalledWith(ctx, 'b', dir)
    })
  }
})
