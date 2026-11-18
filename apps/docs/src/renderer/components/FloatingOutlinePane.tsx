import { memo, useCallback, useMemo, type RefObject } from 'react'
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'
import { FloatingOutline, floatingOutlineLabels, useOutlineScrollSpy } from '@genoffice/ui'
import { collectOutlineItems } from '../editor/headings'
import { useI18n } from '../i18n/locale'

/** Docs supplies TOC destinations and page labels; shared UI owns all interactions. */
export const FloatingOutlinePane = memo(function FloatingOutlinePane({
  editor,
  doc,
  visible,
  onVisibleChange,
  scrollRef,
  zoom,
}: {
  editor: Editor
  doc: PmNode
  visible: boolean
  onVisibleChange: (visible: boolean) => void
  scrollRef: RefObject<HTMLElement | null>
  zoom: number
}) {
  const { lang } = useI18n()
  const { items } = useMemo(() => collectOutlineItems(doc), [doc])
  const anchors = useCallback(() => {
    const scroller = scrollRef.current
    if (!scroller) return []
    const origin = scroller.getBoundingClientRect().top - scroller.scrollTop
    return items.flatMap((item) => {
      const node = editor.view.nodeDOM(item.targetPos ?? item.pos)
      return node instanceof HTMLElement && node.getClientRects().length > 0
        ? [{ id: item.id, top: node.getBoundingClientRect().top - origin }]
        : []
    })
    // Zoom changes display geometry even when the unscaled document is identical.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor, items, scrollRef, zoom])
  const activeId = useOutlineScrollSpy(scrollRef, editor.view.dom, anchors)
  return (
    <FloatingOutline
      {...floatingOutlineLabels(lang)}
      items={items}
      activeId={activeId}
      open={visible}
      onOpenChange={onVisibleChange}
      onSelect={(id) => {
        const item = items.find((item) => item.id === id)
        if (!item) return
        const pos = item.targetPos ?? item.pos
        const selection = TextSelection.near(
          editor.state.doc.resolve(Math.min(pos + 1, editor.state.doc.content.size)),
        )
        editor.view.dispatch(editor.state.tr.setSelection(selection))
        editor.view.dom.focus({ preventScroll: true })
        const node = editor.view.nodeDOM(pos)
        const scroller = scrollRef.current
        if (node instanceof HTMLElement && scroller) {
          scroller.scrollTo({
            top: Math.max(
              0,
              scroller.scrollTop +
                node.getBoundingClientRect().top -
                scroller.getBoundingClientRect().top -
                24,
            ),
            behavior: 'smooth',
          })
        }
      }}
    />
  )
})
