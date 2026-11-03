import { memo, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'
import { collectOutlineItems, type OutlineItem } from '../editor/headings'
import { useI18n } from '../i18n/locale'
import { IconClose, IconOutlineView, IconSearch } from './icons'

interface FloatingOutlinePaneProps {
  editor: Editor | null
  doc: PmNode | null
  visible: boolean
  isDark?: boolean
  canFitRightGutter?: boolean
  onClose: () => void
  onActiveHeadingChange?: (text: string | null) => void
}

/**
 * Floating outline & Table of Contents (TOC) navigation pane:
 * 1. Automatically parses the document's Table of Contents (TOC), falling back to headings.
 * 2. In wide windows with right-side gray gutter: stays fixed on the right; hover expands, mouse-leave collapses.
 * 3. In narrow windows without right-side gray space: hides from canvas and pops up from the bottom status bar.
 */
export const FloatingOutlinePane = memo(function FloatingOutlinePane({
  editor,
  doc,
  visible,
  isDark = false,
  canFitRightGutter = true,
  onClose,
  onActiveHeadingChange,
}: FloatingOutlinePaneProps) {
  const { t, lang } = useI18n()
  const isZhTW = lang === 'zh-TW'
  const isZh = lang === 'zh'

  const { items } = useMemo(() => {
    if (!doc) return { items: [] as OutlineItem[] }
    return collectOutlineItems(doc)
  }, [doc])

  const hasToc = useMemo(() => items.some((i) => i.isToc), [items])

  const ui = useMemo(
    () => ({
      title: hasToc
        ? isZhTW
          ? '文件目錄'
          : isZh
            ? '文档目录'
            : 'Table of Contents'
        : isZhTW
          ? '標題大綱'
          : isZh
            ? '标题大纲'
            : 'Document Outline',
      close: t('appClose'),
      searchPlaceholder: hasToc
        ? isZhTW
          ? '搜尋目錄項目...'
          : isZh
            ? '搜索目录项目...'
            : 'Search TOC...'
        : isZhTW
          ? '搜尋標題大綱...'
          : isZh
            ? '搜索标题大纲...'
            : 'Search outline...',
      noHeadings: hasToc
        ? isZhTW
          ? '無目錄項目'
          : isZh
            ? '无目录项目'
            : 'No TOC items'
        : t('appNavNoHeadings'),
      noResults: isZhTW ? '無符合的搜尋結果' : isZh ? '无符合的搜索结果' : 'No matching results',
    }),
    [hasToc, isZhTW, isZh, t],
  )

  const [isHoverExpanded, setIsHoverExpanded] = useState(false)
  const [isCollapsing, setIsCollapsing] = useState(false)
  const hoverLeaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const collapseEndTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  // Animated popup entry/exit state for bottom popup mode
  const [shouldRenderPopup, setShouldRenderPopup] = useState(visible)
  const [isPopupClosing, setIsPopupClosing] = useState(false)
  const popupCloseTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  useEffect(() => {
    if (visible) {
      if (popupCloseTimerRef.current) {
        clearTimeout(popupCloseTimerRef.current)
        popupCloseTimerRef.current = null
      }
      setIsPopupClosing(false)
      setShouldRenderPopup(true)
    } else if (shouldRenderPopup && !isPopupClosing) {
      setIsPopupClosing(true)
      popupCloseTimerRef.current = setTimeout(() => {
        setShouldRenderPopup(false)
        setIsPopupClosing(false)
      }, 240)
    }
  }, [visible, shouldRenderPopup, isPopupClosing])

  useEffect(() => {
    return () => {
      if (hoverLeaveTimerRef.current) clearTimeout(hoverLeaveTimerRef.current)
      if (collapseEndTimerRef.current) clearTimeout(collapseEndTimerRef.current)
      if (popupCloseTimerRef.current) clearTimeout(popupCloseTimerRef.current)
    }
  }, [])

  const handleMouseEnter = useCallback(() => {
    if (hoverLeaveTimerRef.current) {
      clearTimeout(hoverLeaveTimerRef.current)
      hoverLeaveTimerRef.current = null
    }
    if (collapseEndTimerRef.current) {
      clearTimeout(collapseEndTimerRef.current)
      collapseEndTimerRef.current = null
    }
    setIsCollapsing(false)
    setIsHoverExpanded(true)
  }, [])

  const handleMouseLeave = useCallback(() => {
    if (hoverLeaveTimerRef.current) {
      clearTimeout(hoverLeaveTimerRef.current)
      hoverLeaveTimerRef.current = null
    }
    // Retract immediately upon mouse leave without delay
    setIsCollapsing(true)
    if (collapseEndTimerRef.current) clearTimeout(collapseEndTimerRef.current)
    collapseEndTimerRef.current = setTimeout(() => {
      collapseEndTimerRef.current = null
      setIsHoverExpanded(false)
      setIsCollapsing(false)
    }, 160)
  }, [])

  const handleClose = useCallback(
    (isBottomPopup: boolean) => {
      if (isBottomPopup) {
        setIsPopupClosing(true)
        if (popupCloseTimerRef.current) clearTimeout(popupCloseTimerRef.current)
        popupCloseTimerRef.current = setTimeout(() => {
          setShouldRenderPopup(false)
          setIsPopupClosing(false)
          onClose()
        }, 200)
      } else {
        setIsCollapsing(true)
        if (collapseEndTimerRef.current) clearTimeout(collapseEndTimerRef.current)
        collapseEndTimerRef.current = setTimeout(() => {
          collapseEndTimerRef.current = null
          setIsHoverExpanded(false)
          setIsCollapsing(false)
        }, 160)
      }
    },
    [onClose],
  )

  const [activePos, setActivePos] = useState<number | null>(null)
  const [hoveredItemId, setHoveredItemId] = useState<string | null>(null)
  const [search, setSearch] = useState('')

  // Track active heading/TOC entry based on editor caret position
  useEffect(() => {
    if (!editor) return

    const updateActiveFromSelection = () => {
      const caret = editor.state.selection.from
      if (items.length === 0) {
        setActivePos(null)
        onActiveHeadingChange?.(null)
        return
      }

      let closest: OutlineItem | null = null
      for (const item of items) {
        const itemPos = item.targetPos ?? item.pos
        if (itemPos <= caret) {
          closest = item
        } else {
          break
        }
      }
      const active = closest ?? items[0] ?? null
      setActivePos(active?.pos ?? null)
      onActiveHeadingChange?.(active?.text ?? null)
    }

    updateActiveFromSelection()
    editor.on('selectionUpdate', updateActiveFromSelection)
    editor.on('transaction', updateActiveFromSelection)
    return () => {
      editor.off('selectionUpdate', updateActiveFromSelection)
      editor.off('transaction', updateActiveFromSelection)
    }
  }, [editor, items, onActiveHeadingChange])

  // Jump to item position or target section in the document on click and scroll it to top
  const handleItemClick = useCallback(
    (item: OutlineItem) => {
      if (!editor) return
      const jumpPos = item.targetPos ?? item.pos

      // 1. Move selection to target position without default browser bottom-scroll
      try {
        const docSize = editor.state.doc.content.size
        const safePos = Math.min(Math.max(0, jumpPos + 1), docSize)
        const tr = editor.state.tr.setSelection(
          TextSelection.near(editor.state.doc.resolve(safePos)),
        )
        editor.view.dispatch(tr)
        editor.view.dom.focus({ preventScroll: true })
      } catch {
        // Fallback for atomic or protected nodes
      }

      // 2. Locate target element in the DOM
      let targetEl: HTMLElement | null = null
      try {
        const dom = editor.view.nodeDOM(jumpPos)
        if (dom instanceof HTMLElement) {
          targetEl = dom
        } else {
          const { node } = editor.view.domAtPos(jumpPos)
          targetEl = (node instanceof HTMLElement ? node : node.parentElement) as HTMLElement | null
        }
      } catch {
        // Fallback if jumpPos is outside valid bounds
      }

      if (!targetEl && item.anchor) {
        targetEl = editor.view.dom.querySelector<HTMLElement>(`[data-toc-anchor="${item.anchor}"]`)
      }
      if (!targetEl) {
        const cleanText = item.text.replace(/\s+/g, '')
        const allHeadings = editor.view.dom.querySelectorAll<HTMLElement>(
          'h1, h2, h3, h4, h5, h6, .doc-li, p',
        )
        for (const el of allHeadings) {
          if ((el.textContent ?? '').replace(/\s+/g, '').includes(cleanText)) {
            targetEl = el
            break
          }
        }
      }

      // 3. Smoothly scroll target element to the very top of the editor scroll viewport
      if (targetEl) {
        const scroller = (editor.view.dom.closest('.editor-scroll') ??
          document.querySelector('.editor-scroll')) as HTMLElement | null
        if (scroller) {
          const targetRect = targetEl.getBoundingClientRect()
          const scrollerRect = scroller.getBoundingClientRect()
          // 24px comfortable breathing room below top ruler/toolbar
          const topOffset = 24
          const targetScrollTop = scroller.scrollTop + (targetRect.top - scrollerRect.top) - topOffset
          scroller.scrollTo({
            top: Math.max(0, targetScrollTop),
            behavior: 'smooth',
          })
        } else {
          targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' })
        }
      }

      setActivePos(item.pos)
      onActiveHeadingChange?.(item.text)
    },
    [editor, onActiveHeadingChange],
  )

  if (items.length === 0) return null

  const effectiveIsDark =
    isDark ||
    (typeof document !== 'undefined' &&
      (document.documentElement.getAttribute('data-theme') === 'dark' ||
        document.body.classList.contains('dark-theme')))

  // Renders the full expanded outline card
  const renderExpandedCard = (isBottomPopup: boolean) => {
    const filteredItems = search.trim()
      ? items.filter((item) => item.text.toLowerCase().includes(search.trim().toLowerCase()))
      : items

    return (
      <aside
        className={`floating-outline-card ${isBottomPopup ? 'bottom-popup' : 'in-gutter'} ${
          isBottomPopup && isPopupClosing ? 'is-closing' : ''
        } ${!isBottomPopup && isCollapsing ? 'is-collapsing' : ''} ${
          effectiveIsDark ? 'theme-dark' : 'theme-light'
        }`}
        aria-label={ui.title}
      >
        {/* Header bar */}
        <div className="floating-outline-header">
          <div className="floating-outline-title-wrap">
            <IconOutlineView size={15} />
            <span>{ui.title}</span>
            <span className="floating-outline-count-badge">{items.length}</span>
          </div>
          <div className="floating-outline-actions">
            <button
              className="floating-outline-btn"
              aria-label={ui.close}
              data-tip={ui.close}
              onClick={() => handleClose(isBottomPopup)}
            >
              <IconClose size={13} />
            </button>
          </div>
        </div>

        {/* Search filter if more than 5 items */}
        {items.length > 5 && (
          <div className="floating-outline-search-wrap">
            <IconSearch size={12} />
            <input
              type="text"
              className="floating-outline-search"
              placeholder={ui.searchPlaceholder}
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
            {search && (
              <button
                className="floating-outline-btn"
                style={{ width: 16, height: 16, fontSize: 11 }}
                onClick={() => setSearch('')}
              >
                ×
              </button>
            )}
          </div>
        )}

        {/* Outline item list */}
        <div className="floating-outline-list">
          {filteredItems.map((item) => {
            const isActive = activePos === item.pos
            const levelCls = `floating-outline-l${Math.min(item.level, 4)}`
            return (
              <button
                key={item.id}
                className={`floating-outline-item ${levelCls} ${isActive ? 'is-active' : ''}`}
                data-tip={item.text}
                onClick={() => handleItemClick(item)}
              >
                <span className="floating-outline-item-text">{item.text}</span>
                {item.page && <span className="floating-outline-item-page">{item.page}</span>}
              </button>
            )
          })}

          {filteredItems.length === 0 && (
            <div className="floating-outline-empty">
              {search ? ui.noResults : ui.noHeadings}
            </div>
          )}
        </div>
      </aside>
    )
  }

  // Renders the sleek Notion-style right mini-bar
  const renderMiniBar = () => {
    const miniItems = items.slice(0, 36)
    return (
      <aside
        className={`notion-outline-bar ${effectiveIsDark ? 'theme-dark' : 'theme-light'}`}
        aria-label={ui.title}
      >
        <div className="notion-bar-expand-btn" aria-label={ui.title} data-tip={ui.title}>
          <IconOutlineView size={13} />
        </div>
        <div className="notion-bar-tracks">
          {miniItems.map((item) => {
            const isActive = activePos === item.pos
            const isHovered = hoveredItemId === item.id
            const levelCls = `level-${Math.min(item.level, 4)}`
            return (
              <div
                key={item.id}
                className={`notion-bar-item ${levelCls} ${isActive ? 'is-active' : ''}`}
                onClick={() => handleItemClick(item)}
                onMouseEnter={() => setHoveredItemId(item.id)}
                onMouseLeave={() => setHoveredItemId(null)}
              >
                <div className="notion-bar-dash" />
                {isHovered && (
                  <div className="notion-bar-tooltip">
                    {item.text}
                    {item.page ? ` (${item.page})` : ''}
                  </div>
                )}
              </div>
            )
          })}
        </div>
      </aside>
    )
  }

  // Mode A: When window is narrow (not enough right gray space), only show as bottom popup when requested
  if (!canFitRightGutter) {
    if (!shouldRenderPopup) return null
    return renderExpandedCard(true /* isBottomPopup */)
  }

  // Mode B: When window is wide (right gray space is present), stay fixed on the right side.
  // Hovering automatically expands the full TOC card; moving mouse out collapses it back to mini bar smoothly.
  return (
    <div
      className={`notion-outline-container ${isHoverExpanded ? 'is-expanded' : 'is-collapsed'}`}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
    >
      {isHoverExpanded ? renderExpandedCard(false) : renderMiniBar()}
    </div>
  )
})
