/**
 * Right-clicking a sheet tab must not activate that sheet (Univer does, twice).
 *
 * Univer's sheet bar activates a non-active tab on the right button's
 * `pointerdown` (SlideTabBar._downAction → onChangeTab) and, if that was
 * skipped, again in its `contextmenu` handler (SetWorksheetActiveOperation,
 * then opens the menu). Then every menu item runs with
 * `{ value, subUnitId: <active sheet> }`, so the menu only ever acted on the
 * active sheet.
 *
 * Here:
 *  - a right-button `pointerdown` on a non-selected tab is swallowed in the
 *    capture phase, so the sheet bar never sees it;
 *  - during the tab-list `contextmenu` dispatch, SetWorksheetActiveOperation is
 *    turned into a no-op, so Univer simply opens the menu next to the tab;
 *  - while that menu is the open one, the item's command is re-aimed at the
 *    right-clicked sheet: delete / copy / tab colour / hide take a `subUnitId`
 *    and just get the clicked sheet's id; rename and the protection commands
 *    work on the ACTIVE sheet (the tab editor, the permission panel), so for
 *    those the clicked sheet is activated first, at the moment the item runs;
 *  - the menu's bottom edge is pinned to the top of the tab bar. Univer's Popup
 *    measures the menu's height once, when it opens, but the items it shows
 *    (Unhide, protection) settle a moment later, so a menu that ends up
 *    shorter floats above the bar with a gap (up to ~80 px).
 */
import { ICommandService, IUniverInstanceService, UniverInstanceType } from '@univerjs/core'
import type { Workbook } from '@univerjs/core'

import type { UniverRuntime } from './univer-state'

const SET_ACTIVE_COMMAND = 'sheet.operation.set-worksheet-active'
const TAB_ITEM_SELECTOR = '[data-u-comp="slide-tab-item"]'
const TAB_LIST_SELECTOR = '[data-u-comp="slide-tab-bar"]'
/** containers whose clicks belong to the open tab menu (menu panel, submenus, colour picker) */
const MENU_SUBMENU_SELECTOR = '[data-u-context-menu-submenu]'
const MENU_SURFACE_SELECTOR = '.univer-popup, [data-u-context-menu-submenu]'

/** commands whose handler takes `subUnitId`: only the target id has to change */
const RETARGET_COMMANDS = new Set([
  'sheet.command.remove-sheet-confirm',
  'sheet.command.copy-sheet',
  'sheet.command.set-tab-color',
  'sheet.command.set-worksheet-hidden',
])
/** commands that act on the active sheet: it has to be the clicked one first */
const ACTIVATE_FIRST_COMMANDS = new Set([
  'sheet.operation.rename-sheet',
  'sheet.command.add-range-protection-from-sheet-bar',
  'sheet.command.delete-worksheet-protection-from-sheet-bar',
  'sheet.command.change-sheet-protection-from-sheet-bar',
  'sheet.command.view-sheet-permission-from-sheet-bar',
])

/** only a right-button press on a tab that is not the selected one is swallowed */
export function shouldSwallowTabPointerDown(button: number, tabIsSelected: boolean): boolean {
  return button === 2 && !tabIsSelected
}

export type TabMenuPlan =
  | { kind: 'passthrough' }
  | { kind: 'retarget'; params: Record<string, unknown> }
  | { kind: 'activate-first'; params: Record<string, unknown> }

/**
 * What to do with a command while a tab menu opened on `menuTarget` is up.
 * Menu items always arrive as `{ value, subUnitId: <active sheet id> }`.
 */
export function planTabMenuCommand(
  menuTarget: string | null,
  activeSheetId: string | null,
  commandId: string,
  params: unknown,
): TabMenuPlan {
  if (!menuTarget || !activeSheetId || menuTarget === activeSheetId) return { kind: 'passthrough' }
  if (!params || typeof params !== 'object') return { kind: 'passthrough' }
  const p = params as Record<string, unknown>
  if (p.subUnitId !== activeSheetId || !('value' in p)) return { kind: 'passthrough' }
  if (RETARGET_COMMANDS.has(commandId)) {
    return { kind: 'retarget', params: { ...p, subUnitId: menuTarget } }
  }
  if (ACTIVATE_FIRST_COMMANDS.has(commandId)) {
    return { kind: 'activate-first', params: { ...p, subUnitId: menuTarget } }
  }
  return { kind: 'passthrough' }
}

/**
 * `bottom` (px from the window's bottom) that puts a menu's lower edge on the
 * bar top. Anchoring by the bottom needs no height: measuring the menu and
 * moving it feeds Univer's own over-viewport max-height clamp and ratchets the
 * menu shorter.
 */
export function pinnedMenuBottom(windowHeight: number, barTop: number): number {
  return Math.max(0, Math.round(windowHeight - barTop))
}

type ExecuteCommand = (id: string, params?: unknown, options?: unknown) => Promise<boolean>

export function installSheetTabContextMenu(
  runtime: UniverRuntime,
  doc: Document = document,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const commandService = injector.get(ICommandService) as unknown as {
    executeCommand: ExecuteCommand
  }
  const instances = injector.get(IUniverInstanceService)
  const workbook = () =>
    instances.getCurrentUnitOfType<Workbook>(UniverInstanceType.UNIVER_SHEET) ?? null
  const activeSheetId = () => workbook()?.getActiveSheet()?.getSheetId() ?? null

  /** sheet whose tab opened the menu that is currently showing (null: none or the active one) */
  let menuTarget: string | null = null
  /** true only while Univer's tab-list contextmenu handler runs */
  let inTabContextMenu = false
  let resetTimer: ReturnType<typeof setTimeout> | null = null

  const original = commandService.executeCommand
  commandService.executeCommand = function (this: unknown, id, params, options) {
    if (id === SET_ACTIVE_COMMAND) {
      if (inTabContextMenu) return Promise.resolve(true)
      // any real sheet switch (left click, keyboard) ends the tab menu
      menuTarget = null
      return original.call(this, id, params, options)
    }
    const plan = planTabMenuCommand(menuTarget, activeSheetId(), id, params)
    if (plan.kind === 'passthrough') return original.call(this, id, params, options)
    const target = menuTarget as string
    menuTarget = null
    if (plan.kind === 'retarget') return original.call(this, id, plan.params, options)
    const unitId = workbook()?.getUnitId()
    return original
      .call(this, SET_ACTIVE_COMMAND, { unitId, subUnitId: target })
      .then(
        // let the sheet bar re-render so its active tab (the rename editor's
        // host) is the clicked one before the command looks for it
        () => new Promise<void>((resolve) => setTimeout(resolve, 80)),
      )
      .then(() => original.call(this, id, plan.params, options))
  }

  const onPointerDown = (event: PointerEvent) => {
    const el = event.target instanceof Element ? event.target : null
    if (!el) return
    const tab = el.closest<HTMLElement>(TAB_ITEM_SELECTOR)
    if (tab) {
      if (shouldSwallowTabPointerDown(event.button, tab.getAttribute('aria-selected') === 'true')) {
        event.stopPropagation()
      }
      return
    }
    // a click outside the tab menu ends it (Escape/outside close have no hook)
    if (!el.closest(MENU_SURFACE_SELECTOR)) menuTarget = null
  }

  const onContextMenu = (event: MouseEvent) => {
    const el = event.target instanceof Element ? event.target : null
    const tab = el?.closest<HTMLElement>(TAB_ITEM_SELECTOR)
    if (!tab || !tab.closest(TAB_LIST_SELECTOR)) return
    const id = tab.dataset.id ?? null
    menuTarget = id && id !== activeSheetId() ? id : null
    inTabContextMenu = true
    if (resetTimer) clearTimeout(resetTimer)
    // React's handler runs later in this same dispatch; reset once it is done
    resetTimer = setTimeout(() => {
      inTabContextMenu = false
      resetTimer = null
    }, 0)
    pinMenuToBar()
  }

  let pinFrame = 0
  /** after a tab right-click, keep the menu glued to the bar while it is open */
  const pinMenuToBar = () => {
    if (pinFrame) cancelAnimationFrame(pinFrame)
    let seen = false
    let misses = 0
    const step = () => {
      pinFrame = 0
      const bar = doc.querySelector<HTMLElement>(TAB_LIST_SELECTOR)
      const menu = [...doc.querySelectorAll<HTMLElement>('.univer-popup')].find(
        (el) =>
          !el.matches(MENU_SUBMENU_SELECTOR) &&
          !el.closest(MENU_SUBMENU_SELECTOR) &&
          el.getBoundingClientRect().top > -1000 &&
          getComputedStyle(el).display !== 'none',
      )
      if (bar && menu) {
        seen = true
        misses = 0
        const barTop = bar.getBoundingClientRect().top
        const bottom = `${pinnedMenuBottom(doc.defaultView?.innerHeight ?? 0, barTop)}px`
        if (menu.style.bottom !== bottom) menu.style.setProperty('bottom', bottom, 'important')
        if (menu.style.top !== 'auto') menu.style.setProperty('top', 'auto', 'important')
        // Univer caps the panel at (root container height - top), a number that
        // is unrelated to the window and shrinks a menu that sits low; the
        // stylesheet lets it grow up to the bar top instead (styles.css)
        menu.setAttribute('data-sheet-tab-menu', '')
        menu.style.setProperty('--sheet-tab-menu-max', `${Math.max(120, barTop - 8)}px`)
      } else if (++misses > (seen ? 5 : 40)) {
        return
      }
      pinFrame = requestAnimationFrame(step)
    }
    pinFrame = requestAnimationFrame(step)
  }

  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key === 'Escape') menuTarget = null
  }

  doc.addEventListener('pointerdown', onPointerDown, true)
  doc.addEventListener('contextmenu', onContextMenu, true)
  doc.addEventListener('keydown', onKeyDown, true)
  return {
    dispose() {
      doc.removeEventListener('pointerdown', onPointerDown, true)
      doc.removeEventListener('contextmenu', onContextMenu, true)
      doc.removeEventListener('keydown', onKeyDown, true)
      if (resetTimer) clearTimeout(resetTimer)
      if (pinFrame) cancelAnimationFrame(pinFrame)
      commandService.executeCommand = original
    },
  }
}
