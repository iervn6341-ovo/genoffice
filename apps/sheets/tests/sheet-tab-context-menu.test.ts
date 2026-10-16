/**
 * Right-clicking a non-active sheet tab must not switch sheets, and the menu's
 * commands must act on the clicked sheet. planTabMenuCommand decides how a
 * menu command (always sent as { value, subUnitId: <active sheet> }) is aimed.
 */
import { describe, expect, it } from 'vitest'
import {
  pinnedMenuBottom,
  planTabMenuCommand,
  shouldSwallowTabPointerDown,
} from '../src/renderer/sheet-tab-context-menu'

describe('shouldSwallowTabPointerDown', () => {
  it('swallows only the right button on a non-selected tab', () => {
    expect(shouldSwallowTabPointerDown(2, false)).toBe(true)
    expect(shouldSwallowTabPointerDown(2, true)).toBe(false)
    expect(shouldSwallowTabPointerDown(0, false)).toBe(false)
  })
})

describe('planTabMenuCommand', () => {
  const params = { value: undefined, subUnitId: 'active' }
  it('retargets id-based commands to the clicked sheet', () => {
    for (const id of [
      'sheet.command.remove-sheet-confirm',
      'sheet.command.copy-sheet',
      'sheet.command.set-tab-color',
      'sheet.command.set-worksheet-hidden',
    ]) {
      expect(planTabMenuCommand('clicked', 'active', id, params)).toEqual({
        kind: 'retarget',
        params: { value: undefined, subUnitId: 'clicked' },
      })
    }
  })
  it('activates the clicked sheet first for rename and protection', () => {
    expect(
      planTabMenuCommand('clicked', 'active', 'sheet.operation.rename-sheet', params).kind,
    ).toBe('activate-first')
  })
  it('leaves everything else alone', () => {
    expect(planTabMenuCommand(null, 'active', 'sheet.command.copy-sheet', params).kind).toBe(
      'passthrough',
    )
    expect(planTabMenuCommand('active', 'active', 'sheet.command.copy-sheet', params).kind).toBe(
      'passthrough',
    )
    expect(
      planTabMenuCommand('clicked', 'active', 'sheet.command.copy-sheet', {
        subUnitId: 'other',
        value: 1,
      }).kind,
    ).toBe('passthrough')
    expect(
      planTabMenuCommand('clicked', 'active', 'sheet.command.copy-sheet', { subUnitId: 'active' })
        .kind,
    ).toBe('passthrough')
    expect(
      planTabMenuCommand('clicked', 'active', 'ui-sheet.command.show-menu-list', params).kind,
    ).toBe('passthrough')
  })
})

describe('pinnedMenuBottom', () => {
  it('is the distance from the window bottom to the bar top', () => {
    expect(pinnedMenuBottom(760, 700)).toBe(60)
    expect(pinnedMenuBottom(660, 600)).toBe(60)
    expect(pinnedMenuBottom(100, 300)).toBe(0)
  })
})
