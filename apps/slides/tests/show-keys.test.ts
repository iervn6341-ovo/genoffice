import { describe, it, expect } from 'vitest'
import { showKeyCommand, toggleShowScreen } from '../src/renderer/slideshow-utils'

const key = (k: string, buffer = '', mods = {}) => showKeyCommand(k, buffer, mods)

describe('slide show keys (PowerPoint)', () => {
  it('next: N, Enter, Space, →, ↓, PageDown', () => {
    for (const k of ['n', 'N', 'Enter', ' ', 'ArrowRight', 'ArrowDown', 'PageDown']) {
      expect(key(k).command, k).toBe('next')
    }
  })

  it('previous: P, Backspace, ←, ↑, PageUp', () => {
    for (const k of ['p', 'P', 'Backspace', 'ArrowLeft', 'ArrowUp', 'PageUp']) {
      expect(key(k).command, k).toBe('prev')
    }
  })

  it('Home / End / Esc', () => {
    expect(key('Home').command).toBe('first')
    expect(key('End').command).toBe('last')
    expect(key('Escape').command).toBe('exit')
  })

  it('B and . black out, W and , whiten', () => {
    for (const k of ['b', 'B', '.']) expect(key(k).command, k).toBe('black')
    for (const k of ['w', 'W', ',']) expect(key(k).command, k).toBe('white')
  })

  it('digits then Enter go to that slide; Enter alone still means next', () => {
    let r = key('1')
    expect(r).toMatchObject({ buffer: '1', command: null, handled: true })
    r = key('2', r.buffer)
    expect(r.buffer).toBe('12')
    r = key('Enter', r.buffer)
    expect(r).toMatchObject({ goto: 12, command: null, buffer: '', handled: true })
    expect(key('Enter').command).toBe('next')
  })

  it('a pending number is edited by Backspace and cancelled by Esc instead of turning pages or ending the show', () => {
    expect(key('Backspace', '12')).toMatchObject({ buffer: '1', command: null, handled: true })
    expect(key('Escape', '12')).toMatchObject({ buffer: '', command: null, handled: true })
    expect(key('Escape').command).toBe('exit')
  })

  it('keeps at most four digits', () => {
    expect(key('5', '1234').buffer).toBe('2345')
  })

  it('navigating clears a pending number', () => {
    expect(key('ArrowRight', '12')).toMatchObject({ command: 'next', buffer: '' })
  })

  it('leaves ⌘ / Ctrl / Alt combos and unrelated keys alone', () => {
    expect(key('n', '', { meta: true }).handled).toBe(false)
    expect(key('p', '', { ctrl: true }).handled).toBe(false)
    expect(key('z').handled).toBe(false)
    expect(key('Tab').handled).toBe(false)
  })

  it('a screen cover toggles off with the same key and switches with the other', () => {
    expect(toggleShowScreen('none', 'black')).toBe('black')
    expect(toggleShowScreen('black', 'black')).toBe('none')
    expect(toggleShowScreen('black', 'white')).toBe('white')
  })
})
