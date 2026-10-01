import { describe, expect, it } from 'vitest'
import { isNotesDefaultColor, syncNotesDisplayColors } from '../src/renderer/notes-color'

describe('notes theme colors', () => {
  it('marks inherited neutral colors while preserving colored children and authored styles', () => {
    const root = document.createElement('div')
    root.innerHTML =
      '<div><span style="color: black">Default <b>bold</b><span style="color: red">red <i>italic</i></span></span><span style="color: white">White</span></div>'
    const before = [...root.querySelectorAll('*')].map((el) => el.getAttribute('style'))
    syncNotesDisplayColors(root)
    expect(root.querySelector('b')!.hasAttribute('data-notes-default-color')).toBe(true)
    expect(root.querySelector('i')!.hasAttribute('data-notes-default-color')).toBe(false)
    expect([...root.querySelectorAll('*')].map((el) => el.getAttribute('style'))).toEqual(before)
    const red = root.querySelector('i')!.parentElement!
    red.style.color = 'white'
    syncNotesDisplayColors(root)
    expect(root.querySelector('i')!.hasAttribute('data-notes-default-color')).toBe(true)
  })
  it.each([
    undefined,
    '',
    '#000',
    '#000000',
    '#FFF',
    '#FFFFFF',
    'black',
    'white',
    'rgb(0, 0, 0)',
    'rgb(255, 255, 255)',
    'rgba(0,0,0,1)',
  ])('uses the readable default for %s', (color) => {
    expect(isNotesDefaultColor(color)).toBe(true)
  })
  it.each(['#ff0000', '#123456', '#fefefe', '#010101', 'rgb(255, 0, 0)'])(
    'preserves authored color %s',
    (color) => {
      expect(isNotesDefaultColor(color)).toBe(false)
    },
  )
})
