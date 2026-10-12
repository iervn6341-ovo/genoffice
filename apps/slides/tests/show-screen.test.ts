import { describe, expect, it } from 'vitest'

import {
  preselectedScreen,
  shouldAskScreen,
  showCandidates,
  type ShowScreen,
} from '../src/renderer/show-screen'

const screen = (id: number, label: string, current = false): ShowScreen => ({
  id,
  label,
  primary: id === 1,
  current,
  width: 1920,
  height: 1080,
})
const laptop = screen(1, 'Built-in Retina Display', true)
const asus = screen(2, 'ASUS VS229')
const side = screen(3, 'SideScreen')

describe('choosing the screen for a show', () => {
  it('offers every screen but the presenter’s', () => {
    expect(showCandidates([laptop, asus, side]).map((s) => s.label)).toEqual([
      'ASUS VS229',
      'SideScreen',
    ])
  })

  it('asks every time another screen is connected, and never with only the laptop', () => {
    expect(shouldAskScreen([laptop])).toBe(false)
    expect(shouldAskScreen([laptop, asus])).toBe(true)
    expect(shouldAskScreen([laptop, asus, side])).toBe(true)
  })

  it('starts the dialog on last time’s screen, else the first', () => {
    const candidates = showCandidates([laptop, asus, side])
    expect(preselectedScreen(candidates, 'SideScreen')?.id).toBe(3)
    expect(preselectedScreen(candidates, 'Gone')?.id).toBe(2)
    expect(preselectedScreen(candidates, null)?.id).toBe(2)
  })
})
