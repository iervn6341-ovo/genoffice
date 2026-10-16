import { describe, expect, it } from 'vitest'
import { printEntryFor } from '../src/renderer/print-entry'

describe('printEntryFor', () => {
  it('hands the job to the native print panel on macOS', () => {
    expect(printEntryFor('MacIntel')).toBe('native-panel')
    expect(printEntryFor('macOS')).toBe('native-panel')
  })
  it('prints from the in-app dialog on Windows and Linux', () => {
    expect(printEntryFor('Win32')).toBe('in-app')
    expect(printEntryFor('Linux x86_64')).toBe('in-app')
  })
  it('lets tests force the in-app dialog on macOS', () => {
    expect(printEntryFor('MacIntel', true)).toBe('in-app')
  })
})
