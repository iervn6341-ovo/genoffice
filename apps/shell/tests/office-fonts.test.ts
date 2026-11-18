import { describe, expect, it, vi } from 'vitest'
import {
  getUserFontsDirectory,
  handleAcquireOfficeFont,
  handleOpenFontsFolder,
} from '../src/main/office-fonts'

describe('office-fonts acquisition', () => {
  it('resolves the user fonts directory for the current platform', () => {
    const dir = getUserFontsDirectory()
    expect(typeof dir).toBe('string')
    expect(dir.length).toBeGreaterThan(0)
    if (process.platform === 'darwin') {
      expect(dir).toContain('/Library/Fonts')
    }
  })

  it('rejects empty font family names', async () => {
    const result = await handleAcquireOfficeFont('   ')
    expect(result.success).toBe(false)
    expect(result.reason).toBe('error')
    expect(result.message).toContain('empty')
  })

  it('handles non-existent font families gracefully', async () => {
    const result = await handleAcquireOfficeFont('NonExistentCrazyFont98765')
    expect(result.success).toBe(false)
    expect(['no-office', 'not-found']).toContain(result.reason)
  })

  it('successfully acquires an Office font if local Office installation is present', async () => {
    // Test with Calibri or Microsoft JhengHei
    const result = await handleAcquireOfficeFont('Calibri')
    if (result.success) {
      expect(result.copiedCount).toBeGreaterThan(0)
      expect(result.fontName).toBe('Calibri')
    } else {
      // If Office is not installed in the CI / test runner environment, reason must be 'no-office'
      expect(['no-office', 'not-found']).toContain(result.reason)
    }
  })

  it('can invoke handleOpenFontsFolder without error', async () => {
    await expect(handleOpenFontsFolder()).resolves.not.toThrow()
  })
})
