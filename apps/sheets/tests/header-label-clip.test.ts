import { describe, expect, it } from 'vitest'
import { forceHeaderLabelClip } from '../src/renderer/header-label-clip'

describe('header label clip', () => {
  it('keeps the merged config but always asks Univer to clip the label', () => {
    const ext = {
      getCfgOfCurrentRow: () => [{ text: '21' }, false] as [unknown, boolean],
    }
    expect(forceHeaderLabelClip(ext, 'getCfgOfCurrentRow')).toBe(true)
    expect(ext.getCfgOfCurrentRow()).toEqual([{ text: '21' }, true])
  })

  it('patches columns too and is idempotent', () => {
    const ext = {
      getCfgOfCurrentColumn: () => [{ text: 'B' }, false] as [unknown, boolean],
    }
    expect(forceHeaderLabelClip(ext, 'getCfgOfCurrentColumn')).toBe(true)
    expect(forceHeaderLabelClip(ext, 'getCfgOfCurrentColumn')).toBe(false)
    expect(ext.getCfgOfCurrentColumn()[1]).toBe(true)
  })

  it('ignores an extension without the method', () => {
    expect(forceHeaderLabelClip({}, 'getCfgOfCurrentRow')).toBe(false)
  })
})
