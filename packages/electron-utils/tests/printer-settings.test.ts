import { describe, expect, it } from 'vitest'
import { directPrintOptions, listSystemPrinters } from '../src/printer-settings'

describe('directPrintOptions', () => {
  it('opens the system dialog when no printer is chosen', () => {
    expect(directPrintOptions(undefined)).toEqual({ silent: false })
    expect(directPrintOptions({ copies: 3 })).toEqual({ silent: false })
  })

  it('prints straight to the chosen printer with its settings', () => {
    expect(
      directPrintOptions({ printer: 'Office', copies: 2, duplex: 'long', collate: false }),
    ).toEqual({
      silent: true,
      deviceName: 'Office',
      copies: 2,
      duplexMode: 'longEdge',
      collate: false,
    })
    expect(directPrintOptions({ printer: 'P', duplex: 'short' })).toMatchObject({
      duplexMode: 'shortEdge',
      collate: true,
      copies: 1,
    })
    expect(directPrintOptions({ printer: 'P' })).toMatchObject({ duplexMode: 'simplex' })
  })

  it('clamps copies to 1…999', () => {
    expect(directPrintOptions({ printer: 'P', copies: 0 }).copies).toBe(1)
    expect(directPrintOptions({ printer: 'P', copies: 5000 }).copies).toBe(999)
    expect(directPrintOptions({ printer: 'P', copies: Number.NaN }).copies).toBe(1)
  })
})

describe('listSystemPrinters', () => {
  it('maps printers and survives a failing query', async () => {
    const list = await listSystemPrinters({
      getPrintersAsync: async () => [{ name: 'a', isDefault: true } as never, { name: 'b' }],
    })
    expect(list).toEqual([
      { name: 'a', displayName: 'a', isDefault: true },
      { name: 'b', displayName: 'b', isDefault: false },
    ])
    expect(
      await listSystemPrinters({
        getPrintersAsync: async () => {
          throw new Error('x')
        },
      }),
    ).toEqual([])
  })
})
