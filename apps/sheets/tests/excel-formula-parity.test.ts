/**
 * Formulas whose Excel results were captured from a real Excel (see
 * docs/qa/office-parity-recheck-2026-10-07.md) and used to differ here.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import {
  addMonths,
  excelNumberText,
  installExcelFormulaCompat,
  matchDescending,
  negativeBasePower,
  normalizeNumberToken,
  rewriteUnaryMinus,
  searchPattern,
  timePart,
} from '../src/renderer/excel-formula-compat'
import { createFormulaEngine, type FormulaEngine } from './formula-engine'

describe('pure helpers', () => {
  it('excelNumberText keeps 15 significant digits', () => {
    expect(excelNumberText(1 / 3)).toBe('0.333333333333333')
    expect(excelNumberText(0.1 + 0.2)).toBe('0.3')
    expect(excelNumberText(1e-7)).toBe('0.0000001')
    expect(excelNumberText(1e-10)).toBe('1E-10')
    expect(excelNumberText(100)).toBe('100')
    expect(excelNumberText(-12.5)).toBe('-12.5')
    expect(excelNumberText(123456789012345)).toBe('123456789012345')
    expect(excelNumberText(1e15)).toBe('1E+15')
    expect(excelNumberText(1234567890123456)).toBe('1.23456789012346E+15')
  })

  it('normalizeNumberToken only touches numeric tokens', () => {
    expect(normalizeNumberToken('1.0')).toBe('1')
    expect(normalizeNumberToken('1E-7')).toBe('1e-7')
    expect(normalizeNumberToken('"1.0"')).toBe('"1.0"')
    expect(normalizeNumberToken('A1')).toBe('A1')
    expect(normalizeNumberToken('TRUE')).toBe('TRUE')
  })

  it('searchPattern supports ? * ~', () => {
    expect(searchPattern('b*c').exec('abxxc')?.index).toBe(1)
    expect(searchPattern('~*').exec('a*b')?.index).toBe(1)
    expect(searchPattern('a?c').test('abc')).toBe(true)
  })

  it('addMonths clamps to month end', () => {
    expect(addMonths(45322, 1)).toBe(45351) // 2024-01-31 → 2024-02-29
    expect(addMonths(45382, -1)).toBe(45351) // 2024-03-31 → 2024-02-29
  })

  it('timePart rounds to the second', () => {
    expect(timePart(0.7534, 's')).toBe(54)
    expect(timePart(0.75, 'h')).toBe(18)
    expect(timePart(0.99999999, 'h')).toBe(0)
  })

  it('negativeBasePower / matchDescending', () => {
    expect(negativeBasePower(-8, 1 / 3)).toBeCloseTo(-2)
    expect(negativeBasePower(-8, 0.5)).toBeNull()
    expect(matchDescending(35, [10, 20, 30, 40, 50])).toBeNull()
    expect(matchDescending(35, [50, 40, 30, 20, 10])).toBe(2)
  })
})

describe('rewriteUnaryMinus', () => {
  it('spells stacked and literal unary minus as (0+x)', () => {
    expect(rewriteUnaryMinus('=1+--1')).toBe('=1+(0+1)')
    expect(rewriteUnaryMinus('=2*--(A1>1)')).toBe('=2*(0+(A1>1))')
    expect(rewriteUnaryMinus('=-"3"')).toBe('=-(0+"3")')
    expect(rewriteUnaryMinus('=---A1')).toBe('=-(0+A1)')
    expect(rewriteUnaryMinus('=--SUM(A1:A3)')).toBe('=(0+SUM(A1:A3))')
  })

  it('leaves everything else alone', () => {
    expect(rewriteUnaryMinus('=1--1')).toBe('=1--1')
    expect(rewriteUnaryMinus('=A1-"3"')).toBe('=A1-"3"')
    expect(rewriteUnaryMinus('="a--b"')).toBe('="a--b"')
    expect(rewriteUnaryMinus('=-A1')).toBe('=-A1')
    expect(rewriteUnaryMinus('=1E-7+2')).toBe('=1E-7+2')
  })
})

describe('engine results match Excel', () => {
  let engine: FormulaEngine
  const cases: Array<[string, unknown]> = [
    ['=1.0=1', true],
    ['=ISNUMBER(1.0)', true],
    ['=3>=2.0', true],
    ['=1.0&""', '1'],
    ['=1E-7&""', '0.0000001'],
    ['=1/3&""', '0.333333333333333'],
    ['=CONCATENATE(1/3,"x")', '0.333333333333333x'],
    ['=IF("a","y","n")', '#VALUE!'],
    ['=IF("TRUE","y","n")', 'y'],
    ['=SEARCH("b*c","abxxc")', 2],
    ['=SEARCH("~*","a*b")', 2],
    ['=SEARCH("c","abc")', 3],
    ['=SEARCH("b*c",B2)', 2],
    ['=IF(B1,"y","n")', '#VALUE!'],
    ['=1+--1', 2],
    ['=2*--3', 6],
    ['=--TRUE', 1],
    ['=--"3"', 3],
    ['=-"3"', -3],
    ['=--B3', 12],
    ['=2*--(1=1)', 2],
    ['=1--1', 2],
    ['=--"1e3"', 1000],
    ['=ISNUMBER(--"1e3")', true],
    ['=SUMPRODUCT(--(M1:M5>20))', 3],
    ['=EDATE(DATE(2024,1,31),1)', 45351],
    ['=EDATE(DATE(2024,3,31),-1)', 45351],
    ['=SECOND(0.7534)', 54],
    ['=POWER(-8,1/3)', -2],
    ['=TEXT(TRUE,"0")', 'TRUE'],
    ['=DOLLAR(1234.567,2)', '$1,234.57'],
    ['=MATCH(35,M1:M5,-1)', '#N/A'],
  ]

  beforeAll(async () => {
    engine = await createFormulaEngine((univer) => {
      installExcelFormulaCompat(univer)
    })
    const cells: Record<string, string | number> = {
      B1: 'a',
      B2: 'abxxc',
      B3: '12',
      M1: 10,
      M2: 20,
      M3: 30,
      M4: 40,
      M5: 50,
    }
    cases.forEach(([formula], index) => {
      cells[`A${index + 1}`] = formula
    })
    await engine.set(cells)
  }, 30_000)

  afterAll(() => engine?.dispose())

  cases.forEach(([formula, expected], index) => {
    it(formula, () => {
      expect(engine.value(`A${index + 1}`)).toEqual(expected)
    })
  })
})
