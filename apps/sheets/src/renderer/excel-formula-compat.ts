/**
 * Excel-parity fixes for Univer formula behaviour found by comparing 447
 * formulas against real Excel. Each fix wraps the engine's own executor (or
 * node factory) and only takes over the cases Excel defines differently;
 * everything else — arrays, errors, odd argument types — is delegated.
 *
 *  - numeric literals written `1.0` / `1E-7` were kept as text, so
 *    `=3>=2.0` and `ISNUMBER(1.0)` were FALSE
 *  - number → text (`&`, CONCAT) used JS digits instead of Excel's 15
 *  - IF("text",…) is #VALUE!, SEARCH wildcards, EDATE month-end clamp,
 *    HOUR/MINUTE/SECOND round to the second, (-8)^(1/3) = -2,
 *    TEXT(TRUE,…) = "TRUE", DOLLAR has no trailing pad, MATCH(x,r,-1)
 *    binary-searches like Excel
 */
import {
  ArrayValueObject,
  BaseFunction,
  ErrorType,
  ErrorValueObject,
  IFunctionService,
  LexerTreeBuilder,
  NumberValueObject,
  StringValueObject,
  ValueNodeFactory,
  type BaseValueObject,
} from '@univerjs/engine-formula'
import type { Univer } from '@univerjs/core'

type Value = BaseValueObject
type Impl = (original: BaseFunction, self: BaseFunction, args: Value[]) => Value | null

/// Excel's number → text: 15 significant digits, no exponent for 1e-9 ≤ |x| < 1e15.
export function excelNumberText(value: number): string {
  if (!Number.isFinite(value)) return String(value)
  if (value === 0) return '0'
  const [mantissa, exponentText] = Math.abs(value).toExponential(14).split('e')
  const exponent = Number(exponentText)
  const digits = (mantissa ?? '').replace('.', '').replace(/0+$/, '') || '0'
  const sign = value < 0 ? '-' : ''
  if (exponent >= 15 || exponent < -9) {
    const head = digits.length > 1 ? `${digits[0]}.${digits.slice(1)}` : digits
    const abs = Math.abs(exponent)
    return `${sign}${head}E${exponent < 0 ? '-' : '+'}${abs < 10 ? '0' : ''}${abs}`
  }
  if (exponent < 0) return `${sign}0.${'0'.repeat(-exponent - 1)}${digits}`
  if (digits.length <= exponent + 1)
    return `${sign}${digits}${'0'.repeat(exponent + 1 - digits.length)}`
  return `${sign}${digits.slice(0, exponent + 1)}.${digits.slice(exponent + 1)}`
}

/// `1.0` → `1`, `1E-7` → `1e-7`: the form the engine recognises as a number.
export function normalizeNumberToken(token: string): string {
  const trimmed = token.trim()
  if (trimmed === '' || !/^[0-9.][0-9.eE+-]*$/.test(trimmed)) return token
  const number = Number(trimmed)
  return Number.isFinite(number) ? String(number) : token
}

/// Excel SEARCH pattern → RegExp: `?` any char, `*` any run, `~` escapes.
export function searchPattern(find: string): RegExp {
  let source = ''
  for (let i = 0; i < find.length; i += 1) {
    const char = find[i]!
    if (char === '~' && i + 1 < find.length) {
      i += 1
      source += find[i]!.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    } else if (char === '*') source += '[\\s\\S]*'
    else if (char === '?') source += '[\\s\\S]'
    else source += char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  }
  return new RegExp(source, 'i')
}

const OPERAND_END = /[A-Za-z0-9_.$)\]}"%']/

function skipQuoted(text: string, at: number, quote: string): number {
  let k = at + 1
  while (k < text.length) {
    if (text[k] === quote) {
      if (text[k + 1] === quote) k += 2
      else return k + 1
    } else k += 1
  }
  return k
}

function matchClose(text: string, at: number): number {
  const open = text[at]!
  const close = open === '(' ? ')' : '}'
  let depth = 0
  let k = at
  while (k < text.length) {
    const c = text[k]!
    if (c === '"' || c === "'") {
      k = skipQuoted(text, k, c)
      continue
    }
    if (c === open) depth += 1
    else if (c === close && (depth -= 1) === 0) return k + 1
    k += 1
  }
  return k
}

/// End of the primary operand (literal, reference, call, parenthesised) at `at`.
function primaryEnd(text: string, at: number): number {
  let k = at
  const first = text[k]
  if (first === '"') k = skipQuoted(text, k, '"')
  else if (first === '(' || first === '{') k = matchClose(text, k)
  else {
    while (k < text.length) {
      const c = text[k]!
      if (c === "'") k = skipQuoted(text, k, "'")
      else if (c === '(') k = matchClose(text, k)
      else if (/[+\-*/^&=<>,;)%{}\s]/.test(c)) {
        if ((c === '+' || c === '-') && /^[0-9.]+[eE]$/.test(text.slice(at, k))) k += 1
        else break
      } else k += 1
    }
  }
  while (text[k] === '%') k += 1
  return k
}

/// Univer's lexer mis-parses stacked unary minus (`1+--1` → 0, `2*--3` →
/// #VALUE!) and cannot negate a string/boolean literal (`-"3"`, `--TRUE` →
/// #NAME?). `--x` is Excel's number coercion, so spell it `(0+x)`; an odd run
/// keeps one leading minus. Evaluation path only — never the editor's text.
export function rewriteUnaryMinus(formula: string): string {
  if (!/--|-\s*"|-\s*(?:TRUE|FALSE)\b/i.test(formula)) return formula
  let out = ''
  let prev = ''
  let at = 0
  while (at < formula.length) {
    const ch = formula[at]!
    if (ch === '"' || ch === "'") {
      const end = skipQuoted(formula, at, ch)
      out += formula.slice(at, end)
      prev = ch
      at = end
      continue
    }
    if (ch === '-' && !OPERAND_END.test(prev)) {
      let run = 0
      let k = at
      while (k < formula.length && (formula[k] === '-' || formula[k] === ' ')) {
        if (formula[k] === '-') run += 1
        k += 1
      }
      const end = primaryEnd(formula, k)
      const operand = formula.slice(k, end)
      const literal = /^"/.test(operand) || /^(?:TRUE|FALSE)$/i.test(operand)
      if (end > k && (run >= 2 || literal)) {
        out += `${run % 2 === 1 ? '-' : ''}(0+${rewriteUnaryMinus(operand)})`
        prev = ')'
        at = end
        continue
      }
    }
    out += ch
    if (ch !== ' ') prev = ch
    at += 1
  }
  return out
}

const EXCEL_EPOCH = Date.UTC(1899, 11, 30)
const DAY_MS = 86_400_000

/// EDATE over serials ≥ 61 (after Excel's phantom 1900-02-29), month-end clamped.
export function addMonths(serial: number, months: number): number | null {
  const start = Math.trunc(serial)
  if (start < 61) return null
  const date = new Date(EXCEL_EPOCH + start * DAY_MS)
  const total = date.getUTCFullYear() * 12 + date.getUTCMonth() + Math.trunc(months)
  const year = Math.floor(total / 12)
  const month = total - year * 12
  const last = new Date(Date.UTC(year, month + 1, 0)).getUTCDate()
  const result = Date.UTC(year, month, Math.min(date.getUTCDate(), last))
  const out = Math.round((result - EXCEL_EPOCH) / DAY_MS)
  return out < 61 ? null : out
}

export function timePart(serial: number, part: 'h' | 'm' | 's'): number {
  const seconds = Math.round((serial - Math.floor(serial)) * 86400) % 86400
  if (part === 'h') return Math.floor(seconds / 3600)
  if (part === 'm') return Math.floor(seconds / 60) % 60
  return seconds % 60
}

/// Excel's POWER for a negative base: defined only for exponents 1/odd.
export function negativeBasePower(base: number, exponent: number): number | null {
  const inverse = 1 / exponent
  const rounded = Math.round(inverse)
  if (Math.abs(inverse - rounded) > 1e-9 || rounded % 2 === 0) return null
  return -Math.pow(-base, exponent)
}

/// Excel's binary search for MATCH(x, range, -1) over a column/row of numbers.
export function matchDescending(lookup: number, values: number[]): number | null {
  let low = 0
  let high = values.length - 1
  let found = -1
  while (low <= high) {
    const mid = (low + high) >> 1
    const value = values[mid]!
    if (value === lookup) return mid + 1
    if (value > lookup) {
      found = mid
      low = mid + 1
    } else high = mid - 1
  }
  return found < 0 ? null : found + 1
}

/// A single-cell reference reaches executors as a 1×1 array: unwrap it.
function scalar(value: Value | undefined): Value | undefined {
  if (!value || !value.isArray()) return value
  const array = value as ArrayValueObject
  return array.getRowCount() === 1 && array.getColumnCount() === 1
    ? (array.get(0, 0) as Value)
    : undefined
}

function numberOf(value: Value | undefined): number | undefined {
  const single = scalar(value)
  return single && single.isNumber() ? Number(single.getValue()) : undefined
}

function stringOf(value: Value | undefined): string | undefined {
  const single = scalar(value)
  return single && single.isString() ? String(single.getValue()) : undefined
}

function textifyNumbers(value: Value): Value {
  if (value.isArray()) {
    return (value as ArrayValueObject).mapValue((item) => textifyNumbers(item as Value)) as Value
  }
  return value.isNumber()
    ? StringValueObject.create(excelNumberText(Number(value.getValue())))
    : value
}

const IMPLEMENTATIONS: Record<string, Impl> = {
  CONCATENATE: (original, self, args) =>
    original.calculate.apply(self, args.map(textifyNumbers) as never) as Value,
  CONCAT: (original, self, args) =>
    original.calculate.apply(self, args.map(textifyNumbers) as never) as Value,
  IF: (_original, _self, args) => {
    const test = stringOf(args[0])
    if (test === undefined) return null
    return /^(true|false)$/i.test(test.trim()) ? null : ErrorValueObject.create(ErrorType.VALUE)
  },
  SEARCH: (_original, _self, args) => {
    const find = stringOf(args[0])
    const within = stringOf(args[1])
    if (find === undefined || within === undefined || !/[*?~]/.test(find)) return null
    const startArg = args[2]
    const start = startArg === undefined ? 1 : numberOf(startArg)
    if (start === undefined || start < 1 || start > within.length + 1) {
      return ErrorValueObject.create(ErrorType.VALUE)
    }
    const hit = searchPattern(find).exec(within.slice(start - 1))
    return hit
      ? NumberValueObject.create(hit.index + start)
      : ErrorValueObject.create(ErrorType.VALUE)
  },
  EDATE: (_original, _self, args) => {
    const start = numberOf(args[0])
    const months = numberOf(args[1])
    if (start === undefined || months === undefined) return null
    const result = addMonths(start, months)
    return result === null ? null : NumberValueObject.create(result)
  },
  HOUR: (_original, _self, args) => timeResult(args, 'h'),
  MINUTE: (_original, _self, args) => timeResult(args, 'm'),
  SECOND: (_original, _self, args) => timeResult(args, 's'),
  POWER: (_original, _self, args) => {
    const base = numberOf(args[0])
    const exponent = numberOf(args[1])
    if (base === undefined || exponent === undefined) return null
    if (base >= 0 || Number.isInteger(exponent) || exponent === 0) return null
    const result = negativeBasePower(base, exponent)
    return result === null ? null : NumberValueObject.create(result)
  },
  TEXT: (_original, _self, args) => {
    const first = scalar(args[0])
    return first && first.isBoolean()
      ? StringValueObject.create(first.getValue() ? 'TRUE' : 'FALSE')
      : null
  },
  DOLLAR: (original, self, args) => {
    const result = original.calculate.apply(self, args as never) as Value
    return !result.isArray() && result.isString() && /\s$/.test(String(result.getValue()))
      ? StringValueObject.create(String(result.getValue()).trimEnd())
      : result
  },
  MATCH: (_original, _self, args) => {
    const lookup = numberOf(args[0])
    const type = numberOf(args[2])
    const range = args[1]
    if (lookup === undefined || type !== -1 || !range || !range.isArray()) return null
    const array = range as ArrayValueObject
    if (array.getRowCount() !== 1 && array.getColumnCount() !== 1) return null
    const values: number[] = []
    for (let r = 0; r < array.getRowCount(); r += 1) {
      for (let c = 0; c < array.getColumnCount(); c += 1) {
        const cell = array.get(r, c) as Value
        if (!cell.isNumber()) return null
        values.push(Number(cell.getValue()))
      }
    }
    const index = matchDescending(lookup, values)
    return index === null ? ErrorValueObject.create(ErrorType.NA) : NumberValueObject.create(index)
  },
}

function timeResult(args: Value[], part: 'h' | 'm' | 's'): Value | null {
  const serial = numberOf(args[0])
  if (serial === undefined) return null
  if (serial < 0) return ErrorValueObject.create(ErrorType.NUM)
  return NumberValueObject.create(timePart(serial, part))
}

const MARK = '__excelFormulaCompat'

function wrapExecutor(original: BaseFunction, impl: Impl): BaseFunction {
  const wrapper = Object.create(original) as BaseFunction
  wrapper.calculate = function (this: BaseFunction, ...args: Value[]) {
    const handled = impl(original, this, args)
    return handled ?? (original.calculate.apply(this, args as never) as Value)
  } as BaseFunction['calculate']
  Object.defineProperty(wrapper, MARK, { value: original })
  return wrapper
}

export function installExcelFormulaCompat(univer: Univer): { dispose(): void } {
  const injector = univer.__getInjector()
  const functionService = injector.get(IFunctionService)
  // Builtins register during plugin start-up, possibly after this runs, and a
  // later registration replaces ours: wrap whatever is current, again later.
  const ensure = () => {
    for (const [name, impl] of Object.entries(IMPLEMENTATIONS)) {
      const current = functionService.getExecutor(name)
      if (!current || Object.prototype.hasOwnProperty.call(current, MARK)) continue
      functionService.registerExecutors(wrapExecutor(current, impl))
    }
  }
  ensure()
  const timers = [0, 50, 500, 3000].map((ms) => setTimeout(ensure, ms))

  const valueFactory = injector.get(ValueNodeFactory) as unknown as {
    create(param: string): unknown
  }
  const originalCreate = valueFactory.create
  valueFactory.create = function (this: unknown, param: string) {
    return originalCreate.call(this, normalizeNumberToken(param))
  }

  const lexer = injector.get(LexerTreeBuilder) as unknown as {
    treeBuilder(formula: string, ...rest: unknown[]): unknown
  }
  const originalTreeBuilder = lexer.treeBuilder
  lexer.treeBuilder = function (this: unknown, formula: string, ...rest: unknown[]) {
    return originalTreeBuilder.call(this, rewriteUnaryMinus(formula), ...rest)
  }

  return {
    dispose() {
      timers.forEach(clearTimeout)
      lexer.treeBuilder = originalTreeBuilder
      valueFactory.create = originalCreate
      for (const name of Object.keys(IMPLEMENTATIONS)) {
        const current = functionService.getExecutor(name) as
          (BaseFunction & { [MARK]?: BaseFunction }) | undefined
        const original =
          current && Object.prototype.hasOwnProperty.call(current, MARK) ? current[MARK] : undefined
        if (original) functionService.registerExecutors(original)
      }
    },
  }
}
