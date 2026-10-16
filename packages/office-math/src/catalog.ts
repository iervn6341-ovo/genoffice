/**
 * Equation structure gallery shared by Docs, Sheets and Slides.
 *
 * Eleven categories mirror Word for Mac's Equation tab (Fraction, Script, Radical, Integral, Large
 * Operator, Bracket, Function, Accent, Limit and Log, Operator, Matrix). Every template is a LaTeX
 * source that `latexToOmml` writes as native OMML and `ommlToLatex` reopens; `{{selection}}` is the
 * slot filled with the current selection (or `x`). Template ids are stable (`<category><n>`).
 */

export interface EquationTemplate {
  id: string
  latex: string
}

export interface EquationStructure {
  id: string
  labelKey: string
  templates: readonly EquationTemplate[]
}

const SLOT = '{{{selection}}}'

function category(id: string, latex: readonly string[]): EquationStructure {
  const seen = new Set<string>()
  // templates wrap the slot in their own braces; collapse accidental double wrapping
  const normalized = latex.map((source) =>
    source.split('{{{{selection}}}}').join('{{{selection}}}'),
  )
  const templates = normalized.filter((source) => !seen.has(source) && seen.add(source))
  return {
    id,
    labelKey: id,
    templates: templates.map((source, index) => ({ id: `${id}${index + 1}`, latex: source })),
  }
}

const fraction = category('fraction', [
  `\\frac{${SLOT}}{y}`,
  `\\nicefrac{${SLOT}}{y}`,
  `\\linfrac{${SLOT}}{y}`,
  `\\genfrac{}{}{0pt}{}{${SLOT}}{y}`,
  '\\frac{dy}{dx}',
  '\\frac{\\Delta y}{\\Delta x}',
  '\\frac{\\partial y}{\\partial x}',
  '\\frac{\\delta y}{\\delta x}',
  '\\frac{\\pi}{2}',
])

const script = category('script', [
  `{${SLOT}}^{n}`,
  `{${SLOT}}_{i}`,
  `{${SLOT}}_{i}^{n}`,
  `{}_{a}^{b}{${SLOT}}`,
  '{e}^{-i\\omega t}',
  '{x}^{2}',
  '{}_{n}^{k}{C}',
  '{y}^{{x}^{2}}',
])

const radical = category('radical', [
  `\\sqrt{${SLOT}}`,
  `\\sqrt[n]{${SLOT}}`,
  `\\sqrt[2]{${SLOT}}`,
  `\\sqrt[3]{${SLOT}}`,
  '\\frac{-b\\pm \\sqrt{{b}^{2}-4ac}}{2a}',
  '\\sqrt{{a}^{2}+{b}^{2}}',
])

/** one integral glyph in the three limit placements Word offers: side limits, over/under limits, none */
function integralForms(command: string): string[] {
  return [
    `\\${command}_{a}^{b} ${SLOT}`,
    `\\${command}\\limits_{a}^{b} ${SLOT}`,
    `\\${command} ${SLOT}`,
  ]
}

const integral = category('integral', [
  ...integralForms('int'),
  ...integralForms('iint'),
  ...integralForms('iiint'),
  ...integralForms('oint'),
  ...integralForms('oiint'),
  ...integralForms('oiiint'),
  `\\int ${SLOT} \\,dx`,
  '\\int \\sin \\theta \\,d\\theta',
  '\\int {\\cos 2\\theta} \\,d\\theta',
])

/** big operators: both limits above/below, a lower limit only, side limits, and none */
function operatorForms(command: string): string[] {
  return [
    `\\${command}_{i=1}^{n} ${SLOT}`,
    `\\${command}_{i} ${SLOT}`,
    `\\${command}\\nolimits_{i=1}^{n} ${SLOT}`,
    `\\${command} ${SLOT}`,
  ]
}

const largeOperator = category('largeOperator', [
  ...operatorForms('sum'),
  ...operatorForms('prod'),
  ...operatorForms('coprod'),
  ...operatorForms('bigcup'),
  ...operatorForms('bigcap'),
  ...operatorForms('bigvee'),
  ...operatorForms('bigwedge'),
  ...operatorForms('bigoplus'),
  '\\sum_{k=1}^{n} {a}_{k}',
  '\\sum_{i=0}^{n} \\binom{n}{i} {x}^{i}',
  '\\prod_{k=1}^{n} {x}_{k}',
  ...operatorForms('bigotimes').slice(0, 2),
  `\\bigodot_{i=1}^{n} ${SLOT}`,
  `\\biguplus_{i=1}^{n} ${SLOT}`,
  `\\bigsqcup_{i=1}^{n} ${SLOT}`,
])

/** paired delimiters Word lists in its Bracket gallery */
const BRACKET_PAIRS: Array<[string, string]> = [
  ['(', ')'],
  ['[', ']'],
  ['\\{', '\\}'],
  ['\\langle', '\\rangle'],
  ['|', '|'],
  ['\\|', '\\|'],
  ['\\lfloor', '\\rfloor'],
  ['\\lceil', '\\rceil'],
  ['\\llbracket', '\\rrbracket'],
  ['\\ulcorner', '\\urcorner'],
  ['\\Langle', '\\Rangle'],
]

const delimiter = category('delimiter', [
  ...BRACKET_PAIRS.map(([open, close]) => `\\left${open} ${SLOT} \\right${close}`),
  // one-sided brackets
  ...['(', '[', '\\{', '\\langle', '|', '\\|'].map((open) => `\\left${open} ${SLOT} \\right.`),
  ...[')', ']', '\\}', '\\rangle', '|'].map((close) => `\\left. ${SLOT} \\right${close}`),
  // separators between several items
  `\\left( ${SLOT} \\middle| y \\right)`,
  `\\left\\langle ${SLOT} \\middle| y \\right\\rangle`,
  `\\left\\{ ${SLOT} \\middle| y \\right\\}`,
  `\\left[ ${SLOT} \\middle| y \\middle| z \\right]`,
  `\\left( ${SLOT} \\middle| y \\middle| z \\right)`,
  `\\left\\langle ${SLOT} \\middle| y \\middle| z \\right\\rangle`,
  `\\left. ${SLOT} \\right|_{a}^{b}`,
  `{\\left( ${SLOT} \\right)}^{n}`,
  '\\left( \\frac{a}{b} \\right)',
  // case analysis and stacks
  `\\begin{cases} ${SLOT} & y \\\\ z & w \\end{cases}`,
  `\\begin{cases} ${SLOT} & y \\\\ z & w \\\\ u & v \\end{cases}`,
  `\\binom{${SLOT}}{k}`,
  `\\left\\{ \\begin{aligned} ${SLOT} \\\\ y \\end{aligned} \\right.`,
  // common formulas
  '{\\left( a+b \\right)}^{2}',
  '\\left| x \\right|',
  '\\left\\| v \\right\\|',
  '\\left\\lfloor x \\right\\rfloor',
  '\\left\\lceil x \\right\\rceil',
  '\\binom{n}{k}',
])

const TRIG = ['sin', 'cos', 'tan', 'csc', 'sec', 'cot']
const HYPERBOLIC = ['sinh', 'cosh', 'tanh', 'csch', 'sech', 'coth']
const fn = (name: string) => (name === 'csch' || name === 'sech' ? `\\${name}` : `\\${name}`)

const func = category('function', [
  ...TRIG.map((name) => `${fn(name)} {${SLOT}}`),
  ...TRIG.map((name) => `${fn(name)}^{-1} {${SLOT}}`),
  ...HYPERBOLIC.map((name) => `${fn(name)} {${SLOT}}`),
  ...HYPERBOLIC.map((name) => `${fn(name)}^{-1} {${SLOT}}`),
  '\\sin {\\theta}',
  '\\cos {2x}',
  '\\tan {\\theta} = \\frac{\\sin \\theta}{\\cos \\theta}',
])

const accent = category('accent', [
  `\\dot{${SLOT}}`,
  `\\ddot{${SLOT}}`,
  `\\dddot{${SLOT}}`,
  `\\hat{${SLOT}}`,
  `\\check{${SLOT}}`,
  `\\acute{${SLOT}}`,
  `\\grave{${SLOT}}`,
  `\\breve{${SLOT}}`,
  `\\tilde{${SLOT}}`,
  `\\bar{${SLOT}}`,
  `\\vec{${SLOT}}`,
  `\\overleftarrow{${SLOT}}`,
  `\\overleftrightarrow{${SLOT}}`,
  `\\mathring{${SLOT}}`,
  `\\overline{${SLOT}}`,
  `\\underline{${SLOT}}`,
  `\\overbrace{${SLOT}}`,
  `\\underbrace{${SLOT}}`,
  `\\overparen{${SLOT}}`,
  `\\underparen{${SLOT}}`,
  `\\overbracket{${SLOT}}`,
  `\\underbracket{${SLOT}}`,
  `\\boxed{${SLOT}}`,
  `\\cancel{${SLOT}}`,
  `\\bcancel{${SLOT}}`,
  `\\xcancel{${SLOT}}`,
  '\\overline{A \\cup B}',
])

const limitLog = category('limitLog', [
  `\\lim_{x\\to 0} {${SLOT}}`,
  `\\lim_{n\\to\\infty} {${SLOT}}`,
  `\\max_{x} {${SLOT}}`,
  `\\min_{x} {${SLOT}}`,
  `\\log {${SLOT}}`,
  `\\log_{b} {${SLOT}}`,
  `\\ln {${SLOT}}`,
  '\\lim_{n\\to\\infty} {{(1+\\frac{1}{n})}^{n}}',
])

const operator = category('operator', [
  `{${SLOT}} \\pm y`,
  `{${SLOT}} \\times y`,
  `{${SLOT}} \\div y`,
  `{${SLOT}} \\cdot y`,
  `{${SLOT}} := y`,
  `{${SLOT}} == y`,
  `{${SLOT}} += y`,
  `{${SLOT}} -= y`,
  `\\overset{def}{${SLOT}}`,
  `\\overset{\\Delta}{${SLOT}}`,
  `\\overset{m}{${SLOT}}`,
  `\\underset{n\\to\\infty}{${SLOT}}`,
  `\\xrightarrow{${SLOT}}`,
  `\\xleftarrow{${SLOT}}`,
  `\\xleftrightarrow{${SLOT}}`,
  `\\xRightarrow{${SLOT}}`,
  `\\xLeftarrow{${SLOT}}`,
  `\\xLeftrightarrow{${SLOT}}`,
  `\\overset{${SLOT}}{\\to}`,
  `\\underset{${SLOT}}{\\to}`,
  `\\stackrel{x}{=}`,
])

const MATRIX_ENVS = ['matrix', 'pmatrix', 'bmatrix', 'Bmatrix', 'vmatrix', 'Vmatrix']
const matrix = category('matrix', [
  ...MATRIX_ENVS.map((env) => `\\begin{${env}} ${SLOT} & b \\\\ c & d \\end{${env}}`),
  ...MATRIX_ENVS.map(
    (env) => `\\begin{${env}} ${SLOT} & b & c \\\\ d & e & f \\\\ g & h & i \\end{${env}}`,
  ),
  '\\begin{pmatrix} 1 & 0 \\\\ 0 & 1 \\end{pmatrix}',
  '\\begin{pmatrix} 1 & 0 & 0 \\\\ 0 & 1 & 0 \\\\ 0 & 0 & 1 \\end{pmatrix}',
  `\\begin{matrix} ${SLOT} & \\cdots & c \\end{matrix}`,
  `\\begin{matrix} ${SLOT} \\\\ \\vdots \\\\ c \\end{matrix}`,
  `\\begin{matrix} ${SLOT} & b & c \\\\ d & \\ddots & f \\\\ g & h & i \\end{matrix}`,
  `\\begin{pmatrix} ${SLOT} & \\cdots & b \\\\ \\vdots & \\ddots & \\vdots \\\\ c & \\cdots & d \\end{pmatrix}`,
  `\\begin{bmatrix} ${SLOT} & 0 \\\\ 0 & d \\end{bmatrix}`,
  `\\begin{pmatrix} ${SLOT} & b & c \\end{pmatrix}`,
  `\\begin{pmatrix} ${SLOT} \\\\ b \\\\ c \\end{pmatrix}`,
  `\\begin{pmatrix} ${SLOT} & b \\end{pmatrix}`,
])

export const EQUATION_STRUCTURES: readonly EquationStructure[] = [
  fraction,
  script,
  radical,
  integral,
  largeOperator,
  delimiter,
  func,
  accent,
  limitLog,
  operator,
  matrix,
]
