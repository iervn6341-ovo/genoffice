/**
 * Picture Styles: each preset is a plain composition of native picture
 * formatting ops (preset outline geometry, picture border, shadow / reflection /
 * soft-edge effects). They are named after what they write, not a copy of
 * PowerPoint's gallery XML; "none" restores the plain rectangle.
 */
export interface PictureStyle {
  id: string
  labelKey:
    | 'ribbonPicStyleNone'
    | 'ribbonPicStyleSimpleFrame'
    | 'ribbonPicStyleDropShadow'
    | 'ribbonPicStyleRounded'
    | 'ribbonPicStyleReflectedRounded'
    | 'ribbonPicStyleSoftEdge'
    | 'ribbonPicStyleOval'
    | 'ribbonPicStyleSoftEdgeOval'
    | 'ribbonPicStyleBlackFrame'
  prst: string
  stroke: { color: string; widthEmu: number } | null
  effects: {
    shadow: {
      color: string
      blurRad: number
      dist: number
      dirDeg: number
      algn?: string
    } | null
    reflection: { blurRad: number; startA: number; endPos: number; dist: number } | null
    softEdge: number | null
  }
}

const NO_EFFECTS = { shadow: null, reflection: null, softEdge: null } as const
const SOFT_SHADOW = { color: '#00000066', blurRad: 50800, dist: 38100, dirDeg: 45, algn: 'tl' }

export const PICTURE_STYLES: readonly PictureStyle[] = [
  { id: 'none', labelKey: 'ribbonPicStyleNone', prst: 'rect', stroke: null, effects: NO_EFFECTS },
  {
    id: 'simple-frame',
    labelKey: 'ribbonPicStyleSimpleFrame',
    prst: 'rect',
    stroke: { color: '#FFFFFF', widthEmu: 88900 },
    effects: {
      shadow: { color: '#00000066', blurRad: 55000, dist: 18000, dirDeg: 90, algn: 'ctr' },
      reflection: null,
      softEdge: null,
    },
  },
  {
    id: 'black-frame',
    labelKey: 'ribbonPicStyleBlackFrame',
    prst: 'rect',
    stroke: { color: '#000000', widthEmu: 38100 },
    effects: NO_EFFECTS,
  },
  {
    id: 'drop-shadow',
    labelKey: 'ribbonPicStyleDropShadow',
    prst: 'rect',
    stroke: null,
    effects: { shadow: SOFT_SHADOW, reflection: null, softEdge: null },
  },
  {
    id: 'rounded',
    labelKey: 'ribbonPicStyleRounded',
    prst: 'roundRect',
    stroke: null,
    effects: NO_EFFECTS,
  },
  {
    id: 'reflected-rounded',
    labelKey: 'ribbonPicStyleReflectedRounded',
    prst: 'roundRect',
    stroke: null,
    effects: {
      shadow: null,
      reflection: { blurRad: 6350, startA: 0.5, endPos: 0.9, dist: 0 },
      softEdge: null,
    },
  },
  {
    id: 'soft-edge',
    labelKey: 'ribbonPicStyleSoftEdge',
    prst: 'rect',
    stroke: null,
    effects: { shadow: null, reflection: null, softEdge: 63500 },
  },
  {
    id: 'oval',
    labelKey: 'ribbonPicStyleOval',
    prst: 'ellipse',
    stroke: null,
    effects: NO_EFFECTS,
  },
  {
    id: 'soft-edge-oval',
    labelKey: 'ribbonPicStyleSoftEdgeOval',
    prst: 'ellipse',
    stroke: null,
    effects: { shadow: null, reflection: null, softEdge: 63500 },
  },
]

/** Op list that applies a style to one picture (clears whatever the style does not set). */
export function pictureStyleOps(
  slide: number,
  el: string,
  style: PictureStyle,
): Array<{ op: string; [key: string]: unknown }> {
  const target = { slide, el }
  return [
    { op: 'setShapeGeometry', target, prst: style.prst },
    { op: 'setStroke', target, stroke: style.stroke },
    { op: 'setEffects', target, effects: style.effects },
  ]
}
