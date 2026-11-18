/**
 * Transition and animation actions extracted from App.tsx.
 * Functions take the ActionCtx built fresh per call.
 */
import type { ShapeRenderNode } from '@genoffice/pptx-render'
import type { AnimEffectKind, AnimTrigger, AnimationItem, TransitionKind } from '../shared/ipc'
import type { ActionCtx } from './action-context'
import { t } from './i18n/locale'
import { showToast } from './components/toast-bus'

function sameDocument(ctx: ActionCtx): boolean {
  return !ctx.getLatestContext || ctx.getLatestContext().path === ctx.path
}

function displayAnimations(ctx: ActionCtx, items: AnimationItem[]): void {
  const active = ctx.getLatestContext?.() ?? ctx
  if (active.path === ctx.path && active.current === ctx.current) active.setAnimations(items)
}

function animationFailure(ctx: ActionCtx, error: unknown): void {
  if (!sameDocument(ctx)) return
  const message = t('paneAnimSaveFailed', {
    error: error instanceof Error ? error.message : String(error),
  })
  ctx.setStatus(message)
  showToast(message, 'error')
}

export async function applyTransition(
  ctx: ActionCtx,
  kind: TransitionKind,
  allSlides: boolean,
): Promise<void> {
  const ok = await window.slidesApi.setTransition({
    slideIndex: allSlides ? -1 : ctx.current,
    kind,
  })
  if (ok) {
    ctx.setTransition(kind)
    ctx.setDirty(true)
    ctx.setStatus(allSlides ? t('appStatusTransitionAll') : t('appStatusTransitionSet'))
  }
}

/** Default durations for each PowerPoint effect. */
export function animDefaultDur(effect: AnimEffectKind): number {
  return effect === 'appear' || effect === 'disappear'
    ? 0
    : effect === 'spin' || effect === 'grow' || effect === 'bounce' || effect === 'motionPath'
      ? 2000
      : effect === 'pulse' || effect === 'teeter'
        ? 1000
        : 500
}

/** Hover preview: play the hovered effect on the selected shapes without committing it */
export function hoverPreviewAnimation(
  ctx: ActionCtx,
  effect: AnimEffectKind,
  motionPath?: string,
): void {
  if (ctx.selectedIds.length === 0) return
  const items: AnimationItem[] = ctx.selectedIds.map((id) => ({
    sourceId: id,
    targetName: '',
    effect,
    trigger: 'onClick' as AnimTrigger,
    durationMs: animDefaultDur(effect),
    delayMs: 0,
    ...(motionPath ? { motionPath } : {}),
  }))
  ctx.setHoverAnim({ nonce: Date.now(), items })
}

/** Commit the whole page's animation list overwrite-style, then read back on success (target names/stale items per the main process). */
export async function commitAnimations(ctx: ActionCtx, items: AnimationItem[]): Promise<boolean> {
  const payload = {
    slideIndex: ctx.current,
    items: items.map(({ targetName: _t, ...rest }) => rest),
  }
  try {
    if (!sameDocument(ctx)) return false
    const ok = await window.slidesApi.setAnimations(payload)
    if (!ok) {
      const diagnostic = await window.slidesApi.applyTxn({
        ops: [{ op: 'setAnimations', target: { slide: ctx.current }, items: payload.items }],
        dryRun: true,
      })
      throw new Error(diagnostic?.failures?.[0]?.error ?? t('appUnknownError'))
    }
    if (sameDocument(ctx)) ctx.setDirty(true)
    displayAnimations(ctx, await window.slidesApi.getAnimations(ctx.current))
    return true
  } catch (error) {
    if (sameDocument(ctx)) {
      const actual = await window.slidesApi.getAnimations(ctx.current).catch(() => ctx.animations)
      displayAnimations(ctx, actual)
    }
    animationFailure(ctx, error)
    return false
  }
}

/** Paragraph count of a text-bearing node ("by paragraph" animation splitting). */
export function paraCountOf(ctx: ActionCtx, id: string): number {
  const node = ctx.findNodeCtx(id)?.node
  const text =
    node && (node.type === 'text' || node.type === 'shape')
      ? (node as ShapeRenderNode).text
      : undefined
  if (!text) return 0
  return text.lines.filter((l) => l.paraStart !== false).length
}

/** Gallery click: edit the selected effect, or replace a shape's effects when no effect is selected. */
export function applyAnimation(ctx: ActionCtx, effect: AnimEffectKind | 'none'): void {
  // A pane/number selection addresses one effect, including stacked effects on
  // the same shape. The gallery must not replace its siblings in that case.
  if (ctx.selAnim >= 0 && ctx.selAnim < ctx.animations.length) {
    if (effect === 'none') deleteAnimation(ctx, ctx.selAnim)
    else patchAnimation(ctx, ctx.selAnim, { effect })
    return
  }
  if (ctx.selectedIds.length === 0) return
  const sel = new Set(ctx.selectedIds)
  if (effect === 'none') {
    const kept = ctx.animations.filter((a) => !sel.has(a.sourceId))
    ctx.setSelAnim(-1)
    void commitAnimations(ctx, kept)
    ctx.setStatus(t('appStatusAnimRemoved'))
    return
  }
  const firstExistingIdx = ctx.animations.findIndex((a) => sel.has(a.sourceId))
  const added = ctx.selectedIds.flatMap((id) => {
    const existing = ctx.animations.find((a) => a.sourceId === id)
    const base = {
      sourceId: id,
      targetName: existing?.targetName || '',
      effect,
      trigger: (existing?.trigger ?? 'onClick') as AnimTrigger,
      durationMs: animDefaultDur(effect),
      delayMs: existing?.delayMs ?? 0,
      ...(existing?.after ? { after: existing.after } : {}),
      ...((['flyIn', 'flyOut', 'wipe', 'wipeOut'].includes(effect) && existing?.direction)
        ? { direction: existing.direction }
        : {}),
    }
    const n = ctx.animByParagraph ? paraCountOf(ctx, id) : 0
    return n > 1 ? Array.from({ length: n }, (_, i) => ({ ...base, paragraph: i })) : [base]
  })
  if (firstExistingIdx < 0) {
    ctx.setSelAnim(ctx.animations.length)
    void commitAnimations(ctx, [...ctx.animations, ...added])
  } else {
    const before = ctx.animations.slice(0, firstExistingIdx).filter((a) => !sel.has(a.sourceId))
    const after = ctx.animations.slice(firstExistingIdx).filter((a) => !sel.has(a.sourceId))
    ctx.setSelAnim(before.length)
    void commitAnimations(ctx, [...before, ...added, ...after])
  }
  ctx.setStatus(t('appStatusAnimSet'))
}

/** "Add animation": append at the end of the list (existing animations kept). */
export function addAnimation(ctx: ActionCtx, effect: AnimEffectKind): void {
  if (ctx.selectedIds.length === 0) return
  const added = ctx.selectedIds.flatMap((id) => {
    const base = {
      sourceId: id,
      targetName: '',
      effect,
      trigger: 'onClick' as AnimTrigger,
      durationMs: animDefaultDur(effect),
      delayMs: 0,
    }
    const n = ctx.animByParagraph ? paraCountOf(ctx, id) : 0
    return n > 1 ? Array.from({ length: n }, (_, i) => ({ ...base, paragraph: i })) : [base]
  })
  void commitAnimations(ctx, [...ctx.animations, ...added])
  ctx.setStatus(t('appStatusAnimAdded'))
}

/** "Motion path": append an animation moving the selected shape along a preset path. */
export function applyMotionPath(ctx: ActionCtx, path: string): void {
  if (ctx.selectedIds.length === 0) return
  const added = ctx.selectedIds.map((id) => ({
    sourceId: id,
    targetName: '',
    effect: 'motionPath' as AnimEffectKind,
    trigger: 'onClick' as AnimTrigger,
    durationMs: animDefaultDur('motionPath'),
    delayMs: 0,
    motionPath: path,
  }))
  void commitAnimations(ctx, [...ctx.animations, ...added])
  ctx.setStatus(t('appStatusMotionPathAdded'))
}

export function patchAnimTiming(
  ctx: ActionCtx,
  patch: {
    trigger?: AnimTrigger
    durationMs?: number
    delayMs?: number
    direction?: AnimationItem['direction']
    after?: AnimationItem['after'] | null
  },
): void {
  patchAnimation(ctx, ctx.timingIdx, patch)
}

export type AnimationEditPatch = {
  effect?: AnimEffectKind
  trigger?: AnimTrigger
  durationMs?: number
  delayMs?: number
  direction?: AnimationItem['direction']
  after?: AnimationItem['after'] | null
}

/** Change just the addressed effect; untouched imported effects retain their native IDs. */
export function editedAnimation(item: AnimationItem, patch: AnimationEditPatch): AnimationItem {
  const { after, ...rest } = patch
  const updated: AnimationItem = { ...item, ...rest }
  if (after === null) delete updated.after
  else if (after !== undefined) updated.after = after
  if (patch.effect && (patch.effect !== item.effect || item.customPreset)) {
    updated.durationMs = animDefaultDur(patch.effect)
    delete updated.customPreset
    delete updated.nativeId // deliberate replacement must not retain an approximated imported preset
    if (!['flyIn', 'flyOut', 'wipe', 'wipeOut'].includes(patch.effect)) delete updated.direction
    if (patch.effect === 'motionPath') updated.motionPath ??= 'M 0 0 L 0.25 0.25'
    else delete updated.motionPath
  }
  return updated
}

let animationEditTail: Promise<void> = Promise.resolve()

/** Save and history commands wait for the choice that was just made in the pane. */
export async function flushAnimationEdits(): Promise<void> {
  await animationEditTail
}

/** Locate the intended effect after native cTn IDs have changed during an earlier edit/save. */
function refreshedIndex(before: AnimationItem[], index: number, latest: AnimationItem[]): number {
  const target = before[index]!
  if (target.nativeId != null) {
    const native = latest.findIndex(
      (a) =>
        a.nativeId === target.nativeId &&
        (a.sourceId === target.sourceId || a.targetName === target.targetName),
    )
    if (native >= 0) return native
  }
  const sameTarget = (a: AnimationItem) =>
    a.sourceId === target.sourceId && (a.paragraph ?? null) === (target.paragraph ?? null)
  const oldSiblings = before.flatMap((a, i) => (sameTarget(a) ? [i] : []))
  const currentSiblings = latest.flatMap((a, i) => (sameTarget(a) ? [i] : []))
  // A removed effect must not turn a delayed edit into a change to its sibling.
  if (oldSiblings.length !== currentSiblings.length) return -1
  return currentSiblings[oldSiblings.indexOf(index)] ?? -1
}

export function patchAnimation(
  ctx: ActionCtx,
  index: number,
  patch: AnimationEditPatch,
): Promise<void> {
  if (index < 0 || index >= ctx.animations.length) return Promise.resolve()
  const before = ctx.animations
  const run = async (): Promise<void> => {
    if (!sameDocument(ctx)) return
    try {
      const latest = await window.slidesApi.getAnimations(ctx.current)
      if (!sameDocument(ctx)) return
      const currentIndex = refreshedIndex(before, index, latest)
      if (currentIndex < 0) {
        displayAnimations(ctx, latest)
        throw new Error(t('paneAnimSelectHint'))
      }
      const fresh = { ...ctx, animations: latest }
      const next = latest.map((a, i) => (i === currentIndex ? editedAnimation(a, patch) : a))
      // Keep the chosen value visible while the main process validates/writes it.
      displayAnimations(fresh, next)
      await commitAnimations(fresh, next)
    } catch (error) {
      animationFailure(ctx, error)
    }
  }
  const work = animationEditTail.then(run)
  animationEditTail = work.catch(() => undefined)
  return work
}

export function moveAnimation(ctx: ActionCtx, idx: number, dir: -1 | 1): void {
  const j = idx + dir
  if (idx < 0 || j < 0 || idx >= ctx.animations.length || j >= ctx.animations.length) return
  const next = [...ctx.animations]
  ;[next[idx], next[j]] = [next[j]!, next[idx]!]
  if (ctx.selAnim === idx) ctx.setSelAnim(j)
  else if (ctx.selAnim === j) ctx.setSelAnim(idx)
  void commitAnimations(ctx, next)
}

export function deleteAnimation(ctx: ActionCtx, idx: number): void {
  if (idx < 0 || idx >= ctx.animations.length) return
  if (ctx.selAnim === idx) ctx.setSelAnim(-1)
  else if (ctx.selAnim > idx) ctx.setSelAnim(ctx.selAnim - 1)
  void commitAnimations(
    ctx,
    ctx.animations.filter((_, i) => i !== idx),
  )
}
