/**
 * Column / row border drags of a table. The renderer streams live `preview` frames through the same
 * op (the model follows the pointer, the rebuilt slide re-flows the cell text) and settles with a plain
 * commit — the whole gesture is ONE undo step, exactly like a Transformer resize (`transformPreview`).
 * `cancel` rolls the frames back and leaves no step.
 */
import { runTxn, type TxnRequest, type TxnResult } from '@genoffice/pptx-ops'
import {
  pushHistory,
  restoreSnapshot,
  scheduleDeckBroadcast,
  scheduleHistoryNotify,
  type Session,
} from './session-state'

/** The journaling transaction runner of slides-main (every settled commit lands in the op journal) */
export type JournaledTxn = (session: Session, source: 'edit', req: TxnRequest) => TxnResult

/** Apply one gesture request; false = refused (nothing changed, no history step left behind) */
export function applyTableGridTxn(
  session: Session,
  req: TxnRequest,
  op: { preview?: boolean; cancel?: boolean },
  journaledTxn: JournaledTxn,
): boolean {
  if (op.cancel) {
    if (session.transformPreview) {
      session.transformPreview = false
      const snap = session.undoStack.pop()
      if (snap) restoreSnapshot(session, snap)
      scheduleHistoryNotify(session)
      scheduleDeckBroadcast(session)
    }
    return true
  }
  // validate before the bookkeeping: a failed first frame must not set the flag or clear redo
  const plan = runTxn(session.opened, { ...req, dryRun: true })
  if (plan.failures?.length) return false
  let pushed = false
  if (op.preview) {
    if (!session.transformPreview) {
      pushHistory(session)
      pushed = true
      session.transformPreview = true
    }
  } else if (session.transformPreview) {
    session.transformPreview = false
  } else {
    pushHistory(session)
    pushed = true
  }
  const r = op.preview ? runTxn(session.opened, req) : journaledTxn(session, 'edit', req)
  if (!r.applied) {
    if (pushed) {
      session.undoStack.pop()
      if (op.preview) session.transformPreview = false
    }
    return false
  }
  return true
}
