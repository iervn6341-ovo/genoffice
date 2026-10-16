import { useEffect, useRef, useState } from 'react'

import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

import type { WorkbookSheetThread, WorkbookThreadPerson } from '../shared/desktop-api'
import { useI18n } from './i18n/locale'
import type { ThreadEditError } from './threaded-comments-actions'

/// Excel 365's Comments pane: every thread of the sheet with replies, Resolve / Reopen, edit and
/// delete, plus a composer for the selected cell. Notes (legacy comments) are a separate path.
export function CommentsPanel({
  threads,
  persons,
  activeCell,
  activeCellHasNote,
  focusThreadId,
  author,
  onAuthorChange,
  onCreate,
  onReply,
  onResolve,
  onDelete,
  onEdit,
  onGoto,
  onClose,
}: {
  readonly threads: readonly WorkbookSheetThread[]
  readonly persons: readonly WorkbookThreadPerson[]
  readonly activeCell: { readonly row: number; readonly column: number } | null
  readonly activeCellHasNote: boolean
  readonly focusThreadId: string | null
  readonly author: string
  readonly onAuthorChange: (name: string) => void
  readonly onCreate: (text: string) => Promise<ThreadEditError | null>
  readonly onReply: (threadId: string, text: string) => Promise<ThreadEditError | null>
  readonly onResolve: (threadId: string, done: boolean) => Promise<ThreadEditError | null>
  readonly onDelete: (threadId: string, commentId: string) => Promise<ThreadEditError | null>
  readonly onEdit: (
    threadId: string,
    commentId: string,
    text: string,
  ) => Promise<ThreadEditError | null>
  readonly onGoto: (thread: WorkbookSheetThread) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [draft, setDraft] = useState('')
  const [replies, setReplies] = useState<Record<string, string>>({})
  const [editing, setEditing] = useState<{ thread: string; comment: string; text: string } | null>(
    null,
  )
  const [error, setError] = useState<ThreadEditError | null>(null)
  const [noteError, setNoteError] = useState(false)
  const composerRef = useRef<HTMLTextAreaElement>(null)
  const focusRef = useRef<HTMLElement | null>(null)

  useEffect(() => composerRef.current?.focus(), [])
  useEffect(() => focusRef.current?.scrollIntoView?.({ block: 'nearest' }), [focusThreadId])

  const nameOf = (personId: string): string =>
    persons.find((person) => person.id === personId)?.displayName ?? '—'
  const cellLabel = activeCell ? formatAddress(activeCell.row, activeCell.column) : ''
  const existing = activeCell
    ? threads.find((thread) => thread.row === activeCell.row && thread.column === activeCell.column)
    : undefined
  const errorText = (code: ThreadEditError): string =>
    code === 'nofile'
      ? t('appCmtNoFile')
      : code === 'occupied'
        ? t('appCmtOccupied')
        : t('appCmtEmptyText')
  const run = async (work: Promise<ThreadEditError | null>): Promise<boolean> => {
    const failure = await work
    setError(failure)
    return failure === null
  }
  const sorted = [...threads].sort((a, b) => a.row - b.row || a.column - b.column)

  return (
    <aside className="comments-panel" role="complementary" aria-label={t('appCmtTitle')}>
      <header>
        <span className="comments-title">{t('appCmtTitle')}</span>
        <button className="comments-close" data-tip={t('appClose')} onClick={onClose}>
          ✕
        </button>
      </header>
      <label className="comments-author">
        {t('appCmtAuthor')}
        <input
          value={author}
          maxLength={80}
          aria-label={t('appCmtAuthor')}
          onChange={(event) => onAuthorChange(event.target.value)}
        />
      </label>
      <div className="comments-composer">
        <label>
          {activeCell ? t('appCmtNewOn', { cell: cellLabel }) : t('appCmtSelectCell')}
          <textarea
            ref={composerRef}
            value={draft}
            rows={3}
            placeholder={t('appCmtPlaceholder')}
            disabled={!activeCell}
            onChange={(event) => setDraft(event.target.value)}
          />
        </label>
        <button
          className="primary"
          disabled={!activeCell || draft.trim() === '' || existing !== undefined}
          onClick={() => {
            if (!activeCell) return
            if (activeCellHasNote) {
              setError(null)
              setNoteError(true)
              return
            }
            setNoteError(false)
            void run(onCreate(draft)).then((ok) => {
              if (ok) setDraft('')
            })
          }}
        >
          {t('appCmtPost')}
        </button>
      </div>
      {existing && <p className="comments-note">{t('appCmtOccupied')}</p>}
      {noteError && <p role="alert">{t('appCmtHasNote')}</p>}
      {error && <p role="alert">{errorText(error)}</p>}
      {sorted.length === 0 ? (
        <p className="comments-empty">{t('appCmtEmpty')}</p>
      ) : (
        <ul className="comments-list">
          {sorted.map((thread) => (
            <li
              key={thread.id}
              className={`comments-thread${thread.done ? ' resolved' : ''}${
                thread.id === focusThreadId ? ' focused' : ''
              }`}
              ref={(node) => {
                if (thread.id === focusThreadId) focusRef.current = node
              }}
            >
              <div className="comments-thread-head">
                <button
                  className="comments-goto"
                  data-tip={t('appCmtGoto', { cell: formatAddress(thread.row, thread.column) })}
                  onClick={() => onGoto(thread)}
                >
                  {formatAddress(thread.row, thread.column)}
                </button>
                {thread.done && <span className="comments-badge">{t('appCmtResolved')}</span>}
                <span className="comments-spacer" />
                <button onClick={() => void run(onResolve(thread.id, !thread.done))}>
                  {t(thread.done ? 'appCmtReopen' : 'appCmtResolve')}
                </button>
                <button onClick={() => void run(onDelete(thread.id, thread.comments[0]?.id ?? ''))}>
                  {t('appCmtDeleteThread')}
                </button>
              </div>
              {thread.comments.map((comment) => (
                <div className="comments-item" key={comment.id}>
                  <div className="comments-meta">
                    <strong>{nameOf(comment.personId)}</strong>
                    <time>{comment.time.replace('T', ' ').replace(/\.\d+$/, '')}</time>
                  </div>
                  {editing?.comment === comment.id ? (
                    <>
                      <textarea
                        rows={3}
                        value={editing.text}
                        aria-label={t('appCmtEdit')}
                        onChange={(event) => setEditing({ ...editing, text: event.target.value })}
                      />
                      <div className="comments-actions">
                        <button
                          className="primary"
                          disabled={editing.text.trim() === ''}
                          onClick={() =>
                            void run(onEdit(thread.id, comment.id, editing.text)).then((ok) => {
                              if (ok) setEditing(null)
                            })
                          }
                        >
                          {t('appCmtSave')}
                        </button>
                        <button onClick={() => setEditing(null)}>{t('appCancel')}</button>
                      </div>
                    </>
                  ) : (
                    <>
                      <p className="comments-text">{comment.text}</p>
                      <div className="comments-actions">
                        <button
                          onClick={() =>
                            setEditing({
                              thread: thread.id,
                              comment: comment.id,
                              text: comment.text,
                            })
                          }
                        >
                          {t('appCmtEdit')}
                        </button>
                        {comment.id !== thread.comments[0]?.id && (
                          <button onClick={() => void run(onDelete(thread.id, comment.id))}>
                            {t('appCmtDelete')}
                          </button>
                        )}
                      </div>
                    </>
                  )}
                </div>
              ))}
              <div className="comments-reply">
                <textarea
                  rows={2}
                  value={replies[thread.id] ?? ''}
                  placeholder={t('appCmtReplyPlaceholder')}
                  aria-label={t('appCmtReplyPlaceholder')}
                  onChange={(event) => setReplies({ ...replies, [thread.id]: event.target.value })}
                />
                <button
                  disabled={(replies[thread.id] ?? '').trim() === ''}
                  onClick={() =>
                    void run(onReply(thread.id, replies[thread.id] ?? '')).then((ok) => {
                      if (ok) setReplies({ ...replies, [thread.id]: '' })
                    })
                  }
                >
                  {t('appCmtReply')}
                </button>
              </div>
            </li>
          ))}
        </ul>
      )}
    </aside>
  )
}
