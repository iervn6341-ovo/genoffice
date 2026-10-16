import { describe, expect, it } from 'vitest'

import {
  addReply,
  addThread,
  deleteComment,
  editComment,
  ensurePerson,
  setThreadDone,
  shiftThreads,
  threadAt,
  threadTimestamp,
} from '../src/renderer/threaded-comments'

const base = addThread([], {
  row: 2,
  column: 1,
  text: 'first',
  personId: '{P}',
  time: '2026-10-15T08:00:00.00',
  id: '{T1}',
}).threads

describe('thread operations', () => {
  it('adds one thread per cell and rejects empty text', () => {
    expect(base).toHaveLength(1)
    expect(
      addThread(base, { row: 2, column: 1, text: 'x', personId: 'p', time: 't', id: '{T2}' }).error,
    ).toBe('occupied')
    expect(
      addThread([], { row: 0, column: 0, text: '   ', personId: 'p', time: 't', id: '{T3}' }).error,
    ).toBe('empty')
    expect(threadAt(base, 2, 1)?.id).toBe('{T1}')
  })

  it('replies, resolves, edits and deletes with the root owning the thread', () => {
    let threads = addReply(base, '{T1}', { id: '{R1}', text: 'reply', personId: '{Q}', time: 't2' })
    expect(threads[0]!.comments.map((c) => c.parentId)).toEqual([undefined, '{T1}'])
    threads = setThreadDone(threads, '{T1}', true)
    expect(threads[0]!.done).toBe(true)
    // a reply reopens a resolved thread
    threads = addReply(threads, '{T1}', { id: '{R2}', text: 'again', personId: '{Q}', time: 't3' })
    expect(threads[0]!.done).toBe(false)
    threads = editComment(threads, '{T1}', '{R1}', 'edited')
    expect(threads[0]!.comments[1]!.text).toBe('edited')
    threads = deleteComment(threads, '{T1}', '{R1}')
    expect(threads[0]!.comments.map((c) => c.id)).toEqual(['{T1}', '{R2}'])
    expect(deleteComment(threads, '{T1}', '{T1}')).toEqual([])
  })

  it('drops raw xml only for the comment that changed', () => {
    const withRaw = [
      {
        ...base[0]!,
        comments: [
          { ...base[0]!.comments[0]!, raw: '<threadedComment id="{T1}"/>' },
          { id: '{R1}', personId: 'p', time: 't', text: 'r', parentId: '{T1}', raw: '<r/>' },
        ],
      },
    ]
    const resolved = setThreadDone(withRaw, '{T1}', true)
    expect(resolved[0]!.comments[0]!.raw).toBeUndefined()
    expect(resolved[0]!.comments[1]!.raw).toBe('<r/>')
    expect(editComment(withRaw, '{T1}', '{R1}', 'new')[0]!.comments[1]!.raw).toBeUndefined()
  })

  it('keeps one person per display name', () => {
    let result = ensurePerson([], 'Ann', () => '{A}')
    result = ensurePerson(result.persons, 'Ann', () => '{B}')
    expect(result.personId).toBe('{A}')
    expect(result.persons).toHaveLength(1)
  })

  it('formats Excel local timestamps', () => {
    expect(threadTimestamp(new Date(2026, 9, 5, 7, 8, 9, 450))).toBe('2026-10-05T07:08:09.45')
  })
})

describe('thread positions follow row and column edits', () => {
  it('shifts on insert and delete, and Undo of a delete restores the cell', () => {
    const insert = shiftThreads('s', base, [], { kind: 'insert-rows', index: 1, count: 2 }, false)
    expect(insert.threads[0]!.row).toBe(4)
    const removed = shiftThreads('s', base, [], { kind: 'remove-rows', index: 2, count: 1 }, false)
    expect(removed.threads).toEqual([])
    expect(removed.stash).toHaveLength(1)
    // a user insert of the same shape is not an undo: nothing comes back
    const stray = shiftThreads(
      's',
      [],
      removed.stash,
      { kind: 'insert-rows', index: 2, count: 1 },
      false,
    )
    expect(stray.threads).toEqual([])
    const undone = shiftThreads(
      's',
      [],
      removed.stash,
      { kind: 'insert-rows', index: 2, count: 1 },
      true,
    )
    expect(undone.threads[0]!.row).toBe(2)
    expect(undone.stash).toEqual([])
    const below = shiftThreads('s', base, [], { kind: 'remove-rows', index: 0, count: 1 }, false)
    expect(below.threads[0]!.row).toBe(1)
  })

  it('handles columns and row moves', () => {
    expect(
      shiftThreads('s', base, [], { kind: 'insert-cols', index: 0, count: 1 }, false).threads[0],
    ).toMatchObject({ row: 2, column: 2 })
    // move row 2 up to before row 0
    expect(
      shiftThreads('s', base, [], { kind: 'move-rows', index: 2, count: 1, before: 0 }, false)
        .threads[0]!.row,
    ).toBe(0)
    // moving row 0 down to before row 4 pulls row 2 up by one
    expect(
      shiftThreads('s', base, [], { kind: 'move-rows', index: 0, count: 1, before: 4 }, false)
        .threads[0]!.row,
    ).toBe(1)
  })
})
