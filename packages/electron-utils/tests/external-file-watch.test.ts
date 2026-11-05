import { afterEach, it, expect, vi } from 'vitest'
const fs = vi.hoisted(() => ({
  event: undefined as undefined | ((kind: string, filename: string | null) => void),
  close: vi.fn(),
  watch: vi.fn(),
}))
vi.mock('node:fs', () => ({ watch: (...args: any[]) => fs.watch(...args) }))
import { ExternalFileWatch } from '../src/external-file-watch'
afterEach(() => vi.useRealTimers())

it('debounces matching events and keeps watching after atomic replacement', () => {
  vi.useFakeTimers()
  fs.watch.mockImplementation((_path, _options, event) => {
    fs.event = event
    return { on: vi.fn(), close: fs.close }
  })
  const watcher = new ExternalFileWatch()
  const changed = vi.fn()
  watcher.remember(1, '/tmp/deck/file.pptx', changed)
  expect(fs.watch).toHaveBeenCalledWith('/tmp/deck', { persistent: false }, expect.any(Function))
  fs.event!('change', 'other.pptx')
  vi.advanceTimersByTime(300)
  expect(changed).not.toHaveBeenCalled()
  for (let i = 0; i < 2; i++) {
    fs.event!('rename', 'file.pptx')
    fs.event!('change', 'file.pptx')
    vi.advanceTimersByTime(250)
    expect(changed).toHaveBeenCalledTimes(i + 1)
  }
  expect(fs.watch).toHaveBeenCalledTimes(1)
  fs.event!('change', null)
  watcher.forget(1)
  vi.advanceTimersByTime(300)
  expect(changed).toHaveBeenCalledTimes(2)
  expect(fs.close).toHaveBeenCalledTimes(1)
})

it('uses focus/save fallback if OS notifications cannot be registered', () => {
  fs.watch.mockImplementation(() => {
    throw new Error('unsupported volume')
  })
  const watcher = new ExternalFileWatch()
  expect(() => watcher.remember(2, '/tmp/deck/file.pptx', vi.fn())).not.toThrow()
  watcher.forget(2)
})
