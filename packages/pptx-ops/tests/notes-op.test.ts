/**
 * Verifies the setNotes op executed by the pptx-ops executor.
 * Tests plain text speaker notes, multiline notes, formatted paragraphs,
 * clearing notes, batch execution alongside element ops, and error validation.
 */
import { describe, it, expect, beforeEach } from 'vitest'
import {
  createBlankPptx,
  openPptx,
  getSlideNotes,
  getSlideNotesParagraphs,
  addElement,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

describe('setNotes op in pptx-ops executor', () => {
  let opened: OpenedPptx

  beforeEach(async () => {
    opened = await openPptx(await createBlankPptx())
  })

  it('sets plain text speaker notes on slide 0', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setNotes',
          target: { slide: 0 },
          text: 'Welcome everyone to the quarterly presentation.',
        },
      ],
    })

    expect(r.applied).toBe(true)
    expect(r.failures).toBeUndefined()
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe(
      'Welcome everyone to the quarterly presentation.',
    )
  })

  it('supports multiline speaker notes', () => {
    const notesText = 'Point 1: Q1 revenue exceeded expectations\nPoint 2: Expansion plans for Q2'
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setNotes',
          target: { slide: 0 },
          text: notesText,
        },
      ],
    })

    expect(r.applied).toBe(true)
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe(notesText)
  })

  it('supports rich paragraphs with formatting', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setNotes',
          target: { slide: 0 },
          paragraphs: [
            {
              runs: [
                { text: 'Important Note: ', bold: true },
                { text: 'Do not share confidential figures.' },
              ],
            },
          ],
        },
      ],
    })

    expect(r.applied).toBe(true)
    const readBack = getSlideNotes(opened.archive, opened.deck.slides[0]!.path)
    expect(readBack).toBe('Important Note: Do not share confidential figures.')

    const paras = getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)
    expect(paras.length).toBeGreaterThanOrEqual(1)
    const fullText = paras[0]!.runs.map((r) => r.text).join('')
    expect(fullText).toBe('Important Note: Do not share confidential figures.')
  })

  it('clears notes when given an empty string', () => {
    // First set a note
    runTxn(opened, {
      ops: [
        {
          op: 'setNotes',
          target: { slide: 0 },
          text: 'Temporary speaker note',
        },
      ],
    })
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe(
      'Temporary speaker note',
    )

    // Then clear it
    const clearResult = runTxn(opened, {
      ops: [
        {
          op: 'setNotes',
          target: { slide: 0 },
          text: '',
        },
      ],
    })

    expect(clearResult.applied).toBe(true)
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe('')
  })

  it('runs setNotes alongside canvas element ops in an atomic transaction', () => {
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
      paragraphs: [{ runs: [{ text: 'Slide Title' }] }],
    })

    const r = runTxn(opened, {
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el: el.id },
          paragraphs: [{ runs: [{ text: 'Updated Slide Title' }] }],
        },
        {
          op: 'setNotes',
          target: { slide: 0 },
          text: 'Speaker prompt for slide title',
        },
      ],
    })

    expect(r.applied).toBe(true)
    expect(r.records).toHaveLength(2)
    expect(getSlideNotes(opened.archive, slide.path)).toBe('Speaker prompt for slide title')
  })

  it('rejects setNotes without text or paragraphs', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setNotes',
          target: { slide: 0 },
        },
      ],
    })

    expect(r.applied).toBe(false)
    expect(r.failures).toHaveLength(1)
    expect(r.failures![0]!.error).toContain('needs "text" (or "paragraphs"')
  })

  it('rejects setNotes with out-of-range slide index', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setNotes',
          target: { slide: 99 },
          text: 'Out of range',
        },
      ],
    })

    expect(r.applied).toBe(false)
    expect(r.failures).toHaveLength(1)
    expect(r.failures![0]!.error).toContain('slide index 99 is out of range')
  })
})
