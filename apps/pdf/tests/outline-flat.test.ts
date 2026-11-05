import { describe, expect, it } from 'vitest'
import { flattenOutline } from '../src/renderer/OutlinePanel'

describe('floating PDF outline adapter', () => {
  it('preserves tree order, hierarchy, named destinations and external links', () => {
    const nodes = [
      {
        title: 'Chapter',
        dest: 'chapter-one',
        items: [{ title: 'Section', dest: [2, { name: 'Fit' }] }],
      },
      { title: 'Reference', url: 'https://example.com' },
    ]
    const flat = flattenOutline(nodes)
    expect(flat.map(({ id, text, level }) => ({ id, text, level }))).toEqual([
      { id: '0', text: 'Chapter', level: 1 },
      { id: '0.0', text: 'Section', level: 2 },
      { id: '1', text: 'Reference', level: 1 },
    ])
    expect(flat[0]!.node).toBe(nodes[0])
    expect(flat[1]!.node.dest).toEqual([2, { name: 'Fit' }])
    expect(flat[2]!.node.url).toBe('https://example.com')
  })

  it('omits blank labels while retaining their children and unique paths', () => {
    expect(
      flattenOutline([{ title: ' ', items: [{ title: 'Kept' }] }, { title: 'Kept' }]).map(
        ({ id, text }) => ({ id, text }),
      ),
    ).toEqual([
      { id: '0.0', text: 'Kept' },
      { id: '1', text: 'Kept' },
    ])
  })
})
