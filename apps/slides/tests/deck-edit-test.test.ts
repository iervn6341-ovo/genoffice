import { describe, it, expect } from 'vitest';
import fs from 'fs';
import { openPptx, savePptx, reparseDeck } from '@genoffice/pptx-engine';
import { runTxn } from '@genoffice/pptx-ops';
import { buildRenderSlide } from '@genoffice/pptx-render';
import { populateEditorDom, extractParagraphs, ensureEditorSelection } from '../src/renderer/TextEditOverlay';
import { registerTextCommit, flushTextCommit } from '../src/renderer/active-text-edit';

import path from 'path';

describe('第五次估咪簡報.pptx text editing & font color', () => {
  it('tests setFont and setText on elements of 第五次估咪簡報.pptx', async () => {
    const candidates = [
      '第五次估咪簡報.pptx',
      '第四次估咪簡報.pptx',
      '10:19 精誠開會簡報.pptx',
    ];
    let filePath: string | undefined;
    for (const name of candidates) {
      const p1 = path.resolve(process.cwd(), name);
      const p2 = path.resolve(__dirname, '../../../', name);
      const p3 = path.resolve(__dirname, '../../', name);
      if (fs.existsSync(p1)) { filePath = p1; break; }
      if (fs.existsSync(p2)) { filePath = p2; break; }
      if (fs.existsSync(p3)) { filePath = p3; break; }
    }
    if (!filePath) throw new Error('No test presentation file found');
    const buf = fs.readFileSync(filePath);
    const opened = await openPptx(buf);

    // Let's test on every slide with text shapes
    for (let sIdx = 0; sIdx < Math.min(5, opened.deck.slides.length); sIdx++) {
      const slide = opened.deck.slides[sIdx];
      for (const el of slide.elements) {
        const textEl = el as any;
        if (!textEl.text?.paragraphs?.length) continue;
        console.log(`Testing Slide ${sIdx + 1} Element ${el.id} (${el.name})`);
        
        // Test setFont (element font color)
        const resColor = runTxn(opened, {
          ops: [
            {
              op: 'setFont',
              target: { slide: sIdx, el: el.id },
              font: { color: '#00AA55' },
            },
          ],
        });
        expect(resColor.applied).toBe(true);

        // Test setText
        const originalText = textEl.text.paragraphs[0]?.runs[0]?.text ?? '';
        const resText = runTxn(opened, {
          ops: [
            {
              op: 'setText',
              target: { slide: sIdx, el: el.id },
              paragraphs: [
                {
                  runs: [
                    {
                      text: originalText + ' [EDITED]',
                      color: '#FF0000',
                    },
                  ],
                },
              ],
            },
          ],
        });
        expect(resText.applied).toBe(true);
      }
    }

    // Save and re-open to verify XML generation and reparsing
    const savedBytes = await savePptx(opened);
    const reopened = await openPptx(savedBytes);
    expect(reopened.deck.slides.length).toBe(opened.deck.slides.length);

    // Check if the edited text and colors are present
    const s0 = reopened.deck.slides[0];
    const el0 = s0.elements[0] as any;
    console.log('Reopened Slide 1 Element 0 text:', el0.text?.paragraphs[0]?.runs[0]?.text);
    console.log('Reopened Slide 1 Element 0 color:', el0.text?.paragraphs[0]?.runs[0]?.color);
    expect(el0.text?.paragraphs[0]?.runs[0]?.text).toContain('[EDITED]');
    expect(el0.text?.paragraphs[0]?.runs[0]?.color).toBe('#FF0000');

    // Also check buildRenderSlide
    const rendered = buildRenderSlide(s0, reopened.deck.size, { fitWidthPx: 1280 });
    const node0 = rendered.nodes.find((n) => n.sourceId === el0.id) as any;
    console.log('Rendered node 0 text runs:', node0?.text?.lines[0]?.runs[0]);
    expect(node0?.text?.lines[0]?.runs[0]?.color).toBe('#FF0000');
  });

  it('tests populateEditorDom, DOM formatting and extractParagraphs', async () => {
    const candidates = [
      '第五次估咪簡報.pptx',
      '第四次估咪簡報.pptx',
      '10:19 精誠開會簡報.pptx',
    ];
    let filePath: string | undefined;
    for (const name of candidates) {
      const p1 = path.resolve(process.cwd(), name);
      const p2 = path.resolve(__dirname, '../../../', name);
      const p3 = path.resolve(__dirname, '../../', name);
      if (fs.existsSync(p1)) { filePath = p1; break; }
      if (fs.existsSync(p2)) { filePath = p2; break; }
      if (fs.existsSync(p3)) { filePath = p3; break; }
    }
    const buf = fs.readFileSync(filePath!);
    const opened = await openPptx(buf);
    const s0 = opened.deck.slides[0];
    const el0 = s0.elements[0];
    const rendered = buildRenderSlide(s0, opened.deck.size, { fitWidthPx: 1280 });
    const node0 = rendered.nodes.find((n) => n.sourceId === el0.id) as any;

    const div = document.createElement('div');
    div.className = 'slide-text-editor';
    div.contentEditable = 'true';
    document.body.appendChild(div);
    populateEditorDom(div, node0.text.lines);

    const initialParas = extractParagraphs(div, 1);
    console.log('Initial extracted paras:', JSON.stringify(initialParas));

    // Test ensureEditorSelection on collapsed selection:
    div.focus();
    const sel = window.getSelection();
    const range = document.createRange();
    range.setStart(div.firstChild!, 0);
    range.collapse(true);
    sel?.removeAllRanges();
    sel?.addRange(range);
    expect(sel?.isCollapsed).toBe(true);

    const expandedRange = ensureEditorSelection(div);
    expect(expandedRange).not.toBeNull();
    // After ensureEditorSelection, the selection is no longer collapsed
    expect(sel?.isCollapsed).toBe(false);

    // Test extractParagraphs with style.color (#0000FF) and font tag color
    const span = document.createElement('span');
    span.style.color = '#0000FF';
    span.textContent = 'Custom Blue Text';
    div.appendChild(span);

    const fontTag = document.createElement('font');
    fontTag.setAttribute('color', '#00FF00');
    fontTag.textContent = 'Custom Green Text';
    div.appendChild(fontTag);

    const parasAfterColor = extractParagraphs(div, 1);
    const blueRun = parasAfterColor.flatMap((p) => p.runs).find((r) => r.text.includes('Blue'));
    const greenRun = parasAfterColor.flatMap((p) => p.runs).find((r) => r.text.includes('Green'));
    expect(blueRun?.color).toBe('#0000FF');
    expect(greenRun?.color).toBe('#00FF00');

    div.remove();
  });

  it('verifies flushTextCommit flushes pending text edits during save or slide switch', async () => {
    let committedText = '';
    const unregister = registerTextCommit(() => {
      committedText = 'Committed Text Content';
    });

    await flushTextCommit();
    expect(committedText).toBe('Committed Text Content');
    unregister();
  });
});
