import { describe, expect, it } from 'vitest';
import type { SlidesModel } from '../src/shared/model.ts';
import { PptxAdapter } from '../src/webview/adapters/pptxAdapter.ts';
import { makePptx } from './helpers/fixtures.ts';

describe('pptx adapter', () => {
  it('returns slides in presentation order with titles, text, notes and pictures', async () => {
    const a = new PptxAdapter();
    const model = (await a.load(makePptx())) as SlidesModel;
    expect(model.type).toBe('slides');
    expect(model.slides.map((s) => s.title)).toEqual(['Opening slide', 'Second in order']);
    expect(model.slides.map((s) => s.number)).toEqual([1, 2]);
    expect(model.slides[0]!.paragraphs).toEqual(['Point A', 'Point B']);
    expect(model.slides[0]!.notes).toBe('Remember to smile');
    expect(model.slides[0]!.imageDataUri).toMatch(/^data:image\/png;base64,/);
    expect(model.slides[1]!.paragraphs).toEqual(['Detail line']);
    expect(model.slides[1]!.imageDataUri).toBeUndefined();
  });

  it('is read-only', async () => {
    const a = new PptxAdapter();
    expect(a.readonly).toBe(true);
    await a.load(makePptx());
    expect(() => a.applyEdit({ kind: 'paragraph', index: 0, text: 'x', old: '' })).toThrow(/read-only/);
  });
});
