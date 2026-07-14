import { describe, expect, it } from 'vitest';
import { testFontSnapshot } from './test-font-snapshot.js';

describe('explicit deterministic test font snapshot', () => {
  it('does not discover families or invent unavailable style tuples', () => {
    const snapshot = testFontSnapshot([
      { family: 'Regular Only' },
      { family: 'Styled', weight: 700, style: 'italic' },
      { family: 'Synthesized', weight: 700, synthesized: true },
    ]);

    expect(Object.keys(snapshot)).toEqual([
      'regular only',
      'styled:700:italic',
      'synthesized:700:normal',
    ]);
    expect(snapshot['regular only']).toMatchObject({ synthesized: false, weight: 400, style: 'normal' });
    expect(snapshot['styled:700:italic']).toMatchObject({ synthesized: false, weight: 700, style: 'italic' });
    expect(snapshot['synthesized:700:normal']).toMatchObject({ synthesized: true });
    expect(snapshot['regular only:700:normal']).toBeUndefined();
  });
});
