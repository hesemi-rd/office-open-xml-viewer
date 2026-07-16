import { describe, expect, it } from 'vitest';
import {
  defineCompatibilityRule,
  WORD_SECTION_BTLR_TBRL_PAGE_FRAME,
} from './compatibility.js';

describe('defineCompatibilityRule', () => {
  it('retains explicit Microsoft evidence as immutable data', () => {
    const rule = defineCompatibilityRule({
      id: 'word-example',
      evidence: { kind: 'microsoft-note', reference: '[MS-OE376] §2.1' },
      description: 'Synthetic compatibility boundary',
    });

    expect(rule.evidence).toEqual({ kind: 'microsoft-note', reference: '[MS-OE376] §2.1' });
    expect(Object.isFrozen(rule)).toBe(true);
    expect(Object.isFrozen(rule.evidence)).toBe(true);
  });

  it('rejects observation evidence without a reproducible fixture identity', () => {
    expect(() => defineCompatibilityRule({
      id: 'word-observation',
      evidence: {
        kind: 'office-observation',
        syntheticFixtureId: '',
        application: 'Word',
        version: 'current',
        platform: 'Windows',
      },
      description: 'Unreproducible observation',
    })).toThrow(/syntheticFixtureId/);
  });

  it('records the approved Word section btLr page-frame difference', () => {
    expect(WORD_SECTION_BTLR_TBRL_PAGE_FRAME).toEqual({
      id: 'word-section-btlr-tbrl-page-frame',
      evidence: {
        kind: 'office-observation',
        syntheticFixtureId: 'issue-988-batch-3-section-btlr',
        application: 'Microsoft Word',
        version: 'not recorded',
        platform: 'not recorded',
      },
      description: expect.stringMatching(
        /Issue #988 comment 4950296007.*normative.*lr.*page frame.*glyph orientation.*paint-owned/i,
      ),
    });
    expect(Object.isFrozen(WORD_SECTION_BTLR_TBRL_PAGE_FRAME)).toBe(true);
    expect(Object.isFrozen(WORD_SECTION_BTLR_TBRL_PAGE_FRAME.evidence)).toBe(true);
  });
});
