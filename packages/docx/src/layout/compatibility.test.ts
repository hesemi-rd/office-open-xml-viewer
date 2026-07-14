import { describe, expect, it } from 'vitest';
import { defineCompatibilityRule } from './compatibility.js';

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
});
