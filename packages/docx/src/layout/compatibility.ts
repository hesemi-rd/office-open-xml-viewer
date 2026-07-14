import type { CompatibilityRule, DeepReadonly } from './types.js';

function requireText(value: string, field: string): void {
  if (value.trim() === '') throw new Error(`CompatibilityRule.${field} must not be empty`);
}

export function defineCompatibilityRule(rule: CompatibilityRule): DeepReadonly<CompatibilityRule> {
  requireText(rule.id, 'id');
  requireText(rule.description, 'description');
  if (rule.evidence.kind === 'microsoft-note') {
    requireText(rule.evidence.reference, 'evidence.reference');
  } else {
    requireText(rule.evidence.syntheticFixtureId, 'evidence.syntheticFixtureId');
    requireText(rule.evidence.application, 'evidence.application');
    requireText(rule.evidence.version, 'evidence.version');
    requireText(rule.evidence.platform, 'evidence.platform');
  }
  Object.freeze(rule.evidence);
  return Object.freeze(rule);
}
