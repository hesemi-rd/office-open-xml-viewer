/**
 * PPTX hyperlink classification (IX1).
 *
 * The Rust parser hands the TS side a single resolved target string per run /
 * shape (`hyperlink`) plus, for shapes, the raw `<a:hlinkClick @action>`
 * (`hyperlinkAction`). This module turns that pair into the format-agnostic
 * {@link HyperlinkTarget} shape the overlay + viewer consume.
 *
 * ECMA-376 §21.1.2.3.5 (CT_Hyperlink): a `<a:hlinkClick>` is EXTERNAL when its
 * relationship `TargetMode="External"` (an absolute URL — http/https/mailto/tel
 * etc.), and INTERNAL when it names an action verb (`ppaction://hlinksldjump`
 * → the rel is `TargetMode="Internal"` and resolves to a slide part such as
 * `../slides/slide3.xml`).
 *
 * Classification is done purely from the resolved target string (the core
 * `TextRun` type only carries `hyperlink`, so text runs cannot smuggle a second
 * field through — see the pptx types re-export). The rule:
 *   - an explicit `hyperlinkAction` (any `ppaction://…`), OR
 *   - a target whose URL scheme is NOT one of the navigable external schemes
 *     (http/https/mailto/tel — {@link DEFAULT_ALLOWED_HYPERLINK_SCHEMES}),
 *     which covers `ppaction://…` targets and bare internal part names like
 *     `../slides/slide3.xml` (no scheme),
 * ⇒ INTERNAL. Everything else is EXTERNAL.
 *
 * Keeping the scheme predicate in core (`hyperlinkUrlScheme`) means the
 * external-vs-internal boundary matches the sanitiser's allowlist exactly: a
 * scheme the viewer would refuse to open externally is treated as internal here
 * rather than silently dropped.
 */
import {
  type HyperlinkTarget,
  DEFAULT_ALLOWED_HYPERLINK_SCHEMES,
  hyperlinkUrlScheme,
} from '@silurus/ooxml-core';

/**
 * Classify a resolved pptx hyperlink target into a {@link HyperlinkTarget}, or
 * `undefined` when there is no link.
 *
 * @param target the resolved `hyperlink` string (external URL or internal part
 *               name), or undefined/empty when the run/shape has no hlinkClick.
 * @param action the raw `<a:hlinkClick @action>` string when present (shapes
 *               carry this; text runs pass `undefined`).
 */
export function classifyPptxHyperlink(
  target: string | undefined,
  action?: string,
): HyperlinkTarget | undefined {
  // Narrow to non-empty strings so the returned `url` / `ref` are never empty.
  const t = target !== undefined && target !== '' ? target : undefined;
  const a = action !== undefined && action !== '' ? action : undefined;
  if (t === undefined && a === undefined) return undefined;

  // An action verb (ppaction://…) is always internal. `ref` is the resolved
  // internal part name when we have one, else the raw action verb.
  if (a !== undefined) {
    return { kind: 'internal', ref: t ?? a };
  }

  // No action: classify by the target's URL scheme. A navigable external scheme
  // (http/https/mailto/tel) ⇒ external; anything else (a bare internal part
  // name with no scheme, or a non-navigable scheme) ⇒ internal. `t` is defined
  // here (a was undefined and at least one of t/a is set).
  const url = t as string;
  const scheme = hyperlinkUrlScheme(url);
  if (scheme !== null && DEFAULT_ALLOWED_HYPERLINK_SCHEMES.includes(scheme)) {
    return { kind: 'external', url };
  }
  return { kind: 'internal', ref: url };
}
