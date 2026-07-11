/**
 * Same-font Canvas-vs-Word horizontal advance bias (Regime A).
 *
 * This module corrects the measurement gap between Canvas `measureText` and
 * Word's own layout advances for the SAME resolved face. It deliberately does
 * NOT emulate the metrics of a font the host does not have: when a requested
 * face is absent, the viewer substitutes the host's fallback face and reflows —
 * exactly like Word does when it opens a document whose fonts are missing.
 * Line breaks and page counts may then differ from the authoring machine; that
 * is accepted behavior, not a defect (product decision 2026-07-12; issue #855
 * closed as won't-fix). Cross-font per-script advance profiles that faked the
 * requested face's metrics on top of a substitute (PR #979 "Regime B") were
 * reverted because per-family metric tables cannot scale across the open set
 * of document fonts.
 *
 * Values in this module describe fonts, not samples. A family belongs here only
 * when its bias has a documented provenance; unknown families remain a strict
 * zero no-op so existing layout is byte-stable.
 */

interface FontBiasProfile {
  readonly test: (family: string) => boolean;
  readonly biasEm: number;
}

/**
 * Canvas-vs-Word horizontal advance bias in em per glyph. This is a line-fit
 * allowance for the gap between Canvas `measureText` advances and Word's own
 * layout advances for the SAME face, not a glyph transform. The allowance is
 * backend-agnostic by design; the committed Georgia value below is calibrated
 * on the Chromium VRT.
 *
 * Georgia: the tracked public demo's justified body face (issue #794). The
 * The Chromium VRT's Canvas-vs-Word accumulated excess measures at roughly
 * 0.1–0.3 px per glyph
 * at the demo's 10–11 px body em — an em-fraction band of ~0.009–0.028. The
 * committed value is fixed INSIDE that measured band by the public demo's
 * Word-reference wrap positions (demo/sample-1 fidelity ratchet, scanned at
 * 0.009/0.0105/0.0115/0.012/0.013/0.02): 0.009 under-admits words Word keeps
 * (pages 4–5 regress) and 0.012+ over-admits words Word wraps (pages 3/5);
 * 0.0105 reproduces every Word-verified wrap. Times New Roman (measures
 * ~0.03 px/glyph, effectively zero), CSS generics, and unknown families
 * intentionally fall through to zero.
 */
const FONT_BIAS_PROFILES: ReadonlyArray<FontBiasProfile> = [
  {
    test: (family) => family === 'georgia',
    biasEm: 0.0105,
  },
];

function normalizeFamily(family: string | null | undefined): string {
  return (family ?? '')
    .trim()
    .replace(/^["']|["']$/g, '')
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

/** Expected Canvas-over-Word advance bias, in em per glyph. */
export function fontAdvanceBiasEm(family: string | null | undefined): number {
  const normalized = normalizeFamily(family);
  for (const profile of FONT_BIAS_PROFILES) {
    if (profile.test(normalized)) return profile.biasEm;
  }
  return 0;
}
