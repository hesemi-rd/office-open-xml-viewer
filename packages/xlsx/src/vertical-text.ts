// ECMA-376 §18.8.1 alignment `textRotation="255"` — Excel "stacked" vertical
// text: each character is stacked upright, one per line, top to bottom. Which
// glyphs stay literally upright and which must be substituted / rotated is decided
// by the Unicode UAX#50 Vertical_Orientation (vo) property (core
// `verticalOrientation`), the SAME classifier the pptx eaVert and docx tbRl paths
// use (issue #790) — only the geometry differs (xlsx has no page rotation).
//
//   • vo=U  (CJK / kana / Hangul / fullwidth forms & digits) → upright, drawn as
//     the caller set up (center / top). Excel stacks these upright.
//   • vo=R  (Latin letters, ASCII digits, Latin punctuation) → upright, unchanged.
//     Unlike pptx/docx eaVert (where Latin rotates), Excel stacks Latin UPRIGHT,
//     one letter per line.
//   • vo=Tu with a U+FE10–FE12 vertical form (、。，) → SUBSTITUTE that form
//     (core `verticalFormSubstitute`), drawn upright so the font hangs it in the
//     cell's upper-right corner; ！？ and small kana have no form and draw upright.
//   • vo=Tr with a U+FE1x/FE3x vertical form present in the substitute fonts — the
//     brackets （）「」〈〉【】… and the white lenticular brackets 〖〗 (#969) →
//     SUBSTITUTE that form (core `verticalBracketFormSubstitute`), drawn upright
//     (UAX#50 §5 "substitute a vertical glyph, rotate only as fallback"). XLSX has
//     no Excel ground-truth image, so 〖〗 follow the Word/PowerPoint verdict.
//   • vo=Tr with NO substituted form → a geometric fallback: ROTATE 90° CW (the
//     quotes “” and the colon ： — a rotation turns the base ：'s stacked dots into
//     FE13's side-by-side dots), ROTATE + REFLECT (the long-stroke marks ー 〜 ～,
//     whose font-designed vertical form is a horizontal mirror of the rotation, not
//     the rotation — core `verticalTrMirrorFallback`; add `scale(1, -1)`), or UPRIGHT
//     (the semicolon ；, whose FE14 form is an upright dot-over-comma, not a rotation;
//     core `verticalTrUprightFallback`). FE13/FE14 are absent from most render fonts,
//     so ：；take this geometric path rather than substitution (issue #969 follow-up).

import {
  verticalOrientation,
  verticalFormSubstitute,
  verticalBracketFormSubstitute,
  verticalTrUprightFallback,
  verticalTrMirrorFallback,
  withVertFeature,
  verticalFallbackShearCoefficient,
} from '@silurus/ooxml-core';

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/**
 * Draw one character of an Excel stacked (`textRotation="255"`) cell into its
 * vertical cell slot. The caller stacks the cells top-to-bottom and sets
 * `textAlign='center'` / `textBaseline='top'`; this routes each glyph through the
 * UAX#50 classifier to draw it upright, substituted, or rotated.
 *
 * Substitution and rotation are GLYPH-only: the caller's cell height / advance and
 * everything the model reports (selection, find) keep the ORIGINAL character.
 *
 * @param ctx       2D context; `ctx.font`/`ctx.fillStyle` and (for the upright
 *                  path) `textAlign='center'` / `textBaseline='top'` are the
 *                  caller's.
 * @param ch        One character (single code point) of the cell text.
 * @param centerX   Horizontal centre of the cell (px) — the upright glyph's x.
 * @param cellTopY  Top y of this character's cell slot (px) — the upright glyph's
 *                  `top`-baseline y.
 * @param charH     Height of one stacked cell slot (px) — used to centre a rotated
 *                  glyph within its slot.
 */
export function drawStackedVerticalChar(
  ctx: Ctx2D,
  ch: string,
  centerX: number,
  cellTopY: number,
  charH: number,
  vertCapable = false,
): void {
  const cp = ch.codePointAt(0) ?? 0;
  const vo = verticalOrientation(cp);

  if (vertCapable && verticalTrMirrorFallback(cp)) {
    ctx.save();
    ctx.translate(centerX, cellTopY + charH / 2);
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    withVertFeature(ctx, () => ctx.fillText(ch, 0, 0));
    ctx.restore();
    return;
  }

  if (vo === 'Tr') {
    // Substitute-first: a fullwidth bracket or white lenticular bracket 〖〗 (#969)
    // has a U+FE3x/FE1x vertical form drawn UPRIGHT (UAX#50 §5). It fills the cell
    // like an ideograph, so the caller's center/top placement lands it correctly.
    const bracket = verticalBracketFormSubstitute(cp);
    if (bracket !== null) {
      ctx.fillText(String.fromCodePoint(bracket), centerX, cellTopY);
      return;
    }
    // Semicolon ；(FF1B): its FE14 vertical form is an upright dot-over-comma (not a
    // rotation), so draw it UPRIGHT like an ideograph — NOT rotated (issue #969
    // follow-up; FE13/FE14 were dropped from the core substitute map because those
    // forms are absent from most render fonts). The colon ：falls through to rotate,
    // reproducing FE13's side-by-side dots.
    if (verticalTrUprightFallback(cp)) {
      ctx.fillText(ch, centerX, cellTopY);
      return;
    }
    // Fallback: no vertical form (ー U+30FC, wave dash / tilde 〜 ～, quotes, colon ：)
    // → rotate 90° CW, centred in the cell slot. save/restore also restores the
    // caller's textAlign/textBaseline.
    //
    // The long-stroke marks ー and 〜 ～ (core `verticalTrMirrorFallback`) additionally
    // REFLECT: their font-designed vertical form is the HORIZONTAL MIRROR of the +90°
    // rotation, not the rotation (Word/PowerPoint + font `vert` glyph verified — a plain
    // rotation of ー bulges LEFT, the designed form bulges RIGHT). After the +90° rotate,
    // y'=−y is the on-screen horizontal mirror. For U+30FC only, the shared
    // runtime coefficient adds y'=m·x−y to cancel the horizontal glyph's measured
    // drift. The wave marks keep their designed drift; quotes and colon do not
    // reflect. Same fix as docx `drawVerticalRun` / pptx `drawEaVertRun`.
    ctx.save();
    ctx.translate(centerX, cellTopY + charH / 2);
    ctx.rotate(Math.PI / 2);
    if (verticalTrMirrorFallback(cp)) {
      ctx.transform(1, verticalFallbackShearCoefficient(ctx, cp), 0, -1, 0, 0);
    }
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(ch, 0, 0);
    ctx.restore();
    return;
  }

  // vo=U / vo=R → upright unchanged; vo=Tu → substitute a vertical form when one
  // exists (、。，), else upright unchanged (！？, small kana).
  const sub = vo === 'Tu' ? verticalFormSubstitute(cp) : null;
  ctx.fillText(sub !== null ? String.fromCodePoint(sub) : ch, centerX, cellTopY);
}
