/**
 * Internal-hyperlink target resolution shared by docx / pptx / xlsx (IX-nav).
 *
 * An *internal* hyperlink jumps within the document itself rather than opening a
 * URL: docx `w:anchor` -> a `<w:bookmarkStart w:name>` destination (§17.16.23),
 * pptx `action="ppaction://hlinksldjump"` -> a slide part, pptx
 * `action="ppaction://hlinkshowjump?jump=firstslide|…"` -> a relative slide.
 * Turning a raw target into a *destination page / slide index* is the piece the
 * viewers were missing (the IX1 internal-nav no-op): this module supplies the
 * pure resolution primitives, and each viewer owns the map it resolves against
 * (docx bookmark→page, pptx slidePart→index).
 *
 * Scope is deliberately "pure predicate + arithmetic": no DOM, no model access.
 * `resolveOpcPartName` mirrors the Rust `ooxml_common::rels::resolve_target`
 * byte-for-byte so a slide-rel target like `../slides/slide3.xml` resolves to the
 * same normalized part name (`ppt/slides/slide3.xml`) the parser keys its slide
 * list by — keeping the TS index lookup and the Rust part naming in lockstep.
 */

/**
 * Resolve an OPC relationship `Target` to a normalized, root-relative zip part
 * name — the TS mirror of the Rust `ooxml_common::rels::resolve_target`
 * (ECMA-376 Part 2 §9.3). Kept identical so a pptx internal slide target
 * (`../slides/slide3.xml`, authored relative to `ppt/slides`) normalizes to the
 * exact `ppt/slides/slide3.xml` part name the parser tags each slide with.
 *
 *   - **Root-absolute** (`target` starts with `/`, e.g. `/ppt/slides/slide1.xml`):
 *     resolved from the package root, ignoring `baseDir`; the leading slash is
 *     dropped.
 *   - **Relative** (`../slides/slide3.xml`, `slide1.xml`): resolved against
 *     `baseDir` — the *directory* of the source part (e.g. `ppt/slides`). `..`
 *     pops one segment; `.` and empty segments are skipped.
 *
 * @param baseDir directory of the source part (no trailing slash needed).
 * @param target  the relationship `Target` verbatim.
 * @returns the normalized part name with no relative components.
 */
export function resolveOpcPartName(baseDir: string, target: string): string {
  const parts: string[] = target.startsWith('/')
    ? [] // Root-absolute part name: ignore baseDir entirely.
    : baseDir.split('/').filter((s) => s !== '');
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg === '.' || seg === '') continue;
    else parts.push(seg);
  }
  return parts.join('/');
}

/**
 * The four relative slide-show jump verbs of `ppaction://hlinkshowjump`
 * (ECMA-376 §21.1.2.3.5, the `<a:hlinkClick @action>` action list). Each names a
 * slide position relative to the current one rather than a specific slide part.
 */
export type RelativeSlideJump = 'firstslide' | 'lastslide' | 'nextslide' | 'previousslide';

/**
 * Parse the `jump` query of a `ppaction://hlinkshowjump?jump=…` action into its
 * {@link RelativeSlideJump} verb, or `null` when the action is not a
 * show-jump / carries an unrecognized verb. Case-insensitive on the verb (the
 * action URI itself is authored lowercase, but we normalize defensively).
 *
 * Only the four navigation verbs are recognized here; other show actions
 * (`endshow`, `lastslideviewed`, custom shows) have no positional slide index
 * and return `null` so the caller no-ops rather than jumping somewhere wrong.
 */
export function parseRelativeSlideJump(action: string): RelativeSlideJump | null {
  const m = /[?&]jump=([a-zA-Z]+)/.exec(action);
  if (!m) return null;
  const verb = m[1].toLowerCase();
  if (
    verb === 'firstslide' ||
    verb === 'lastslide' ||
    verb === 'nextslide' ||
    verb === 'previousslide'
  ) {
    return verb;
  }
  return null;
}

/**
 * Resolve a {@link RelativeSlideJump} to a 0-based slide index, given the
 * `current` index and the total slide `count`. `nextslide` / `previousslide`
 * clamp at the deck ends (a "next" past the last slide, or a "previous" before
 * the first, stays put) — matching PowerPoint, which simply does nothing at a
 * boundary rather than wrapping. Returns `undefined` when `count` is 0 (no slide
 * to land on).
 *
 * @param jump    the parsed relative verb.
 * @param current the 0-based index the jump is relative to.
 * @param count   total slide count.
 */
export function resolveRelativeSlideJump(
  jump: RelativeSlideJump,
  current: number,
  count: number,
): number | undefined {
  if (count <= 0) return undefined;
  switch (jump) {
    case 'firstslide':
      return 0;
    case 'lastslide':
      return count - 1;
    case 'nextslide':
      return Math.min(current + 1, count - 1);
    case 'previousslide':
      return Math.max(current - 1, 0);
  }
}
