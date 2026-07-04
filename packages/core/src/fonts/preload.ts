/**
 * Google Fonts preload utility shared by docx / pptx / xlsx viewers.
 *
 * The contract is intentionally narrow: callers pass the set of font-family
 * names they want available, plus a static map from a lower-cased key to a
 * Google Fonts CSS URL (and optionally an alternate FontFaceSet family name
 * for Office substitutes such as Calibri → Carlito). Names without a map
 * entry are skipped (the renderer falls back to the system font).
 *
 * Rather than inject a `<link rel="stylesheet">` and read `document.fonts`
 * (which is impossible inside a Web Worker), this fetches the Google Fonts CSS
 * directly, parses its `@font-face` rules, and registers `FontFace` objects
 * into whichever FontFaceSet exists in the current JS context —
 * `document.fonts` on the main thread, `self.fonts` in a worker. This keeps the
 * loader FontFaceSet-agnostic so both the main-thread and worker rendering
 * modes share one code path.
 *
 * Font load is forced via `face.load()` rather than `FontFaceSet.load()`
 * because canvas-only rendering does not put glyphs into the DOM, so the
 * unicode-range gating in modern Google Fonts CSS would otherwise leave the
 * `FontFace` entries in the `unloaded` state — the first paint would then
 * use a system fallback and shift once a later interaction re-rasterized
 * the canvas after the font landed.
 */
export interface FontPreloadEntry {
  /** Google Fonts CSS URL — `display=swap` recommended. */
  url: string;
  /**
   * Family name to drive {@link FontFaceSet} loading when the substitute
   * differs from the requested face (e.g. Calibri → Carlito). Defaults to
   * the requested name when omitted.
   */
  loadFamily?: string;
}

/**
 * Hard ceiling so a wedged network (a stylesheet fetch that never settles, or a
 * `FontFace.load()` that never resolves) cannot hang the caller forever. This is
 * a SAFETY NET, not the normal exit: on a reachable network every awaited
 * promise settles well within it, so first paint is deterministic. It is
 * intentionally generous — the previous 3 s timeout RACED the font loads and
 * could resolve while faces were still downloading, which is exactly the
 * cold-cache flicker this function must prevent.
 */
const HARD_CEILING_MS = 15000;

/** Race a font-load promise against a generous hard ceiling so a wedged network
 *  or a `FontFace.load()` that never settles cannot hang the caller forever.
 *  Shared by the Google-Fonts and embedded-font loaders (same first-paint
 *  determinism contract). */
export function withFontCeiling<T>(p: Promise<T>): Promise<T | void> {
  return Promise.race([
    p,
    new Promise<void>((resolve) => setTimeout(resolve, HARD_CEILING_MS)),
  ]);
}

/** In-flight (or completed) fetch-and-register promise per CSS url, keyed by url.
 *  Resolves with the `FontFace` objects this loader added for that stylesheet
 *  (empty on fetch failure). Storing the PROMISE (not just a flag) lets a
 *  concurrent caller with the same url JOIN the first registration instead of
 *  seeing a cache hit, skipping registration, and resolving while the first
 *  fetch is still in flight (which would defeat the first-paint determinism this
 *  function exists to provide). It also hands step 2 the exact FontFace
 *  references to load — see why that matters there. The stored promise ALWAYS
 *  resolves: failure handling (recording the failed families + deleting the
 *  entry so a later call can retry) happens inside the producing call's own
 *  logic, so a cache-hit awaiter never throws. */
const cssRegistrations = new Map<string, Promise<FontFace[]>>();

/** Test hook — clears the per-context CSS registration cache. */
export function _resetCssCacheForTests(): void {
  cssRegistrations.clear();
}

export interface ParsedFontFace {
  family: string;
  src: string;
  descriptors: FontFaceDescriptors;
}

/** Extract @font-face rules from a Google Fonts stylesheet. Deliberately
 *  minimal: Google's CSS is machine-generated (one declaration per line, no
 *  nesting), so a brace-block regex is sufficient and avoids a CSS parser. */
export function parseFontFaceRules(css: string): ParsedFontFace[] {
  const faces: ParsedFontFace[] = [];
  const blockRe = /@font-face\s*\{([^}]*)\}/g;
  let m: RegExpExecArray | null;
  while ((m = blockRe.exec(css))) {
    const body = m[1];
    const prop = (name: string): string | undefined =>
      body.match(new RegExp(`(?:^|;|\\n)\\s*${name}\\s*:\\s*([^;]+)`, 'i'))?.[1].trim();
    const familyRaw = prop('font-family');
    const src = prop('src');
    if (!familyRaw || !src) continue;
    const descriptors: FontFaceDescriptors = {};
    const style = prop('font-style');
    if (style) descriptors.style = style;
    const weight = prop('font-weight');
    if (weight) descriptors.weight = weight;
    const stretch = prop('font-stretch');
    if (stretch) descriptors.stretch = stretch;
    const unicodeRange = prop('unicode-range');
    if (unicodeRange) descriptors.unicodeRange = unicodeRange;
    faces.push({ family: familyRaw.replace(/^['"]|['"]$/g, ''), src, descriptors });
  }
  return faces;
}

/** The FontFaceSet of the current context: `document.fonts` on the main
 *  thread, `self.fonts` in a worker, null elsewhere (Node without a shim).
 *  Exported so the embedded-font loader shares one FontFaceSet-resolution rule
 *  with the Google-Fonts loader (both must register into the SAME set). */
export function activeFontSet(): FontFaceSet | null {
  if (typeof document !== 'undefined' && document && document.fonts) return document.fonts;
  if (typeof self !== 'undefined' && self && 'fonts' in self) {
    return (self as unknown as { fonts: FontFaceSet }).fonts;
  }
  return null;
}

export async function preloadGoogleFonts(
  fontNames: Iterable<string | null | undefined>,
  map: Record<string, FontPreloadEntry>,
): Promise<void> {
  const fonts = activeFontSet();
  if (!fonts || typeof FontFace === 'undefined' || typeof fetch === 'undefined') return;

  const seen = new Set<string>();
  const targetFamilies = new Set<string>();
  const cssUrls = new Set<string>();
  // url → lower-cased target families requested from that stylesheet. Built in
  // the main loop so the failure path can attribute a failed fetch to the
  // families this call actually asked for, without re-scanning `map` under the
  // implicit assumption that a map key equals the lower-cased requested name.
  const urlTargets = new Map<string, Set<string>>();
  // Families that could not be made available (stylesheet fetch failed, or every
  // matching FontFace.load() rejected). Surfaced once at the end via a single
  // console.warn so a failed web font no longer falls back to a system face
  // completely silently. The renderer still degrades gracefully — this is
  // diagnostics only, the return contract stays `Promise<void>`.
  const failedFamilies = new Set<string>();

  for (const name of fontNames) {
    if (!name) continue;
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const entry = map[key];
    if (!entry) continue;
    cssUrls.add(entry.url);
    const family = (entry.loadFamily ?? name).toLowerCase();
    targetFamilies.add(family);
    let targets = urlTargets.get(entry.url);
    if (!targets) {
      targets = new Set<string>();
      urlTargets.set(entry.url, targets);
    }
    targets.add(family);
  }
  if (targetFamilies.size === 0) return;

  // 1) Fetch each stylesheet once per JS context and register its FontFace
  //    entries into the active FontFaceSet, keeping references to the faces we
  //    add. Canvas rendering never puts glyphs into the DOM, so registration
  //    alone is not enough — see (2).
  const faceGroups = await withFontCeiling(
    Promise.all(
      [...cssUrls].map((url) => {
        // Cache hit: AWAIT the in-flight registration so concurrent callers join
        // the first fetch rather than racing past it. The stored promise never
        // rejects (see cssRegistrations), so a joined-then-failed registration
        // just yields no faces in step 2.
        const inFlight = cssRegistrations.get(url);
        if (inFlight) return inFlight;
        const registration = (async (): Promise<FontFace[]> => {
          try {
            const res = await fetch(url);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const added: FontFace[] = [];
            for (const f of parseFontFaceRules(await res.text())) {
              const face = new FontFace(f.family, f.src, f.descriptors);
              fonts.add(face);
              added.push(face);
            }
            return added;
          } catch {
            cssRegistrations.delete(url); // free the slot so a later call retries
            for (const family of urlTargets.get(url) ?? []) {
              failedFamilies.add(family);
            }
            return [];
          }
        })();
        cssRegistrations.set(url, registration);
        return registration;
      }),
    ),
  );

  // 2) Force-load the FontFaces we added and AWAIT them all — no timeout race
  //    that could resolve mid-download. `face.load()` is required because
  //    unicode-range gating would otherwise leave the faces `unloaded` and the
  //    first canvas paint would use a system fallback, shifting once a later
  //    interaction re-rasterized. We load the FontFace objects WE created (held
  //    via cssRegistrations) rather than re-selecting them from the set by
  //    family name: `FontFace.family` serializes a multi-word name back WITH
  //    quotes (e.g. `"Nunito Sans"`), so a `family`-string filter silently
  //    matches nothing and the fonts never load — the bug this avoids.
  const addedFaces = (Array.isArray(faceGroups) ? faceGroups : []).flat();
  await withFontCeiling(
    Promise.allSettled(addedFaces.map((f) => f.load())).then((results) => {
      results.forEach((res, i) => {
        if (res.status === 'rejected') {
          failedFamilies.add(addedFaces[i].family.replace(/['"]/g, '').toLowerCase());
        }
      });
      return fonts.ready;
    }),
  );

  if (failedFamilies.size > 0) {
    console.warn(
      `[ooxml] failed to preload web font(s): ${[...failedFamilies].join(', ')}; ` +
        `falling back to system fonts (text may shift or differ).`,
    );
  }
}
