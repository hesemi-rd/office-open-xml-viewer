/**
 * Google Fonts preload utility shared by docx / pptx / xlsx viewers.
 *
 * The contract is intentionally narrow: callers pass the set of font-family
 * names they want available, plus a static map from a lower-cased key to a
 * Google Fonts CSS URL (and optionally an alternate FontFaceSet family name
 * for Office substitutes such as Calibri → Carlito). Names without a map
 * entry are skipped (the renderer falls back to the system font).
 *
 * Font load is forced via `face.load()` rather than `document.fonts.load()`
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
 * Hard ceiling so a wedged network (the stylesheet `<link>` never firing
 * load/error, or a `FontFace.load()` that never settles) cannot hang the
 * caller forever. This is a SAFETY NET, not the normal exit: on a reachable
 * network every awaited promise settles well within it, so first paint is
 * deterministic. It is intentionally generous — the previous 3 s timeout RACED
 * the font loads and could resolve while faces were still downloading, which is
 * exactly the cold-cache flicker this function must prevent.
 */
const HARD_CEILING_MS = 15000;

/** Resolve when the stylesheet `<link>` has finished loading (load or error),
 *  so the FontFace entries it defines are registered in `document.fonts`.
 *  Resolves immediately for a `<link>` that already loaded. */
function linkLoaded(link: HTMLLinkElement): Promise<void> {
  // A stylesheet that is already applied exposes its `sheet`; treat as loaded.
  if (link.sheet) return Promise.resolve();
  return new Promise<void>((resolve) => {
    const done = () => {
      link.removeEventListener('load', done);
      link.removeEventListener('error', done);
      resolve();
    };
    link.addEventListener('load', done);
    link.addEventListener('error', done);
  });
}

function withCeiling<T>(p: Promise<T>): Promise<T | void> {
  return Promise.race([
    p,
    new Promise<void>((resolve) => setTimeout(resolve, HARD_CEILING_MS)),
  ]);
}

export async function preloadGoogleFonts(
  fontNames: Iterable<string | null | undefined>,
  map: Record<string, FontPreloadEntry>,
): Promise<void> {
  if (typeof document === 'undefined') return;

  const seen = new Set<string>();
  const targetFamilies = new Set<string>();
  const linkPromises: Promise<void>[] = [];

  for (const name of fontNames) {
    if (!name) continue;
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const entry = map[key];
    if (!entry) continue;

    let link = document.querySelector<HTMLLinkElement>(`link[href="${entry.url}"]`);
    if (!link) {
      try {
        link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = entry.url;
        document.head.appendChild(link);
      } catch {
        // Network or DOM error — silently skip; renderer falls back to system.
        link = null;
      }
    }
    // Wait for the stylesheet to actually parse so its FontFace entries are
    // registered before we enumerate `document.fonts`. The previous version
    // polled for ~500 ms and gave up loading NOTHING on a slow/cold network,
    // so the next (warm) reload rendered with different fonts — the flicker.
    if (link) linkPromises.push(linkLoaded(link));

    targetFamilies.add((entry.loadFamily ?? name).toLowerCase());
  }

  if (targetFamilies.size === 0) return;

  // 1) Wait for every stylesheet to register its FontFace entries.
  await withCeiling(Promise.allSettled(linkPromises));

  // 2) Force-load every matching FontFace and AWAIT them all — no timeout race
  //    that could resolve mid-download. `face.load()` is required because
  //    canvas rendering never puts glyphs in the DOM, so unicode-range gating
  //    would otherwise leave the faces `unloaded` and the first paint would use
  //    a system fallback, shifting once a later interaction re-rasterized.
  const registered = [...document.fonts].filter((f) =>
    targetFamilies.has(f.family.toLowerCase()),
  );
  await withCeiling(
    Promise.allSettled(registered.map((f) => f.load())).then(() =>
      document.fonts.ready,
    ),
  );
}
