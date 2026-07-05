/**
 * Refcounted FontFace registry shared by the embedded-font loader
 * ({@link ./embedded.ts}) and the Google-Fonts preloader ({@link ./preload.ts}).
 *
 * `document.fonts` (the FontFaceSet) is a process-global singleton, so opening
 * the same font in two documents — or the same document twice in an SPA — would
 * otherwise add a second, byte- (or url-) identical `FontFace` every time and
 * leak them all (nothing ever removes them). Both loaders share this dedup +
 * refcount so a face is added to the set once, shared by every holder, and
 * removed from the set only when the LAST holder releases it (e.g. from every
 * open document's `destroy()`).
 *
 * The registry is deliberately format-agnostic: callers own how they compute a
 * face's signature (embedded fonts hash the de-obfuscated bytes; Google Fonts
 * key on the CSS url + family + descriptors) and how they build the `FontFace`.
 * The registry owns only the shared concern — *this exact face is referenced by
 * N holders; delete it from its set at N = 0* — so neither loader duplicates the
 * refcount / last-release / double-release-safety logic.
 */

/** One shared, refcounted FontFace registration. */
interface FontRegistration {
  face: FontFace;
  /** The FontFaceSet the face was added to (`document.fonts` / `self.fonts`).
   *  Held so release can `delete()` from the SAME set the retain added to, and
   *  so a stale signature colliding across two different sets never mixes. */
  set: FontFaceSet;
  refs: number;
}

/** signature → the single shared registration. Module-global to mirror the
 *  process-global FontFaceSet: two callers computing the same signature in the
 *  same set share one `FontFace`. */
const registry = new Map<string, FontRegistration>();

/** Test hook — clears the shared refcount registry (does NOT touch any
 *  FontFaceSet; tests install a fresh fake set per case). */
export function _resetFontRegistryForTests(): void {
  registry.clear();
}

/** Result of {@link retainFace}: the shared `FontFace` this caller now holds a
 *  reference to, and whether THIS call created it (so the caller knows it must
 *  force-`load()` the face — an already-shared face was loaded by its first
 *  holder). */
export interface RetainResult {
  face: FontFace;
  /** `true` when this retain created + added the face (first holder); `false`
   *  when it reused an existing shared registration (refs bumped). */
  isNew: boolean;
}

/**
 * Retain a shared FontFace for `sig` in `set`, bumping its refcount.
 *
 * - First holder of `sig` (in this set): `create()` builds the `FontFace`, it is
 *   added to the set, and `{ face, isNew: true }` is returned — the caller must
 *   force-`load()` it.
 * - A later holder of the same `sig`: the existing shared face is reused, its
 *   refcount bumped, and `{ face, isNew: false }` returned — already loaded by
 *   the first holder, so the caller skips the load.
 *
 * A signature whose registration lives in a DIFFERENT set (e.g. a stale
 * cross-context collision) is treated as absent: a fresh registration replaces
 * it, so a face is never handed back from a set the caller is not adding to.
 *
 * `create()` must both construct the `FontFace` AND add it to `set` (the two are
 * inseparable — the registry cannot know a loader's `set.add` semantics), then
 * return the face. It runs ONLY on the first-holder path.
 */
export function retainFace(sig: string, set: FontFaceSet, create: () => FontFace): RetainResult {
  const existing = registry.get(sig);
  if (existing && existing.set === set) {
    existing.refs++;
    return { face: existing.face, isNew: false };
  }
  const face = create();
  registry.set(sig, { face, set, refs: 1 });
  return { face, isNew: true };
}

/**
 * Release a set of shared `FontFace` objects (as returned by the loaders'
 * retain paths). Each face's refcount is decremented; the face is removed from
 * its FontFaceSet only when the last holder releases it, so a font shared by two
 * open documents survives until both are destroyed. Safe to call with faces the
 * registry does not know (no-op).
 *
 * **Idempotent / double-release safe (refs are never over-decremented).** Two
 * independent guards protect a font another document is still using from being
 * evicted by a stray double-release:
 *
 * - *Within one call*: the same `FontFace` appearing twice in `faces` (a caller
 *   passing a list with duplicates) is decremented AT MOST ONCE — a per-call
 *   `seen` set skips repeats. Without this, `release([F, F])` would drop refs by
 *   2 and could delete `F` while a second holder still references it.
 * - *Across calls*: once a face's refcount reaches 0 its registry entry is
 *   removed, so a later call that passes the same (now-unregistered) face finds
 *   no entry and is a no-op — it can never push another registration's refs
 *   negative.
 */
export function releaseFaces(faces: Iterable<FontFace>): void {
  // Guard against the same face appearing more than once in THIS call so a
  // duplicate cannot decrement a shared registration's refcount twice.
  const seen = new Set<FontFace>();
  for (const face of faces) {
    if (seen.has(face)) continue;
    seen.add(face);
    // Find the registration for this face (identity match). The registry is
    // small (one entry per distinct face), so a linear scan is fine. A face that
    // was already fully released has no entry → this loop finds nothing and the
    // release is a no-op (cross-call idempotency).
    for (const [sig, reg] of registry) {
      if (reg.face !== face) continue;
      reg.refs--;
      if (reg.refs <= 0) {
        try {
          reg.set.delete(face);
        } catch {
          /* a set without delete() (older shim / mock): drop the entry anyway */
        }
        registry.delete(sig);
      }
      break;
    }
  }
}
