/**
 * DrawingML 3D bevel shading (Phase B) — Canvas 2D, no WebGL.
 *
 * ECMA-376 references:
 *   - §20.1.5.12 `sp3d` (CT_Shape3D): bevelT / bevelB / extrusionH / contour.
 *   - §20.1.5.3  `bevel` (CT_Bevel): bevel width `w`, height `h`, preset `prst`.
 *   - §20.1.10.9 ST_BevelPresetType: the bevel cross-section profiles
 *     (relaxedInset, circle, slope, cross, angle, softRound, convex, coolSlant,
 *     divot, riblet, hardEdge, artDeco).
 *   - §20.1.5.9  `lightRig` (CT_LightRig): rig preset + `dir` (ST_LightRigDirection).
 *
 * ## Method (Canvas 2D, skia-canvas compatible — no WebGL per renderer constraints)
 * A bevel is a raised lip running along the shape's silhouette. We synthesise it
 * as a height field driven by the *distance to the silhouette boundary*:
 *
 *   1. Euclidean distance transform (EDT) of the shape's alpha mask gives, for
 *      every interior pixel, the distance `d` (px) to the nearest boundary pixel.
 *   2. A per-preset 1-D cross-section `height(d/w)` maps that distance to a
 *      surface height in [0,1]: 0 at the rim (d=0), 1 once `d ≥ w` (the flat top).
 *      `circle` is a quarter-circle (rounded lip), `hardEdge` a linear chamfer,
 *      etc. (§20.1.10.9).
 *   3. The surface normal combines two sources, both grounded in the exact EDT
 *      distance (see `computeBevelNormals`): its TILT MAGNITUDE comes from the
 *      profile's local slope at the exact distance, and its in-plane AZIMUTH from
 *      the gradient of a Gaussian-blurred copy of that same distance field. The
 *      distance is exact; the direction is analytically smooth at every raster scale
 *      AND free of the high-curvature "apex plateau" a blurred-coverage gradient
 *      suffers (#418). On the flat interior the tilt is zero → n = (0,0,1); on the
 *      lip it tilts outward toward the rim along the silhouette's outward normal.
 *   4. Lambert diffuse + a weak specular term against the light-rig direction
 *      give a per-pixel brightness factor, applied as a multiply/screen over the
 *      already-painted body.
 *
 * SPEC GAP (documented — see CLAUDE.md "spec fidelity"): ECMA-376 names the
 * bevel presets and the light-rig presets/directions but gives **no numeric
 * cross-section, no light vector, no intensity**. The profile shapes here are
 * the geometrically obvious reading of each preset name; the light vectors and
 * the ambient/diffuse/specular weights are CALIBRATED against the PowerPoint
 * ground truth (sample-11.pdf page 3 — `circle` bevel, `matte`, lightRig threePt
 * dir="t"). The calibration procedure and resulting constants are documented at
 * `THREE_PT_*` and `MATERIAL` below.
 */

export interface Vec3 {
  x: number;
  y: number;
  z: number;
}

/** Surface material reflectivity class we distinguish in Phase B. */
export type BevelMaterial = 'matte' | 'plastic';

/**
 * In-plane revolution (and lat/lon, currently a SPEC GAP) of a light rig, in
 * DEGREES. 1:1 with the parser's Rot3d (§20.1.5.11 CT_SphereCoords).
 */
export interface LightRigRot {
  /** Latitude — rotation about the horizontal (X) axis, degrees. */
  lat: number;
  /** Longitude — rotation about the vertical (Y) axis, degrees. */
  lon: number;
  /** Revolution — in-plane rotation about the view (Z) axis, degrees. */
  rev: number;
}

export interface BevelShadeParams {
  /** Unit KEY light direction (points FROM surface TOWARD the light). */
  light: Vec3;
  material: BevelMaterial;
  /** Ambient term — base brightness factor where no light reaches. */
  ambient: number;
  /** Diffuse weight (Lambert) of the KEY light. */
  diffuse: number;
  /** Specular weight (Blinn-Phong-ish; 0 for pure matte). */
  specular: number;
  /** Specular exponent. */
  shininess: number;
  /**
   * Optional FILL light of a threePt rig (§20.1.5.9). A softer light roughly
   * opposite the key that lifts surfaces facing away from the key out of the pure
   * ambient floor. Omitted → single-key behaviour (the legacy model).
   */
  fillLight?: Vec3;
  /** Diffuse weight of the fill light. 0 / omitted → no fill. */
  fillDiffuse?: number;
}

/**
 * 1-D bevel cross-section. Returns the surface height in [0,1] for a pixel at
 * distance `d` (px) from the silhouette boundary, given the bevel band width `w`
 * (px). h(0)=0 (rim, fully turned-down), h(w)=1 (flat top). Beyond the band the
 * height stays 1 (the flat interior of the shape).
 *
 * SPEC GAP: §20.1.10.9 only names the presets. These profiles are the direct
 * geometric reading of each name (a `circle` lip is a quarter circle, a
 * `hardEdge` lip a straight chamfer, etc.). Presets we don't model individually
 * fall back to `relaxedInset` (the OOXML default bevel) — a smooth smoothstep
 * lip — which is visually close for the gentle bevels these presets describe.
 */
export function bevelHeightProfile(prst: string, w: number): (d: number) => number {
  if (w <= 0) {
    // No bevel band — everything is flat top.
    return () => 1;
  }
  const norm = (d: number) => Math.max(0, Math.min(1, d / w));
  switch (prst) {
    case 'hardEdge': {
      // `hardEdge` is the picture-frame bevel: a SHORT turned-down lip at the very
      // rim, then flat — its defining "hard edge" is the crease near the rim, not a
      // gentle full-width ramp. Concentrating the rise in a rim shelf (height
      // reaches 1 by `HARD_EDGE_SHELF_FRACTION` of the band) makes the lit lip a
      // thin rim with the rest of the band at the flat-top height (dh/dd → 0
      // inward), matching PowerPoint (sample-11.pdf p6: the photo fills the ellipse
      // to a thin rim, NOT a wide darkened band). The rise is a smoothstep, so the
      // profile is C¹ — tangent-flat where the shelf meets the top — and the lip
      // azimuth stays facet-free (no slope discontinuity to kink the normal). SPEC
      // GAP (§20.1.10.9 names the preset only): this is the geometric reading of the
      // name. A plain straight chamfer is the separate `angle` preset below.
      const shelf = HARD_EDGE_SHELF_FRACTION;
      return (d) => {
        const u = Math.min(1, norm(d) / shelf);
        return u * u * (3 - 2 * u); // smoothstep 0→1 within the shelf, flat after
      };
    }
    case 'angle':
    case 'slope':
      // Straight chamfer: linear ramp from rim to top.
      return (d) => norm(d);
    case 'circle':
    case 'convex':
    case 'softRound': {
      // Quarter-circle lip: h = sqrt(1 - (1-t)^2). Convex (bulges up early),
      // tangent to the flat top at t=1 → smooth meeting with the interior.
      return (d) => {
        const t = norm(d);
        const u = 1 - t;
        return Math.sqrt(Math.max(0, 1 - u * u));
      };
    }
    case 'coolSlant':
    case 'divot':
    case 'riblet':
    case 'cross':
    case 'artDeco':
    case 'relaxedInset':
    default: {
      // Smoothstep S-curve (the relaxedInset default and a reasonable stand-in
      // for the remaining presets we don't model individually): tangent-flat at
      // both ends, so the lip eases out of the rim and into the top.
      return (d) => {
        const t = norm(d);
        return t * t * (3 - 2 * t);
      };
    }
  }
}

/**
 * 1-D squared Euclidean distance transform (Felzenszwalb & Huttenlocher 2012,
 * "Distance Transforms of Sampled Functions"). Input `f[i]` is the cost at i
 * (0 at a seed, +∞ elsewhere); output is min_j (f[j] + (i-j)²). Running it once
 * per row then once per column gives the exact 2-D squared EDT. O(n) per line.
 */
export function edt1d(f: ArrayLike<number>): Float64Array {
  const n = f.length;
  const d = new Float64Array(n);
  if (n === 0) return d;
  const v = new Int32Array(n); // locations of parabolas in the lower envelope
  const z = new Float64Array(n + 1); // boundaries between parabolas
  let k = 0;
  v[0] = 0;
  z[0] = -Infinity;
  z[1] = Infinity;
  for (let q = 1; q < n; q++) {
    let s =
      (f[q] + q * q - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
    while (s <= z[k]) {
      k--;
      s = (f[q] + q * q - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
    }
    k++;
    v[k] = q;
    z[k] = s;
    z[k + 1] = Infinity;
  }
  k = 0;
  for (let q = 0; q < n; q++) {
    while (z[k + 1] < q) k++;
    const dq = q - v[k];
    d[q] = dq * dq + f[v[k]];
  }
  return d;
}

/**
 * Box sizes whose 3-fold convolution approximates a Gaussian of std-dev `sigma`
 * (Wells 1986, "Efficient synthesis of Gaussian filters by cascaded uniform
 * filters"). Three equal-radius box blurs already give a good Gaussian; mixing two
 * adjacent odd widths (`wl` and `wl+2`) matches the target variance more exactly.
 * Returns `n` odd box widths. Used by `gaussianBlur` below.
 */
function boxSizesForGaussian(sigma: number, n = 3): number[] {
  if (sigma <= 0) return new Array(n).fill(1);
  const wIdeal = Math.sqrt((12 * sigma * sigma) / n + 1);
  let wl = Math.floor(wIdeal);
  if (wl % 2 === 0) wl--;
  const wu = wl + 2;
  const mIdeal =
    (12 * sigma * sigma - n * wl * wl - 4 * n * wl - 3 * n) / (-4 * wl - 4);
  const m = Math.round(mIdeal);
  const sizes: number[] = [];
  for (let i = 0; i < n; i++) sizes.push(i < m ? wl : wu);
  return sizes;
}

/**
 * One separable box-blur pass with ZERO padding (out-of-bounds treated as 0).
 * `horizontal` selects the axis. Uses a running sum so cost is O(N) independent
 * of the radius.
 *
 * Zero padding (NOT clamp-to-edge) is the correct boundary condition for a COVERAGE
 * field: a silhouette flush with the canvas edge has transparent (coverage 0) just
 * outside the bounds, so the blurred coverage falls off across that edge and ∇C
 * points outward there — exactly as `distanceToEdge` folds in the out-of-bounds
 * transparent region for the EDT. Clamp-to-edge would instead make the field flat
 * at the border, giving a zero gradient and no lip along a clipped edge.
 */
function boxBlurPass(
  src: Float64Array,
  dst: Float64Array,
  w: number,
  h: number,
  r: number,
  horizontal: boolean,
): void {
  const norm = 1 / (2 * r + 1);
  if (horizontal) {
    for (let y = 0; y < h; y++) {
      const row = y * w;
      // Seed the window [−r, r]: indices < 0 contribute 0 (zero padding).
      let acc = 0;
      for (let k = 0; k <= r; k++) if (k < w) acc += src[row + k];
      for (let x = 0; x < w; x++) {
        dst[row + x] = acc * norm;
        const ai = x + r + 1;
        const si = x - r;
        if (ai < w) acc += src[row + ai];
        if (si >= 0) acc -= src[row + si];
      }
    }
  } else {
    for (let x = 0; x < w; x++) {
      let acc = 0;
      for (let k = 0; k <= r; k++) if (k < h) acc += src[k * w + x];
      for (let y = 0; y < h; y++) {
        dst[y * w + x] = acc * norm;
        const ai = y + r + 1;
        const si = y - r;
        if (ai < h) acc += src[ai * w + x];
        if (si >= 0) acc -= src[si * w + x];
      }
    }
  }
}

/**
 * Approximate Gaussian blur of std-dev `sigma` over a scalar field, via three
 * cascaded separable box blurs (zero-padded). O(N) total, independent of sigma.
 *
 * This is the SAME code path in every environment (browser, OffscreenCanvas,
 * node/skia headless), so the bevel azimuth never depends on `ctx.filter='blur'`
 * — whose support and exact kernel differ across canvas backends. The blur is run
 * on the in-memory coverage field, not on a canvas, so it is fully unit-testable.
 */
export function gaussianBlur(
  src: ArrayLike<number>,
  w: number,
  h: number,
  sigma: number,
): Float64Array {
  const a = Float64Array.from(src as ArrayLike<number> & Iterable<number>);
  if (sigma <= 0 || w <= 0 || h <= 0) return a;
  const scratch = new Float64Array(w * h);
  // Each box does a horizontal pass into `scratch`, then a vertical pass back into
  // `a`, so the result always lands in `a` (no buffer swap to track).
  for (const size of boxSizesForGaussian(sigma, 3)) {
    const r = Math.max(1, (size - 1) / 2);
    boxBlurPass(a, scratch, w, h, r, true);
    boxBlurPass(scratch, a, w, h, r, false);
  }
  return a;
}

/**
 * Exact 2-D Euclidean distance (in px) from every pixel to the nearest pixel
 * whose alpha is below `threshold` (i.e. the boundary of the silhouette). The
 * region OUTSIDE the image bounds is also treated as below-threshold, so a
 * silhouette that runs to the canvas edge still gets a finite edge distance
 * (its bevel lip is drawn along that clipped edge). Pixels at or below threshold
 * get distance 0. Returns a W·H Float64Array.
 */
export function distanceToEdge(
  alpha: ArrayLike<number>,
  w: number,
  h: number,
  threshold = 128,
): Float64Array {
  const INF = 1e20;
  const f = new Float64Array(w * h);
  for (let i = 0; i < w * h; i++) {
    f[i] = (alpha[i] ?? 0) >= threshold ? INF : 0;
  }
  // The implicit boundary just outside the image: a pixel at row 0 is one px
  // from the (transparent) row −1, so its column-pass cost is capped at that
  // distance. Seed the borders by treating index −1 / N as zero-cost via the
  // edt1d padding below — done by capping the first/last interior value.
  // (edt1d alone can't see outside the line, so fold the out-of-bounds seed in
  //  as min(current, distance-to-border²) after each separable pass.)
  // Transform along columns, then rows (separable).
  const col = new Float64Array(h);
  for (let x = 0; x < w; x++) {
    for (let y = 0; y < h; y++) col[y] = f[y * w + x];
    const dc = edt1d(col);
    for (let y = 0; y < h; y++) f[y * w + x] = dc[y];
  }
  const row = new Float64Array(w);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) row[x] = f[y * w + x];
    const dr = edt1d(row);
    for (let x = 0; x < w; x++) f[y * w + x] = dr[x];
  }
  // Fold in the implicit transparent region beyond the image edges: row −1 and
  // row h, column −1 and column w are all "outside", so a pixel at (x,y) is at
  // most (y+1) px from the top edge, (h−y) from the bottom, (x+1) from the left,
  // (w−x) from the right. Take the min of the interior EDT and these border
  // distances (squared) so a silhouette flush with the canvas edge still beveled.
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (f[i] === 0) continue; // already a boundary pixel
      const bt = (y + 1) * (y + 1);
      const bb = (h - y) * (h - y);
      const bl = (x + 1) * (x + 1);
      const br = (w - x) * (w - x);
      const border = Math.min(bt, bb, bl, br);
      if (border < f[i]) f[i] = border;
    }
  }
  // f now holds squared distances; sqrt to px.
  for (let i = 0; i < w * h; i++) f[i] = Math.sqrt(f[i]);
  return f;
}

/**
 * Fraction of the band width used as the σ of the Gaussian applied to the EXACT
 * EDT distance field before its gradient is taken for the lip azimuth. A quarter
 * of the band is large enough to wash out the per-pixel EDT quantisation (the raw
 * `∇dist` Voronoi facets) yet small enough to keep the gradient local to one rim.
 * See `computeBevelNormals` for why the azimuth comes from ∇(blurred DISTANCE) and
 * not ∇(blurred coverage) — the latter plateaus at a high-curvature apex (#418).
 */
const DISTANCE_SIGMA_FRACTION = 0.25;

/**
 * Fraction of the band width over which the lip's shading eases back to the flat
 * top at the INNER crease (d → bandPx). Without it the band→top transition is a
 * hard 0/1 step in `bandMask`: every band pixel is fully shaded, the next pixel
 * inward is the untouched face, so a visible brightness discontinuity traces the
 * band's inner boundary. For an ellipse that inner boundary is the exact Euclidean
 * inward offset, which is naturally *flatter than the silhouette at the major-axis
 * tips* (the offset of a high-curvature convex arc recedes faster) — so the hard
 * step reads as the "flat-top + 45° chamfer" cut users reported on slide 6, even
 * though the boundary curve itself is smooth and correct.
 *
 * PowerPoint fillets the bevel/face junction, so the lip fades into the top rather
 * than ending on a crisp edge. We reproduce that by ramping the lip's coverage —
 * and hence its shading weight — from 1 to 0 over the inner `BEVEL_INNER_FEATHER`
 * of the band with a smoothstep (tangent-flat at both ends). The geometry of the
 * band (the EDT iso-distance set) is unchanged; only the *visibility* of the lip
 * tapers, removing the discontinuity that was the visible artifact. `circle` and
 * the smoothstep presets already have slope → 0 at the inner edge so their shading
 * was near-zero there anyway — the feather is (correctly) almost a no-op for them
 * and the PDF-calibrated slide-3 `circle` brightness is preserved. It is the
 * straight-chamfer presets (`angle`/`slope`) and the rim-shelf `hardEdge`, whose
 * shading does NOT vanish at the inner edge, that the feather rescues.
 */
const BEVEL_INNER_FEATHER = 0.35;

/**
 * `hardEdge` rises to full height within this fraction of the band: a short, sharp
 * turned-down shelf at the rim, flat thereafter. See `bevelHeightProfile`.
 */
const HARD_EDGE_SHELF_FRACTION = 0.5;

/**
 * Compute per-pixel surface normals for the bevel lip of a silhouette.
 *
 * @param alpha     W·H alpha mask of the painted body (≥128 = inside).
 * @param w,h       mask dimensions.
 * @param bandPx    bevel band width in px (the EMU `w` scaled to device px).
 * @param prst      bevel preset (ST_BevelPresetType) → cross-section profile.
 * @param heightPx  bevel height in px (the EMU `h` scaled). Controls how steeply
 *                  the lip rises: the height field is `bevelH/bevelW · profile`,
 *                  so a taller bevel tilts the normals more (stronger shading).
 *
 * Returns `normals` (W·H·3 floats, unit vectors, +Z toward viewer) and
 * `bandMask` (W·H of 0/1; 1 where the pixel is inside the shape and within the
 * bevel band, i.e. where shading should be applied).
 *
 * ## Distance-gradient azimuth (issue lineage #410→#413→#415→#416→#417→#418)
 *
 * The lip normal has two parts derived from two INDEPENDENT sources so each is
 * correct on its own terms, BOTH ultimately grounded in the exact EDT distance:
 *
 *   • TILT MAGNITUDE (how steeply the lip rises) — from the cross-section profile's
 *     local slope at the EXACT EDT distance `d`. Distance is the geometrically
 *     meaningful quantity for the height (a `circle` lip at d=band/2 is at a precise
 *     height), and the EDT gives it exactly. The PDF-calibrated slide-3 brightness
 *     curve depends on this, so it is preserved.
 *
 *   • IN-PLANE AZIMUTH (which way the lip faces) — from −∇C where C is the EDT
 *     distance itself, Gaussian-blurred by σ = 0.25·bandPx. The distance field's
 *     level sets ARE the silhouette's offset curves, so −∇C is the true outward
 *     radial normal everywhere; the blur is a smooth (C^∞) convolution, so the
 *     direction rotates continuously and is scale-invariant (σ and the geometry
 *     scale with devScale together).
 *
 * ### Why each prior fix was insufficient (the history this comment preserves)
 *
 *  - #410 (first bevel): azimuth from the 1px finite-difference gradient of the EDT
 *    HEIGHT field. ∇(distance-to-point-set) is piecewise-constant — each band pixel's
 *    nearest boundary sample dominates, so the direction snaps to that sample's
 *    Voronoi-cell direction and chords the lip into facets (worse as the shape grows).
 *  - #413: box-blurred the scalar HEIGHT → smoothed gradient MAGNITUDE, not
 *    direction. Fixed devScale ≤ 2 only.
 *  - #415: tangential low-pass of the normal VECTOR (radius 0.25·band, gated ≥24px).
 *    Fixed devScale 4, REGRESSED at devScale 8 — every band-proportional post-blur
 *    chases a Voronoi cell whose angular width grows with the band.
 *  - #416: moved the azimuth source to −∇(blurred COVERAGE) `G_σ * alpha`. This
 *    killed the Voronoi facets (no discrete cell field left) and the scale sweep
 *    went green — BUT coverage FLAT-TOPS at a high-curvature convex apex: the blur
 *    of a sharply curved tip plateaus, so ∇C there is near-constant over a wide
 *    angular span. The azimuth rotated TOO SLOWLY across the apex → the lit factor
 *    went flat → the "flat horizontal cut" at the top of the slide-6 ellipse. The
 *    ring/ellipse scale-sweep (a JUMP detector) never caught it because a plateau is
 *    a SMOOTH error, not a jump.
 *  - #417: corrected the `hardEdge` profile (rim shelf) and added the inner-crease
 *    feather. Necessary for the band GEOMETRY, but orthogonal to the azimuth — the
 *    flat cut persisted because its cause was the coverage plateau, not the profile.
 *  - #418 (this): source the azimuth from −∇(blurred DISTANCE). Distance does NOT
 *    plateau (it grows linearly inward with the same radial direction at the apex),
 *    so the lit factor follows the true elliptical curve; the blur still removes the
 *    raw-EDT Voronoi facets, keeping the scale-sweep green. One field, both
 *    pathologies gone.
 *
 * The distance blur is O(N) (three cascaded box passes, running-sum, independent of
 * σ), runs once per bevel on the device offscreen — same order as the EDT beside it.
 */
export function computeBevelNormals(
  alpha: ArrayLike<number>,
  w: number,
  h: number,
  bandPx: number,
  prst: string,
  heightPx: number,
): { normals: Float32Array; bandMask: Uint8Array; bandWeight: Float32Array } {
  const normals = new Float32Array(w * h * 3);
  const bandMask = new Uint8Array(w * h);
  // Feathered shading weight in [0,1]: 1 over the outer band, easing to 0 at the
  // inner crease so the lip's brightness change fades into the flat top instead of
  // ending on a hard step. `bandMask` stays a crisp 0/1 membership for callers that
  // need the band footprint; `bandWeight` is what the compositor multiplies the
  // shade excess by. See BEVEL_INNER_FEATHER.
  const bandWeight = new Float32Array(w * h);
  if (w <= 0 || h <= 0) return { normals, bandMask, bandWeight };
  const dist = distanceToEdge(alpha, w, h);
  const profile = bevelHeightProfile(prst, bandPx);
  const inside = (x: number, y: number) => (alpha[y * w + x] ?? 0) >= 128;

  // Height-field aspect scale: the lip rises `heightPx` over `bandPx` of run, so a
  // unit of profile change maps to heightScale·bandPx of surface height (the same
  // mapping the old height field used — preserves the calibrated tilt magnitude).
  const heightScale = bandPx > 0 ? heightPx / bandPx : 0;
  const tiltScale = heightScale * bandPx;

  // ── Azimuth source: gradient of a Gaussian-blurred EXACT DISTANCE field ─────
  // The lip azimuth is the silhouette's outward normal. We read it from ∇C where C
  // is the EDT distance `dist` smoothed by σ = 0.25·band (a fraction of the bevel's
  // own length scale). Two pathologies are avoided BY CONSTRUCTION:
  //   • Voronoi facets (raw ∇dist): the EDT's nearest-seed direction is piecewise
  //     constant, so a 1px finite difference of the bare distance snaps to facet
  //     directions. The blur is a smooth (C^∞) convolution, so ∇C rotates
  //     continuously around the rim — no facets at any raster scale.
  //   • Apex plateau (∇ of blurred COVERAGE, the #416 design): coverage `G_σ*alpha`
  //     FLAT-TOPS at a high-curvature convex apex, so its gradient points near-
  //     constant over a wide angular span there and the lit factor goes flat — the
  //     "flat horizontal cut" users reported at the top of the slide-6 ellipse. The
  //     DISTANCE field does NOT plateau: its level sets are the silhouette's offset
  //     curves, so it keeps growing linearly inward with the SAME radial direction
  //     everywhere, apex included. Blurring distance therefore preserves the true
  //     normal direction while only removing the per-pixel quantisation.
  // σ stays local to a single rim (a larger σ would blur a thin ring's inner and
  // outer rims together and flip the azimuth at the medial axis). The spatial
  // STRUCTURE of the shading (band membership, the inner-crease feather, the tilt
  // magnitude) is taken from the exact `dist` below; only the DIRECTION uses C.
  // See DISTANCE_SIGMA_FRACTION.
  const sigma = Math.max(1, bandPx * DISTANCE_SIGMA_FRACTION);
  const C = gaussianBlur(dist, w, h, sigma);

  // Local profile slope dh/dd at distance d (per px), via a centred 1px difference
  // of the profile. Drives the lip's tilt magnitude; the height stays distance-exact.
  const slopeAt = (d: number): number => {
    const a = profile(Math.max(0, d - 0.5));
    const b = profile(d + 0.5);
    return b - a; // change in profile over 1 px of distance
  };

  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (!inside(x, y)) {
        normals[i * 3 + 2] = 1; // flat ground outside the silhouette
        continue;
      }
      const d = dist[i];
      const inBand = d > 0 && d < bandPx;
      bandMask[i] = inBand ? 1 : 0;
      if (!inBand) {
        normals[i * 3 + 2] = 1; // flat top of the shape
        continue;
      }
      // Inner-crease feather: t01 is the fractional distance across the band
      // (0 at the rim, 1 at the inner edge). Over the inner BEVEL_INNER_FEATHER of
      // the band the shading weight smoothsteps from 1 down to 0 so the lip eases
      // into the flat top with no hard discontinuity (the visible "flat cut").
      const t01 = d / bandPx;
      const fadeStart = 1 - BEVEL_INNER_FEATHER;
      let weight = 1;
      if (t01 > fadeStart) {
        const u = Math.min(1, (t01 - fadeStart) / BEVEL_INNER_FEATHER);
        weight = 1 - u * u * (3 - 2 * u); // 1→0 smoothstep
      }
      bandWeight[i] = weight;
      // Outward azimuth = −∇C/|∇C|. C is the blurred DISTANCE, which DECREASES
      // toward the rim (distance 0 at the boundary), so −∇C points outward along
      // the silhouette's outward normal — the same sign the blurred-coverage field
      // used (coverage also decreases outward), so only the SOURCE of C changed.
      const xl = x > 0 ? x - 1 : x;
      const xr = x < w - 1 ? x + 1 : x;
      const yu = y > 0 ? y - 1 : y;
      const yd = y < h - 1 ? y + 1 : y;
      const gx = (C[y * w + xr] - C[y * w + xl]) / (xr - xl || 1);
      const gy = (C[yd * w + x] - C[yu * w + x]) / (yd - yu || 1);
      const gm = Math.hypot(gx, gy);
      let ax = 0;
      let ay = 0;
      if (gm > 1e-9) {
        ax = -gx / gm;
        ay = -gy / gm;
      }
      // Tilt magnitude from the exact-distance profile slope; azimuth from ∇C.
      const s = slopeAt(d) * tiltScale;
      let nx = s * ax;
      let ny = s * ay;
      let nz = 1;
      const m = Math.hypot(nx, ny, nz) || 1;
      nx /= m;
      ny /= m;
      nz /= m;
      normals[i * 3] = nx;
      normals[i * 3 + 1] = ny;
      normals[i * 3 + 2] = nz;
    }
  }
  return { normals, bandMask, bandWeight };
}

// ── Light rig ───────────────────────────────────────────────────────────────
//
// SPEC GAP: §20.1.5.9 enumerates the rig presets and the 8 `dir` octants but
// gives no light vector. We map `dir` to an azimuth on the screen plane and lift
// the KEY light OUT of the screen toward the viewer by a fixed elevation so the
// lip catches light on the side facing `dir`. CALIBRATION (sample-11.pdf p3): the
// card's bevel rim is brightest along its upper edges and dims toward the lower
// edges, with `dir="t"` → light from straight above. An elevation that puts the
// light ~35° above the screen plane reproduces that gradient (a steeper, more
// overhead light flattens the contrast; a shallower one over-darkens the lower
// rim vs the PDF).
//
// `<a:rot lat lon rev>` (§20.1.5.11 CT_SphereCoords, on the lightRig):
//   • rev — IMPLEMENTED. The in-plane revolution about the view (+Z) axis. It
//     rotates the screen-plane azimuth by the standard 2-D rotation matrix (screen
//     +Y is DOWN). GROUND TRUTH: sample-11 slide-6 carries dir="t" rev=320°; the
//     PDF shows the ellipse's LEFT shoulder markedly brighter than the right
//     (+21.6% vs +6.3% over the face), i.e. the key azimuth points UPPER-LEFT.
//     R(320°)·(0,−1) = (−sin320°, −cos320°) = (+0.643, −0.766)?  No — with the
//     screen-down sign the up-vector (0,−1) maps to (−sin θ, −cos θ); at 320° that
//     is (+0.643, −0.766) which is upper-RIGHT, so we use the NEGATED screen-x sign
//     that puts it upper-left and matches the PDF. The exact mapping is in
//     `rotateAzimuth` below with the sign pinned by the slide-6 measurement.
//   • lat / lon — SPEC GAP (NOT IMPLEMENTED). Latitude/longitude would tilt the
//     rig OUT of the view plane (changing the key's elevation/horizon). No
//     calibration sample exercises a non-zero lat/lon on a lightRig (slide-6 has
//     lat=lon=0; slide-3 has no lightRig rot at all), so honouring them would be
//     an un-grounded guess. They are parsed and carried but ignored here; revisit
//     when a sample with lat/lon≠0 and a PDF to calibrate against appears.
//
// THREE-POINT FILL: a threePt rig is key + fill + back (§20.1.5.9). The fill is a
// softer light roughly OPPOSITE the key that lifts the surfaces backing the key
// out of the pure ambient floor. The PDF (sample-11 p6) shows the ellipse's BOTTOM
// lip — whose outward normal fully backs the upper-left key — still ABOVE the face
// (+3.5%), not at the ambient floor a lone key would give (≈ −29%). The fill term
// (see FILL_*) reproduces that lift. Calibration of all weights is documented at
// MATERIAL / FILL_* below against BOTH p3 (circle, no rot) and p6 (hardEdge, rev).

/** Key-light elevation above the screen plane, radians. Calibrated vs p3 (≈35°). */
const LIGHT_ELEVATION = (35 * Math.PI) / 180;

/**
 * Fill-light elevation above the screen plane, radians. Much shallower than the
 * key so the fill grazes the surfaces facing away from the key (raising them out of
 * deep shadow) while barely touching the flat top face — keeping it from inflating
 * the face reference and flattening the key-lit lip. CALIBRATED (with FILL_DIFFUSE)
 * by a real-Chrome sweep over the sample-11 p6 columns at devScale 2: every
 * elevation above ≈12° dimmed the top/left lit lip (the fill leaked onto the +Z
 * face) without further lifting the shadow side, so the grazing 12° is the floor of
 * the swept range and the best fit.
 */
const FILL_ELEVATION = (12 * Math.PI) / 180;

/** Screen-plane azimuth (x,y) per ST_LightRigDirection octant. y is screen-down. */
const DIR_AZIMUTH: Record<string, { x: number; y: number }> = {
  t: { x: 0, y: -1 },
  b: { x: 0, y: 1 },
  l: { x: -1, y: 0 },
  r: { x: 1, y: 0 },
  tl: { x: -1, y: -1 },
  tr: { x: 1, y: -1 },
  bl: { x: -1, y: 1 },
  br: { x: 1, y: 1 },
};

/**
 * Rotate a screen-plane azimuth (screen +Y down) by `revDeg` degrees about the
 * view (+Z) axis (§20.1.5.9 `<a:rot rev>`), using the standard 2-D rotation matrix
 *   x' = x·cosθ − y·sinθ
 *   y' = x·sinθ + y·cosθ
 * The sign is PINNED by the sample-11 slide-6 measurement: dir="t" (the up-vector
 * (0,−1)) at rev=320° maps to (−sin320°, −cos320°) = (−0.643, −0.766) — UPPER-LEFT,
 * which is exactly the side the PDF shows brightest (the +21.6% left shoulder vs the
 * +6.3% right). The unit test `a:rot rev rotates the screen-plane azimuth` asserts
 * this sign so it cannot silently flip.
 */
function rotateAzimuth(ax: number, ay: number, revDeg: number): { x: number; y: number } {
  const th = (revDeg * Math.PI) / 180;
  const c = Math.cos(th);
  const s = Math.sin(th);
  return { x: ax * c - ay * s, y: ax * s + ay * c };
}

/**
 * Unit KEY light direction (FROM surface TOWARD light) for a light rig. `dir` sets
 * the screen-plane azimuth; `rot.rev` rotates that azimuth in-plane; the light is
 * lifted toward the viewer (+Z) by LIGHT_ELEVATION. See the SPEC GAP / calibration
 * note above. `rot.lat`/`rot.lon` are a documented SPEC GAP (ignored).
 */
export function lightDirFromRig(_rig: string, dir: string, rot?: LightRigRot): Vec3 {
  let az = DIR_AZIMUTH[dir] ?? DIR_AZIMUTH.t;
  if (rot && rot.rev) az = rotateAzimuth(az.x, az.y, rot.rev);
  return liftAzimuth(az.x, az.y, LIGHT_ELEVATION);
}

/**
 * Build the FILL light of a threePt rig from the key's screen azimuth: it sits
 * roughly OPPOSITE the key on the screen plane, lifted toward the viewer by the
 * (shallower) FILL_ELEVATION. See the THREE-POINT FILL note above.
 */
export function fillDirFromKey(key: Vec3): Vec3 {
  // Project the key onto the screen plane and negate to get the fill azimuth.
  const plen = Math.hypot(key.x, key.y) || 1;
  return liftAzimuth(-key.x / plen, -key.y / plen, FILL_ELEVATION);
}

/** Lift a (normalised-or-not) screen azimuth to a unit 3-D light at `elev` rad. */
function liftAzimuth(ax: number, ay: number, elev: number): Vec3 {
  const planeLen = Math.hypot(ax, ay) || 1;
  const cosE = Math.cos(elev);
  const sinE = Math.sin(elev);
  const x = (ax / planeLen) * cosE;
  const y = (ay / planeLen) * cosE;
  const z = sinE;
  const m = Math.hypot(x, y, z) || 1;
  return { x: x / m, y: y / m, z: z / m };
}

/**
 * Per-material ambient/diffuse/specular weights for Phase B. We only distinguish
 * matte vs plastic (the brief's stated scope). `warmMatte`/`matte`/`flat` →
 * matte (no specular); everything glossier (`plastic`/`metal`/`...`) → plastic
 * with a modest specular lobe.
 *
 * SPEC GAP / CALIBRATION: ECMA-376 names ~15 ST_PresetMaterialType values with
 * no reflectance numbers. The matte weights are calibrated against sample-11 p3
 * (`prstMaterial="matte"`, lightRig threePt): the bevel rim there shows a clear
 * but soft light/dark gradient with no hard hot-spot, i.e. a high ambient floor
 * (the rig's fill+back lights) plus a moderate Lambert term and no specular.
 * ambient 0.62 / diffuse 0.45 reproduces the rim's lightest band ≈ +28% and the
 * darkest ≈ −12% relative to the flat face, matching the PDF rim sampling.
 */
/**
 * Screen-blend weight for the lit lip's highlight, per unit of shade excess over
 * the flat face. A chamfer facing the key light reads as a bright edge even on a
 * dark texture, which a pure multiply can't produce; the screen term lifts it
 * toward white. CALIBRATED vs sample-11.pdf p3: the lit top rim there sits a
 * clear step above the face but well short of white. With the matte weights the
 * lit lip's normalised factor peaks near 1.15, so a gain of 2.0 yields a screen
 * weight ≈0.3 — a visible but soft highlight matching the PDF (a higher gain
 * blows the rim to pure white, which p3 does not show).
 */
const SCREEN_GAIN = 2.0;

const MATERIAL: Record<BevelMaterial, { ambient: number; diffuse: number; specular: number; shininess: number }> = {
  matte: { ambient: 0.62, diffuse: 0.45, specular: 0.0, shininess: 8 },
  plastic: { ambient: 0.55, diffuse: 0.5, specular: 0.35, shininess: 22 },
};

/**
 * Diffuse weight of the threePt FILL light (§20.1.5.9), as a fraction of the key
 * diffuse. The fill is the softer light opposite the key that lifts the surfaces
 * backing the key out of the pure ambient floor.
 *
 * SPEC GAP / CALIBRATION: ECMA-376 gives no fill intensity. Pinned by a real-Chrome
 * sweep (devScale 2) over the sample-11.pdf p6 columns — the canonical threePt rig
 * with `hardEdge`, dir="t" rev=320° → upper-left key. PDF band/face targets and the
 * after-fit (amb 0.62, keyDiff 0.45, FILL_DIFFUSE 0.8, FILL_ELEVATION 12°):
 *
 *   point          PDF target   baseline (no fill)   after fill
 *   p6 top         +9.9%        +9.4%                +7.8%
 *   p6 left  rim   +21.6%       +17.2%               +16.0%   (rev asymmetry intact)
 *   p6 right rim   +6.3%        −6.5%                ±0%      (lifted to the face)
 *   p6 bottom      +5.3%        −18.6%               ±0%      (lifted out of shadow)
 *   p3 top   rim   +28.1%       +28.3%               +28.3%   (no regression)
 *
 * The fill at 0.8·keyDiff with the grazing 12° elevation lifts the key-backing
 * bottom/right rims from deep shadow (−18.6% / −6.5%) up to the face level and keeps
 * the rev-driven left≫right asymmetry — the defining feature of this slide. It does
 * NOT reach the PDF's modest +5–6% POSITIVE lift on those rims: the faceFactor-
 * normalised multiply/screen compositor (shadow side darkens, lit side screens)
 * cannot push a key-backing lip ABOVE the face without a stronger fill, and a
 * stronger fill inflates the face reference and flattens the key-lit top/left below
 * their PDF targets (verified across the sweep). The ≈5–6% residual on bottom/right
 * is the accepted trade-off; closing it fully would need a different compositing
 * model (a separate additive back-light pass) — out of scope here and documented as
 * a known limitation rather than papered over with a per-rim hack.
 */
const FILL_DIFFUSE = 0.8;

/** Map a ST_PresetMaterialType name to the matte/plastic class. */
export function materialClass(prstMaterial: string | undefined): BevelMaterial {
  switch (prstMaterial) {
    case 'plastic':
    case 'metal':
    case 'clear':
    case 'softEdge':
    case 'shiny':
    case 'softmetal':
      return 'plastic';
    case 'matte':
    case 'warmMatte':
    case 'flat':
    case 'dkEdge':
    case 'powder':
    default:
      return 'matte';
  }
}

/**
 * Build shade params for a material + KEY light. When `fill` is true (the threePt
 * rig default) a fill light opposite the key is added at FILL_DIFFUSE·keyDiffuse,
 * lifting the surfaces backing the key out of the pure ambient floor (see the
 * THREE-POINT FILL note and FILL_DIFFUSE). Pass `fill=false` for a single-key rig.
 */
export function shadeParamsFor(material: BevelMaterial, light: Vec3, fill = true): BevelShadeParams {
  const m = MATERIAL[material];
  const params: BevelShadeParams = {
    light,
    material,
    ambient: m.ambient,
    diffuse: m.diffuse,
    specular: m.specular,
    shininess: m.shininess,
  };
  if (fill) {
    params.fillLight = fillDirFromKey(light);
    params.fillDiffuse = params.diffuse * FILL_DIFFUSE;
  }
  return params;
}

/**
 * Brightness multiplier for a surface normal under the light. Lambert diffuse from
 * the KEY light, an optional softer Lambert term from the threePt FILL light (lifts
 * surfaces backing the key), and (for plastic) a Blinn-Phong specular against the
 * half-vector with the view direction (0,0,1). Returns a factor ≥ 0 to multiply the
 * body colour by (1.0 = unchanged; >1 brightens via the screen-side blend the
 * caller applies).
 */
export function shadePixel(n: Vec3, p: BevelShadeParams): number {
  const ndotl = n.x * p.light.x + n.y * p.light.y + n.z * p.light.z;
  const diff = p.diffuse * Math.max(0, ndotl);
  let fill = 0;
  if (p.fillLight && p.fillDiffuse) {
    const ndotf = n.x * p.fillLight.x + n.y * p.fillLight.y + n.z * p.fillLight.z;
    fill = p.fillDiffuse * Math.max(0, ndotf);
  }
  let spec = 0;
  if (p.specular > 0) {
    // Half-vector between light and view (view = +Z).
    const hx = p.light.x;
    const hy = p.light.y;
    const hz = p.light.z + 1;
    const hm = Math.hypot(hx, hy, hz) || 1;
    const ndoth = (n.x * hx + n.y * hy + n.z * hz) / hm;
    spec = p.specular * Math.pow(Math.max(0, ndoth), p.shininess);
  }
  return Math.max(0, p.ambient + diff + fill + spec);
}

/** A minimal 2D context surface the bevel compositor needs. */
export interface BevelCtx {
  canvas: { width: number; height: number };
  getImageData(sx: number, sy: number, sw: number, sh: number): ImageData;
  putImageData(data: ImageData, dx: number, dy: number): void;
}

/**
 * A sub-rectangle of the body bitmap that the shading actually touches (perf: A3).
 * The distance transform + normal loop are O(w·h); when the painted silhouette
 * occupies only part of a large offscreen, restricting the `getImageData` window
 * and the loop to `bbox ⊕ effect-reach` (clamped to the canvas) avoids running
 * the whole thing over transparent margin.
 *
 * CORRECTNESS: `distanceToEdge` treats the area OUTSIDE the passed alpha plane as
 * transparent (below-threshold), so a silhouette flush with the plane edge still
 * gets a finite edge distance. Passing a sub-window therefore gives IDENTICAL
 * results to the full canvas **provided the window edge lies in transparent
 * territory** — i.e. the window must contain the silhouette grown by the effect's
 * spatial reach (the bevel band, or the extrusion offset). Callers size the
 * region that way; when omitted the whole canvas is used (unchanged behaviour).
 */
export interface BevelRegion {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Clamp a requested region to the canvas bounds; returns null if it collapses. */
function clampRegion(
  region: BevelRegion | undefined,
  canvasW: number,
  canvasH: number,
): { x: number; y: number; w: number; h: number } {
  if (!region) return { x: 0, y: 0, w: canvasW, h: canvasH };
  const x0 = Math.max(0, Math.floor(region.x));
  const y0 = Math.max(0, Math.floor(region.y));
  const x1 = Math.min(canvasW, Math.ceil(region.x + region.w));
  const y1 = Math.min(canvasH, Math.ceil(region.y + region.h));
  return { x: x0, y: y0, w: Math.max(0, x1 - x0), h: Math.max(0, y1 - y0) };
}

export interface BevelInput {
  /** Bevel band width in device px (EMU `w` × scale × devScale). */
  widthPx: number;
  /** Bevel height in device px (EMU `h` × scale × devScale). */
  heightPx: number;
  /** Bevel preset (ST_BevelPresetType). */
  prst: string;
  /** Surface material class. */
  material: BevelMaterial;
  /** Light direction (FROM surface TOWARD light), already built from the rig. */
  light: Vec3;
  /**
   * When true the bevel runs along the BOTTOM face (bevelB): the height field is
   * the same but the light hits the underside, so we flip the normal's screen-Y
   * and Z handling by treating the lip as facing away from the viewer. In Phase
   * B we approximate bevelB as bevelT with an inverted vertical light response
   * (the lip on a back face catches the opposite side of the key light).
   */
  bottom?: boolean;
}

/**
 * Bake bevel shading into an already-painted body bitmap held in `ctx`.
 *
 * Reads the body's pixels, computes the lip normals from its alpha silhouette,
 * and multiplies the RGB of each band pixel by the per-pixel light factor
 * (factors >1 brighten the lit edge, <1 darken the shadowed edge). The shading
 * is baked into the same bitmap so that, when the caller later warps the bitmap
 * through the scene3d camera, the bevel rides the projection automatically.
 *
 * No-op when the band is sub-pixel (nothing visible to shade).
 */
export function applyBevelShading(ctx: BevelCtx, input: BevelInput, region?: BevelRegion): void {
  const canvasW = ctx.canvas.width;
  const canvasH = ctx.canvas.height;
  if (canvasW <= 0 || canvasH <= 0) return;
  const bandPx = input.widthPx;
  if (bandPx < 0.75) return; // sub-pixel lip — skip.

  // Restrict the distance-transform + normal loop to the region the lip can
  // occupy (perf: A3). Equivalent to the full canvas because the region contains
  // the silhouette ⊕ bandPx, so every band pixel (dist < bandPx) is inside it and
  // the region edge sits in transparent territory — matching `distanceToEdge`'s
  // out-of-bounds-is-transparent boundary. See BevelRegion.
  const { x: rx, y: ry, w, h } = clampRegion(region, canvasW, canvasH);
  if (w <= 0 || h <= 0) return;

  const img = ctx.getImageData(rx, ry, w, h);
  const px = img.data;
  // Alpha plane for the distance transform.
  const alpha = new Uint8ClampedArray(w * h);
  for (let i = 0; i < w * h; i++) alpha[i] = px[i * 4 + 3];

  const { bandMask, bandWeight, normals } = computeBevelNormals(
    alpha,
    w,
    h,
    bandPx,
    input.prst,
    input.heightPx,
  );
  const params = shadeParamsFor(input.material, input.light);

  // Reference brightness of the FLAT top face (normal = +Z). The flat interior
  // of the shape is left untouched (multiplier 1.0), so we divide every band
  // pixel's absolute shade by this reference: the lip's lit side then comes out
  // > 1 (brighter than the face) and the shadowed side < 1, reproducing the
  // raised-relief look. Without this normalisation the whole band would read as
  // a *darker* outline because the flat-face Lambert term (ambient + diffuse·L.z)
  // is < 1 yet we never applied it to the face.
  const faceFactor = shadePixel({ x: 0, y: 0, z: 1 }, params) || 1;

  for (let i = 0; i < w * h; i++) {
    if (bandMask[i] === 0) continue;
    const wt = bandWeight[i];
    if (wt <= 0) continue; // fully faded inner crease — leave the face untouched.
    let nx = normals[i * 3];
    let ny = normals[i * 3 + 1];
    const nz = normals[i * 3 + 2];
    // bevelB: the lip belongs to the back face, lit from the opposite vertical
    // sense — mirror the in-plane normal so the light/shadow sides swap.
    if (input.bottom) {
      nx = -nx;
      ny = -ny;
    }
    // Per-pixel shade factor, then ease it back toward 1.0 (no change) by the
    // inner-crease feather weight so the lip fades into the flat top with no hard
    // step. wt=1 keeps the full lip near the rim; wt→0 at the inner edge.
    const fRaw = shadePixel({ x: nx, y: ny, z: nz }, params) / faceFactor;
    const f = 1 + (fRaw - 1) * wt;
    const o = i * 4;
    if (f >= 1) {
      // Lit lip: a chamfer facing the light reads as a bright highlight even
      // over a dark texture (in PowerPoint the bevel is a separate lit surface,
      // not a darkened copy of the front-face image). Multiply alone can only
      // brighten a dark pixel a little, so we ALSO screen-blend toward white by
      // the excess over the face brightness — the multiply/screen composite the
      // brief calls for. `hi` is the screen weight, capped so a strong lit lip
      // approaches (but doesn't blow past) white.
      const hi = Math.min(1, (f - 1) * SCREEN_GAIN);
      for (let c = 0; c < 3; c++) {
        const base = Math.min(255, px[o + c] * f);
        px[o + c] = base + (255 - base) * hi;
      }
    } else {
      // Shadowed lip: straight multiply darkening.
      px[o] = Math.max(0, px[o] * f);
      px[o + 1] = Math.max(0, px[o + 1] * f);
      px[o + 2] = Math.max(0, px[o + 2] * f);
    }
  }
  ctx.putImageData(img, rx, ry);
}

export interface ExtrusionInput {
  /** Screen-space depth offset of the back face (device px), from the camera. */
  offsetX: number;
  offsetY: number;
  /** Side-wall colour as [r,g,b] (extrusionClr, or a darkened body colour). */
  rgb: [number, number, number];
}

/**
 * Bake an extrusion side-wall band into a painted body bitmap (§20.1.5.12
 * `extrusionH`). The front face sits at z=0 and the back face is pushed to
 * −depth; the screen projection of that push is `(offsetX, offsetY)` device px
 * (from `computeDepthOffset`). The visible side wall is the region swept between
 * the front silhouette and the offset back silhouette.
 *
 * APPROXIMATION (Phase B, documented): we sweep the front silhouette by the
 * single centre depth-offset vector and fill the newly-uncovered pixels (those
 * transparent in the body but covered by an offset copy of the silhouette) with
 * the side-wall colour, UNDER the front face. This is exact for a parallel sweep
 * and a good approximation for small extrusions / gentle tilts; it does NOT
 * model per-silhouette-point frustum divergence or self-occlusion of a deep
 * extrusion under a strong perspective. When the offset is sub-pixel (face-on
 * camera) there is no visible wall and this is a no-op.
 */
export function applyExtrusion(ctx: BevelCtx, input: ExtrusionInput, region?: BevelRegion): void {
  const canvasW = ctx.canvas.width;
  const canvasH = ctx.canvas.height;
  if (canvasW <= 0 || canvasH <= 0) return;
  const dx = input.offsetX;
  const dy = input.offsetY;
  const len = Math.hypot(dx, dy);
  if (len < 0.75) return; // sub-pixel — no visible side wall.

  // Restrict to the region the wall can occupy (perf: A3). A wall pixel is a
  // transparent pixel within `len` px of the front silhouette along the depth
  // vector, so it lies inside the silhouette bbox ⊕ |offset|. When the region
  // contains that (callers size it so), the result is identical to the full
  // canvas: the walk-back always moves TOWARD the silhouette (into the region),
  // so no in-region wall pixel needs an out-of-region sample. See BevelRegion.
  const { x: rx, y: ry, w, h } = clampRegion(region, canvasW, canvasH);
  if (w <= 0 || h <= 0) return;

  const img = ctx.getImageData(rx, ry, w, h);
  const px = img.data;
  // Source alpha snapshot (the front face silhouette), region-local.
  const srcA = new Uint8ClampedArray(w * h);
  for (let i = 0; i < w * h; i++) srcA[i] = px[i * 4 + 3];

  const steps = Math.max(1, Math.ceil(len));
  const [r, g, b] = input.rgb;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (srcA[i] >= 128) continue; // covered by the front face — wall hidden.
      // Walk back toward the front face along the depth vector; if any sample
      // along the sweep lands on the front silhouette, this pixel is side wall.
      // Coords are region-local (the walk stays in-region — see the note above).
      let onWall = false;
      for (let s = 1; s <= steps; s++) {
        const t = s / steps;
        const sx = Math.round(x - dx * t);
        const sy = Math.round(y - dy * t);
        if (sx < 0 || sy < 0 || sx >= w || sy >= h) continue;
        if (srcA[sy * w + sx] >= 128) {
          onWall = true;
          break;
        }
      }
      if (!onWall) continue;
      const o = i * 4;
      px[o] = r;
      px[o + 1] = g;
      px[o + 2] = b;
      px[o + 3] = 255;
    }
  }
  ctx.putImageData(img, rx, ry);
}
