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
 *   3. The surface normal combines two INDEPENDENT sources (the scale-invariant
 *      redesign — see `computeBevelNormals`): its TILT MAGNITUDE comes from the
 *      profile's local slope at the exact EDT distance, and its in-plane AZIMUTH
 *      comes from the gradient of a Gaussian-blurred coverage field. Distance is
 *      exact; direction is analytically smooth at every raster scale. On the flat
 *      interior the tilt is zero → n = (0,0,1); on the lip it tilts outward toward
 *      the rim along the silhouette's macroscopic normal.
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

export interface BevelShadeParams {
  /** Unit light direction (points FROM surface TOWARD the light). */
  light: Vec3;
  material: BevelMaterial;
  /** Ambient term — base brightness factor where no light reaches. */
  ambient: number;
  /** Diffuse weight (Lambert). */
  diffuse: number;
  /** Specular weight (Blinn-Phong-ish; 0 for pure matte). */
  specular: number;
  /** Specular exponent. */
  shininess: number;
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

/** Fraction of the band width used as the coverage-blur σ. See COVERAGE_SIGMA. */
const COVERAGE_SIGMA_FRACTION = 0.25;

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
 * ## Scale-invariant azimuth (this redesign — issue lineage #410→#413→#415→here)
 *
 * The lip normal has two parts that we derive from two INDEPENDENT, decoupled
 * sources so each is correct on its own terms:
 *
 *   • TILT MAGNITUDE (how steeply the lip rises) — from the cross-section profile's
 *     local slope at the EXACT EDT distance `d`. Distance is the geometrically
 *     meaningful quantity for the height (a `circle` lip at d=band/2 is at a precise
 *     height), and the EDT gives it exactly. This is what the PDF-calibrated slide-3
 *     brightness curve depends on, so it is preserved bit-for-bit in spirit.
 *
 *   • IN-PLANE AZIMUTH (which way the lip faces) — from the gradient of a Gaussian-
 *     blurred COVERAGE field `C = G_σ * alpha`, σ = 0.25·bandPx. This is the load-
 *     bearing change. `∇C = (∇G_σ) * alpha` is a convolution with the smooth kernel
 *     ∇G_σ, so C ∈ C^∞ regardless of how the silhouette was rasterised: the azimuth
 *     field rotates continuously around any smooth boundary, with a per-step turn
 *     bounded only by the silhouette's curvature. Because σ scales with the band and
 *     the geometry scales with devScale TOGETHER, the smoothness is identical at
 *     every raster scale — it is scale-invariant by construction.
 *
 * ### Why this finally fixes the size-dependent facet (and the patches didn't)
 *
 * The OLD method took BOTH parts from the finite-difference gradient of the EDT
 * height field. But ∇(distance-to-point-set) is piecewise-constant: each band
 * pixel's nearest boundary sample dominates, so the gradient DIRECTION snaps to
 * that sample's Voronoi-cell direction, chording the lip into facets. The screen-
 * blend highlight amplifies the chord at the lit/shadow terminator. Both prior
 * fixes were post-filters on that same broken field:
 *   - #413: box-blur the scalar height → smooths gradient MAGNITUDE, not direction
 *     (fixed devScale ≤ 2 only).
 *   - #415: tangential low-pass of the normal VECTOR over radius 0.25·bandPx, gated
 *     at bandPx ≥ 24 → fixed devScale 4, REGRESSED at devScale 8.
 * Every band-proportional blur radius chases a moving target: the Voronoi cell's
 * angular width depends on radial distance from the rim, which grows with the band,
 * so a large enough band always re-opens the chord. Sourcing the azimuth from ∇C
 * removes the Voronoi structure at the SOURCE — there is no discrete cell field
 * left to facet — so no gate and no post-blur are needed at any scale.
 *
 * The coverage blur is O(N) (three cascaded box passes, running-sum, independent of
 * σ), runs once per bevel on the device offscreen — the same order as the EDT it
 * sits beside, and cheaper than the two post-filter passes it replaces.
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

  // ── Azimuth source: gradient of a Gaussian-blurred coverage field ──────────
  // Blur the raw coverage (alpha/255) by σ ∝ band. σ is a quarter of the band: a
  // fraction of the bevel's OWN length scale, large enough to wash out the per-
  // pixel Voronoi noise yet small enough to stay local to a single rim (a larger σ
  // would blur the inner and outer rims of a thin ring together and flip the
  // azimuth at the medial axis). See COVERAGE_SIGMA_FRACTION.
  const coverage = new Float64Array(w * h);
  for (let i = 0; i < w * h; i++) coverage[i] = (alpha[i] ?? 0) / 255;
  const sigma = Math.max(1, bandPx * COVERAGE_SIGMA_FRACTION);
  const C = gaussianBlur(coverage, w, h, sigma);

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
      // Outward azimuth = −∇C/|∇C| (coverage decreases toward the rim, so the
      // negative gradient points outward, along the silhouette's outward normal).
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
// the light OUT of the screen toward the viewer by a fixed elevation so the lip
// catches light on the side facing `dir`. CALIBRATION (sample-11.pdf p3): the
// card's bevel rim is brightest along its upper edges and dims toward the lower
// edges, with `dir="t"` → light from straight above. An elevation that puts the
// light ~35° above the screen plane reproduces that gradient (a steeper, more
// overhead light flattens the contrast; a shallower one over-darkens the lower
// rim vs the PDF). The threePt rig is treated as a single dominant key light in
// the `dir` octant for Phase B; the fill/back lights of a true three-point rig
// are folded into the ambient term (see MATERIAL).

/** Light elevation above the screen plane, radians. Calibrated vs p3 (≈35°). */
const LIGHT_ELEVATION = (35 * Math.PI) / 180;

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
 * Unit light direction (FROM surface TOWARD light) for a light rig. `dir` sets
 * the screen-plane azimuth; the light is lifted toward the viewer (+Z) by
 * LIGHT_ELEVATION. See the SPEC GAP / calibration note above.
 */
export function lightDirFromRig(_rig: string, dir: string): Vec3 {
  const az = DIR_AZIMUTH[dir] ?? DIR_AZIMUTH.t;
  // Normalise the screen-plane component, then split between plane and +Z by
  // the elevation angle.
  const planeLen = Math.hypot(az.x, az.y) || 1;
  const cosE = Math.cos(LIGHT_ELEVATION);
  const sinE = Math.sin(LIGHT_ELEVATION);
  const x = (az.x / planeLen) * cosE;
  const y = (az.y / planeLen) * cosE;
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

/** Build shade params for a material + light rig. */
export function shadeParamsFor(material: BevelMaterial, light: Vec3): BevelShadeParams {
  const m = MATERIAL[material];
  return { light, material, ambient: m.ambient, diffuse: m.diffuse, specular: m.specular, shininess: m.shininess };
}

/**
 * Brightness multiplier for a surface normal under the light. Lambert diffuse +
 * (for plastic) a Blinn-Phong specular against the half-vector with the view
 * direction (0,0,1). Returns a factor ≥ 0 to multiply the body colour by
 * (1.0 = unchanged; >1 brightens via the screen-side blend the caller applies).
 */
export function shadePixel(n: Vec3, p: BevelShadeParams): number {
  const ndotl = n.x * p.light.x + n.y * p.light.y + n.z * p.light.z;
  const diff = p.diffuse * Math.max(0, ndotl);
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
  return Math.max(0, p.ambient + diff + spec);
}

/** A minimal 2D context surface the bevel compositor needs. */
export interface BevelCtx {
  canvas: { width: number; height: number };
  getImageData(sx: number, sy: number, sw: number, sh: number): ImageData;
  putImageData(data: ImageData, dx: number, dy: number): void;
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
export function applyBevelShading(ctx: BevelCtx, input: BevelInput): void {
  const w = ctx.canvas.width;
  const h = ctx.canvas.height;
  if (w <= 0 || h <= 0) return;
  const bandPx = input.widthPx;
  if (bandPx < 0.75) return; // sub-pixel lip — skip.

  const img = ctx.getImageData(0, 0, w, h);
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
  ctx.putImageData(img, 0, 0);
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
export function applyExtrusion(ctx: BevelCtx, input: ExtrusionInput): void {
  const w = ctx.canvas.width;
  const h = ctx.canvas.height;
  if (w <= 0 || h <= 0) return;
  const dx = input.offsetX;
  const dy = input.offsetY;
  const len = Math.hypot(dx, dy);
  if (len < 0.75) return; // sub-pixel — no visible side wall.

  const img = ctx.getImageData(0, 0, w, h);
  const px = img.data;
  // Source alpha snapshot (the front face silhouette).
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
  ctx.putImageData(img, 0, 0);
}
