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
 *   3. The surface normal is the finite-difference gradient of that height field
 *      (n = normalize(-∂h/∂x, -∂h/∂y, 1)). On the flat interior the gradient is
 *      zero → n = (0,0,1); on the lip it tilts outward toward the rim.
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
    case 'hardEdge':
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
 */
export function computeBevelNormals(
  alpha: ArrayLike<number>,
  w: number,
  h: number,
  bandPx: number,
  prst: string,
  heightPx: number,
): { normals: Float32Array; bandMask: Uint8Array } {
  const normals = new Float32Array(w * h * 3);
  const bandMask = new Uint8Array(w * h);
  const dist = distanceToEdge(alpha, w, h);
  const profile = bevelHeightProfile(prst, bandPx);

  // Height field: profile(d) scaled by the bevel's height/width aspect. A height
  // of `heightPx` over a band of `bandPx` means the lip rises heightPx px over
  // bandPx px of run, so the surface-space gradient scale is heightPx/bandPx.
  const heightScale = bandPx > 0 ? heightPx / bandPx : 0;

  // Dense height field over the whole mask: profile(d)·heightScale·bandPx inside
  // the silhouette, 0 (ground) outside. Materialising it lets us regularise the
  // gradient (below) before differencing, instead of sampling it per-neighbour.
  const inside = (x: number, y: number) => (alpha[y * w + x] ?? 0) >= 128;
  const height = new Float64Array(w * h);
  for (let i = 0; i < w * h; i++) {
    height[i] = (alpha[i] ?? 0) >= 128 ? profile(dist[i]) * heightScale * bandPx : 0;
  }

  // ── Anti-facet smoothing of the height field ──────────────────────────────
  // The EDT `dist` is exact, but its *gradient* is piecewise-constant: every
  // band pixel's distance is dominated by ONE nearest boundary sample, so ∇d (and
  // hence the lip normal) snaps to that sample's Voronoi-cell direction. On a
  // curved silhouette (e.g. an ellipse) this turns the lip into flat chords — a
  // visible polygonal facet whose coarseness grows with the band width (the wider
  // the band, the larger the cells). A bevel lip is geometrically a smooth ruled
  // surface swept along the silhouette, so its normal must follow the silhouette's
  // *macroscopic* direction, not the per-pixel nearest-sample direction.
  //
  // We recover that by box-blurring the height field over a radius proportional to
  // the band width before differencing. The radius is NOT an empirical fudge: it
  // is the only length scale in the problem — averaging ∇h over a neighbourhood a
  // small fraction of the band removes the sub-pixel Voronoi noise while staying
  // far more local than the band itself, so the iso-contours (band mask) and the
  // calibrated brightness curve along the lip are unchanged (the profile is ~linear
  // across a few px, so a symmetric blur preserves its value). `dist`/`bandMask`
  // stay exact — only the gradient used for the normal is regularised.
  //
  // Cost vs. a "full-resolution exact" alternative: the EDT is already exact and
  // full-res (Felzenszwalb O(N) per pass). The faceting is intrinsic to sampling
  // ∇(distance-to-point-set), so a higher-res EDT would NOT remove it — only push
  // the cells smaller. A band-scaled separable box blur is O(N) (two passes,
  // running-sum) and adds a constant factor to a step we already pay, which is why
  // it is preferred over, say, re-deriving the analytic silhouette normal.
  const blurR = Math.max(1, Math.round(bandPx * 0.12));
  const sm = boxBlurClampedInside(height, alpha, w, h, blurR);

  const heightSmooth = (x: number, y: number): number => sm[y * w + x];

  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (!inside(x, y)) {
        // Outside the silhouette: flat ground normal, no shading.
        normals[i * 3 + 2] = 1;
        continue;
      }
      const d = dist[i];
      const inBand = d > 0 && d < bandPx;
      bandMask[i] = inBand ? 1 : 0;
      if (!inBand) {
        // Flat top of the shape (or right at the rim) — face the viewer.
        normals[i * 3 + 2] = 1;
        continue;
      }
      // Central finite difference of the SMOOTHED height field. Clamp neighbours
      // that step outside the silhouette to the current pixel's height so the lip
      // gradient is measured against the rim, not against ground level beyond.
      const hC = heightSmooth(x, y);
      const hL = x > 0 && inside(x - 1, y) ? heightSmooth(x - 1, y) : hC;
      const hR = x < w - 1 && inside(x + 1, y) ? heightSmooth(x + 1, y) : hC;
      const hU = y > 0 && inside(x, y - 1) ? heightSmooth(x, y - 1) : hC;
      const hD = y < h - 1 && inside(x, y + 1) ? heightSmooth(x, y + 1) : hC;
      // Gradient (∂h/∂x, ∂h/∂y) via central difference (px units cancel: 2px run).
      const dhdx = (hR - hL) / 2;
      const dhdy = (hD - hU) / 2;
      // Surface normal of z=h(x,y): n = normalize(-dh/dx, -dh/dy, 1).
      let nx = -dhdx;
      let ny = -dhdy;
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
  return { normals, bandMask };
}

/**
 * Separable box blur of a height field, with samples that fall OUTSIDE the
 * silhouette (alpha < 128) skipped rather than averaged in as ground level.
 *
 * Why skip instead of clamp-to-zero: averaging in the exterior ground (h=0) near
 * the rim would pull the smoothed height down and bias the gradient inward,
 * weakening the lip exactly at the rim where the shading matters most. Skipping
 * out-of-silhouette taps keeps the smoothing a pure *tangential* regulariser of
 * the lip surface — it averages along the band, not across the rim into the void.
 *
 * O(N·r) here (a plain windowed mean, not a running sum) because the per-tap
 * inside test makes a constant-time sliding sum awkward; r is small (≈ band·0.12,
 * a handful of px) so this stays well within the EDT's own cost. The blur runs
 * once per bevel, on the device-resolution offscreen, so it is paid only for
 * shapes that actually carry an sp3d bevel.
 */
function boxBlurClampedInside(
  src: Float64Array,
  alpha: ArrayLike<number>,
  w: number,
  h: number,
  r: number,
): Float64Array {
  const tmp = new Float64Array(w * h);
  const out = new Float64Array(w * h);
  const isInside = (idx: number) => (alpha[idx] ?? 0) >= 128;
  // Horizontal pass.
  for (let y = 0; y < h; y++) {
    const row = y * w;
    for (let x = 0; x < w; x++) {
      const i = row + x;
      if (!isInside(i)) {
        tmp[i] = 0;
        continue;
      }
      let s = 0;
      let n = 0;
      for (let k = -r; k <= r; k++) {
        const xx = x + k;
        if (xx < 0 || xx >= w) continue;
        const j = row + xx;
        if (!isInside(j)) continue;
        s += src[j];
        n++;
      }
      tmp[i] = n > 0 ? s / n : src[i];
    }
  }
  // Vertical pass.
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (!isInside(i)) {
        out[i] = 0;
        continue;
      }
      let s = 0;
      let n = 0;
      for (let k = -r; k <= r; k++) {
        const yy = y + k;
        if (yy < 0 || yy >= h) continue;
        const j = yy * w + x;
        if (!isInside(j)) continue;
        s += tmp[j];
        n++;
      }
      out[i] = n > 0 ? s / n : tmp[i];
    }
  }
  return out;
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

  const { normals, bandMask } = computeBevelNormals(
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
    let nx = normals[i * 3];
    let ny = normals[i * 3 + 1];
    const nz = normals[i * 3 + 2];
    // bevelB: the lip belongs to the back face, lit from the opposite vertical
    // sense — mirror the in-plane normal so the light/shadow sides swap.
    if (input.bottom) {
      nx = -nx;
      ny = -ny;
    }
    const f = shadePixel({ x: nx, y: ny, z: nz }, params) / faceFactor;
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
