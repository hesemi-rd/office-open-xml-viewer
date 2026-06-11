/**
 * DrawingML 3D camera — perspective projection of a planar shape (Phase A).
 *
 * ECMA-376 references:
 *   - §20.1.5.5  `camera` (CT_Camera): prst / fov / zoom / rot.
 *   - §20.1.5.11 `rot`    (CT_SphereCoords): lat / lon / rev rotation.
 *   - §20.1.10.47 ST_PresetCameraType: the 62 preset cameras.
 *
 * SPEC GAP (documented assumption — see CLAUDE.md "spec fidelity"):
 * ECMA-376 specifies the camera *model* (rotation order, fov meaning, the
 * preset enumeration) but gives **no numeric definitions** for the presets —
 * no base orientation, no per-preset field of view, no near/far. There is no
 * MS-OI29500 in this repo's spec/ tree either. The numeric constants below are
 * therefore the de-facto values from the OOXML implementer consensus, i.e. the
 * orientation/FOV PowerPoint emits and that LibreOffice's `oox` import maps in
 * oox/source/drawingml/scene3dcontext.cxx + the preset-camera table. They are
 * flagged here as a derived assumption, not a spec quotation.
 *
 * Phase A scope: we model the shape as a flat z=0 rectangle (the picture / 2D
 * drawing) and compute the 2D homography that maps the un-projected rectangle
 * to the projected quad. Bevel / extrusion / lightRig shading is Phase B.
 */

/**
 * Geometry-only camera inputs the projection math needs. Defined locally so the
 * core helper has no dependency on the pptx (or any format) type layer; the
 * pptx `Camera3d` / `Rot3d` are structurally assignable to these. Angles are in
 * degrees (the parser already converted from 60000ths).
 */
export interface CameraInput {
  /** Preset camera name (`ST_PresetCameraType`). */
  prst: string;
  /** Field-of-view override in degrees. Omitted = preset default. */
  fov?: number;
  /** Zoom factor as a unit ratio (1.0 = 100%). Omitted = 1.0. */
  zoom?: number;
  /** Camera rotation override. Omitted = preset base orientation. */
  rot?: RotInput;
}

/** Sphere-coordinate rotation in degrees — ECMA-376 §20.1.5.11. */
export interface RotInput {
  lat: number;
  lon: number;
  rev: number;
}

/** A 2D point in the projected (screen) plane, pixels relative to bbox origin. */
export interface Vec2 {
  x: number;
  y: number;
}

/** Result of projecting a planar shape through the camera. */
export interface Scene3dQuad {
  /**
   * The four projected corners in CSS pixels, relative to the element's
   * bounding-box top-left, in order: top-left, top-right, bottom-right,
   * bottom-left (matching the source rectangle (0,0)-(w,0)-(w,h)-(0,h)).
   *
   * The quad is re-centred and scaled to fit the SAME bounding box (w×h) the
   * un-projected shape occupied, so the renderer can blit a w×h offscreen
   * bitmap onto it without changing the element's footprint.
   */
  corners: [Vec2, Vec2, Vec2, Vec2];
  /**
   * True when the projection is (numerically) an affine map — i.e. the four
   * corners still form a parallelogram. Orthographic / isometric presets and a
   * pure in-plane `rev` are affine; perspective presets with lat/lon tilt are
   * not. Callers can take a cheaper single-`setTransform` path when affine.
   */
  isAffine: boolean;
  /**
   * True when the projection is (numerically) the identity within the bbox —
   * the renderer should skip projection entirely and draw normally. Happens for
   * `orthographicFront` with no rot override (or rot all-zero).
   */
  isIdentity: boolean;
}

/** Preset camera kinds that determine the projection math. */
type ProjectionKind = 'perspective' | 'orthographic';

/**
 * Per-preset definition: the base camera orientation (applied before the user
 * `rot` override) and the default field of view in degrees for perspective
 * presets. Orthographic / isometric presets use parallel projection (fov is
 * ignored). See the SPEC GAP note above for the source of these numbers.
 *
 * Base orientation is expressed as the same lat/lon/rev sphere coordinates the
 * file uses, in degrees, so it composes with the user rot in one place.
 */
interface PresetDef {
  kind: ProjectionKind;
  /** Base latitude (X-axis tilt), degrees. */
  baseLat: number;
  /** Base longitude (Y-axis turn), degrees. */
  baseLon: number;
  /** Base revolution (Z-axis spin), degrees. */
  baseRev: number;
  /** Default field of view for perspective presets, degrees. */
  fovDeg: number;
}

// Default perspective FOV. PowerPoint's perspective presets cluster around a
// ~45° vertical field of view; the "relaxed" variants soften it. These match
// the values LibreOffice oox uses for the corresponding presets.
const DEFAULT_PERSP_FOV = 45;

/**
 * Preset camera table. Only the presets we can render in Phase A (planar
 * homography) are enumerated with their orientation; everything else falls back
 * to `orthographicFront` (identity) so an unknown/legacy preset never throws.
 *
 * Orientation sign convention (object axes): +X right, +Y down, +Z toward the
 * viewer. `lat` tilts the top toward (-) / away (+) from the viewer, `lon`
 * turns the right edge toward (+) / away (-) from the viewer, `rev` spins in
 * the screen plane clockwise (+).
 */
const PRESETS: Record<string, PresetDef> = {
  // ---- Orthographic / front-facing ----
  orthographicFront: { kind: 'orthographic', baseLat: 0, baseLon: 0, baseRev: 0, fovDeg: 0 },

  // ---- Perspective family (front + heroic + relaxed) ----
  // perspectiveFront: looking straight on, but with perspective foreshortening.
  perspectiveFront: { kind: 'perspective', baseLat: 0, baseLon: 0, baseRev: 0, fovDeg: DEFAULT_PERSP_FOV },
  // perspectiveRelaxed / perspectiveRelaxedModerately: a gentle downward tilt so
  // the top edge recedes. PowerPoint's "Perspective Relaxed" tips the shape
  // back ~10–18°; "Moderately" is roughly half that. These are the de-facto
  // angles (see SPEC GAP). The file-level <a:rot> overrides them when present.
  perspectiveRelaxed: { kind: 'perspective', baseLat: 0, baseLon: 0, baseRev: 0, fovDeg: DEFAULT_PERSP_FOV },
  perspectiveRelaxedModerately: { kind: 'perspective', baseLat: 0, baseLon: 0, baseRev: 0, fovDeg: DEFAULT_PERSP_FOV },
  // perspectiveAbove / perspectiveBelow: camera looks down / up at the shape.
  perspectiveAbove: { kind: 'perspective', baseLat: 20, baseLon: 0, baseRev: 0, fovDeg: DEFAULT_PERSP_FOV },
  perspectiveBelow: { kind: 'perspective', baseLat: -20, baseLon: 0, baseRev: 0, fovDeg: DEFAULT_PERSP_FOV },
  // Left / right facing perspective variants: shape turned about the Y axis.
  perspectiveLeft: { kind: 'perspective', baseLat: 0, baseLon: -20, baseRev: 0, fovDeg: DEFAULT_PERSP_FOV },
  perspectiveRight: { kind: 'perspective', baseLat: 0, baseLon: 20, baseRev: 0, fovDeg: DEFAULT_PERSP_FOV },
};

/**
 * Note on `perspectiveRelaxed` base orientation: the most common authoring path
 * (and the sample-11 case) supplies an explicit `<a:rot>` that fully defines the
 * orientation, so the preset's own base tilt is only a fallback. We keep the
 * base tilt at 0 for the relaxed presets and let the file's rot drive the
 * orientation; this is the spec-faithful choice because the rot override, when
 * present, *replaces* the preset rotation (§20.1.5.5: "rotation … overrides …
 * that further rotate the camera"). When no rot is present we fall back to the
 * front view rather than inventing an angle we cannot cite.
 */

/** 3×3 matrix in row-major order. */
type Mat3 = [number, number, number, number, number, number, number, number, number];

function mul3(a: Mat3, b: Mat3): Mat3 {
  const r = new Array(9).fill(0) as number[];
  for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
      let s = 0;
      for (let k = 0; k < 3; k++) s += a[i * 3 + k] * b[k * 3 + j];
      r[i * 3 + j] = s;
    }
  }
  return r as Mat3;
}

function rotX(deg: number): Mat3 {
  const t = (deg * Math.PI) / 180;
  const c = Math.cos(t);
  const s = Math.sin(t);
  // X-axis rotation (latitude tilt).
  return [1, 0, 0, 0, c, -s, 0, s, c];
}

function rotY(deg: number): Mat3 {
  const t = (deg * Math.PI) / 180;
  const c = Math.cos(t);
  const s = Math.sin(t);
  // Y-axis rotation (longitude turn).
  return [c, 0, s, 0, 1, 0, -s, 0, c];
}

function rotZ(deg: number): Mat3 {
  const t = (deg * Math.PI) / 180;
  const c = Math.cos(t);
  const s = Math.sin(t);
  // Z-axis rotation (revolution / in-plane spin).
  return [c, -s, 0, s, c, 0, 0, 0, 1];
}

function applyMat3(m: Mat3, x: number, y: number, z: number): [number, number, number] {
  return [
    m[0] * x + m[1] * y + m[2] * z,
    m[3] * x + m[4] * y + m[5] * z,
    m[6] * x + m[7] * y + m[8] * z,
  ];
}

/**
 * Compose the full object rotation from a base orientation and an optional file
 * `rot` override. Per §20.1.5.11 the revolution is applied "about the axis as
 * the latitude and longitude coordinates", i.e. rev is the last (outermost)
 * rotation about the resulting view axis. We therefore compose
 *   R = Rz(rev) · Rx(lat) · Ry(lon).
 * When an explicit `<a:rot>` is present it supplies lat/lon/rev directly
 * (replacing the preset base, per the override semantics in §20.1.5.5); when
 * absent we use the preset's base angles.
 */
function buildRotation(preset: PresetDef, rot: RotInput | undefined): Mat3 {
  const lat = rot ? rot.lat : preset.baseLat;
  const lon = rot ? rot.lon : preset.baseLon;
  const rev = rot ? rot.rev : preset.baseRev;
  return mul3(rotZ(rev), mul3(rotX(lat), rotY(lon)));
}

/** Look up a preset def, falling back to orthographicFront for unknown names. */
function presetDef(prst: string): PresetDef {
  return PRESETS[prst] ?? PRESETS.orthographicFront;
}

/**
 * Project a planar shape (the w×h rectangle) through the camera and return the
 * four projected corners, refit into the original w×h bounding box.
 *
 * @param camera  parsed <a:camera> (prst / fov / zoom / rot, angles in degrees).
 * @param w       shape width in CSS pixels.
 * @param h       shape height in CSS pixels.
 */
export function computeScene3dQuad(camera: CameraInput, w: number, h: number): Scene3dQuad {
  const def = presetDef(camera.prst);
  const R = buildRotation(def, camera.rot);

  // Degenerate sizes can't be projected meaningfully.
  if (w <= 0 || h <= 0) {
    return {
      corners: [
        { x: 0, y: 0 },
        { x: w, y: 0 },
        { x: w, y: h },
        { x: 0, y: h },
      ],
      isAffine: true,
      isIdentity: true,
    };
  }

  // Source rectangle corners centred at the origin in shape-local space, with
  // +X right, +Y down. We work in a unit-ish space (half-extents) so the FOV /
  // camera distance is independent of pixel size, then rescale at the end.
  const hw = w / 2;
  const hh = h / 2;
  const srcCorners: Array<[number, number]> = [
    [-hw, -hh], // top-left
    [hw, -hh], // top-right
    [hw, hh], // bottom-right
    [-hw, hh], // bottom-left
  ];

  const zoom = camera.zoom ?? 1;

  // Camera distance for perspective. The shape's larger half-extent subtends
  // the full field of view at the camera plane: d = halfMax / tan(fov/2). This
  // is the standard pinhole relation; it keeps the shape filling the frame
  // before tilting, matching how PowerPoint frames the un-tilted shape.
  const halfMax = Math.max(hw, hh);
  let projected: Array<[number, number]>;

  if (def.kind === 'perspective') {
    const fovDeg = camera.fov ?? def.fovDeg;
    const fov = (Math.max(1, Math.min(179, fovDeg)) * Math.PI) / 180;
    const d = halfMax / Math.tan(fov / 2);
    projected = srcCorners.map(([sx, sy]) => {
      const [rx, ry, rz] = applyMat3(R, sx, sy, 0);
      // Camera at +Z = d looking toward -Z (origin). A point at object-z rz is
      // at camera-relative depth (d - rz); perspective divide by that depth.
      const depth = d - rz;
      // Guard against a corner crossing the camera plane (shouldn't happen for
      // sane tilts, but clamp to keep the map finite).
      const safeDepth = Math.abs(depth) < 1e-6 ? 1e-6 * Math.sign(depth || 1) : depth;
      const f = d / safeDepth;
      return [rx * f, ry * f] as [number, number];
    });
  } else {
    // Parallel (orthographic / isometric) projection: drop z after rotation.
    projected = srcCorners.map(([sx, sy]) => {
      const [rx, ry] = applyMat3(R, sx, sy, 0);
      return [rx, ry] as [number, number];
    });
  }

  // Apply zoom about the centre.
  projected = projected.map(([px, py]) => [px * zoom, py * zoom] as [number, number]);

  // Refit the projected quad into the original w×h box: scale uniformly so the
  // quad's bounding box matches w×h, then translate so its centre sits at the
  // box centre (w/2, h/2). Uniform scale preserves the projected shape's aspect
  // (we never stretch x and y independently).
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const [px, py] of projected) {
    if (px < minX) minX = px;
    if (py < minY) minY = py;
    if (px > maxX) maxX = px;
    if (py > maxY) maxY = py;
  }
  const projW = maxX - minX || 1;
  const projH = maxY - minY || 1;
  const fit = Math.min(w / projW, h / projH);
  // Centre of the projected quad's bounding box; this is the point we pin to the
  // element's bbox centre so the projected shape stays inside its footprint.
  const projCx = (minX + maxX) / 2;
  const projCy = (minY + maxY) / 2;

  const corners = projected.map(([px, py]) => ({
    x: w / 2 + (px - projCx) * fit,
    y: h / 2 + (py - projCy) * fit,
  })) as [Vec2, Vec2, Vec2, Vec2];

  // Affine test: a quad is a parallelogram iff (C0 + C2) == (C1 + C3) (the
  // diagonals bisect). Use a relative epsilon vs. the bbox size.
  const eps = 1e-3 * Math.max(w, h);
  const sumDiagX = corners[0].x + corners[2].x - (corners[1].x + corners[3].x);
  const sumDiagY = corners[0].y + corners[2].y - (corners[1].y + corners[3].y);
  const isAffine = Math.abs(sumDiagX) < eps && Math.abs(sumDiagY) < eps;

  // Identity test: corners within epsilon of the un-projected rectangle.
  const ref: Array<[number, number]> = [
    [0, 0],
    [w, 0],
    [w, h],
    [0, h],
  ];
  let isIdentity = true;
  for (let i = 0; i < 4; i++) {
    if (Math.abs(corners[i].x - ref[i][0]) > eps || Math.abs(corners[i].y - ref[i][1]) > eps) {
      isIdentity = false;
      break;
    }
  }

  return { corners, isAffine, isIdentity };
}

/**
 * True when a scene3d camera will visibly change the shape, i.e. the renderer
 * should take the projection path. A camera whose preset is unknown/front-facing
 * with no rot (or an all-zero rot) is the identity and can be skipped.
 */
export function isScene3dNonIdentity(camera: CameraInput): boolean {
  // A 1×1 probe is enough to detect identity vs. projection; the result scales.
  const { isIdentity } = computeScene3dQuad(camera, 1000, 1000);
  return !isIdentity;
}
