//! OOXML color transforms (lumMod, lumOff, satMod, satOff, hueMod, hueOff,
//! shade, tint, alpha and friends) shared between the docx and pptx parsers.
//!
//! Word and PowerPoint diverge on the `tint` transform — Word reads `val` as
//! the *retained fraction of the input color* (the literal ECMA-376
//! §20.1.2.3.34 reading: `result = val·input + (1-val)·white`), while
//! PowerPoint applies it as a `lerp(input, white, val)` in linear sRGB. Empirical
//! comparison against PDF exports confirms each app does its own thing — see
//! `TintMode` and the per-app `apply_color_transforms_with` flag.
//!
//! Everything else (shade, lumMod/Off, satMod/Off, hueMod/Off, alpha
//! family) is identical between the two and lives here uncopied.

use roxmltree::Node;

/// Selects the formula applied to `<a:tint val>` modifiers. The OOXML spec
/// is consistent (val = retained input), but the two desktop apps render
/// templates differently in practice — see ECMA-376 §20.1.2.3.34 and the
/// commit history of pptx-parser for the linear-sRGB derivation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TintMode {
    /// Word: `result = val·input + (1-val)·white` in sRGB. Matches Word's
    /// rendering of resume / cover templates that use accent recolors with
    /// tint values.
    WordLiteral,
    /// PowerPoint: `lerp(input, white, val)` in linear sRGB. Matches
    /// PowerPoint's rendering of SmartArt accent recolors pixel-for-pixel.
    PowerPointLinear,
}

/// Apply OOXML color transforms to `hex` based on the modifier elements
/// declared as direct children of `node`. Returns 6-char hex when fully
/// opaque, or 8-char hex (RRGGBBAA) when alpha < 1.
pub fn apply_color_transforms(hex: &str, node: Node, tint_mode: TintMode) -> String {
    if hex.len() < 6 {
        return hex.to_owned();
    }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);

    let mut rf = r as f64 / 255.0;
    let mut gf = g as f64 / 255.0;
    let mut bf = b as f64 / 255.0;
    let mut alpha = if hex.len() >= 8 {
        u8::from_str_radix(&hex[6..8], 16).unwrap_or(255) as f64 / 255.0
    } else {
        1.0
    };

    let attr_pct = |t: &Node, name: &str, default: f64| -> f64 {
        t.attribute(name)
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(default)
            / 100_000.0
    };

    for t in node.children().filter(|n| n.is_element()) {
        match t.tag_name().name() {
            "lumMod" => {
                let val = attr_pct(&t, "val", 100_000.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, (l * val).min(1.0), s);
                rf = nr; gf = ng; bf = nb;
            }
            "lumOff" => {
                let val = attr_pct(&t, "val", 0.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, (l + val).clamp(0.0, 1.0), s);
                rf = nr; gf = ng; bf = nb;
            }
            "satMod" => {
                let val = attr_pct(&t, "val", 100_000.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, l, (s * val).clamp(0.0, 1.0));
                rf = nr; gf = ng; bf = nb;
            }
            "satOff" => {
                let val = attr_pct(&t, "val", 0.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb(h, l, (s + val).clamp(0.0, 1.0));
                rf = nr; gf = ng; bf = nb;
            }
            "hueMod" => {
                let val = attr_pct(&t, "val", 100_000.0);
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb((h * val).rem_euclid(1.0), l, s);
                rf = nr; gf = ng; bf = nb;
            }
            "hueOff" => {
                // hueOff is in 60000ths of a degree per ECMA-376 §20.1.2.3.16.
                let val_deg = t.attribute("val").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 60_000.0;
                let (h, l, s) = rgb_to_hls(rf, gf, bf);
                let (nr, ng, nb) = hls_to_rgb((h + val_deg / 360.0).rem_euclid(1.0), l, s);
                rf = nr; gf = ng; bf = nb;
            }
            "shade" => {
                // ECMA-376 §20.1.2.3.31: result = val·input + (1-val)·black.
                let val = attr_pct(&t, "val", 100_000.0);
                rf *= val; gf *= val; bf *= val;
            }
            "tint" => {
                let val = attr_pct(&t, "val", 0.0);
                match tint_mode {
                    TintMode::WordLiteral => {
                        // `result = val·input + (1-val)·white` per literal spec.
                        rf = val * rf + (1.0 - val);
                        gf = val * gf + (1.0 - val);
                        bf = val * bf + (1.0 - val);
                    }
                    TintMode::PowerPointLinear => {
                        // PowerPoint reads val as the lerp fraction toward
                        // white in LINEAR sRGB. Verified against PDF
                        // exports of SmartArt accent recolors.
                        let lr = srgb_to_linear(rf);
                        let lg = srgb_to_linear(gf);
                        let lb = srgb_to_linear(bf);
                        rf = linear_to_srgb((lr + (1.0 - lr) * val).clamp(0.0, 1.0));
                        gf = linear_to_srgb((lg + (1.0 - lg) * val).clamp(0.0, 1.0));
                        bf = linear_to_srgb((lb + (1.0 - lb) * val).clamp(0.0, 1.0));
                    }
                }
            }
            "alpha" => {
                // ECMA-376 §20.1.2.3.1 — sets absolute alpha.
                alpha = attr_pct(&t, "val", 100_000.0);
            }
            "alphaModFix" => {
                // ECMA-376 §20.1.8.4 — fixed (absolute) alpha modulation.
                alpha = attr_pct(&t, "amt", 100_000.0);
            }
            "alphaMod" => {
                // ECMA-376 §20.1.2.3.2 — multiply current alpha by val/100000.
                alpha *= attr_pct(&t, "val", 100_000.0);
            }
            "alphaOff" => {
                // ECMA-376 §20.1.2.3.3 — additive offset to alpha.
                alpha += attr_pct(&t, "val", 0.0);
            }
            _ => {}
        }
    }

    let r = (rf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let g = (gf.clamp(0.0, 1.0) * 255.0).round() as u8;
    let b = (bf.clamp(0.0, 1.0) * 255.0).round() as u8;
    if (alpha - 1.0).abs() < 0.004 {
        format!("{:02X}{:02X}{:02X}", r, g, b)
    } else {
        let a = (alpha.clamp(0.0, 1.0) * 255.0).round() as u8;
        format!("{:02X}{:02X}{:02X}{:02X}", r, g, b, a)
    }
}

/// sRGB → linear light. IEC 61966-2-1 transfer function.
pub fn srgb_to_linear(c: f64) -> f64 {
    if c <= 0.04045 { c / 12.92 } else { ((c + 0.055) / 1.055).powf(2.4) }
}

/// Linear light → sRGB.
pub fn linear_to_srgb(c: f64) -> f64 {
    if c <= 0.0031308 { 12.92 * c } else { 1.055 * c.powf(1.0 / 2.4) - 0.055 }
}

/// RGB → HLS conversion (lightness in the middle, matches Python's colorsys).
/// Returns (h, l, s) with each component in [0, 1].
pub fn rgb_to_hls(r: f64, g: f64, b: f64) -> (f64, f64, f64) {
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let d = max - min;
    if d < 1e-10 {
        return (0.0, l, 0.0);
    }
    let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
    let h = if (max - r).abs() < 1e-10 {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if (max - g).abs() < 1e-10 {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    };
    (h / 6.0, l, s)
}

/// HLS → RGB conversion. (h, l, s) are each in [0, 1]; returns linear-RGB
/// triple in [0, 1].
pub fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s < 1e-10 {
        return (l, l, l);
    }
    fn hue2rgb(p: f64, q: f64, mut t: f64) -> f64 {
        if t < 0.0 { t += 1.0; }
        if t > 1.0 { t -= 1.0; }
        if t < 1.0 / 6.0 { return p + (q - p) * 6.0 * t; }
        if t < 0.5        { return q; }
        if t < 2.0 / 3.0  { return p + (q - p) * (2.0 / 3.0 - t) * 6.0; }
        p
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    (
        hue2rgb(p, q, h + 1.0 / 3.0),
        hue2rgb(p, q, h),
        hue2rgb(p, q, h - 1.0 / 3.0),
    )
}
