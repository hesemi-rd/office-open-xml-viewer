use crate::styles::{parse_run_fmt, RunFmt};
use crate::xml_util::*;
use ooxml_common::blip::mime_from_ext;
use roxmltree::Document as XmlDoc;
use std::collections::HashMap;

/// Parse a single VML CSS length (e.g. `width:9pt`) from a `style` attribute
/// into pt. Supports the units Word emits for picture-bullet shapes: `pt`
/// (1pt), `in` (72pt), `pc` (12pt), `cm` (28.3465pt), `mm` (2.83465pt). A bare
/// number with no unit is treated as pt (VML's default user unit for shapes is
/// the point). Returns `None` when the property is absent or unparseable.
fn vml_style_len(style: &str, prop: &str) -> Option<f64> {
    for decl in style.split(';') {
        let (k, v) = decl.split_once(':')?;
        if k.trim() != prop {
            continue;
        }
        let v = v.trim();
        let (num, factor) = if let Some(n) = v.strip_suffix("pt") {
            (n, 1.0)
        } else if let Some(n) = v.strip_suffix("in") {
            (n, 72.0)
        } else if let Some(n) = v.strip_suffix("pc") {
            (n, 12.0)
        } else if let Some(n) = v.strip_suffix("mm") {
            (n, 2.834_645_7)
        } else if let Some(n) = v.strip_suffix("cm") {
            (n, 28.346_457)
        } else {
            (v, 1.0)
        };
        return num.trim().parse::<f64>().ok().map(|n| n * factor);
    }
    None
}

#[derive(Debug, Clone)]
pub struct LevelDef {
    pub format: String,   // "decimal" | "bullet" | etc.
    pub text: String,     // lvlText val, e.g. "%1." or "•"
    pub indent_left: f64, // pt — w:ind@left ≡ logical START indent (§17.3.1.12)
    /// pt — w:ind@right ≡ logical END indent (Part 4 §14.11.2). RTL list levels
    /// carry their indent here (the renderer maps it to the physical left side).
    pub indent_right: Option<f64>,
    pub tab: f64, // pt
    /// ECMA-376 §17.9.28 `<w:suff>` — what follows the number text: "tab"
    /// (default), "space", or "nothing". Controls where the body text starts
    /// relative to the marker on the first line.
    pub suff: String,
    pub start: u32,
    /// ECMA-376 §17.9.6 `<w:lvl><w:rPr>` — the level's run (character) properties
    /// for the number/bullet glyph itself. Merged OVER the paragraph's resolved
    /// run formatting at use-site so the marker's font axes (ascii/eastAsia)
    /// resolve through the same chain a body run uses. Often only carries a bare
    /// `<w:rFonts w:hint="eastAsia"/>` (no explicit typeface), in which case every
    /// axis is `None` and the marker simply inherits the paragraph's fonts.
    pub rpr: RunFmt,
    /// ECMA-376 §17.9.9 `<w:lvlPicBulletId w:val="N"/>` — when present, the
    /// level's marker is the image defined by the `<w:numPicBullet>` whose
    /// `numPicBulletId` is N (§17.9.20), drawn in place of `text`. Resolved at
    /// parse time to the bullet image's zip path (+ MIME + pt size from the
    /// `<v:shape style>`). `None` ⇒ ordinary text/glyph marker.
    pub pic_bullet: Option<PicBullet>,
}

/// ECMA-376 §17.9.20 `<w:numPicBullet>` — an image used as a list marker. The
/// image is defined by a VML `<w:pict><v:shape><v:imagedata r:id="…"/>` whose
/// `r:id` resolves through `word/_rels/numbering.xml.rels` to a media part, and
/// whose `<v:shape style="width:..;height:..">` carries the marker size.
#[derive(Debug, Clone)]
pub struct PicBullet {
    /// Zip path of the bullet image (e.g. `word/media/image1.gif`), resolved
    /// from the `<v:imagedata r:id>` via the numbering part's relationships.
    pub image_path: String,
    /// MIME type derived from the part extension (e.g. `image/gif`).
    pub mime_type: String,
    /// Marker width in pt, from the `<v:shape style="width:..">` (default 9pt
    /// when the style omits a width — the common Word default for bullet shapes).
    pub width_pt: f64,
    /// Marker height in pt, from the `<v:shape style="height:..">`.
    pub height_pt: f64,
}

impl Default for LevelDef {
    fn default() -> Self {
        LevelDef {
            format: "decimal".to_string(),
            text: "%1.".to_string(),
            indent_left: 36.0,
            indent_right: None,
            tab: 36.0,
            suff: "tab".to_string(),
            start: 1,
            rpr: RunFmt::default(),
            pic_bullet: None,
        }
    }
}

#[derive(Default)]
pub struct NumberingMap {
    /// abstractNumId → [level0..level8]
    abstract_nums: HashMap<u32, Vec<LevelDef>>,
    /// numId → abstractNumId
    num_to_abstract: HashMap<u32, u32>,
    /// numId → level override starts
    num_overrides: HashMap<u32, HashMap<u32, u32>>,
    /// per-numId per-level counter
    pub counters: HashMap<u32, HashMap<u32, u32>>,
}

impl NumberingMap {
    /// Parse `word/numbering.xml`. `media_map` is the numbering part's own
    /// relationship table (rId → zip media path, built from
    /// `word/_rels/numbering.xml.rels`); it is required to resolve the
    /// `<w:numPicBullet>` images (§17.9.20) — an empty map simply yields no
    /// picture bullets, leaving levels on their text/glyph markers.
    pub fn parse(xml: &str, media_map: &HashMap<String, String>) -> Self {
        let mut map = NumberingMap::default();
        let doc = match XmlDoc::parse(xml) {
            Ok(d) => d,
            Err(_) => return map,
        };
        let root = doc.root_element();

        // ECMA-376 §17.9.20 — collect `<w:numPicBullet>` definitions first so
        // each level's `<w:lvlPicBulletId>` (§17.9.9) can resolve against them.
        // The bullet image is a VML `<v:shape><v:imagedata r:id>`; the r:id maps
        // to a media part through the numbering part's rels (`media_map`), and
        // the `<v:shape style="width:..;height:..">` carries the marker size.
        let mut pic_bullets: HashMap<u32, PicBullet> = HashMap::new();
        for pb_node in children_w(root, "numPicBullet") {
            let Some(id) = attr_w(pb_node, "numPicBulletId").and_then(|v| v.parse::<u32>().ok())
            else {
                continue;
            };
            let Some(imagedata) = pb_node
                .descendants()
                .find(|n| n.tag_name().name() == "imagedata")
            else {
                continue;
            };
            // `r:id` lives in the relationships namespace; fall back to the
            // unqualified attribute for defensiveness.
            let Some(rid) = imagedata
                .attribute((R_NS, "id"))
                .or_else(|| imagedata.attribute("id"))
            else {
                continue;
            };
            let Some(image_path) = media_map.get(rid).cloned() else {
                continue;
            };
            // `<v:shape style="width:9pt;height:9pt">` — VML CSS lengths. Word
            // always emits explicit pt for bullet shapes; default to 9pt (Word's
            // standard picture-bullet size) when a dimension is absent.
            let shape_style = pb_node
                .descendants()
                .find(|n| n.tag_name().name() == "shape")
                .and_then(|n| n.attribute("style"))
                .unwrap_or("");
            let width_pt = vml_style_len(shape_style, "width").unwrap_or(9.0);
            let height_pt = vml_style_len(shape_style, "height").unwrap_or(9.0);
            let mime_type = mime_from_ext(&image_path).to_string();
            pic_bullets.insert(
                id,
                PicBullet {
                    image_path,
                    mime_type,
                    width_pt,
                    height_pt,
                },
            );
        }

        // Parse abstractNum definitions
        for abs_node in children_w(root, "abstractNum") {
            let Some(abs_id_s) = attr_w(abs_node, "abstractNumId") else {
                continue;
            };
            let abs_id: u32 = abs_id_s.parse().unwrap_or(0);
            let mut levels = vec![];
            for lvl_node in children_w(abs_node, "lvl") {
                let start = child_w(lvl_node, "start")
                    .and_then(|n| attr_w(n, "val"))
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(1);
                let format = child_w(lvl_node, "numFmt")
                    .and_then(|n| attr_w(n, "val"))
                    .unwrap_or_else(|| "decimal".to_string());
                let text = child_w(lvl_node, "lvlText")
                    .and_then(|n| attr_w(n, "val"))
                    .unwrap_or_else(|| "%1.".to_string());
                let ind_node = child_w(lvl_node, "pPr").and_then(|p| child_w(p, "ind"));
                // When the level defines a w:ind, a missing @left means "no
                // start indent from this source" (an RTL level carries its
                // indent in @right ≡ end instead); the per-level depth default
                // applies only when no w:ind exists at all.
                let indent_left = ind_node
                    .and_then(|i| attr_w(i, "left"))
                    .map(|v| twips_to_pt(&v))
                    .unwrap_or(if ind_node.is_some() {
                        0.0
                    } else {
                        720.0 / 20.0 * (levels.len() as f64 + 1.0)
                    });
                let indent_right = ind_node
                    .and_then(|i| attr_w(i, "right"))
                    .map(|v| twips_to_pt(&v));
                let tab = ind_node
                    .and_then(|i| attr_w(i, "hanging").or_else(|| attr_w(i, "firstLine")))
                    .map(|v| twips_to_pt(&v))
                    .unwrap_or(36.0);
                // §17.9.28: absent <w:suff> means "tab".
                let suff = child_w(lvl_node, "suff")
                    .and_then(|n| attr_w(n, "val"))
                    .unwrap_or_else(|| "tab".to_string());
                // §17.9.6 — the level's run properties for the marker glyph.
                // Parsed with the SAME `parse_run_fmt` body runs use; theme refs
                // stay as "@theme:<ref>" markers and are resolved at use-site once
                // merged over the paragraph's run formatting.
                let rpr = child_w(lvl_node, "rPr")
                    .map(parse_run_fmt)
                    .unwrap_or_default();
                // §17.9.9 — resolve the level's picture bullet (if any) against
                // the `<w:numPicBullet>` definitions collected above.
                let pic_bullet = child_w(lvl_node, "lvlPicBulletId")
                    .and_then(|n| attr_w(n, "val"))
                    .and_then(|v| v.parse::<u32>().ok())
                    .and_then(|id| pic_bullets.get(&id).cloned());
                levels.push(LevelDef {
                    format,
                    text,
                    indent_left,
                    indent_right,
                    tab,
                    suff,
                    start,
                    rpr,
                    pic_bullet,
                });
            }
            map.abstract_nums.insert(abs_id, levels);
        }

        // Parse num → abstractNum
        for num_node in children_w(root, "num") {
            let Some(num_id_s) = attr_w(num_node, "numId") else {
                continue;
            };
            let num_id: u32 = num_id_s.parse().unwrap_or(0);
            if let Some(abs_ref) = child_w(num_node, "abstractNumId").and_then(|n| attr_w(n, "val"))
            {
                let abs_id: u32 = abs_ref.parse().unwrap_or(0);
                map.num_to_abstract.insert(num_id, abs_id);
            }
            // Level overrides
            let mut overrides = HashMap::new();
            for lvl_ov in children_w(num_node, "lvlOverride") {
                let ilvl: u32 = attr_w(lvl_ov, "ilvl")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                if let Some(start_ov) =
                    child_w(lvl_ov, "startOverride").and_then(|n| attr_w(n, "val"))
                {
                    overrides.insert(ilvl, start_ov.parse().unwrap_or(1));
                }
            }
            if !overrides.is_empty() {
                map.num_overrides.insert(num_id, overrides);
            }
        }

        map
    }

    pub fn get_level(&self, num_id: u32, level: u32) -> Option<&LevelDef> {
        let abs_id = self.num_to_abstract.get(&num_id)?;
        let levels = self.abstract_nums.get(abs_id)?;
        levels.get(level as usize)
    }

    pub fn get_start(&self, num_id: u32, level: u32) -> u32 {
        if let Some(ov) = self.num_overrides.get(&num_id).and_then(|m| m.get(&level)) {
            return *ov;
        }
        self.get_level(num_id, level).map(|l| l.start).unwrap_or(1)
    }

    /// Advance counter for (numId, level), resetting deeper levels.
    ///
    /// `counters` stores each level's CURRENT displayed value (not the next):
    /// a level's first appearance shows its `start`, each later advance adds
    /// one, and advancing a level clears all deeper levels (§17.9.25 default
    /// `lvlRestart`). Shallower levels are seeded to their `start` so an
    /// ancestor that only prefixes the marker (e.g. `%1.%2`) still resolves
    /// when it is never advanced on its own. Returns the value to display.
    pub fn advance(&mut self, num_id: u32, level: u32) -> u32 {
        // Pre-compute start values to avoid borrow conflicts
        let starts: Vec<u32> = (0..=level).map(|l| self.get_start(num_id, l)).collect();

        let entry = self.counters.entry(num_id).or_default();

        // Reset deeper levels
        let keys: Vec<u32> = entry.keys().copied().filter(|&l| l > level).collect();
        for k in keys {
            entry.remove(&k);
        }

        // Seed shallower levels to their start (their displayed value when they
        // are never advanced themselves).
        for (lvl, &start) in starts.iter().enumerate().take(level as usize) {
            entry.entry(lvl as u32).or_insert(start);
        }

        // Current level: first appearance shows start, otherwise increment.
        let val = match entry.get(&level) {
            Some(&v) => v + 1,
            None => starts[level as usize],
        };
        entry.insert(level, val);
        val
    }

    /// Resolve the display text for a counter value in the given level.
    ///
    /// ECMA-376 §17.9.11 (`<w:lvlText>`): each `%N` placeholder is the counter
    /// of level `N-1`, formatted with THAT level's own `<w:numFmt>`. A
    /// multi-level marker such as `%1.%2` therefore needs every ancestor
    /// counter, not just the current level's. The current level uses `counter`
    /// (the value `advance` just returned); ancestor levels read their live
    /// counter from `self.counters` — `advance` seeds every shallower level to
    /// its start, so an ancestor that is never itself advanced (e.g. a list
    /// whose level 0 only exists to prefix subsection numbers with a fixed
    /// `start`) still resolves to its start value.
    pub fn resolve_text(&self, num_id: u32, level: u32, counter: u32) -> String {
        let Some(lvl) = self.get_level(num_id, level) else {
            return format!("{}.", counter);
        };

        let mut text = lvl.text.clone();
        // Replace from the deepest placeholder down so "%1" can never partially
        // match a two-digit "%1N" (Word caps lists at 9 levels, so this is
        // belt-and-braces — but cheap).
        for k in (0..=level).rev() {
            let val = if k == level {
                counter
            } else {
                self.counters
                    .get(&num_id)
                    .and_then(|m| m.get(&k))
                    .copied()
                    .unwrap_or_else(|| self.get_start(num_id, k))
            };
            let fmt = self
                .get_level(num_id, k)
                .map(|l| l.format.as_str())
                .unwrap_or(lvl.format.as_str());
            text = text.replace(&format!("%{}", k + 1), &format_counter(val, fmt));
        }
        text
    }
}

fn format_counter(n: u32, format: &str) -> String {
    match format {
        "decimal" => n.to_string(),
        "bullet" => "•".to_string(),
        "lowerLetter" => {
            let c = (b'a' + ((n - 1) % 26) as u8) as char;
            c.to_string()
        }
        "upperLetter" => {
            let c = (b'A' + ((n - 1) % 26) as u8) as char;
            c.to_string()
        }
        "lowerRoman" => to_roman(n).to_lowercase(),
        "upperRoman" => to_roman(n),
        _ => n.to_string(),
    }
}

fn to_roman(n: u32) -> String {
    let vals = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut n = n;
    let mut s = String::new();
    for (v, r) in &vals {
        while n >= *v {
            s.push_str(r);
            n -= v;
        }
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: &str = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    fn map(body: &str) -> NumberingMap {
        NumberingMap::parse(
            &format!("<w:numbering {W}>{body}</w:numbering>"),
            &HashMap::new(),
        )
    }

    /// §17.9.20 / §17.9.9 — a `<w:numPicBullet>` image resolves through the
    /// numbering part's rels (`media_map`), and the level's `<w:lvlPicBulletId>`
    /// picks it up with the `<v:shape style>` size (here width:9pt;height:9pt).
    #[test]
    fn picture_bullet_resolves_image_path_and_size() {
        const R: &str =
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        const V: &str = "xmlns:v=\"urn:schemas-microsoft-com:vml\"";
        const O: &str = "xmlns:o=\"urn:schemas-microsoft-com:office:office\"";
        let media: HashMap<String, String> =
            [("rId1".to_string(), "word/media/image1.gif".to_string())]
                .into_iter()
                .collect();
        let xml = format!(
            r#"<w:numbering {W} {R} {V} {O}>
                 <w:numPicBullet w:numPicBulletId="0">
                   <w:pict>
                     <v:shape id="x" style="width:9pt;height:9pt" o:bullet="t">
                       <v:imagedata r:id="rId1" o:title="BD"/>
                     </v:shape>
                   </w:pict>
                 </w:numPicBullet>
                 <w:abstractNum w:abstractNumId="8">
                   <w:lvl w:ilvl="0">
                     <w:numFmt w:val="bullet"/><w:lvlText w:val=""/>
                     <w:lvlPicBulletId w:val="0"/>
                   </w:lvl>
                 </w:abstractNum>
                 <w:num w:numId="3"><w:abstractNumId w:val="8"/></w:num>
               </w:numbering>"#
        );
        let m = NumberingMap::parse(&xml, &media);
        let lvl = m.get_level(3, 0).expect("level 0");
        let pb = lvl.pic_bullet.as_ref().expect("picture bullet resolved");
        assert_eq!(pb.image_path, "word/media/image1.gif");
        assert_eq!(pb.mime_type, "image/gif");
        assert!((pb.width_pt - 9.0).abs() < 1e-6);
        assert!((pb.height_pt - 9.0).abs() < 1e-6);
    }

    /// An unresolvable `r:id` (no matching rel) yields no picture bullet — the
    /// level falls back to its ordinary text marker.
    #[test]
    fn picture_bullet_missing_rel_falls_back() {
        const R: &str =
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        const V: &str = "xmlns:v=\"urn:schemas-microsoft-com:vml\"";
        let xml = format!(
            r#"<w:numbering {W} {R} {V}>
                 <w:numPicBullet w:numPicBulletId="0">
                   <w:pict><v:shape style="width:9pt;height:9pt">
                     <v:imagedata r:id="rIdX"/>
                   </v:shape></w:pict>
                 </w:numPicBullet>
                 <w:abstractNum w:abstractNumId="8">
                   <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/>
                     <w:lvlPicBulletId w:val="0"/></w:lvl>
                 </w:abstractNum>
                 <w:num w:numId="3"><w:abstractNumId w:val="8"/></w:num>
               </w:numbering>"#
        );
        let m = NumberingMap::parse(&xml, &HashMap::new());
        assert!(m.get_level(3, 0).unwrap().pic_bullet.is_none());
    }

    /// §17.9.11 — a subsection list (`%1.%2`) whose level 0 is never advanced
    /// but starts at 3 must render "3.1", "3.2", … (the bug: only the current
    /// level's placeholder was substituted, leaving a literal "%1").
    #[test]
    fn multilevel_parent_placeholder_uses_level_start_when_not_advanced() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="5">
                 <w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl>
                 <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="5"><w:abstractNumId w:val="5"/></w:num>"#);
        let c1 = m.advance(5, 1);
        assert_eq!(m.resolve_text(5, 1, c1), "3.1");
        let c2 = m.advance(5, 1);
        assert_eq!(m.resolve_text(5, 1, c2), "3.2");
        let c3 = m.advance(5, 1);
        assert_eq!(m.resolve_text(5, 1, c3), "3.3");
    }

    /// Parent counter is tracked live and resets deeper levels: 1, 1.1, 1.2,
    /// 2, 2.1.
    #[test]
    fn multilevel_parent_counter_increments_and_resets() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="0">
                 <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
                 <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#);
        let a = m.advance(1, 0);
        assert_eq!(m.resolve_text(1, 0, a), "1.");
        let b = m.advance(1, 1);
        assert_eq!(m.resolve_text(1, 1, b), "1.1");
        let c = m.advance(1, 1);
        assert_eq!(m.resolve_text(1, 1, c), "1.2");
        let d = m.advance(1, 0);
        assert_eq!(m.resolve_text(1, 0, d), "2.");
        let e = m.advance(1, 1);
        assert_eq!(m.resolve_text(1, 1, e), "2.1"); // deeper level reset on parent advance
    }

    /// Each level's `%N` is formatted with its OWN numFmt (§17.9.11): an
    /// upper-letter parent with a decimal child renders "A.1".
    #[test]
    fn multilevel_per_level_format() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="2">
                 <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1"/></w:lvl>
                 <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>"#);
        m.advance(2, 0);
        let c = m.advance(2, 1);
        assert_eq!(m.resolve_text(2, 1, c), "A.1");
    }
}
