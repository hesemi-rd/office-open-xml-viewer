use crate::styles::{parse_run_fmt, RunFmt};
use crate::xml_util::*;
use roxmltree::Document as XmlDoc;
use std::collections::HashMap;

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
    pub fn parse(xml: &str) -> Self {
        let mut map = NumberingMap::default();
        let doc = match XmlDoc::parse(xml) {
            Ok(d) => d,
            Err(_) => return map,
        };
        let root = doc.root_element();

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
                levels.push(LevelDef {
                    format,
                    text,
                    indent_left,
                    indent_right,
                    tab,
                    suff,
                    start,
                    rpr,
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
        NumberingMap::parse(&format!("<w:numbering {W}>{body}</w:numbering>"))
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
