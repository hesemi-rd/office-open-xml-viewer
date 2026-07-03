//! OMML (Office Math Markup Language, ECMA-376 §22.1) extraction into the shared
//! math AST consumed by `@silurus/ooxml-core`'s math engine. The serialized JSON
//! must match the TS `MathNode` types in `packages/core/src/types/math.ts`:
//! a `kind` discriminator with camelCase variant tags and fields.
//!
//! OMML is the same across WordprocessingML, PresentationML and SpreadsheetML
//! (the `m:` namespace is shared), so this parser is format-agnostic and is used
//! by the docx, pptx and xlsx parsers alike. It matches on element local names,
//! which keeps it independent of how each host document declares the math prefix
//! (e.g. PowerPoint wraps `m:oMathPara` in `a14:m` / `mc:AlternateContent`).

use roxmltree::Node;
use serde::{Deserialize, Serialize};

// `Deserialize` is derived for symmetry with the pptx/xlsx parser type trees
// (every node there derives both); the math AST is only ever serialized in
// practice. `default` on the skip-serialized fields keeps a round-trip valid.
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum MathNode {
    Run {
        text: String,
        /// "roman" | "italic" | "bold" | "boldItalic"
        style: String,
    },
    Fraction {
        num: Vec<MathNode>,
        den: Vec<MathNode>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        bar: Option<bool>,
    },
    Sup {
        base: Vec<MathNode>,
        sup: Vec<MathNode>,
    },
    Sub {
        base: Vec<MathNode>,
        sub: Vec<MathNode>,
    },
    SubSup {
        base: Vec<MathNode>,
        sub: Vec<MathNode>,
        sup: Vec<MathNode>,
    },
    #[serde(rename_all = "camelCase")]
    Nary {
        op: String,
        #[serde(default, skip_serializing_if = "String::is_empty")]
        lim_loc: String,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        sub: Vec<MathNode>,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        sup: Vec<MathNode>,
        body: Vec<MathNode>,
    },
    #[serde(rename_all = "camelCase")]
    Delimiter {
        beg_char: String,
        end_char: String,
        items: Vec<Vec<MathNode>>,
    },
    Radical {
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        index: Vec<MathNode>,
        radicand: Vec<MathNode>,
    },
    Limit {
        base: Vec<MathNode>,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        lower: Vec<MathNode>,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        upper: Vec<MathNode>,
    },
    Array {
        rows: Vec<Vec<Vec<MathNode>>>,
        /// "eq" (alternating right/left) | "center" (matrix) | "left"
        align: String,
    },
    #[serde(rename_all = "camelCase")]
    GroupChr {
        #[serde(rename = "char")]
        chr: String,
        pos: String,
        base: Vec<MathNode>,
    },
    Bar {
        pos: String,
        base: Vec<MathNode>,
    },
    Accent {
        #[serde(rename = "char")]
        chr: String,
        base: Vec<MathNode>,
    },
    Func {
        name: Vec<MathNode>,
        arg: Vec<MathNode>,
    },
    Group {
        items: Vec<MathNode>,
    },
}

/// Local-name child lookup (OMML elements live in the math namespace; this parser
/// matches on local names, mirroring how the rest of the docx parser works).
fn mchild<'a, 'i>(node: Node<'a, 'i>, name: &str) -> Option<Node<'a, 'i>> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == name)
}

fn mval(node: Node, child: &str) -> Option<String> {
    mchild(node, child).and_then(|n| {
        // The `m:val` attribute lives in the math namespace — Transitional or
        // Strict (ISO/IEC 29500) — falling back to the unqualified form.
        crate::ns::attr_ns(
            &n,
            crate::ns::math::TRANSITIONAL,
            crate::ns::math::STRICT,
            "val",
        )
        .map(|s| s.to_string())
    })
}

/// Parse the math children directly under `el` into a node list.
pub fn parse_omath_nodes(el: Node) -> Vec<MathNode> {
    let mut out = Vec::new();
    for child in el.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "r" => {
                let text = run_text(child);
                if !text.is_empty() {
                    out.push(MathNode::Run {
                        text,
                        style: run_style(child),
                    });
                }
            }
            "f" => out.push(MathNode::Fraction {
                num: nodes_in(child, "num"),
                den: nodes_in(child, "den"),
                bar: match mval(mchild(child, "fPr").unwrap_or(child), "type").as_deref() {
                    Some("noBar") => Some(false),
                    _ => None,
                },
            }),
            "sSup" => out.push(MathNode::Sup {
                base: nodes_in(child, "e"),
                sup: nodes_in(child, "sup"),
            }),
            "sSub" => out.push(MathNode::Sub {
                base: nodes_in(child, "e"),
                sub: nodes_in(child, "sub"),
            }),
            "sSubSup" => out.push(MathNode::SubSup {
                base: nodes_in(child, "e"),
                sub: nodes_in(child, "sub"),
                sup: nodes_in(child, "sup"),
            }),
            "nary" => {
                let pr = mchild(child, "naryPr");
                let op = pr
                    .and_then(|p| mval(p, "chr"))
                    .unwrap_or_else(|| "\u{222B}".to_string()); // default ∫
                let lim_loc = pr.and_then(|p| mval(p, "limLoc")).unwrap_or_default();
                out.push(MathNode::Nary {
                    op,
                    lim_loc,
                    sub: nodes_in(child, "sub"),
                    sup: nodes_in(child, "sup"),
                    body: nodes_in(child, "e"),
                });
            }
            "d" => {
                let pr = mchild(child, "dPr");
                let beg_char = pr
                    .and_then(|p| mval(p, "begChr"))
                    .unwrap_or_else(|| "(".to_string());
                let end_char = pr
                    .and_then(|p| mval(p, "endChr"))
                    .unwrap_or_else(|| ")".to_string());
                let items: Vec<Vec<MathNode>> = child
                    .children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "e")
                    .map(parse_omath_nodes)
                    .collect();
                out.push(MathNode::Delimiter {
                    beg_char,
                    end_char,
                    items,
                });
            }
            "rad" => out.push(MathNode::Radical {
                index: nodes_in(child, "deg"),
                radicand: nodes_in(child, "e"),
            }),
            "limLow" => out.push(MathNode::Limit {
                base: nodes_in(child, "e"),
                lower: nodes_in(child, "lim"),
                upper: Vec::new(),
            }),
            "limUpp" => out.push(MathNode::Limit {
                base: nodes_in(child, "e"),
                lower: Vec::new(),
                upper: nodes_in(child, "lim"),
            }),
            "m" => out.push(parse_matrix(child)),
            "eqArr" => out.push(parse_eqarr(child)),
            "groupChr" => {
                let pr = mchild(child, "groupChrPr");
                let pos = pr
                    .and_then(|p| mval(p, "pos"))
                    .unwrap_or_else(|| "bot".to_string());
                let chr = pr.and_then(|p| mval(p, "chr")).unwrap_or_else(|| {
                    if pos == "top" {
                        "\u{23DE}".to_string()
                    } else {
                        "\u{23DF}".to_string()
                    }
                });
                out.push(MathNode::GroupChr {
                    chr,
                    pos,
                    base: nodes_in(child, "e"),
                });
            }
            "bar" => {
                let pr = mchild(child, "barPr");
                let pos = pr
                    .and_then(|p| mval(p, "pos"))
                    .unwrap_or_else(|| "top".to_string());
                out.push(MathNode::Bar {
                    pos,
                    base: nodes_in(child, "e"),
                });
            }
            "acc" => {
                let pr = mchild(child, "accPr");
                let chr = pr
                    .and_then(|p| mval(p, "chr"))
                    .unwrap_or_else(|| "\u{0302}".to_string());
                out.push(MathNode::Accent {
                    chr,
                    base: nodes_in(child, "e"),
                });
            }
            "func" => out.push(MathNode::Func {
                name: nodes_in(child, "fName"),
                arg: nodes_in(child, "e"),
            }),
            // Containers / unknowns: descend so inner runs survive (degrade gracefully).
            _ => out.extend(parse_omath_nodes(child)),
        }
    }
    out
}

/// Flatten a math node list to its literal characters (for plain-text / markdown export).
pub fn nodes_to_text(nodes: &[MathNode]) -> String {
    let mut s = String::new();
    for n in nodes {
        match n {
            MathNode::Run { text, .. } => s.push_str(text),
            MathNode::Fraction { num, den, .. } => {
                s.push_str(&nodes_to_text(num));
                s.push('/');
                s.push_str(&nodes_to_text(den));
            }
            MathNode::Sup { base, sup } => {
                s.push_str(&nodes_to_text(base));
                s.push('^');
                s.push_str(&nodes_to_text(sup));
            }
            MathNode::Sub { base, sub } => {
                s.push_str(&nodes_to_text(base));
                s.push('_');
                s.push_str(&nodes_to_text(sub));
            }
            MathNode::SubSup { base, sub, sup } => {
                s.push_str(&nodes_to_text(base));
                s.push('_');
                s.push_str(&nodes_to_text(sub));
                s.push('^');
                s.push_str(&nodes_to_text(sup));
            }
            MathNode::Nary {
                op, sub, sup, body, ..
            } => {
                s.push_str(op);
                s.push_str(&nodes_to_text(sub));
                s.push_str(&nodes_to_text(sup));
                s.push_str(&nodes_to_text(body));
            }
            MathNode::Delimiter {
                beg_char,
                end_char,
                items,
            } => {
                s.push_str(beg_char);
                for (i, it) in items.iter().enumerate() {
                    if i > 0 {
                        s.push(',');
                    }
                    s.push_str(&nodes_to_text(it));
                }
                s.push_str(end_char);
            }
            MathNode::Radical { index, radicand } => {
                s.push('√');
                if !index.is_empty() {
                    s.push('[');
                    s.push_str(&nodes_to_text(index));
                    s.push(']');
                }
                s.push('(');
                s.push_str(&nodes_to_text(radicand));
                s.push(')');
            }
            MathNode::Func { name, arg } => {
                s.push_str(&nodes_to_text(name));
                s.push(' ');
                s.push_str(&nodes_to_text(arg));
            }
            MathNode::Group { items } => s.push_str(&nodes_to_text(items)),
            MathNode::Limit { base, lower, upper } => {
                s.push_str(&nodes_to_text(base));
                if !lower.is_empty() {
                    s.push('_');
                    s.push_str(&nodes_to_text(lower));
                }
                if !upper.is_empty() {
                    s.push('^');
                    s.push_str(&nodes_to_text(upper));
                }
            }
            MathNode::Array { rows, .. } => {
                for (ri, row) in rows.iter().enumerate() {
                    if ri > 0 {
                        s.push_str("; ");
                    }
                    for (ci, cell) in row.iter().enumerate() {
                        if ci > 0 {
                            s.push(' ');
                        }
                        s.push_str(&nodes_to_text(cell));
                    }
                }
            }
            MathNode::GroupChr { base, .. } => s.push_str(&nodes_to_text(base)),
            MathNode::Bar { base, .. } => s.push_str(&nodes_to_text(base)),
            MathNode::Accent { base, .. } => s.push_str(&nodes_to_text(base)),
        }
    }
    s
}

/// `m:m` matrix -> array of rows (m:mr) of cells (m:e), centered.
fn parse_matrix(el: Node) -> MathNode {
    let mut rows = Vec::new();
    for mr in el
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "mr")
    {
        let cells: Vec<Vec<MathNode>> = mr
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "e")
            .map(parse_omath_nodes)
            .collect();
        rows.push(cells);
    }
    MathNode::Array {
        rows,
        align: "center".to_string(),
    }
}

/// `m:eqArr` -> array where each `m:e` is a row, split into cells at `&` alignment marks.
fn parse_eqarr(el: Node) -> MathNode {
    let mut rows = Vec::new();
    for e in el
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "e")
    {
        rows.push(split_align_cells(parse_omath_nodes(e)));
    }
    MathNode::Array {
        rows,
        align: "eq".to_string(),
    }
}

/// Split a row's nodes into alignment cells at run-text `&` markers.
fn split_align_cells(nodes: Vec<MathNode>) -> Vec<Vec<MathNode>> {
    let mut cells: Vec<Vec<MathNode>> = vec![Vec::new()];
    for node in nodes {
        if let MathNode::Run { text, style } = &node {
            if text.contains('&') {
                for (i, part) in text.split('&').enumerate() {
                    if i > 0 {
                        cells.push(Vec::new());
                    }
                    if !part.is_empty() {
                        cells.last_mut().unwrap().push(MathNode::Run {
                            text: part.to_string(),
                            style: style.clone(),
                        });
                    }
                }
                continue;
            }
        }
        cells.last_mut().unwrap().push(node);
    }
    cells
}

/// Parse the math content of the named child element (`m:<name>`).
fn nodes_in(parent: Node, name: &str) -> Vec<MathNode> {
    match mchild(parent, name) {
        Some(n) => parse_omath_nodes(n),
        None => Vec::new(),
    }
}

/// Concatenate all `m:t` text under a math run.
fn run_text(r: Node) -> String {
    let mut s = String::new();
    for t in r
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "t")
    {
        for tn in t.children() {
            if let Some(txt) = tn.text() {
                s.push_str(txt);
            }
        }
    }
    s
}

/// Map `m:rPr/m:sty` (and `m:nor`) to the TS `MathStyle` strings. Math default is italic.
fn run_style(r: Node) -> String {
    let rpr = mchild(r, "rPr");
    if let Some(rpr) = rpr {
        if mchild(rpr, "nor").is_some() {
            return "roman".to_string();
        }
        match mval(rpr, "sty").as_deref() {
            Some("p") => return "roman".to_string(),
            Some("b") => return "bold".to_string(),
            Some("bi") => return "boldItalic".to_string(),
            Some("i") => return "italic".to_string(),
            _ => {}
        }
    }
    "italic".to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    const M: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    fn parse(xml: &str) -> Vec<MathNode> {
        let doc = roxmltree::Document::parse(xml).unwrap();
        parse_omath_nodes(doc.root_element())
    }

    #[test]
    fn parses_fraction_to_matching_json() {
        let xml = format!(
            r#"<m:oMath xmlns:m="{M}"><m:f>
              <m:num><m:r><m:t>1</m:t></m:r></m:num>
              <m:den><m:r><m:t>x</m:t></m:r></m:den>
            </m:f></m:oMath>"#
        );
        let nodes = parse(&xml);
        let json = serde_json::to_string(&nodes).unwrap();
        // Discriminator + field names must match core/src/types/math.ts.
        assert!(json.contains(r#""kind":"fraction""#), "json: {json}");
        assert!(json.contains(r#""kind":"run""#));
        assert!(json.contains(r#""text":"1""#));
        assert!(json.contains(r#""text":"x""#));
    }

    #[test]
    fn parses_nary_sum_with_limits() {
        let xml = format!(
            r#"<m:oMath xmlns:m="{M}"><m:nary>
              <m:naryPr><m:chr m:val="∑"/></m:naryPr>
              <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
              <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
              <m:e><m:r><m:t>i</m:t></m:r></m:e>
            </m:nary></m:oMath>"#
        );
        let nodes = parse(&xml);
        let json = serde_json::to_string(&nodes).unwrap();
        assert!(json.contains(r#""kind":"nary""#), "json: {json}");
        assert!(
            json.contains(r#""op":"∑""#) || json.contains("∑"),
            "json: {json}"
        );
    }

    // ISO/IEC 29500 Strict: OMML under the Strict math namespace
    // (`http://purl.oclc.org/ooxml/officeDocument/math`) parses identically —
    // notably `m:chr/@m:val` reaches the nary op via the `math::STRICT` branch
    // of `mval`. (Element matching is by local name and was already Strict-safe;
    // this covers the `m:val` attribute read.)
    #[test]
    fn strict_math_ns_nary_val_reaches_op() {
        let m_strict = crate::ns::math::STRICT;
        let xml = format!(
            r#"<m:oMath xmlns:m="{m_strict}"><m:nary>
              <m:naryPr><m:chr m:val="∑"/></m:naryPr>
              <m:e><m:r><m:t>i</m:t></m:r></m:e>
            </m:nary></m:oMath>"#
        );
        let nodes = parse(&xml);
        let json = serde_json::to_string(&nodes).unwrap();
        assert!(json.contains(r#""kind":"nary""#), "json: {json}");
        assert!(
            json.contains(r#""op":"∑""#),
            "Strict m:val must set op; json: {json}"
        );
    }

    #[test]
    fn eqarr_splits_rows_into_aligned_cells() {
        let xml = format!(
            r#"<m:oMath xmlns:m="{M}"><m:eqArr>
              <m:e><m:r><m:t>x&amp;=1+2+3</m:t></m:r></m:e>
              <m:e><m:r><m:t>&amp;=6</m:t></m:r></m:e>
            </m:eqArr></m:oMath>"#
        );
        let nodes = parse(&xml);
        match &nodes[0] {
            MathNode::Array { rows, align } => {
                assert_eq!(align, "eq");
                assert_eq!(rows.len(), 2);
                assert_eq!(rows[0].len(), 2); // "x" | "=1+2+3"
                assert_eq!(rows[1].len(), 2); // ""  | "=6"
                assert!(rows[1][0].is_empty());
            }
            other => panic!("expected array, got {other:?}"),
        }
        // no stray '&' leaks into text
        let json = serde_json::to_string(&nodes).unwrap();
        assert!(!json.contains('&') || !json.contains("\\u0026"));
    }

    #[test]
    fn parses_limlow() {
        let xml = format!(
            r#"<m:oMath xmlns:m="{M}"><m:limLow>
              <m:e><m:r><m:t>lim</m:t></m:r></m:e>
              <m:lim><m:r><m:t>n</m:t></m:r></m:lim>
            </m:limLow></m:oMath>"#
        );
        let nodes = parse(&xml);
        match &nodes[0] {
            MathNode::Limit { lower, upper, .. } => {
                assert!(!lower.is_empty());
                assert!(upper.is_empty());
            }
            other => panic!("expected limit, got {other:?}"),
        }
    }

    #[test]
    fn run_style_defaults_to_italic_and_honors_nor() {
        let xml = format!(
            r#"<m:oMath xmlns:m="{M}">
              <m:r><m:t>a</m:t></m:r>
              <m:r><m:rPr><m:nor/></m:rPr><m:t>b</m:t></m:r>
            </m:oMath>"#
        );
        let nodes = parse(&xml);
        match (&nodes[0], &nodes[1]) {
            (MathNode::Run { style: s0, .. }, MathNode::Run { style: s1, .. }) => {
                assert_eq!(s0, "italic");
                assert_eq!(s1, "roman");
            }
            _ => panic!("expected two runs, got {nodes:?}"),
        }
    }
}
