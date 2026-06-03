//! OMML (Office Math Markup Language, ECMA-376 §22.1) extraction into the shared
//! math AST consumed by `@silurus/ooxml-core`'s math engine. The serialized JSON
//! must match the TS `MathNode` types in `packages/core/src/types/math.ts`:
//! a `kind` discriminator with camelCase variant tags and fields.

use roxmltree::Node;
use serde::Serialize;

#[derive(Debug, Serialize, Clone)]
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
        #[serde(skip_serializing_if = "Option::is_none")]
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
    Nary {
        op: String,
        #[serde(skip_serializing_if = "Vec::is_empty")]
        sub: Vec<MathNode>,
        #[serde(skip_serializing_if = "Vec::is_empty")]
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
        #[serde(skip_serializing_if = "Vec::is_empty")]
        index: Vec<MathNode>,
        radicand: Vec<MathNode>,
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
        n.attribute((
            "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "val",
        ))
        .or_else(|| n.attribute("val"))
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
                out.push(MathNode::Nary {
                    op,
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
            MathNode::Nary { op, sub, sup, body } => {
                s.push_str(op);
                s.push_str(&nodes_to_text(sub));
                s.push_str(&nodes_to_text(sup));
                s.push_str(&nodes_to_text(body));
            }
            MathNode::Delimiter { beg_char, end_char, items } => {
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
        }
    }
    s
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
    for t in r.children().filter(|n| n.is_element() && n.tag_name().name() == "t") {
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
        assert!(json.contains(r#""op":"∑""#) || json.contains("∑"), "json: {json}");
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
