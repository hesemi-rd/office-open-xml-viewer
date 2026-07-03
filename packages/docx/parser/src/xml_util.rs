use ooxml_common::ns::{attr_ns, is_w_ns, wordprocessingml};
use roxmltree::Node;

/// Transitional WordprocessingML URI, used only by test fixtures that build
/// `w:`-namespaced XML; runtime matching goes through [`is_w_ns`], which accepts
/// the Strict URI too.
#[cfg(test)]
pub const W_NS: &str = wordprocessingml::TRANSITIONAL;
/// Transitional relationships URI. See [`W_NS`] for why the Transitional value is
/// kept for tests while runtime matching accepts either class.
#[cfg(test)]
pub const R_NS: &str = ooxml_common::ns::relationships::TRANSITIONAL;

/// Find first child in w: namespace.
pub fn child_w<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Option<Node<'a, 'input>> {
    node.children()
        .find(|n| n.tag_name().name() == name && is_w_ns(n.tag_name().namespace()))
}

/// Collect all children in w: namespace with given name.
pub fn children_w<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Vec<Node<'a, 'input>> {
    node.children()
        .filter(|n| n.tag_name().name() == name && is_w_ns(n.tag_name().namespace()))
        .collect()
}

/// Element children with <w:sdt> wrappers transparently unwrapped. Structured Document
/// Tag (content control) blocks contain their real content inside <w:sdtContent>, and
/// most parsing stages should treat them as inline with the surrounding context.
pub fn element_children_flat<'a, 'input>(node: Node<'a, 'input>) -> Vec<Node<'a, 'input>> {
    let mut out = Vec::new();
    for child in node.children().filter(|n| n.is_element()) {
        let tn = child.tag_name();
        if is_w_ns(tn.namespace()) && tn.name() == "sdt" {
            if let Some(content) = child_w(child, "sdtContent") {
                out.extend(element_children_flat(content));
            }
        } else {
            out.push(child);
        }
    }
    out
}

/// Like children_w but transparently descends into <w:sdt>/<w:sdtContent> wrappers.
pub fn children_w_flat<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Vec<Node<'a, 'input>> {
    element_children_flat(node)
        .into_iter()
        .filter(|n| n.tag_name().name() == name && is_w_ns(n.tag_name().namespace()))
        .collect()
}

/// Get attribute in w: namespace (Transitional or Strict), falling back to
/// no-namespace.
pub fn attr_w(node: Node, name: &str) -> Option<String> {
    attr_ns(
        &node,
        wordprocessingml::TRANSITIONAL,
        wordprocessingml::STRICT,
        name,
    )
    .map(|s| s.to_string())
}

/// Parse twips (1/20 pt) string to f64 pt.
pub fn twips_to_pt(s: &str) -> f64 {
    s.parse::<f64>().unwrap_or(0.0) / 20.0
}

/// Parse half-points string to f64 pt.
pub fn half_pt_to_pt(s: &str) -> f64 {
    s.parse::<f64>().unwrap_or(0.0) / 2.0
}

/// Parse a ST_OnOff-style toggle child element. ECMA-376 §17.3.2.22 allows
/// "true"/"false"/"1"/"0"/"on"/"off" (and absent val attribute = true).
/// Returns None if the element itself is absent so the caller can distinguish
/// "explicitly turned off" from "inherited from parent".
pub fn bool_prop(node: Node, tag: &str) -> Option<bool> {
    let child = child_w(node, tag)?;
    let val = attr_w(child, "val");
    Some(!matches!(
        val.as_deref(),
        Some("0") | Some("false") | Some("off")
    ))
}
