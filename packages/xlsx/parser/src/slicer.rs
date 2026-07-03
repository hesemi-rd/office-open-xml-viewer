use crate::resolve_zip_path;
use crate::types::*;
use ooxml_common::ns::{is_x_ns, is_xdr_ns};
use ooxml_common::zip::read_zip_string;
use std::collections::HashMap;

// ─── Slicer loading ─────────────────────────────────────────────────────────
//
// Office 2010+ extension (`sle:slicer` inside `<mc:AlternateContent>`).
// Resolving one slicer graphicFrame into a drawable anchor takes four
// XML files:
//   1. The sheet's drawing (for the anchor rect + graphicFrame name).
//   2. `xl/slicers/slicerN.xml` — slicer definition: graphicFrame name →
//      caption + cache name.
//   3. `xl/slicerCaches/slicerCacheN.xml` — cache definition: cache name →
//      source field + list of (item index, selected?).
//   4. `xl/pivotCache/pivotCacheDefinitionN.xml` — pivot cache: field name →
//      ordered string values.
// Excel also allows slicers bound to Excel Tables (`tableSlicerCache`), but
// the present sample is pivot-only; we only implement the pivot path.

#[derive(Default)]
pub(crate) struct SlicerCacheInfo {
    source_name: String,
    items: Vec<(u32, bool)>, // (index into pivot field, selected)
}

#[derive(Default)]
pub(crate) struct PivotCacheFields {
    by_name: HashMap<String, Vec<String>>, // field name → ordered string items
}

/// Parse every `xl/pivotCache/pivotCacheDefinition*.xml` and merge its
/// cacheFields (indexed by `@name`) into a single map. Sample workbooks
/// typically have one pivotCache but the loop keeps the code general.
pub(crate) fn load_all_pivot_cache_fields(archive: &mut crate::XlsxZip) -> PivotCacheFields {
    let mut out = PivotCacheFields::default();
    let names: Vec<String> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
        .filter(|n| n.starts_with("xl/pivotCache/pivotCacheDefinition") && n.ends_with(".xml"))
        .collect();
    for name in names {
        let Ok(xml) = read_zip_string(archive, &name) else {
            continue;
        };
        let Ok(doc) = roxmltree::Document::parse(&xml) else {
            continue;
        };
        for field in doc.descendants() {
            if field.tag_name().name() != "cacheField" || !is_x_ns(field.tag_name().namespace()) {
                continue;
            }
            let Some(field_name) = field.attribute("name") else {
                continue;
            };
            let mut items: Vec<String> = Vec::new();
            for shared in field
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "sharedItems")
            {
                for item in shared.children().filter(|n| n.is_element()) {
                    match item.tag_name().name() {
                        "s" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "n" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "d" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "b" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "m" => items.push(String::new()),
                        _ => {}
                    }
                }
            }
            if !items.is_empty() {
                out.by_name.insert(field_name.to_string(), items);
            }
        }
    }
    out
}

/// Parse every `xl/slicerCaches/slicerCache*.xml` and build a map keyed by
/// the slicerCache's `@name` attribute (e.g. `"スライサー_贈答相手1"`). That
/// name is what `<slicer cache="…"/>` in `xl/slicers/slicerN.xml` references.
pub(crate) fn load_all_slicer_caches(
    archive: &mut crate::XlsxZip,
) -> HashMap<String, SlicerCacheInfo> {
    let mut out: HashMap<String, SlicerCacheInfo> = HashMap::new();
    let names: Vec<String> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
        .filter(|n| n.starts_with("xl/slicerCaches/slicerCache") && n.ends_with(".xml"))
        .collect();
    for path in names {
        let Ok(xml) = read_zip_string(archive, &path) else {
            continue;
        };
        let Ok(doc) = roxmltree::Document::parse(&xml) else {
            continue;
        };
        let root = doc.root_element();
        let cache_name = root.attribute("name").unwrap_or("").to_string();
        let source_name = root.attribute("sourceName").unwrap_or("").to_string();
        let mut items: Vec<(u32, bool)> = Vec::new();
        for tabular in doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "tabular")
        {
            for i_el in tabular
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "i")
            {
                let x: u32 = i_el
                    .attribute("x")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                // `s` defaults to "1" (selected) when absent — ECMA-376
                // extension schema for slicer caches.
                let selected = i_el.attribute("s").map(|v| v != "0").unwrap_or(true);
                items.push((x, selected));
            }
        }
        if !cache_name.is_empty() {
            out.insert(cache_name, SlicerCacheInfo { source_name, items });
        }
    }
    out
}

/// Slicer definition (`xl/slicers/slicerN.xml`): maps each graphicFrame name
/// on the sheet to its display caption and the slicerCache it's backed by.
#[derive(Default)]
pub(crate) struct SlicerDef {
    caption: String,
    cache: String,
}

pub(crate) fn parse_slicers_xml(xml: &str) -> HashMap<String, SlicerDef> {
    let mut out: HashMap<String, SlicerDef> = HashMap::new();
    let Ok(doc) = roxmltree::Document::parse(xml) else {
        return out;
    };
    for slicer in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "slicer")
    {
        let name = slicer.attribute("name").unwrap_or("").to_string();
        let caption = slicer.attribute("caption").unwrap_or("").to_string();
        let cache = slicer.attribute("cache").unwrap_or("").to_string();
        if !name.is_empty() {
            out.insert(name, SlicerDef { caption, cache });
        }
    }
    out
}

pub(crate) fn load_sheet_slicers(
    archive: &mut crate::XlsxZip,
    sheet_path: &str, // e.g. "worksheets/sheet1.xml"
) -> Vec<SlicerAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_string(archive, &sheet_rels_path) else {
        return Vec::new();
    };
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else {
        return Vec::new();
    };

    // 1. Collect slicer-definition and drawing targets from the sheet rels.
    let mut drawing_targets: Vec<String> = Vec::new();
    let mut slicer_targets: Vec<String> = Vec::new();
    for rel in rels_doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
    {
        let rel_type = rel.attribute("Type").unwrap_or("");
        let Some(target) = rel.attribute("Target") else {
            continue;
        };
        if rel_type.ends_with("/drawing") {
            drawing_targets.push(target.to_string());
        } else if rel_type.ends_with("/slicer") {
            slicer_targets.push(target.to_string());
        }
    }
    if drawing_targets.is_empty() || slicer_targets.is_empty() {
        return Vec::new();
    }

    // 2. Parse all slicer definitions referenced by this sheet, keyed by
    //    graphicFrame name.
    let mut slicer_defs: HashMap<String, SlicerDef> = HashMap::new();
    for target in &slicer_targets {
        let slicer_path = resolve_zip_path(&format!("xl/{}", sheet_dir), target);
        let Ok(xml) = read_zip_string(archive, &slicer_path) else {
            continue;
        };
        for (k, v) in parse_slicers_xml(&xml) {
            slicer_defs.insert(k, v);
        }
    }
    if slicer_defs.is_empty() {
        return Vec::new();
    }

    // 3. Resolve caches (and their backing pivot fields) once.
    let slicer_caches = load_all_slicer_caches(archive);
    let pivot_fields = load_all_pivot_cache_fields(archive);

    // 4. Walk each drawing and pick up slicer graphicFrames.
    let mut out: Vec<SlicerAnchor> = Vec::new();
    for target in drawing_targets {
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_string(archive, &drawing_path) else {
            continue;
        };
        out.extend(parse_slicer_anchors(
            &drawing_xml,
            &slicer_defs,
            &slicer_caches,
            &pivot_fields,
        ));
    }
    out
}

pub(crate) fn parse_slicer_anchors(
    drawing_xml: &str,
    slicer_defs: &HashMap<String, SlicerDef>,
    slicer_caches: &HashMap<String, SlicerCacheInfo>,
    pivot_fields: &PivotCacheFields,
) -> Vec<SlicerAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else {
        return Vec::new();
    };
    let mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let slicer_uri = "http://schemas.microsoft.com/office/drawing/2010/slicer";
    let mut out: Vec<SlicerAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor" || !is_xdr_ns(anchor.tag_name().namespace())
        {
            continue;
        }

        // Anchor rect.
        let mut from = (0u32, 0i64, 0u32, 0i64);
        let mut to = (0u32, 0i64, 0u32, 0i64);
        for child in anchor.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "from" | "to" => {
                    let is_from = child.tag_name().name() == "from";
                    let mut col: u32 = 0;
                    let mut col_off: i64 = 0;
                    let mut row: u32 = 0;
                    let mut row_off: i64 = 0;
                    for c in child.children() {
                        match (c.tag_name().name(), c.text()) {
                            ("col", Some(t)) => col = t.trim().parse().unwrap_or(0),
                            ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                            ("row", Some(t)) => row = t.trim().parse().unwrap_or(0),
                            ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    if is_from {
                        from = (col, col_off, row, row_off);
                    } else {
                        to = (col, col_off, row, row_off);
                    }
                }
                _ => {}
            }
        }

        // Slicers live inside `<mc:AlternateContent><mc:Choice>` — descend
        // until we find a `<xdr:graphicFrame>` whose graphicData uri is the
        // 2010 slicer namespace, then harvest the graphicFrame's cNvPr name.
        let Some(frame_name) = anchor
            .descendants()
            .filter(|n| {
                n.is_element()
                    && n.tag_name().name() == "Choice"
                    && n.tag_name().namespace() == Some(mc_ns)
            })
            .flat_map(|choice| choice.descendants())
            .find_map(|n| {
                if n.is_element()
                    && n.tag_name().name() == "graphicData"
                    && n.attribute("uri") == Some(slicer_uri)
                {
                    // graphicData → ancestor graphicFrame → nvGraphicFramePr → cNvPr
                    let mut p = n.parent();
                    while let Some(pp) = p {
                        if pp.tag_name().name() == "graphicFrame" {
                            break;
                        }
                        p = pp.parent();
                    }
                    let frame = p?;
                    let cnvpr = frame
                        .descendants()
                        .find(|d| d.is_element() && d.tag_name().name() == "cNvPr")?;
                    cnvpr.attribute("name").map(|s| s.to_string())
                } else {
                    None
                }
            })
        else {
            continue;
        };

        let Some(slicer_def) = slicer_defs.get(&frame_name) else {
            continue;
        };

        // Resolve items via cache → pivot field; fall back to an empty list
        // if any link is broken (still renders the header and box).
        let items: Vec<SlicerItem> = slicer_caches
            .get(&slicer_def.cache)
            .map(|cache| {
                let field_items = pivot_fields.by_name.get(&cache.source_name);
                cache
                    .items
                    .iter()
                    .map(|(x, selected)| {
                        let name = field_items
                            .and_then(|list| list.get(*x as usize))
                            .cloned()
                            .unwrap_or_default();
                        SlicerItem {
                            name,
                            selected: *selected,
                        }
                    })
                    .collect()
            })
            .unwrap_or_default();

        let caption = if !slicer_def.caption.is_empty() {
            slicer_def.caption.clone()
        } else {
            frame_name.clone()
        };

        out.push(SlicerAnchor {
            from_col: from.0,
            from_col_off: from.1,
            from_row: from.2,
            from_row_off: from.3,
            to_col: to.0,
            to_col_off: to.1,
            to_row: to.2,
            to_row_off: to.3,
            caption,
            items,
        });
    }
    out
}

// ─── Chart loading ──────────────────────────────────────────────────────────
