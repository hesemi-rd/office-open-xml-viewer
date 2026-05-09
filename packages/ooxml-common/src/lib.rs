//! Shared OOXML parsing helpers used by the docx and pptx Rust parsers.
//!
//! This crate is intentionally small — only logic that already had two
//! near-duplicate copies in the wild parsers belongs here. Anything
//! schema-specific (DocParagraph, ShapeRun, Slide, etc.) stays in the
//! consuming crate.

pub mod color;
