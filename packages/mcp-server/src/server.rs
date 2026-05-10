use rmcp::{
    ServerHandler,
    handler::server::router::tool::ToolRouter,
    handler::server::wrapper::Parameters,
    model::{ServerCapabilities, ServerInfo},
    tool, tool_handler, tool_router,
};

use crate::tools::{
    docx::DocxTools,
    pptx::PptxTools,
    xlsx::XlsxTools,
};
use crate::tools::docx::{
    DocxImagesParam, DocxIndexParam, DocxPathParam, DocxSearchParam, DocxTableIndexParam,
};
use crate::tools::pptx::{
    PptxOptSlideParam, PptxPathParam, PptxPicturesParam, PptxRelationsParam, PptxSearchParam,
    PptxShapeParam, PptxSlideParam, PptxTextParam,
};
use crate::tools::xlsx::{
    XlsxCellRangeParam, XlsxChartIndexParam, XlsxOptSheetParam, XlsxPathParam, XlsxSearchParam,
    XlsxSheetParam,
};

#[derive(Clone)]
pub struct OoxmlServer {
    #[allow(dead_code)]
    tool_router: ToolRouter<OoxmlServer>,
}

#[tool_router]
impl OoxmlServer {
    pub fn new() -> Self {
        Self {
            tool_router: Self::tool_router(),
        }
    }

    // ── xlsx tools ────────────────────────────────────────────────────────────

    #[tool(description = "Parse an XLSX file and return workbook overview including sheet names and IDs")]
    fn xlsx_parse(&self, Parameters(p): Parameters<XlsxPathParam>) -> String {
        XlsxTools::xlsx_parse(Parameters(p))
    }

    #[tool(description = "Return only the list of sheet names from an XLSX file")]
    fn xlsx_get_sheet_names(&self, Parameters(p): Parameters<XlsxPathParam>) -> String {
        XlsxTools::xlsx_get_sheet_names(Parameters(p))
    }

    #[tool(description = "Return the dimensions (max row and column) of a worksheet")]
    fn xlsx_get_sheet_dimensions(&self, Parameters(p): Parameters<XlsxSheetParam>) -> String {
        XlsxTools::xlsx_get_sheet_dimensions(Parameters(p))
    }

    #[tool(description = "Return cell values and formulas for a given range (e.g. \"A1:C10\") in a worksheet")]
    fn xlsx_get_cell_range(&self, Parameters(p): Parameters<XlsxCellRangeParam>) -> String {
        XlsxTools::xlsx_get_cell_range(Parameters(p))
    }

    #[tool(description = "Return all cells that contain formulas in a worksheet")]
    fn xlsx_get_formulas(&self, Parameters(p): Parameters<XlsxSheetParam>) -> String {
        XlsxTools::xlsx_get_formulas(Parameters(p))
    }

    #[tool(description = "Search for a substring in cell values and formulas across one or all sheets of an XLSX file")]
    fn xlsx_search_cells(&self, Parameters(p): Parameters<XlsxSearchParam>) -> String {
        XlsxTools::xlsx_search_cells(Parameters(p))
    }

    #[tool(description = "List charts on a worksheet (or all sheets if `sheet` is omitted). Returns a summary per chart: anchor cell range, chart type, title, axes, legend, and a series outline (without numeric values)")]
    fn xlsx_get_charts(&self, Parameters(p): Parameters<XlsxOptSheetParam>) -> String {
        XlsxTools::xlsx_get_charts(Parameters(p))
    }

    #[tool(description = "Return one chart's full series data (categories and per-point values) for drill-down. `chartIndex` matches the index from `xlsx_get_charts` for the same sheet")]
    fn xlsx_get_chart_series(&self, Parameters(p): Parameters<XlsxChartIndexParam>) -> String {
        XlsxTools::xlsx_get_chart_series(Parameters(p))
    }

    #[tool(description = "Return all defined names (named ranges) visible in the workbook")]
    fn xlsx_get_named_ranges(&self, Parameters(p): Parameters<XlsxPathParam>) -> String {
        XlsxTools::xlsx_get_named_ranges(Parameters(p))
    }

    #[tool(description = "List Excel Tables (Ctrl+T tables, ECMA-376 §18.5) on a sheet or across all sheets")]
    fn xlsx_get_tables(&self, Parameters(p): Parameters<XlsxOptSheetParam>) -> String {
        XlsxTools::xlsx_get_tables(Parameters(p))
    }

    #[tool(description = "Return all merged cell ranges on a worksheet as A1 strings")]
    fn xlsx_get_merged_cells(&self, Parameters(p): Parameters<XlsxSheetParam>) -> String {
        XlsxTools::xlsx_get_merged_cells(Parameters(p))
    }

    #[tool(description = "Return conditional formatting rules on a worksheet (CellIs, Expression, ColorScale, DataBar, Top10, AboveAverage, IconSet, Other)")]
    fn xlsx_get_conditional_formats(&self, Parameters(p): Parameters<XlsxSheetParam>) -> String {
        XlsxTools::xlsx_get_conditional_formats(Parameters(p))
    }

    #[tool(description = "Return per-sheet layout: explicit column widths, row heights, freeze panes, gridline visibility, default sizes, and tab color")]
    fn xlsx_get_sheet_layout(&self, Parameters(p): Parameters<XlsxSheetParam>) -> String {
        XlsxTools::xlsx_get_sheet_layout(Parameters(p))
    }

    #[tool(description = "Return all `<dataValidation>` rules on a worksheet (ECMA-376 §18.3.1.32)")]
    fn xlsx_get_data_validations(&self, Parameters(p): Parameters<XlsxSheetParam>) -> String {
        XlsxTools::xlsx_get_data_validations(Parameters(p))
    }

    #[tool(description = "Return all comments (text + resolved author) on a worksheet, or across every sheet when `sheet` is omitted")]
    fn xlsx_get_comments(&self, Parameters(p): Parameters<XlsxOptSheetParam>) -> String {
        XlsxTools::xlsx_get_comments(Parameters(p))
    }

    // ── docx tools ────────────────────────────────────────────────────────────

    #[tool(description = "Extract all plain text from a DOCX file")]
    fn docx_extract_text(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_extract_text(Parameters(p))
    }

    #[tool(description = "Return the document structure (paragraphs and tables) of a DOCX file")]
    fn docx_get_structure(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_structure(Parameters(p))
    }

    #[tool(description = "Return all tables from a DOCX file with their cell contents")]
    fn docx_get_tables(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_tables(Parameters(p))
    }

    #[tool(description = "Search for a substring in all paragraph and table text of a DOCX file; returns matching excerpts with their position")]
    fn docx_search_text(&self, Parameters(p): Parameters<DocxSearchParam>) -> String {
        DocxTools::docx_search_text(Parameters(p))
    }

    #[tool(description = "Return one body element's full detail (paragraph or table) including run-level formatting, indents, spacing, numbering, and tab stops")]
    fn docx_get_paragraph(&self, Parameters(p): Parameters<DocxIndexParam>) -> String {
        DocxTools::docx_get_paragraph(Parameters(p))
    }

    #[tool(description = "Return the document's section properties (page size/margins/docGrid) along with default/first/even header and footer body elements")]
    fn docx_get_sections(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_sections(Parameters(p))
    }

    #[tool(description = "Return one table's full detail by index, including cell content, colSpan/vMerge, borders, shading, and row heights")]
    fn docx_get_table(&self, Parameters(p): Parameters<DocxTableIndexParam>) -> String {
        DocxTools::docx_get_table(Parameters(p))
    }

    #[tool(description = "List all images in the document. Set `includeDataUrl=true` to also receive the inline base64 image bytes (large)")]
    fn docx_get_images(&self, Parameters(p): Parameters<DocxImagesParam>) -> String {
        DocxTools::docx_get_images(Parameters(p))
    }

    #[tool(description = "List all drawn shapes embedded in paragraphs. Returns each shape's preset geometry, fill, stroke, dimensions, anchor offsets, rotation, and embedded text blocks")]
    fn docx_get_shapes(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_shapes(Parameters(p))
    }

    #[tool(description = "Return the heading outline of the document built from each paragraph's resolved `outlineLevel`")]
    fn docx_get_outline(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_outline(Parameters(p))
    }

    #[tool(description = "List all comments from word/comments.xml: id, author, initials, date, plain text")]
    fn docx_get_comments(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_comments(Parameters(p))
    }

    #[tool(description = "List footnote and endnote bodies from word/footnotes.xml and word/endnotes.xml")]
    fn docx_get_footnotes(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_footnotes(Parameters(p))
    }

    #[tool(description = "List all track-changes events (insertions and deletions) with author, date, and text")]
    fn docx_get_revisions(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_revisions(Parameters(p))
    }

    // ── pptx tools ────────────────────────────────────────────────────────────

    #[tool(description = "Return the number of slides and each slide's title from a PPTX file")]
    fn pptx_get_slides(&self, Parameters(p): Parameters<PptxPathParam>) -> String {
        PptxTools::pptx_get_slides(Parameters(p))
    }

    #[tool(description = "Extract plain text from a PPTX file; optionally filter to a single slide by 0-based index")]
    fn pptx_extract_text(&self, Parameters(p): Parameters<PptxTextParam>) -> String {
        PptxTools::pptx_extract_text(Parameters(p))
    }

    #[tool(description = "Return the structure (elements with position, size, text) of a single slide")]
    fn pptx_get_slide_structure(&self, Parameters(p): Parameters<PptxSlideParam>) -> String {
        PptxTools::pptx_get_slide_structure(Parameters(p))
    }

    #[tool(description = "Search for a substring across all text in a PPTX file; returns matching slide numbers and the text snippets that matched")]
    fn pptx_search_text(&self, Parameters(p): Parameters<PptxSearchParam>) -> String {
        PptxTools::pptx_search_text(Parameters(p))
    }

    #[tool(description = "Return one shape's full detail by slide and shape index. Includes geometry name, position/size, rotation/flip, fill, stroke (with arrow ends), adjustment values, effects, and text body")]
    fn pptx_get_shape(&self, Parameters(p): Parameters<PptxShapeParam>) -> String {
        PptxTools::pptx_get_shape(Parameters(p))
    }

    #[tool(description = "Return one shape's text body in detail: paragraphs with alignment, list level, bullets, and per-run formatting")]
    fn pptx_get_shape_text(&self, Parameters(p): Parameters<PptxShapeParam>) -> String {
        PptxTools::pptx_get_shape_text(Parameters(p))
    }

    #[tool(description = "List all charts on a slide (or every slide). Each entry exposes type, position, title, categories, and series")]
    fn pptx_get_charts(&self, Parameters(p): Parameters<PptxOptSlideParam>) -> String {
        PptxTools::pptx_get_charts(Parameters(p))
    }

    #[tool(description = "List all tables on a slide (or every slide), including column widths, row heights, and per-cell content with merge information")]
    fn pptx_get_tables(&self, Parameters(p): Parameters<PptxOptSlideParam>) -> String {
        PptxTools::pptx_get_tables(Parameters(p))
    }

    #[tool(description = "List all picture elements on a slide (or every slide). Returns metadata only by default; pass `includeDataUrl=true` to include the inline base64 bytes")]
    fn pptx_get_pictures(&self, Parameters(p): Parameters<PptxPicturesParam>) -> String {
        PptxTools::pptx_get_pictures(Parameters(p))
    }

    #[tool(description = "Return presentation-level metadata: slide width/height (EMU), slide count, default text color, theme major/minor fonts, and hyperlink colors")]
    fn pptx_get_presentation_meta(&self, Parameters(p): Parameters<PptxPathParam>) -> String {
        PptxTools::pptx_get_presentation_meta(Parameters(p))
    }

    #[tool(description = "Return speaker-notes text for one or all slides")]
    fn pptx_get_notes(&self, Parameters(p): Parameters<PptxOptSlideParam>) -> String {
        PptxTools::pptx_get_notes(Parameters(p))
    }

    #[tool(description = "Return legacy slide comments with text, author, and date")]
    fn pptx_get_comments(&self, Parameters(p): Parameters<PptxOptSlideParam>) -> String {
        PptxTools::pptx_get_comments(Parameters(p))
    }

    #[tool(description = "Infer geometric relations between shapes on a slide: connector hookups (with arrow direction when stroke ends are arrows), containment, overlap, and axis-aligned alignment groups. Detection is purely spatial — see `confidence: \"inferred\"` on each emitted relation")]
    fn pptx_get_shape_relations(&self, Parameters(p): Parameters<PptxRelationsParam>) -> String {
        PptxTools::pptx_get_shape_relations(Parameters(p))
    }
}

#[tool_handler]
impl ServerHandler for OoxmlServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo::new(
            ServerCapabilities::builder()
                .enable_tools()
                .build(),
        )
    }
}
