#!/usr/bin/env python3
"""Generate a synthetic .docx that exercises ECMA-376 §17.3.2.14 `w:fitText`
(Manual Run Width) inside §17.3.1.6 `w:bidi` (RTL) paragraphs, plus LTR controls.

This fixture carries NO private content — it is a self-contained, OSS-safe
document whose only purpose is to obtain Word-rendered ground truth for the
RTL fitText gap / residual placement fixed for issue #920.

How to use it for ground truth
------------------------------
1.  python3 packages/docx/scripts/gen-rtl-fittext-fixture.py rtl-fittext.docx
2.  Open `rtl-fittext.docx` in Microsoft Word and export / print to PDF.
3.  Compare against the viewer's render of the same file (Storybook private
    story or the node harness). Each fitText region must occupy exactly
    w:val = 2400 twips = 120 pt (1.667 in). Expected reading-frame layout:
      * RTL multi-run region: the glyphs read right-to-left with EVEN gaps
        between every adjacent pair across the run boundary; the logical-last
        run sits on the visual LEFT.
      * RTL single-char region: the single glyph sits at the region's LEADING
        (right) edge and the residual pad fills to its LEFT.
      * The LTR controls are the mirror image (glyph/pad on the left/right).

Every run property is emitted in CT_RPr (EG_RPrBase) schema order — in
particular `w:fitText` is written BEFORE `w:rtl` (wml.xsd lines 1773/1775) — so
a schema-strict round trip cannot mask a parser that ignores element order.
"""
import sys
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def rpr(fit_val=None, fit_id=None, rtl=False, cs_font="Arial", ea_font="Yu Mincho"):
    """A run property block in strict EG_RPrBase order: rFonts, sz, szCs,
    fitText, rtl. fitText MUST precede rtl (wml.xsd)."""
    parts = [f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="{cs_font}" w:eastAsia="{ea_font}"/>']
    parts.append('<w:sz w:val="22"/>')
    parts.append('<w:szCs w:val="22"/>')
    if fit_val is not None:
        if fit_id is not None:
            parts.append(f'<w:fitText w:val="{fit_val}" w:id="{fit_id}"/>')
        else:
            parts.append(f'<w:fitText w:val="{fit_val}"/>')
    if rtl:
        parts.append('<w:rtl/>')
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"


def run(text, **kw):
    # xml:space="preserve" so trailing/leading spaces survive the round trip.
    return f'<w:r>{rpr(**kw)}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(runs_xml, bidi=False):
    ppr = "<w:pPr><w:bidi/></w:pPr>" if bidi else ""
    return f"<w:p>{ppr}{runs_xml}</w:p>"


def label(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:rPr><w:i/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def build_document():
    body = []

    body.append(label("1) RTL multi-run fitText region (w:id=1, w:val=2400 = 120pt): logical-last run on the visual LEFT, even cross-run gaps"))
    body.append(para(
        run("أبجد", fit_val=2400, fit_id=1, rtl=True)   # أبجد (4)
        + run("هوز", fit_val=2400, fit_id=1, rtl=True),      # هوز (3)
        bidi=True,
    ))

    body.append(label("2) RTL single-char fitText region (w:val=2400 = 120pt): glyph at the LEADING (right) edge, pad to its LEFT"))
    body.append(para(
        run("م", fit_val=2400, rtl=True),                             # م
        bidi=True,
    ))

    body.append(label("3) LTR control multi-run fitText region (w:id=2, w:val=2400 = 120pt)"))
    body.append(para(
        run("氏名又は", fit_val=2400, fit_id=2)            # 氏名又は
        + run("名", fit_val=2400, fit_id=2)                           # 名
        + run("称", fit_val=2400, fit_id=2),                          # 称
    ))

    body.append(label("4) LTR control single-char fitText region (w:val=2400 = 120pt): glyph at the LEFT edge, pad to its RIGHT"))
    body.append(para(run("A", fit_val=2400)))

    sect = (
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
        '</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}">'
        f'<w:body>{"".join(body)}{sect}</w:body>'
        '</w:document>'
    )


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "rtl-fittext.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", build_document())
    print(f"wrote {out}")
    print("Open it in Word and export to PDF to capture RTL fitText ground truth (see module docstring).")


if __name__ == "__main__":
    main()
