"""Generate a pptx exercising WordArt text warps (a:prstTxWarp) for VRT.

ECMA-376 §20.1.9.19 (ST_TextShapeType): a text body's <a:bodyPr> may carry a
<a:prstTxWarp prst="…"> that bends the text along one of 40 preset envelopes.
python-pptx does not expose prstTxWarp, so we reach into the lxml tree and add
the element to each shape's bodyPr by hand.

Run: python3 _gen_wordart_warp.py
Output: wordart-warp.pptx (gitignored — public/private/ is local-only).

Use with a local-only Samples story or the node render probe to eyeball the
per-glyph warp against PowerPoint's own rendering of the same file.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from lxml import etree

OUT = "wordart-warp.pptx"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def set_warp(shape, preset, adjs=None):
    """Add <a:prstTxWarp prst=preset><a:avLst>…</a:avLst> to the shape's bodyPr.

    `adjs` is an optional dict {gd_name: value} for avLst <a:gd fmla="val …"/>.
    """
    txBody = shape.text_frame._txBody
    bodyPr = txBody.find(qn("a:bodyPr"))
    # Remove any existing warp so re-runs are idempotent.
    for old in bodyPr.findall(qn("a:prstTxWarp")):
        bodyPr.remove(old)
    warp = etree.SubElement(bodyPr, qn("a:prstTxWarp"))
    warp.set("prst", preset)
    av = etree.SubElement(warp, qn("a:avLst"))
    if adjs:
        for name, val in adjs.items():
            gd = etree.SubElement(av, qn("a:gd"))
            gd.set("name", name)
            gd.set("fmla", f"val {int(val)}")


def add_wordart(slide, *, left, top, width, height, text, preset, adjs=None,
                size_pt=40, color="1F4E79", bold=True):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = __import__("pptx").dml.color.RGBColor.from_string(color)
    set_warp(tb, preset, adjs)
    return tb


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

# A representative sampler: one arch, one circle, a wave, an inflate/deflate,
# a cascade and a triangle — covering both warp families (single-edge arcs and
# paired-edge envelopes).
SAMPLES = [
    ("textArchUp", "Arch Up", None),
    ("textArchDown", "Arch Down", None),
    ("textCircle", "Circle", None),
    ("textWave1", "Wave One", {"adj1": 12500, "adj2": 10000}),
    ("textInflate", "Inflate", None),
    ("textDeflate", "Deflate", None),
    ("textCascadeUp", "Cascade", None),
    ("textChevron", "Chevron", None),
]

s = prs.slides.add_slide(blank)
title = s.shapes.add_textbox(Inches(0.4), Inches(0.1), Inches(12.5), Inches(0.6))
tp = title.text_frame.paragraphs[0]
r = tp.add_run()
r.text = "WordArt text warps (a:prstTxWarp)"
r.font.size = Pt(24)
r.font.bold = True

cols = 2
cell_w = Inches(6.2)
cell_h = Inches(1.5)
for i, (preset, text, adjs) in enumerate(SAMPLES):
    col = i % cols
    row = i // cols
    add_wordart(
        s,
        left=Inches(0.5) + col * (cell_w + Inches(0.2)),
        top=Inches(0.9) + row * (cell_h + Inches(0.05)),
        width=cell_w,
        height=cell_h,
        text=text,
        preset=preset,
        adjs=adjs,
    )

prs.save(OUT)
print(f"wrote {OUT} with {len(SAMPLES)} WordArt warps")
