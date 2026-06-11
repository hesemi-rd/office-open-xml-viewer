"""Generate pptx files whose slide background is a *tiled* blipFill.

ECMA-376 §20.1.8.14 (a:blipFill) + §20.1.8.58 (a:tile). python-pptx has no API
for a background blipFill, so we:
  1. add the tile image as a *picture* on the slide (this registers the image
     part + a slide relationship we can reuse), then delete the picture shape,
  2. hand-assemble <p:bg><p:bgPr><a:blipFill>…<a:tile/> referencing that rId.

The tile image is an asymmetric "F"-like glyph on a coloured ground so that the
mirror cadence of flip="x"/"y"/"xy" and the algn anchor are visually obvious.

Run: python3 _gen_bg_tile.py
Outputs (gitignored — public/private/ is local-only):
  bg-tile-basic.pptx   — 3 slides: no-flip / flip=x / flip=xy
  bg-tile-algn.pptx    — algn variants (tl / ctr / br) with a scaled tile
"""
import io
from pptx import Presentation
from pptx.util import Inches, Emu
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image, ImageDraw

# ---------------------------------------------------------------------------
# Build a small, deliberately asymmetric tile bitmap (64x48 px).
# ---------------------------------------------------------------------------
def make_tile_png() -> bytes:
    w, h = 64, 48
    img = Image.new("RGB", (w, h), (235, 200, 60))  # gold ground
    d = ImageDraw.Draw(img)
    # An "F" glyph: vertical bar + two horizontal arms (top + middle).
    blue = (25, 110, 202)
    d.rectangle([10, 6, 20, 42], fill=blue)      # vertical stem
    d.rectangle([10, 6, 48, 16], fill=blue)      # top arm (full width)
    d.rectangle([10, 22, 38, 30], fill=blue)     # middle arm (shorter)
    # corner dot marks the tile's top-left for orientation
    d.ellipse([2, 2, 8, 8], fill=(192, 0, 0))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


TILE_PNG = make_tile_png()
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def embed_tile_rid(slide) -> str:
    """Register the tile image as a slide relationship; return its rId.

    We add then immediately remove a picture so python-pptx wires up the image
    part + relationship for us (the rId stays valid for the bg blipFill).
    """
    pic = slide.shapes.add_picture(io.BytesIO(TILE_PNG), Inches(0), Inches(0),
                                   width=Inches(1), height=Inches(0.75))
    rid = pic._element.blipFill.find(qn("a:blip")).get(qn("r:embed"))
    # remove the visible picture; the relationship + part remain
    pic._element.getparent().remove(pic._element)
    return rid


def set_tile_background(slide, rid, *, tx=0, ty=0, sx=None, sy=None,
                        flip="none", algn=None):
    """Insert <p:bg> with a tiled blipFill into the slide's cSld."""
    cSld = slide._element.find(qn("p:cSld"))
    # remove any existing bg
    for old in cSld.findall(qn("p:bg")):
        cSld.remove(old)
    bg = etree.SubElement(cSld, qn("p:bg"))
    # cSld children are ordered: bg must come first.
    cSld.remove(bg)
    cSld.insert(0, bg)
    bgPr = etree.SubElement(bg, qn("p:bgPr"))
    blipFill = etree.SubElement(bgPr, qn("a:blipFill"))
    blip = etree.SubElement(blipFill, qn("a:blip"))
    blip.set(qn("r:embed"), rid)
    tile = etree.SubElement(blipFill, qn("a:tile"))
    if tx:
        tile.set("tx", str(int(tx)))
    if ty:
        tile.set("ty", str(int(ty)))
    if sx is not None:
        tile.set("sx", str(int(sx)))
    if sy is not None:
        tile.set("sy", str(int(sy)))
    tile.set("flip", flip)
    if algn:
        tile.set("algn", algn)
    # bgPr requires an effectLst (or effectDag) sibling per CT_BackgroundProperties
    etree.SubElement(bgPr, qn("a:effectLst"))


def new_prs():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


# ---------------------------------------------------------------------------
# File 1 — flip cadence: none / x / xy
# ---------------------------------------------------------------------------
prs = new_prs()
blank = prs.slide_layouts[6]
for flip in ("none", "x", "xy"):
    s = prs.slides.add_slide(blank)
    rid = embed_tile_rid(s)
    set_tile_background(s, rid, flip=flip)
prs.save("bg-tile-basic.pptx")
print("wrote bg-tile-basic.pptx (flip none / x / xy)")

# ---------------------------------------------------------------------------
# File 2 — algn anchors with a 50% scaled tile + a tx/ty offset slide
# ---------------------------------------------------------------------------
prs = new_prs()
blank = prs.slide_layouts[6]
for algn in ("tl", "ctr", "br"):
    s = prs.slides.add_slide(blank)
    rid = embed_tile_rid(s)
    set_tile_background(s, rid, sx=50000, sy=50000, algn=algn)
# extra slide: tx/ty offset (1 inch right, 0.5 inch down) at native size
s = prs.slides.add_slide(blank)
rid = embed_tile_rid(s)
set_tile_background(s, rid, tx=Emu(Inches(1)), ty=Emu(Inches(0.5)), algn="tl")
prs.save("bg-tile-algn.pptx")
print("wrote bg-tile-algn.pptx (algn tl / ctr / br + tx/ty)")
