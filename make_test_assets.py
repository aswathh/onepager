"""
Builds TEST assets only (you will use your real template):
  test_assets/template_test.pptx           tokenised copy of the layout in your photo
  test_assets/template_test_mirrored.pptx  same tokens, shapes moved around -> proves template changes still fill
  sample/logos/*.png                       placeholder logos (different aspect ratios, one with a white border)
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn
from lxml import etree
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BLUE = RGBColor(0x0A, 0x5D, 0xC8)
W = 13.333


def logos(out_dir: Path):
    out_dir.mkdir(parents=True, exist_ok=True)
    specs = {  # file: (text, bg, fg, size, white_margin)
        "assessli": ("Ac", (0, 0, 0), (255, 255, 255), (500, 500), 0),
        "microsoft": ("Microsoft", (255, 255, 255), (60, 60, 60), (520, 140), 30),
        "sap": ("SAP", (0, 112, 192), (255, 255, 255), (300, 150), 0),
        "notion": ("N", (0, 0, 0), (255, 255, 255), (260, 260), 0),
        "coursera": ("coursera", (0, 86, 210), (255, 255, 255), (500, 160), 0),
        "upgrad": ("upGrad", (220, 40, 50), (255, 255, 255), (420, 200), 0),
        "wysa": ("wysa", (255, 255, 255), (150, 150, 200), (460, 150), 25),
        "smarteye": ("smart eye", (255, 255, 255), (60, 60, 60), (480, 140), 20),
        "panda": ("panda", (120, 160, 220), (0, 0, 0), (300, 240), 0),
        "vw": ("V", (255, 255, 255), (20, 70, 160), (250, 250), 30),
        "dotsin": ("d", (255, 255, 255), (20, 70, 200), (260, 260), 30),
    }
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 64)
    except OSError:
        font = ImageFont.load_default()
    for fn, (text, bg, fg, size, margin) in specs.items():
        im = Image.new("RGB", size, bg)
        d = ImageDraw.Draw(im)
        box = d.textbbox((0, 0), text, font=font)
        d.text(((size[0] - box[2]) / 2, (size[1] - box[3]) / 2), text, fill=fg, font=font)
        if margin:  # white border so the trimmer has something to do
            canvas = Image.new("RGB", (size[0] + 2 * margin, size[1] + 2 * margin), (255, 255, 255))
            canvas.paste(im, (margin, margin))
            im = canvas
        im.save(out_dir / f"{fn}.png")


FONT = "Intel One Display"   # falls back to the default font if it is not installed
LINE = "8EA9C8"              # thin steel-blue grid lines used in the admin template
DARK = RGBColor(0x22, 0x22, 0x22)


def _cell_borders(cell, color=LINE, width_pt=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    for i, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        for old in tcPr.findall(qn(tag)):
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag), w=str(int(width_pt * 12700)), cap="flat", cmpd="sng", algn="ctr")
        fill = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr"), val=color)
        etree.SubElement(ln, qn("a:prstDash"), val="solid")
        tcPr.remove(ln)
        tcPr.insert(i, ln)  # borders must precede the fill element


def _box(slide, x, y, w, h, text, size=12, bold=False, color=None, align=PP_ALIGN.CENTER, fill=None, name=None,
         border=None, dashed=False, hanging=False, underline=False, anchor=MSO_ANCHOR.MIDDLE):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)) if (fill or border) else \
        slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if fill or border:
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = fill
        else:
            shp.fill.background()
        if border:
            shp.line.color.rgb = RGBColor.from_string(border)
            shp.line.width = Pt(0.75)
            if dashed:
                shp.line.dash_style = MSO_LINE.DASH
        else:
            shp.line.fill.background()
        shp.shadow.inherit = False
    if name:
        shp.name = name
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if hanging:  # first line sticks out, like the product description in the admin template
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(Inches(0.4)))
        pPr.set("indent", str(-Inches(0.4)))
    r = p.add_run()
    r.text = text
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.underline = underline
    r.font.color.rgb = color or RGBColor(0, 0, 0)
    return shp


def _table(slide, x, y, w, h, grid, col_w, row_h=None, label_cols=(0,), size=11, merges=(), label_size=10, name=None):
    rows, cols = len(grid), len(grid[0])
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        gf.name = name
    t = gf.table
    tblPr = gf._element.graphic.graphicData.tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    for i, cw in enumerate(col_w):
        t.columns[i].width = Inches(cw)
    for i in range(rows):
        t.rows[i].height = Inches(row_h[i] if row_h else h / rows)
    for (r0, c0, r1, c1) in merges:
        t.cell(r0, c0).merge(t.cell(r1, c1))
    for r in range(rows):
        for c in range(cols):
            cell = t.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
            _cell_borders(cell)
            if cell.is_spanned:
                continue
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.05)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = grid[r][c]
            is_label = c in label_cols
            run.font.name = FONT
            run.font.size = Pt(label_size if is_label else size)
            run.font.bold = is_label
            run.font.color.rgb = DARK
    return gf


def build(path: Path, mirror: bool = False, admin: bool = False):
    """Replica of the admin's 'AI ISV Partnerships' one-pager, with {{tokens}} where the data goes."""
    # admin=True -> NO tokens: sample data + sample logo pictures, exactly like a finished slide from the admin.
    # (use it to test --init-map / label-based filling)
    SAMPLE = {"company_name": "Assessli", "website": "https://assessli.com", "hq": "Bengaluru, Karnataka", "employees": "40",
              "funding": "400K USD", "revenue": "$5.2M", "product_name": "Dotsin", "vertical_market": "EdTech, HRTech, Healthcare",
              "route_to_market": "Direct SaaS-led", "inference_hardware": "Pytorch", "deployment_type": "Cloud",
              "training_hardware": "NVIDIA H100 /A100", "models": "BioBERT +LORA, GNN, TGN", "ai_framework": "Pytorch",
              "additional_ai_stack": "Huggingface, unsloth, OpenVino, SGLang", "serving_framework": "PyTorch, ONNX",
              "company_description": "Assessli founded in 2022 is a It focused on building next-generation behavioral intelligence through its proprietary technology called Large Behavioral Models (LBMs).",
              "product_description": "Dots-In is a foundational AI platform built on a Large Behavioral Model (LBM) that learns how humans behave, decide, and evolve over time."}

    def T(key):
        return SAMPLE[key] if admin else "{{" + key + "}}"

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6])

    def X(x, w):  # mirror = swap left/right to prove a changed template still fills
        return W - x - w if mirror else x

    # ---------------- left half
    _box(s, X(0.45, 2.2), 0.45, 2.2, 1.75, "" if admin else "{{logo:company_logo}}", name="company_logo_slot")
    _box(s, X(2.65, 3.95), 0.4, 3.95, 0.75, T("company_name"), 40)
    _box(s, X(2.65, 3.95), 1.12, 3.95, 0.35, T("website"), 13, color=RGBColor(0x1B, 0x9B, 0xD6), underline=True)
    _table(s, X(2.65, 3.95), 1.68, 3.95, 0.57, [["HQ", T("hq"), "Employees", T("employees")]],
           [0.55, 1.65, 0.95, 0.8], size=13, label_cols=(0, 2), label_size=10, name="table_hq_employees")
    _table(s, X(0.45, 6.15), 2.25, 6.15, 0.62, [["Funding", T("funding"), "Revenue", T("revenue")]],
           [0.8, 2.3, 1.1, 1.95], size=13, label_cols=(0, 2), label_size=10, name="table_funding_revenue")
    _box(s, X(0.45, 6.15), 3.15, 6.15, 1.5, T("company_description"), 13)
    grid = [[lbl] + ["" if admin else f"{{{{logo:{k}[{i}]}}}}" for i in range(4)] for lbl, k in
            (("GTM Partners", "gtm_partners"), ("Key Customers", "key_customers"), ("Main Competitors", "main_competitors"))]
    _table(s, X(0.55, 6.05), 4.7, 6.05, 2.15, grid, [1.05, 1.25, 1.25, 1.25, 1.25], row_h=[0.72, 0.72, 0.71],
           size=11, label_size=10, name="table_logos")
    if admin:  # sample logos sit ON TOP of the table cells, as pictures (a table cell cannot hold a picture)
        from PIL import Image as _I
        lg = Path(__file__).parent / "sample" / "logos"
        rows_ = [["microsoft", "sap", "notion", "dotsin"], ["coursera", "upgrad", "vw", "vw"], ["wysa", "smarteye", "panda", "vw"]]
        for ri_, names in enumerate(rows_):
            for ci_, nm in enumerate(names):
                cx, cy, cw, ch = X(0.55, 6.05) + 1.05 + ci_ * 1.25 if not mirror else X(0.55, 6.05) + 0.0 + (3 - ci_) * 1.25, 4.7 + ri_ * 0.72, 1.25, 0.72
                iw, ih = _I.open(lg / f"{nm}.png").size
                sc = min((cw - 0.1) / iw, (ch - 0.1) / ih)
                s.shapes.add_picture(str(lg / f"{nm}.png"), Inches(cx + (cw - iw * sc) / 2), Inches(cy + (ch - ih * sc) / 2), Inches(iw * sc), Inches(ih * sc))
        pw = 1.75
        s.shapes.add_picture(str(lg / "assessli.png"), Inches(X(0.45, 2.2) + 0.225), Inches(0.45), Inches(1.75), Inches(1.75))

    # ---------------- right half
    _box(s, X(6.85, 6.0), 0.5, 6.0, 0.8, T("product_name"), 28, bold=True, color=RGBColor(255, 255, 255), fill=RGBColor(0x00, 0x53, 0xC8))
    _box(s, X(6.85, 6.0), 1.3, 6.0, 1.95, T("product_description"), 12, align=PP_ALIGN.LEFT, border=LINE, hanging=True)
    rows = [
        ["Vertical Market", T("vertical_market"), "Route to Market", T("route_to_market")],
        ["Current Inference Hardware", T("inference_hardware"), "Deployment Type", T("deployment_type")],
        ["Current Training Hardware", T("training_hardware"), "Model/s:", T("models")],
        ["Current AI Framework", T("ai_framework"), "Additional AI Software Stack", T("additional_ai_stack")],
        ["Current Serving Framework", T("serving_framework"), "", ""],
    ]
    _table(s, X(6.85, 6.0), 3.25, 6.0, 3.6, rows, [1.05, 1.95, 1.05, 1.95], row_h=[0.8, 0.75, 0.75, 0.65, 0.65],
           label_cols=(0, 2), size=11, label_size=10, merges=[(3, 2, 4, 2), (3, 3, 4, 3)], name="table_product_details")

    # ---------------- footer
    _box(s, 0.4, 7.05, 0.6, 0.3, "intel", 12, bold=True, color=RGBColor(0x00, 0x71, 0xC5), align=PP_ALIGN.LEFT)
    _box(s, 1.05, 7.05, 3.0, 0.3, "AI ISV Partnerships", 10, align=PP_ALIGN.LEFT)
    _box(s, 4.9, 6.98, 3.5, 0.38, "Footer", 10, color=RGBColor(0x80, 0x80, 0x80), border="808080", dashed=True)
    _box(s, W - 3.2, 7.05, 2.3, 0.3, "Intel Confidential", 10, align=PP_ALIGN.RIGHT)
    _box(s, W - 0.8, 7.05, 0.4, 0.3, "1", 10, align=PP_ALIGN.RIGHT)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))


if __name__ == "__main__":
    here = Path(__file__).parent
    logos(here / "sample" / "logos")
    build(here / "test_assets" / "template_test.pptx")
    build(here / "test_assets" / "template_test_mirrored.pptx", mirror=True)
    build(here / "test_assets" / "template_admin_style.pptx", admin=True)  # no tokens: use --init-map
    print("test assets written")
