"""uv run run_tests.py  -> end-to-end checks for fill_template.py + logo_extractor.py (no chatbot module needed)."""
import json, sys
from pathlib import Path
from pptx import Presentation
from fill_template import fill_template, validate

here = Path(__file__).parent
T, S, O = here / "test_assets/template_test.pptx", here / "sample", here / "out"
ok = True
def check(name, cond):
    global ok; ok &= bool(cond); print(("PASS " if cond else "FAIL ") + name)

def slide_text(p):
    out = []
    def walk(shapes):
        for s in shapes:
            if s.shape_type == 6: walk(s.shapes)
            elif getattr(s, "has_table", False) and s.has_table:
                out.extend(c.text for r in s.table.rows for c in r.cells)
            elif s.has_text_frame: out.append(s.text_frame.text)
    walk(Presentation(str(p)).slides[0].shapes)
    return "\n".join(out)

def pictures(p): return sum(s.shape_type == 13 for s in Presentation(str(p)).slides[0].shapes)

# 1. valid schema
v = validate(json.load(open(S / "assessli.json")))
check("valid schema has no validation errors", v.ok)
rep = fill_template(T, S / "assessli.json", O / "t_valid.pptx", cache_dir=O / ".c")
txt = slide_text(O / "t_valid.pptx")
check("no leftover {{tokens}}", "{{" not in txt)
check("text fields filled", all(x in txt for x in ["Assessli", "Bengaluru, Karnataka", "400K USD", "$5.2M", "Direct SaaS-led", "BioBERT + LoRA, GNN, TGN"]))
check("list value joined with commas", "EdTech, HRTech, Healthcare, AI Personalized Consumer AI" in txt)
check("12 logo pictures placed", pictures(O / "t_valid.pptx") == 12)
check("missing logo falls back to name text", "Customer C (no logo found)" in txt and rep.logos["key_customers[2]"] == "name-fallback")
check("every schema key was placed", not rep.unplaced_keys)

# 2. changed template layout still fills
fill_template(here / "test_assets/template_test_mirrored.pptx", S / "assessli.json", O / "t_mirror.pptx", cache_dir=O / ".c")
check("mirrored template filled too", "{{" not in slide_text(O / "t_mirror.pptx") and pictures(O / "t_mirror.pptx") == 12)

# 3. invalid schema
rep = fill_template(T, S / "invalid_example.json", O / "t_invalid.pptx", cache_dir=O / ".c")
paths = " ".join(rep.errors)
check("bad website / missing hq / bad employees / bad funding / short description all flagged",
      all(k in paths for k in ["website", "hq", "employees", "funding", "company_description"]))
txt = slide_text(O / "t_invalid.pptx")
check("invalid values are NOT written to the slide", "forty" not in txt and "lots of money" not in txt and "Too short." not in txt)
check("unmapped schema key reported", "extra_unmapped_key" in rep.unplaced_keys)
try:
    fill_template(T, S / "invalid_example.json", O / "t_strict.pptx", strict=True, cache_dir=O / ".c"); check("strict mode raises", False)
except ValueError:
    check("strict mode raises and writes nothing", not (O / "t_strict.pptx").exists())

# 4. logo scraping (offline: fake web server responses)
import io
from PIL import Image
from logo_extractor import get_logo
def _png(color, size=(300, 120)):
    b = io.BytesIO(); Image.new("RGB", size, color).save(b, "PNG"); return b.getvalue()
PAGES = {
 "https://acme.test": b'<html><head><link rel="icon" href="/fav.png"><meta property="og:image" content="/banner.png"></head>'
                      b'<body><header><img class="site-logo" src="/img/logo.png" alt="Acme"></header></body></html>',
 "https://acme.test/img/logo.png": _png((200, 30, 30)), "https://acme.test/fav.png": _png((0, 0, 0), (32, 32)),
 "https://acme.test/banner.png": _png((9, 9, 9), (1200, 630)),
}
def fake(url):
    if url in PAGES: return PAGES[url]
    raise IOError("404 " + url)
import shutil; shutil.rmtree(O / ".lc", ignore_errors=True)
msgs = []
p = get_logo({"name": "Acme", "website": "acme.test"}, O / ".lc", fetch=fake, log=msgs.append)
check("scraper picks the <img class=logo> in the header before favicon/og:image", p and "img/logo.png" in " ".join(msgs))
check("scraped logo saved as normalised PNG", p and Image.open(p).format == "PNG")
check("no logo anywhere -> None (name fallback)", get_logo({"name": "Ghost", "website": "ghost.test"}, O / ".lc", fetch=fake) is None)

# 5. dynamic template: admin-style slide with NO tokens, filled through a map (label / shape-name based)
from fill_template import build_map
AD = here / "test_assets/template_admin_style.pptx"
mp, _ = build_map(Presentation(str(AD)))
check("init-map finds all 13 labels + 3 logo rows", len(mp["labels"]) == 13 and len(mp["logo_rows"]) == 3)
check("init-map guesses company logo / name / website / descriptions / product",
      sorted(mp["shapes"].values()) == sorted(["logo:company_logo", "company_name", "website", "company_description", "product_name", "product_description"]))
rep = fill_template(AD, S / "acme.json", O / "t_admin.pptx", mapping=mp, cache_dir=O / ".c")
txt = slide_text(O / "t_admin.pptx")
check("old sample data (Assessli, Dotsin, Bengaluru, 400K) fully replaced", not any(x in txt for x in ["Assessli", "Dotsin", "Bengaluru", "400K", "BioBERT"]))
check("new company data present", all(x in txt for x in ["Acme Vision", "VisionIQ", "Pune, Maharashtra", "$12M", "YOLOv8, SAM"]))
check("old sample logo pictures removed, only the 8 new ones remain", pictures(O / "t_admin.pptx") == 8)
check("fewer items than slots -> blank slots + warning (no {{token}} left)", "{{" not in txt and any("gtm_partners[2]" in w for w in rep.warnings))
check("more items than slots -> extras dropped + warning", any("key_customers" in w and "slots" in w for w in rep.warnings))
bad = {"labels": {"Nonexistent Label": "hq"}, "shapes": {"NoSuchShape": "company_name"}}
rep = fill_template(AD, S / "acme.json", O / "t_badmap.pptx", mapping=bad, cache_dir=O / ".c")
check("map entries that no longer exist in the template are reported", sum("not found in the template" in w for w in rep.warnings) == 2)
rep = fill_template(AD, S / "acme.json", O / "t_nomap.pptx", mapping={}, cache_dir=O / ".c")
check("no tokens and no map -> clear error, not a silent empty slide", any("nothing was filled" in e for e in rep.errors))

sys.exit(0 if ok else 1)
