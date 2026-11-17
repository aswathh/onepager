"""
Fill a PowerPoint one-pager template from the final schema JSON.

HOW THE TEMPLATE IS WIRED (so a changed template keeps working)
  Put a token wherever a value must appear - in any text box, shape, or table cell:

      {{company_name}}                 text value, path into the schema JSON
      {{company.hq}}                   nested key
      {{gtm_partners[0].name}}         list item field
      {{logo:company_logo}}            a logo slot  (shape or table cell is replaced by the picture)
      {{logo:gtm_partners[0]}}         logo of list item 0 (falls back to its name if no logo found)

  Positions come from the template itself: move / resize / restyle the shapes in PowerPoint and
  nothing in this script changes. Merged table cells are supported.

Run:
  python fill_template.py --template template.pptx --schema sample/assessli.json --out out/assessli.pptx
"""
from __future__ import annotations

import argparse
import difflib
import json
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Optional

from pptx import Presentation
from pptx.util import Pt
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE_TYPE

from logo_extractor import get_logo  # script 2: logo fetching


# =====================================================================
# 1) VALIDATION  - edit DEFAULT_RULES to match your real schema keys
#    Paths: dots for nesting, [*] = every list item, e.g. gtm_partners[*].name
# =====================================================================
URL_RE = re.compile(r"^(https?://)?([\w-]+\.)+[a-z]{2,}(/\S*)?$", re.I)
MONEY_RE = re.compile(
    r"^(?:(?:USD|INR|EUR|GBP|Rs\.?|[$€£₹])\s?)?\d[\d,]*(?:\.\d+)?\s?(?:[KMBkmb]|million|billion|crore|lakh|cr)?\+?"
    r"(?:\s?(?:USD|INR|EUR|GBP))?$|^(?:bootstrapped|undisclosed|n/?a|not available|stealth)$",
    re.I,
)
EMPLOYEES_RE = re.compile(r"^\d[\d,]*\+?$|^\d[\d,]*\s?[-–to]+\s?\d[\d,]*\+?$")


@dataclass
class Rule:
    required: bool = False
    type: type | tuple = str
    max_len: Optional[int] = None
    min_len: int = 1
    pattern: Optional[re.Pattern] = None
    pattern_hint: str = ""
    one_of: Optional[list[str]] = None
    min_items: int = 0
    max_items: Optional[int] = None
    check: Optional[Callable[[Any], Optional[str]]] = None  # returns an error message or None
    default: Optional[str] = "N/A"  # shown in the slide when the value is invalid / missing

    def to_dict(self):
        return {k: (v.pattern if isinstance(v, re.Pattern) else v) for k, v in self.__dict__.items() if k != "check"}


DEFAULT_RULES: dict[str, Rule] = {
    "company_name": Rule(required=True, max_len=40),
    "website": Rule(required=True, pattern=URL_RE, pattern_hint="a URL such as https://example.com", max_len=80),
    "company_logo": Rule(required=False, type=(dict, str)),
    "hq": Rule(required=True, max_len=60),
    "employees": Rule(required=True, type=(str, int), pattern=EMPLOYEES_RE, pattern_hint="'40', '11-50', '1,000+'", max_len=14),
    "funding": Rule(required=True, pattern=MONEY_RE, pattern_hint="'400K USD', '$5M', 'Bootstrapped', 'Undisclosed'", max_len=24),
    "revenue": Rule(required=True, pattern=MONEY_RE, pattern_hint="'$5.2M', 'N/A'", max_len=24),
    "company_description": Rule(required=True, min_len=60, max_len=700),
    "product_name": Rule(required=True, max_len=30),
    "product_description": Rule(required=True, min_len=60, max_len=700),
    "vertical_market": Rule(required=True, type=(str, list), max_len=90),
    "route_to_market": Rule(required=True, max_len=50),
    "inference_hardware": Rule(required=False, type=(str, list), max_len=60),
    "training_hardware": Rule(required=False, type=(str, list), max_len=60),
    "deployment_type": Rule(required=True, one_of=["Cloud", "On-prem", "Edge", "Hybrid", "Cloud + Edge", "On-device"]),
    "models": Rule(required=False, type=(str, list), max_len=90),
    "ai_framework": Rule(required=False, type=(str, list), max_len=60),
    "serving_framework": Rule(required=False, type=(str, list), max_len=60),
    "additional_ai_stack": Rule(required=False, type=(str, list), max_len=160),
    "gtm_partners": Rule(type=list, min_items=1, max_items=4, required=False),
    "key_customers": Rule(type=list, min_items=1, max_items=4, required=False),
    "main_competitors": Rule(type=list, min_items=1, max_items=4, required=False),
    "gtm_partners[*].name": Rule(required=True, max_len=40),
    "key_customers[*].name": Rule(required=True, max_len=40),
    "main_competitors[*].name": Rule(required=True, max_len=40),
}


@dataclass
class Issue:
    path: str
    level: str  # "error" | "warning"
    message: str

    def __str__(self):
        return f"[{self.level.upper()}] {self.path}: {self.message}"


@dataclass
class ValidationResult:
    issues: list[Issue] = field(default_factory=list)
    invalid_paths: set[str] = field(default_factory=set)  # paths that must NOT be written as-is

    @property
    def errors(self):
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self):
        return [i for i in self.issues if i.level == "warning"]

    @property
    def ok(self):
        return not self.errors


# ---------------------------------------------------------------- path helpers
_PART_RE = re.compile(r"([^.\[\]]+)|\[(\d+|\*)\]")


def split_path(path: str) -> list:
    parts: list = []
    for m in _PART_RE.finditer(path):
        if m.group(1) is not None:
            parts.append(m.group(1))
        else:
            parts.append(m.group(2) if m.group(2) == "*" else int(m.group(2)))
    return parts


def get_path(data: Any, path: str, default=None):
    cur = data
    for part in split_path(path):
        try:
            cur = cur[part]
        except (KeyError, IndexError, TypeError):
            return default
    return cur


def flatten(data: Any, prefix: str = "") -> dict[str, Any]:
    """Leaf paths of the schema: {'gtm_partners[0].name': 'Microsoft', ...}"""
    out: dict[str, Any] = {}
    if isinstance(data, dict):
        for k, v in data.items():
            out.update(flatten(v, f"{prefix}.{k}" if prefix else k))
    elif isinstance(data, list) and data and all(isinstance(x, (dict, list)) for x in data):
        for i, v in enumerate(data):
            out.update(flatten(v, f"{prefix}[{i}]"))
    else:
        out[prefix] = data
    return out


def _expand(rule_path: str, data: Any) -> list[str]:
    """Expand 'a[*].b' into concrete paths using the actual list length."""
    if "[*]" not in rule_path:
        return [rule_path]
    head, tail = rule_path.split("[*]", 1)
    lst = get_path(data, head)
    if not isinstance(lst, list):
        return []
    return [f"{head}[{i}]{tail}" for i in range(len(lst))]


def _is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip()) or (isinstance(v, (list, dict)) and not v)


def as_text(v) -> str:
    if isinstance(v, list):
        return ", ".join(as_text(x) for x in v if not _is_blank(x))
    return "" if v is None else str(v).strip()


def validate(data: dict, rules: dict[str, Rule] | None = None) -> ValidationResult:
    rules = rules or DEFAULT_RULES
    res = ValidationResult()
    ruled: set[str] = set()

    for rule_path, rule in rules.items():
        for path in _expand(rule_path, data):
            ruled.add(path)
            value = get_path(data, path)

            if _is_blank(value):
                if rule.required:
                    res.issues.append(Issue(path, "error", "required value is missing/empty"))
                    res.invalid_paths.add(path)
                continue

            types = rule.type if isinstance(rule.type, tuple) else (rule.type,)
            if not isinstance(value, types):
                res.issues.append(Issue(path, "error", f"expected {'/'.join(t.__name__ for t in types)}, got {type(value).__name__}"))
                res.invalid_paths.add(path)
                continue

            if isinstance(value, list):
                if len(value) < rule.min_items:
                    res.issues.append(Issue(path, "error", f"needs at least {rule.min_items} item(s)"))
                if rule.max_items is not None and len(value) > rule.max_items:
                    res.issues.append(Issue(path, "warning", f"{len(value)} items but the template shows {rule.max_items}; extras are dropped"))
                if not rule.max_len:
                    continue

            text = as_text(value)
            if len(text) < rule.min_len:
                res.issues.append(Issue(path, "error", f"too short ({len(text)} < {rule.min_len})"))
                res.invalid_paths.add(path)
            if rule.max_len and len(text) > rule.max_len:
                # overlong text is not fatal: the filler shrinks the font, but we flag it
                res.issues.append(Issue(path, "warning", f"{len(text)} chars (> {rule.max_len}); font will shrink / may look crowded"))
            if rule.pattern and isinstance(value, (str, int)) and not rule.pattern.match(str(value).strip()):
                res.issues.append(Issue(path, "error", f"bad format {text!r}" + (f"; expected {rule.pattern_hint}" if rule.pattern_hint else "")))
                res.invalid_paths.add(path)
            if rule.one_of and text not in rule.one_of:
                res.issues.append(Issue(path, "warning", f"{text!r} not in {rule.one_of}"))
            if rule.check:
                msg = rule.check(value)
                if msg:
                    res.issues.append(Issue(path, "error", msg))
                    res.invalid_paths.add(path)

    for path in flatten(data):
        if path not in ruled and not any(path.startswith(r) for r in ruled):
            res.issues.append(Issue(path, "warning", "no validation rule for this key"))
    return res


def load_rules(path: str) -> dict[str, Rule]:
    """Optional rules.json: {"hq": {"required": true, "max_len": 60}, "website": {"pattern": "^https?://"}}"""
    raw = json.load(open(path, encoding="utf-8"))
    out = {}
    for k, v in raw.items():
        v = dict(v)
        if "pattern" in v:
            v["pattern"] = re.compile(v["pattern"], re.I)
        if "type" in v:
            v["type"] = {"str": str, "int": int, "list": list, "dict": dict}[v["type"]]
        out[k] = Rule(**v)
    return out


# =====================================================================
# 2) TEMPLATE FILLING
# =====================================================================
TOKEN_RE = re.compile(r"\{\{\s*(logo:)?\s*([^{}]+?)\s*\}\}")
LOGO_ONLY_RE = re.compile(r"^\s*\{\{\s*logo:\s*([^{}]+?)\s*\}\}\s*$")
MIN_FONT, DEFAULT_FONT = 8.0, 14.0


@dataclass
class Report:
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    filled: list = field(default_factory=list)
    unresolved_tokens: list = field(default_factory=list)
    unplaced_keys: list = field(default_factory=list)
    logos: dict = field(default_factory=dict)  # slot -> "image" | "name-fallback" | "empty"

    @property
    def ok(self):
        return not self.errors and not self.unresolved_tokens

    def print(self):
        print(f"filled {len(self.filled)} text fields, {sum(v == 'image' for v in self.logos.values())}/{len(self.logos)} logo slots with images")
        for title, items in (("ERRORS", self.errors), ("WARNINGS", self.warnings),
                             ("TOKENS IN TEMPLATE WITH NO SCHEMA VALUE", self.unresolved_tokens),
                             ("SCHEMA KEYS NOT PLACED IN TEMPLATE", self.unplaced_keys)):
            if items:
                print(f"\n{title}:")
                for i in items:
                    print("  -", i)
        fb = [k for k, v in self.logos.items() if v != "image"]
        if fb:
            print("\nLOGO FALLBACKS:", ", ".join(f"{k}={self.logos[k]}" for k in fb))


class Filler:
    def __init__(self, schema: dict, base_dir: Path, cache_dir: Path, rules=None, strict=False, verbose=False):
        self.schema = schema
        self.base_dir = base_dir
        self.cache_dir = cache_dir
        self.strict = strict
        self.verbose = verbose
        self.report = Report()
        self.used: set[str] = set()
        self.vres = validate(schema, rules or DEFAULT_RULES)
        self.rules = rules or DEFAULT_RULES
        self.report.errors += [str(i) for i in self.vres.errors]
        self.report.warnings += [str(i) for i in self.vres.warnings if "no validation rule" not in i.message]
        self._pending_pictures: list = []

    # ------------------------------------------------------------ value lookup
    def _rule_for(self, path: str) -> Rule | None:
        norm = re.sub(r"\[\d+\]", "[*]", path)
        return self.rules.get(path) or self.rules.get(norm)

    def text_for(self, path: str) -> str | None:
        """Validated text for a token; None => token has no schema value."""
        raw = get_path(self.schema, path)
        if raw is None and not any(k == path or k.startswith(path + ".") or k.startswith(path + "[") for k in flatten(self.schema)):
            return None
        self.used.add(path)
        if path in self.vres.invalid_paths:
            rule = self._rule_for(path)
            return (rule.default if rule and rule.default is not None else "N/A")
        text = as_text(raw)
        if not text:
            rule = self._rule_for(path)
            return rule.default if rule and rule.default is not None else ""
        return re.sub(r"\s*\n\s*", " ", text)

    def logo_for(self, path: str):
        """-> (png_path | None, fallback_name)"""
        entry = get_path(self.schema, path)
        if entry is None:
            return None, None
        self.used.add(path)
        if isinstance(entry, str):
            entry = {"name": entry}
        name = entry.get("name") or get_path(self.schema, "company_name") or ""
        msgs = []
        png = get_logo(entry, self.cache_dir, self.base_dir, log=msgs.append)
        if self.verbose:
            print(*msgs, sep="\n")
        return png, name

    # ------------------------------------------------------------ text replacement
    def _replace_in_paragraph(self, p) -> bool:
        runs = p.runs
        full = "".join(r.text for r in runs)
        if "{{" not in full or not runs:
            return False

        def sub(m):
            if m.group(1):  # a logo token inside running text: just print the name
                return (get_path(self.schema, m.group(2)) or {}).get("name", "") if isinstance(get_path(self.schema, m.group(2)), dict) else ""
            val = self.text_for(m.group(2))
            if val is None:
                self.report.unresolved_tokens.append("{{" + m.group(2) + "}}")
                return ""
            self.report.filled.append(m.group(2))
            return val

        new = TOKEN_RE.sub(sub, full)
        runs[0].text = new
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        return True

    def _fit(self, text_frame, width_emu: int, height_emu: int):
        """Shrink font until the text should fit. Approximation: avg glyph width = 0.5 em."""
        paras = text_frame.paragraphs
        sizes = [r.font.size.pt for p in paras for r in p.runs if r.font.size]
        size = max(sizes) if sizes else DEFAULT_FONT
        w_pt = max(width_emu / 12700 - 14, 10)
        h_pt = max(height_emu / 12700 - 8, 10)
        texts = [(p, "".join(r.text for r in p.runs)) for p in paras]

        def needed(sz):
            cpl = max(int(w_pt / (sz * 0.5)), 1)
            lines = sum(max(math.ceil(len(t) / cpl), 1) for _, t in texts)
            return lines * sz * 1.2

        orig = size
        while needed(size) > h_pt and size > MIN_FONT:
            size -= 0.5
        if size < orig:
            for p, _ in texts:
                for r in p.runs:
                    r.font.size = Pt(size)

    # ------------------------------------------------------------ logo placement
    def _queue_logo(self, slide, path, left, top, width, height, slot_id, on_fallback):
        png, name = self.logo_for(path)
        if png is None and name is None:  # template has more slots than the schema has items -> leave blank
            on_fallback("")
            self.report.warnings.append(f"logo slot {path} has no data in the schema; left blank")
            self.report.logos[slot_id] = "empty"
            return
        if png is not None:
            self._pending_pictures.append((slide, png, left, top, width, height))
            self.report.logos[slot_id] = "image"
        else:
            on_fallback(name)
            self.report.logos[slot_id] = "name-fallback"

    def _flush_pictures(self):
        from PIL import Image

        for slide, png, left, top, w, h in self._pending_pictures:
            pad = int(min(w, h) * 0.06)
            bw, bh = w - 2 * pad, h - 2 * pad
            iw, ih = Image.open(png).size
            scale = min(bw / iw, bh / ih)
            pw, ph = int(iw * scale), int(ih * scale)
            slide.shapes.add_picture(str(png), left + pad + (bw - pw) // 2, top + pad + (bh - ph) // 2, pw, ph)
        self._pending_pictures.clear()

    # ------------------------------------------------------------ walkers
    def _fill_shape(self, slide, shape, off=(0, 0)):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for s in shape.shapes:
                self._fill_shape(slide, s, off)
            return
        if getattr(shape, "has_table", False) and shape.has_table:
            self._fill_table(slide, shape)
            return
        if not shape.has_text_frame:
            return
        tf = shape.text_frame
        m = LOGO_ONLY_RE.match(tf.text)
        if m:
            def fallback(name, shape=shape):
                tf = shape.text_frame
                tf.paragraphs[0].runs[0].text = name
                for r in tf.paragraphs[0].runs[1:]:
                    r._r.getparent().remove(r._r)
            self._queue_logo(slide, m.group(1), shape.left, shape.top, shape.width, shape.height, m.group(1), fallback)
            if self.report.logos.get(m.group(1)) == "image":
                shape._element.getparent().remove(shape._element)  # picture replaces the placeholder box
            return
        changed = False
        for p in tf.paragraphs:
            changed |= self._replace_in_paragraph(p)
        if changed:
            self._fit(tf, shape.width, shape.height)

    def _fill_table(self, slide, gframe):
        tbl = gframe.table
        xs = [gframe.left]
        for c in tbl.columns:
            xs.append(xs[-1] + c.width)
        ys = [gframe.top]
        for r in tbl.rows:
            ys.append(ys[-1] + r.height)
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row.cells):
                if cell.is_spanned:
                    continue
                cw = xs[ci + (cell.span_width if cell.is_merge_origin else 1)] - xs[ci]
                ch = ys[ri + (cell.span_height if cell.is_merge_origin else 1)] - ys[ri]
                m = LOGO_ONLY_RE.match(cell.text_frame.text)
                if m:
                    slot = m.group(1)

                    def fallback(name, cell=cell):
                        p = cell.text_frame.paragraphs[0]
                        p.runs[0].text = name
                        for r in p.runs[1:]:
                            r._r.getparent().remove(r._r)

                    self._queue_logo(slide, slot, xs[ci], ys[ri], cw, ch, slot, fallback)
                    if self.report.logos.get(slot) == "image":
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                r.text = ""
                    continue
                changed = False
                for p in cell.text_frame.paragraphs:
                    changed |= self._replace_in_paragraph(p)
                if changed:
                    self._fit(cell.text_frame, cw, ch)

    # ------------------------------------------------------------ run
    def fill(self, prs: Presentation, slide_index: int = 0):
        slide = prs.slides[slide_index]
        for shape in list(slide.shapes):
            self._fill_shape(slide, shape)
        self._flush_pictures()
        for leaf in flatten(self.schema):
            if not any(leaf == u or leaf.startswith(u + ".") or leaf.startswith(u + "[") for u in self.used):
                self.report.unplaced_keys.append(leaf)
        return self.report


# =====================================================================
# 3) DYNAMIC TEMPLATES  - works even when the template has NO {{tokens}}
#    A map file tells the script where each field lives by its LABEL / shape name, so when the
#    admin changes the template you re-run  --init-map  instead of changing code.
#
#    template_map.json
#      "labels":    {"HQ": "hq", "Funding": "funding"}        value goes in the cell to the RIGHT of the label cell
#      "logo_rows": {"GTM Partners": "gtm_partners"}          every cell right of the label gets one logo (old pictures removed)
#      "shapes":    {"TextBox 5": "company_name",             shape name (PowerPoint > Selection Pane) -> schema key
#                    "Picture 3": "logo:company_logo"}        "logo:<key>" = replace that shape by the logo
# =====================================================================
ALIASES = {
    "hq": ["hq", "headquarters", "head quarters", "location"],
    "employees": ["employees", "employee", "headcount", "team size", "no of employees"],
    "funding": ["funding", "funds raised", "total funding"],
    "revenue": ["revenue", "annual revenue", "arr"],
    "vertical_market": ["vertical market", "vertical", "industry", "target market"],
    "route_to_market": ["route to market", "go to market", "gtm model"],
    "inference_hardware": ["current inference hardware", "inference hardware"],
    "training_hardware": ["current training hardware", "training hardware"],
    "deployment_type": ["deployment type", "deployment"],
    "models": ["model/s", "models", "model", "model/s:"],
    "ai_framework": ["current ai framework", "ai framework"],
    "serving_framework": ["current serving framework", "serving framework"],
    "additional_ai_stack": ["additional ai software stack", "additional ai stack", "software stack"],
}
LOGO_ALIASES = {
    "gtm_partners": ["gtm partners", "partners", "key partners"],
    "key_customers": ["key customers", "customers", "main customers"],
    "main_competitors": ["main competitors", "competitors", "key competitors"],
}
URL_TEXT_RE = re.compile(r"(https?://|www\.)\S+|\b[\w-]+\.(com|ai|io|in|org|net|co)\b", re.I)


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _clean(s: str) -> str:
    return " ".join(s.split())


def _match_alias(text: str, table: dict) -> Optional[str]:
    n = _norm(text)
    if not n or len(text) > 40:
        return None
    for key, aliases in table.items():
        if n in {_norm(a) for a in aliases}:
            return key
    for key, aliases in table.items():
        if any(difflib.SequenceMatcher(None, n, _norm(a)).ratio() >= 0.88 for a in aliases):
            return key
    return None


def _grid(gf):
    xs, ys = [gf.left], [gf.top]
    for c in gf.table.columns:
        xs.append(xs[-1] + c.width)
    for r in gf.table.rows:
        ys.append(ys[-1] + r.height)
    return xs, ys


def _cells(gf):
    """yield (row, col, cell, (left, top, width, height)) for every non-spanned cell"""
    xs, ys = _grid(gf)
    for ri, row in enumerate(gf.table.rows):
        for ci, cell in enumerate(row.cells):
            if cell.is_spanned:
                continue
            cs = cell.span_width if cell.is_merge_origin else 1
            rs = cell.span_height if cell.is_merge_origin else 1
            yield ri, ci, cell, (xs[ci], ys[ri], xs[ci + cs] - xs[ci], ys[ri + rs] - ys[ri]), cs


def _tables(slide):
    return [s for s in slide.shapes if getattr(s, "has_table", False) and s.has_table]


def _set_text_keep_format(tf, text: str):
    """Replace all text, keeping the first run's font/size/colour (so the admin's styling survives)."""
    paras = tf.paragraphs
    for p in paras[1:]:
        p._p.getparent().remove(p._p)
    p0 = paras[0]
    for br in p0._p.findall(qn("a:br")):
        p0._p.remove(br)
    runs = p0.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.add_run().text = text


def _pictures_inside(slide, box):
    l, t, w, h = box
    out = []
    for sh in slide.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            cx, cy = sh.left + sh.width / 2, sh.top + sh.height / 2
            if l <= cx <= l + w and t <= cy <= t + h:
                out.append(sh)
    return out


def _find_label(slide, label: str):
    n = _norm(label)
    for gf in _tables(slide):
        for ri, ci, cell, box, cs in _cells(gf):
            if _norm(cell.text_frame.text) == n:
                return gf, ri, ci, cs
    return None


def _find_shape(slide, ref: str):
    for sh in slide.shapes:
        if sh.name == ref:
            return sh
    if ref.startswith("text:"):
        needle = _norm(ref[5:])
        for sh in slide.shapes:
            if sh.has_text_frame and needle and needle in _norm(sh.text_frame.text):
                return sh
    return None


def apply_map(prs, mapping: dict, schema: dict | None = None, slide_index: int = 0) -> list[str]:
    """Pre-pass: write {{tokens}} into the template at the places the map points to. Returns warnings."""
    slide, notes = prs.slides[slide_index], []

    for label, key in mapping.get("labels", {}).items():
        loc = _find_label(slide, label)
        if not loc:
            notes.append(f"map: label {label!r} not found in the template (renamed or removed?)")
            continue
        gf, ri, ci, cs = loc
        if ci + cs >= len(gf.table.columns):
            notes.append(f"map: label {label!r} has no value cell to its right")
            continue
        _set_text_keep_format(gf.table.cell(ri, ci + cs).text_frame, "{{" + key + "}}")

    for label, key in mapping.get("logo_rows", {}).items():
        loc = _find_label(slide, label)
        if not loc:
            notes.append(f"map: logo row {label!r} not found in the template")
            continue
        gf, ri, ci, cs = loc
        slots = [(c, b) for r_, c_, c, b, _ in _cells(gf) if r_ == ri and c_ >= ci + cs]
        n_items = len(get_path(schema, key) or []) if schema else None
        if n_items is not None and n_items > len(slots):
            notes.append(f"{key}: {n_items} items but the template row {label!r} has {len(slots)} slots; extras are dropped")
        for i, (cell, box) in enumerate(slots):
            for pic in _pictures_inside(slide, box):  # remove the admin's sample logo
                pic._element.getparent().remove(pic._element)
            _set_text_keep_format(cell.text_frame, "{{logo:" + f"{key}[{i}]" + "}}")

    for ref, key in mapping.get("shapes", {}).items():
        sh = _find_shape(slide, ref)
        if sh is None:
            notes.append(f"map: shape {ref!r} not found in the template (renamed or removed?)")
            continue
        if key.startswith("logo:"):
            box = slide.shapes.add_textbox(sh.left, sh.top, sh.width, sh.height)
            box.text_frame.text = "{{" + key + "}}"
            sh._element.getparent().remove(sh._element)
        elif sh.has_text_frame:
            _set_text_keep_format(sh.text_frame, "{{" + key + "}}")
        else:
            notes.append(f"map: shape {ref!r} cannot hold text")
    return notes


def build_map(prs, slide_index: int = 0) -> tuple[dict, list[str]]:
    """Scan the template and propose a map. Review it; the right-hand side must be YOUR schema keys."""
    slide = prs.slides[slide_index]
    mp: dict = {"labels": {}, "logo_rows": {}, "shapes": {}}
    msgs: list[str] = []
    logo_boxes = []

    for gf in _tables(slide):
        for ri, ci, cell, box, cs in _cells(gf):
            text = _clean(cell.text_frame.text)
            if not text:
                continue
            if (k := _match_alias(text, LOGO_ALIASES)):
                mp["logo_rows"][text] = k
                slots = [b for r_, c_, c, b, _ in _cells(gf) if r_ == ri and c_ >= ci + cs]
                logo_boxes += slots
                msgs.append(f"logo row  {text!r:28} -> {k}  ({len(slots)} slots)")
            elif (k := _match_alias(text, ALIASES)):
                mp["labels"][text] = k
                msgs.append(f"label     {text!r:28} -> {k}")

    def in_logo_row(sh):
        cx, cy = sh.left + sh.width / 2, sh.top + sh.height / 2
        return any(l <= cx <= l + w and t <= cy <= t + h for l, t, w, h in logo_boxes)

    half = prs.slide_width / 2
    pics = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE and not in_logo_row(s)]
    left_pics = [p for p in pics if p.left + p.width / 2 < half]
    if left_pics:
        p = max(left_pics, key=lambda x: x.width * x.height)
        mp["shapes"][p.name] = "logo:company_logo"
        msgs.append(f"picture   {p.name!r:28} -> logo:company_logo   (GUESS: biggest picture on the left)")

    texts = []
    for sh in slide.shapes:
        if sh.shape_type in (MSO_SHAPE_TYPE.PICTURE,) or getattr(sh, "has_table", False) and sh.has_table:
            continue
        if sh.has_text_frame and sh.text_frame.text.strip():
            sizes = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
            texts.append((sh, _clean(sh.text_frame.text), max(sizes) if sizes else 0))

    def pick(cands, key, note, how):
        if cands and key not in mp["shapes"].values():
            sh = how(cands)
            mp["shapes"][sh[0].name] = key
            msgs.append(f"shape     {sh[0].name!r:28} -> {key}   (GUESS: {note}; text starts {sh[1][:30]!r})")

    urls = [t for t in texts if URL_TEXT_RE.search(t[1]) and len(t[1]) < 80]
    pick(urls, "website", "text that looks like a URL", lambda c: c[0])
    taken = {n for n in mp["shapes"]}
    rest = [t for t in texts if t[0].name not in taken]
    for side, (name_key, desc_key) in (("left", ("company_name", "company_description")), ("right", ("product_name", "product_description"))):
        in_side = [t for t in rest if (t[0].left + t[0].width / 2 < half) == (side == "left")]
        shortish = [t for t in in_side if len(t[1]) <= 40 and t[2] >= 18]
        pick(shortish, name_key, f"largest short text on the {side}", lambda c: max(c, key=lambda t: t[2]))
        longish = [t for t in in_side if len(t[1]) > 80]
        pick(longish, desc_key, f"longest paragraph on the {side}", lambda c: max(c, key=lambda t: len(t[1])))

    unknown = [t[1] for t in texts if t[0].name not in mp["shapes"] and 0 < len(t[1]) <= 40][:12]
    if unknown:
        msgs.append("other text not mapped (static labels/footers are fine): " + "; ".join(repr(u) for u in unknown))
    return mp, msgs


def _load_map(mapping, template) -> dict | None:
    if isinstance(mapping, dict):
        return mapping
    p = Path(mapping) if mapping else Path(template).with_suffix(".map.json")  # auto-pick <template>.map.json
    return json.load(open(p, encoding="utf-8")) if p.exists() else None


def fill_template(template: str | Path, schema: dict | str | Path, out: str | Path, *, rules=None, strict=False,
                  cache_dir: str | Path = ".logo_cache", verbose=False, mapping=None) -> Report:
    base_dir = Path(".")
    if isinstance(schema, (str, Path)):
        base_dir = Path(schema).parent
        schema = json.load(open(schema, encoding="utf-8"))
    prs = Presentation(str(template))
    notes = []
    mp = _load_map(mapping, template)
    if mp:
        notes = apply_map(prs, mp, schema)
    f = Filler(schema, base_dir, Path(cache_dir), rules, strict, verbose)
    f.report.warnings += notes
    report = f.fill(prs)
    if not report.filled and not report.logos:
        report.errors.append("nothing was filled: the template has no {{tokens}} and no map was applied. Run --init-map first.")
    if strict and report.errors:
        raise ValueError("schema validation failed:\n  " + "\n  ".join(report.errors))
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    return report


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--schema")
    ap.add_argument("--out")
    ap.add_argument("--map", help="template map json (default: <template>.map.json if it exists)")
    ap.add_argument("--init-map", action="store_true", help="scan the template and write a starter map, then exit")
    ap.add_argument("--rules", help="optional rules.json overriding DEFAULT_RULES")
    ap.add_argument("--strict", action="store_true", help="refuse to write the file if validation has errors")
    ap.add_argument("--cache-dir", default=".logo_cache")
    ap.add_argument("-v", "--verbose", action="store_true")
    a = ap.parse_args()
    if a.init_map:
        mp, msgs = build_map(Presentation(a.template))
        dest = Path(a.map) if a.map else Path(a.template).with_suffix(".map.json")
        json.dump(mp, open(dest, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
        print("\n".join(msgs), f"\n\nmap written -> {dest}\nReview it: change the right-hand side to YOUR schema keys, fix any GUESS lines.")
        sys.exit(0)
    if not (a.schema and a.out):
        ap.error("--schema and --out are required (unless --init-map)")
    rules = None
    if a.rules:
        rules = load_rules(a.rules)
    try:
        rep = fill_template(a.template, a.schema, a.out, rules=rules, strict=a.strict, cache_dir=a.cache_dir, verbose=a.verbose, mapping=a.map)
    except ValueError as e:
        print(e)
        sys.exit(2)
    rep.print()
    print("\nsaved ->", a.out)
    sys.exit(0 if rep.ok else 1)
