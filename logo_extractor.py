"""
Logo resolution for partners / customers / competitors / the company itself.

Order tried for every entry {name, website?, logo_path?, logo_url?}:
  1. logo_path        local file (relative paths resolve against the schema file)
  2. logo_url         direct image URL
  3. website          scrape the homepage: JSON-LD logo -> <img> with "logo" in src/alt/class
                      -> apple-touch-icon -> <link rel=icon> -> og:image -> /favicon.ico
  4. Google favicon service (low-res, last resort)
  5. None             -> the filler writes the company *name* in the slot instead

Every logo is normalised to a trimmed RGBA PNG and cached, so the slot always receives the same format.
"""
from __future__ import annotations

import hashlib
import io
import json
import re
from pathlib import Path
from typing import Callable, Optional
from urllib.parse import urljoin, urlparse

from PIL import Image, ImageChops

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None
try:
    from bs4 import BeautifulSoup
except ImportError:  # pragma: no cover
    BeautifulSoup = None

UA = {"User-Agent": "Mozilla/5.0 (compatible; OnePagerBot/1.0)"}
MAX_BYTES = 6_000_000
MIN_SIDE = 24


def slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-") or "logo"


def domain_of(url: str) -> str:
    if not url:
        return ""
    if "//" not in url:
        url = "https://" + url
    return urlparse(url).netloc.removeprefix("www.")


# ------------------------------------------------------------------ fetching
def http_get(url: str, timeout: int = 10) -> bytes:
    if requests is None:
        raise RuntimeError("pip install requests")
    r = requests.get(url, headers=UA, timeout=timeout, stream=True)
    r.raise_for_status()
    data = r.raw.read(MAX_BYTES + 1, decode_content=True)
    if len(data) > MAX_BYTES:
        raise ValueError("image too large")
    return data


# Replaceable in tests / behind a corporate proxy
Fetcher = Callable[[str], bytes]


# ------------------------------------------------------------------ image normalising
def _svg_to_png(raw: bytes) -> bytes:
    try:
        import cairosvg  # optional
    except ImportError as e:
        raise RuntimeError("SVG logo found; pip install cairosvg to convert it") from e
    return cairosvg.svg2png(bytestring=raw, output_width=600)


def _trim(im: Image.Image) -> Image.Image:
    """Remove transparent or uniform-colour margins."""
    im = im.convert("RGBA")
    alpha = im.split()[3]
    if alpha.getextrema()[0] < 250:  # real transparency -> trim by alpha
        bbox = alpha.getbbox()
    else:  # opaque -> trim only a white-ish margin; never strip a brand-colour background
        corner = im.getpixel((0, 0))
        if min(corner[:3]) < 235:
            return im
        bg = Image.new("RGBA", im.size, corner)
        diff = ImageChops.difference(im, bg).convert("L").point(lambda p: 255 if p > 12 else 0)
        bbox = diff.getbbox()
    return im.crop(bbox) if bbox else im


def normalise(raw: bytes, max_px: int = 600) -> Image.Image:
    head = raw[:300].lstrip().lower()
    if head.startswith(b"<?xml") or head.startswith(b"<svg") or b"<svg" in head:
        raw = _svg_to_png(raw)
    im = Image.open(io.BytesIO(raw))
    if getattr(im, "n_frames", 1) > 1:  # animated gif / multi-size ico: take the largest frame
        im.seek(0)
    im = _trim(im)
    if min(im.size) < MIN_SIDE:
        raise ValueError(f"logo too small {im.size}")
    im.thumbnail((max_px, max_px))
    return im


# ------------------------------------------------------------------ page scraping
def candidate_urls(website: str, fetch: Fetcher = http_get) -> list[str]:
    base = website if "//" in website else "https://" + website
    try:
        html = fetch(base).decode("utf-8", "ignore")
    except Exception:
        html = ""
    cands: list[tuple[int, str]] = []
    if html and BeautifulSoup:
        soup = BeautifulSoup(html, "html.parser")
        for tag in soup.find_all("script", type="application/ld+json"):
            try:
                blob = json.loads(tag.string or "")
            except Exception:
                continue
            for node in blob if isinstance(blob, list) else [blob]:
                logo = node.get("logo") if isinstance(node, dict) else None
                logo = logo.get("url") if isinstance(logo, dict) else logo
                if isinstance(logo, str):
                    cands.append((0, urljoin(base, logo)))
        for img in soup.find_all("img"):
            blob = " ".join([img.get("src", ""), img.get("alt", ""), " ".join(img.get("class", []))]).lower()
            if "logo" in blob and img.get("src") and not img["src"].startswith("data:"):
                in_header = bool(img.find_parent(["header", "nav"]))
                cands.append((1 if in_header else 2, urljoin(base, img["src"])))
        for link in soup.find_all("link", rel=True):
            rel = " ".join(link["rel"]).lower()
            href = link.get("href")
            if not href:
                continue
            if "apple-touch-icon" in rel:
                cands.append((3, urljoin(base, href)))
            elif "icon" in rel:
                cands.append((4 if href.lower().endswith(".svg") else 5, urljoin(base, href)))
        og = soup.find("meta", property="og:image")
        if og and og.get("content"):
            cands.append((6, urljoin(base, og["content"])))
    cands.append((7, urljoin(base, "/favicon.ico")))
    cands.append((8, f"https://www.google.com/s2/favicons?domain={domain_of(base)}&sz=128"))
    seen, ordered = set(), []
    for _, u in sorted(cands, key=lambda c: c[0]):
        if u not in seen:
            seen.add(u)
            ordered.append(u)
    return ordered


# ------------------------------------------------------------------ public API
def get_logo(
    entry: dict | str,
    cache_dir: str | Path = ".logo_cache",
    base_dir: str | Path = ".",
    fetch: Fetcher = http_get,
    log: Callable[[str], None] = lambda m: None,
) -> Optional[Path]:
    """Return the path of a normalised PNG, or None when nothing usable was found."""
    if isinstance(entry, str):
        entry = {"name": entry}
    name = entry.get("name", "")
    cache = Path(cache_dir)
    cache.mkdir(parents=True, exist_ok=True)
    key = domain_of(entry.get("website", "")) or slug(name)
    out = cache / f"{slug(key)}-{hashlib.md5(json.dumps([entry.get('logo_path'), entry.get('logo_url'), key]).encode()).hexdigest()[:6]}.png"
    if out.exists():
        return out

    def save(raw: bytes, source: str) -> Optional[Path]:
        try:
            normalise(raw).save(out)
            log(f"logo for {name or key!r} from {source}")
            return out
        except Exception as e:
            log(f"  skip {source}: {e}")
            return None

    if entry.get("logo_path"):
        p = Path(entry["logo_path"])
        p = p if p.is_absolute() else Path(base_dir) / p
        if p.exists() and (r := save(p.read_bytes(), str(p))):
            return r
        log(f"  logo_path not usable: {p}")
    if entry.get("logo_url"):
        try:
            if r := save(fetch(entry["logo_url"]), entry["logo_url"]):
                return r
        except Exception as e:
            log(f"  logo_url failed: {e}")
    if entry.get("website"):
        for url in candidate_urls(entry["website"], fetch)[:8]:
            try:
                if r := save(fetch(url), url):
                    return r
            except Exception as e:
                log(f"  {url}: {e}")
    log(f"no logo found for {name or key!r}; the name will be written instead")
    return None


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Fetch + normalise one logo")
    ap.add_argument("--name", default="")
    ap.add_argument("--website", default="")
    ap.add_argument("--logo-url", default="")
    ap.add_argument("--out-dir", default=".logo_cache")
    a = ap.parse_args()
    print(get_logo({"name": a.name, "website": a.website, "logo_url": a.logo_url}, a.out_dir, log=print))
