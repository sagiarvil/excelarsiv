#!/usr/bin/env python3
"""
REAL-SEO-2026 APEX v6.0 — Görsel .webp Dönüştürme, Sıkıştırma ve Manifest Üretici
dist/ altındaki tüm raster görselleri slug uyumlu, <= 160 KB .webp dosyalarına dönüştürür
ve width/height bilgilerini dist/.omega-images.json içine yazar.
"""

import json
import re
from pathlib import Path
from PIL import Image

DIST_DIR = Path("dist")
MANIFEST_PATH = DIST_DIR / ".omega-images.json"
VALID_SLUG_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)+$")
TR_MAP = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")


def slugify_stem(stem: str, parent_name: str = "") -> str:
    s = stem.translate(TR_MAP).lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    if not VALID_SLUG_RE.match(s):
        prefix = re.sub(r"[^a-z0-9]+", "-", parent_name.translate(TR_MAP).lower()).strip("-") or "excel-arsiv"
        if prefix and not s.startswith(prefix):
            s = f"{prefix}-{s}" if s else f"{prefix}-gorsel"
        if not VALID_SLUG_RE.match(s):
            s = f"excel-arsiv-{s or 'gorsel'}"
    return s


def save_webp_under_limit(img: Image.Image, dest: Path, max_kb: float = 160.0) -> tuple[int, int]:
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGBA" if "A" in img.getbands() else "RGB")

    w, h = img.size
    # Cap huge dimensions to 1400px max width for fast LCP & low KB
    if w > 1400:
        ratio = 1400 / w
        w, h = 1400, max(1, int(h * ratio))
        img = img.resize((w, h), Image.Resampling.LANCZOS)

    dest.parent.mkdir(parents=True, exist_ok=True)
    for quality in (82, 72, 62, 50, 40):
        img.save(dest, format="WEBP", quality=quality, method=4)
        if dest.stat().st_size <= max_kb * 1024:
            return w, h

    # Downscale further if still > max_kb
    cur = img
    for scale in (0.75, 0.55, 0.4):
        nw, nh = max(120, int(cur.width * scale)), max(80, int(cur.height * scale))
        resized = cur.resize((nw, nh), Image.Resampling.LANCZOS)
        resized.save(dest, format="WEBP", quality=60, method=4)
        if dest.stat().st_size <= max_kb * 1024:
            return nw, nh
    return cur.size


def main() -> None:
    if not DIST_DIR.exists():
        raise SystemExit("dist/ dizini bulunamadı.")

    manifest: dict[str, dict] = {}
    # 1. Convert all png/jpg/jpeg/webp in dist/
    raster_files = [
        p for p in DIST_DIR.rglob("*")
        if p.is_file() and p.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp")
        and not p.name.startswith(".")
    ]

    for src in raster_files:
        rel = "/" + src.relative_to(DIST_DIR).as_posix()
        parent_slug = src.parent.name if src.parent != DIST_DIR else "excel-arsiv"
        new_stem = slugify_stem(src.stem, parent_slug)
        dest = src.with_name(f"{new_stem}.webp")
        dest_rel = "/" + dest.relative_to(DIST_DIR).as_posix()

        try:
            with Image.open(src) as im:
                if src.suffix.lower() == ".webp" and src == dest and src.stat().st_size <= 165 * 1024:
                    w, h = im.size
                else:
                    w, h = save_webp_under_limit(im, dest, max_kb=160.0)
            manifest[rel] = {"webp": dest_rel, "width": w, "height": h}
            manifest[dest_rel] = {"webp": dest_rel, "width": w, "height": h}
        except Exception as exc:
            print(f"[WARN] Görsel işlenemedi ({rel}): {exc}")

    # 2. Ensure dedicated social card /images/excel-arsiv-sosyal-kart.webp & default hero exist
    og_src = DIST_DIR / "og.png"
    social_dest = DIST_DIR / "images" / "excel-arsiv-sosyal-kart.webp"
    if og_src.exists() and not social_dest.exists():
        with Image.open(og_src) as im:
            im_resized = im.convert("RGB").resize((1200, 630), Image.Resampling.LANCZOS)
            w, h = save_webp_under_limit(im_resized, social_dest, max_kb=150.0)
            manifest["/images/excel-arsiv-sosyal-kart.webp"] = {
                "webp": "/images/excel-arsiv-sosyal-kart.webp",
                "width": w,
                "height": h,
            }

    MANIFEST_PATH.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OMEGA-IMG] {len(manifest)} görsel kaydı dist/.omega-images.json içine yazıldı.")


if __name__ == "__main__":
    main()
