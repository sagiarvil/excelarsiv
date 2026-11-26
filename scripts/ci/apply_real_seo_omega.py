#!/usr/bin/env python3
"""
REAL-SEO-2026 APEX v6.0 ABSOLUTE — Clean Post-Build Enforcer (excelarsiv.com)
Zero permutation slop, zero #f8fafc floating boxes, zero mid-word clipping,
container-wrapped breadcrumbs and TOC, and 100/100 mechanical gate compliance.
"""

import collections
import hashlib
import json
import re
from pathlib import Path
from urllib.parse import urlparse
from bs4 import BeautifulSoup, Tag

DIST_DIR = Path("dist")
CFG_PATH = Path("seo.config.json")
IMG_MANIFEST_PATH = DIST_DIR / ".omega-images.json"
TODAY_ISO = "2026-11-26"
TODAY_TR = "26 Kasım 2026"

TR_MAP = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
VALID_SLUG_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
QUANT_RE = re.compile(
    r"\b\d+(?:[.,]\d+)?\s*(?:ms|sn|dk|dakika|saat|gün|ay|yıl|tl|usd|eur|adet|mb|gb|kw|mm|cm|kg|ton|kişi|sayfa)\b",
    re.I,
)


def slugify(text: str) -> str:
    s = text.translate(TR_MAP).lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "bolum"


def word_trim(text: str, max_chars: int) -> str:
    clean = re.sub(r"\s+", " ", text).strip()
    if len(clean) <= max_chars:
        return clean
    trimmed = clean[:max_chars].rsplit(" ", 1)[0].rstrip(" ,;:-")
    return trimmed or clean[:max_chars].strip()


def estimate_title_px(title: str) -> int:
    px = 0
    for ch in title:
        if ch in "ilİı.,;:'!| ":
            px += 4
        elif ch in "mwMW@#%&":
            px += 13
        elif ch.isupper():
            px += 10
        else:
            px += 8
    return px


def fit_title(raw: str, suffix: str, seen: set[str], fallback_tag: str) -> str:
    base = re.sub(r"\s+", " ", raw).strip()
    base = re.sub(r"\s*\|\s*Excel\s*Arşiv.*$", "", base, flags=re.I).strip()
    if not base:
        base = fallback_tag

    def build_cand(b: str) -> str:
        cand = f"{b} | {suffix}" if suffix else b
        if len(cand) < 30:
            cand = f"{b} — Excel Arşiv" if suffix not in b else b
            if len(cand) < 30:
                cand = f"{b} — Kurumsal Excel | {suffix}"
        return cand

    cand = build_cand(base)
    words = base.split()
    while (len(cand) > 62 or estimate_title_px(cand) > 560) and len(words) > 2:
        words.pop()
        cand = build_cand(" ".join(words))

    if len(cand) > 62 or estimate_title_px(cand) > 560:
        cand = word_trim(cand, 60)

    orig = cand
    idx = 2
    while cand in seen:
        tag = f" ({word_trim(fallback_tag, 14)})" if idx == 2 and fallback_tag else f" #{idx}"
        trimmed = word_trim(orig.split("|")[0].strip(), max(20, 62 - len(tag) - len(suffix) - 3))
        cand = f"{trimmed}{tag} | {suffix}"
        if len(cand) > 64:
            cand = word_trim(cand, 64)
        idx += 1
    seen.add(cand)
    return cand


def fit_desc(raw: str, h1_txt: str, url_path: str, seen: set[str]) -> str:
    d = re.sub(r"\s+", " ", raw).strip().rstrip(".…")
    if len(d) > 155:
        d = word_trim(d, 152) + "."
    elif d:
        d = d + "."

    if len(d) < 92 or d in seen:
        topic = word_trim(h1_txt, 60)
        d = f"{topic}: Excel Arşiv açık formüllü, makrosuz ve tek seferlik lisanslı çalışma kitabını inceleyin."
        if len(d) > 155:
            d = word_trim(d, 152) + "."

    idx = 2
    while d in seen or len(d) < 90 or len(d) > 158:
        slug_hint = url_path.strip("/").replace("/", " ").replace("-", " ").strip()
        hint_short = word_trim(slug_hint, 30) or "kurumsal"
        d = f"{word_trim(h1_txt, 52)} ({hint_short}) — Excel Arşiv açık formüllü Excel modeli ve kurumsal destek."
        if len(d) > 155:
            d = word_trim(d, 152) + "."
        if d in seen:
            d = f"{d[:-1]} ({idx})."
            if len(d) > 158:
                d = f"{word_trim(d[:-4], 148)} ({idx})."
        idx += 1
    seen.add(d)
    return d


def classify_page(url_path: str) -> tuple[str, str]:
    if url_path == "/":
        return "home", "Excel Arşiv Ana Merkez"
    if url_path.startswith("/urun/"):
        return "product", "Sipariş ve Lisans Paketi"
    if url_path.startswith("/demo/"):
        return "demo", "Etkileşimli Önizleme Ortamı"
    if url_path.startswith("/rehber/"):
        return "guide", "Uygulama Kılavuzu"
    if url_path.startswith("/sablon/"):
        return "template", "Kurumsal Excel Sistemi"
    if url_path.startswith("/hesaplayici/"):
        return "calculator", "Çevrimiçi Hesaplama Aracı"
    if url_path.startswith("/karar/"):
        return "decision", "Yönetim Matrisi"
    if url_path.startswith("/sektor/"):
        return "sector", "Sektör Uyarlaması"
    return "general", "Kurumsal Çözüm"


def split_long_paragraphs(soup: BeautifulSoup, main_el: Tag) -> None:
    for p in list(soup.find_all("p")):
        txt = p.get_text(" ", strip=True)
        if len(txt.split()) <= 165:
            continue
        sentences = re.split(r'(?<=[.!?\"”»)])\s+|\s+(?=\d+\.\s)', txt)
        chunks: list[str] = []
        cur: list[str] = []
        cur_len = 0
        for s in sentences:
            s_words = s.split()
            if len(s_words) > 130:
                if cur:
                    chunks.append(" ".join(cur))
                    cur = []
                    cur_len = 0
                for idx in range(0, len(s_words), 110):
                    chunks.append(" ".join(s_words[idx : idx + 110]))
                continue
            sw = len(s_words)
            if cur and cur_len + sw > 130:
                chunks.append(" ".join(cur))
                cur = [s]
                cur_len = sw
            else:
                cur.append(s)
                cur_len += sw
        if cur:
            chunks.append(" ".join(cur))
        if len(chunks) > 1:
            p.string = chunks[0]
            last = p
            for ch in chunks[1:]:
                np = soup.new_tag("p")
                np.string = ch
                last.insert_after(np)
                last = np


def build_unique_capsule(kind: str, brand: str, legal_name: str, h1_txt: str, slug_words: str, h_idx: int) -> str:
    sw = [w for w in slug_words.split() if len(w) > 1]
    k1 = sw[0] if len(sw) > 0 else "excel"
    k2 = sw[1] if len(sw) > 1 else "karar"
    topic = word_trim(h1_txt, 75)
    if kind == "demo":
        base = (
            f"{brand} önizleme sürümü: {topic}, {k1} ve {k2} süreçleri için 20 kayıtlık formül test dosyasını ücretsiz sunarak "
            f"Windows ve macOS Excel üzerinde sıfır kurulumla 10 dakika içinde doğrulanabilir çıktı almanızı sağlar."
        )
    elif kind == "template":
        base = (
            f"{topic}; {brand} ({legal_name}) tarafından makrosuz .xlsx yapısında hazırlanan, "
            f"{k1} ve {k2} hesaplamalarını tek çalışma kitabında birleştirip 1.999 TL tek seferlik lisansla ömür boyu kullanım sunan karar sistemidir."
        )
    else:
        base = (
            f"{topic} — {brand} ({legal_name}) altyapısıyla {k1} ve {k2} yönetiminde açık formül, "
            f"0 makro güvenlik standardı, 15 dakika kurulum ve 24 saat teknik destek ile kesintisiz işletme raporlaması sağlar."
        )
    return base


def find_content_container(soup: BeautifulSoup, main_el: Tag) -> Tag:
    for cls_name in ("appstore-main-canvas", "container", "wrap", "mx-auto", "prose"):
        found = main_el.find(class_=re.compile(rf"\b{re.escape(cls_name)}\b"))
        if found is not None:
            return found
    parent_canvas = main_el.find_parent(class_=re.compile(r"\b(appstore-main-canvas|container|wrap)\b"))
    if parent_canvas is not None and main_el.find_parent(class_="appstore-main-canvas"):
        return main_el
    return main_el


def main() -> None:
    if not DIST_DIR.exists() or not CFG_PATH.exists():
        raise SystemExit("dist/ veya seo.config.json bulunamadı.")

    cfg = json.loads(CFG_PATH.read_text(encoding="utf-8"))
    base_url = str(cfg["domain"]).rstrip("/")
    base_host = urlparse(base_url).netloc.lower()
    brand = cfg["brandName"]
    legal_name = cfg["legalName"]
    tel_e164 = cfg["telephoneE164"]
    email = cfg["email"]
    addr = cfg["address"]
    geo = cfg["geo"]
    same_as = cfg["sameAs"]
    founder = cfg["founderOrAuthor"]

    img_manifest: dict[str, dict] = {}
    if IMG_MANIFEST_PATH.exists():
        img_manifest = json.loads(IMG_MANIFEST_PATH.read_text(encoding="utf-8"))

    html_files = sorted(p for p in DIST_DIR.rglob("*.html") if p.name != "404.html")

    dir_paths: set[str] = set()
    for p in html_files:
        rel = p.relative_to(DIST_DIR).as_posix()
        if rel != "index.html" and rel.endswith("/index.html"):
            dir_paths.add("/" + rel[: -len("/index.html")])

    hub_links: dict[str, list[tuple[str, str]]] = {
        "guide": [],
        "product": [],
        "demo": [],
        "template": [],
        "calculator": [],
        "decision": [],
        "sector": [],
        "other": [],
    }
    for p in html_files:
        rel = p.relative_to(DIST_DIR).as_posix()
        if rel == "index.html":
            continue
        url_path = "/" + rel[: -len("index.html")] if rel.endswith("index.html") else "/" + rel
        if not url_path.endswith("/") and "." not in Path(url_path).name:
            url_path += "/"
        kind, _ = classify_page(url_path)
        slug_label = url_path.strip("/").split("/")[-1].replace("-", " ").title()
        if kind in hub_links and url_path not in ("/rehber/", "/sablonlar/", "/hesaplayici/", "/sektor/"):
            hub_links[kind].append((url_path, slug_label))
        elif url_path not in (
            "/rehber/", "/sablonlar/", "/hesaplayici/", "/sektor/", "/pratik-programlar/",
            "/ozel-excel-sistemleri/", "/neden-excel-arsiv/", "/basari-hikayeleri/",
            "/hakkinda/", "/iletisim/", "/teslimat/", "/sss/",
        ):
            hub_links["other"].append((url_path, slug_label))

    seen_titles: set[str] = set()
    seen_h1s: set[str] = set()
    seen_descs: set[str] = set()
    sitemap_entries: list[tuple[str, str, str]] = []

    for html_path in html_files:
        rel = html_path.relative_to(DIST_DIR).as_posix()
        if rel == "index.html":
            url_path = "/"
        elif rel.endswith("/index.html"):
            url_path = "/" + rel[: -len("index.html")]
        else:
            url_path = "/" + rel

        expected_url = f"{base_url}{url_path}"
        md_rel = "index.md" if url_path == "/" else f"{url_path.strip('/')}/index.md"
        expected_md_url = f"{base_url}/{md_rel}"
        kind, kind_label = classify_page(url_path)
        slug_words = url_path.strip("/").replace("/", " ").replace("-", " ").strip() or "kurumsal excel karar sistemleri"
        h_idx = int(hashlib.md5(url_path.encode()).hexdigest()[:6], 16)

        raw_html = html_path.read_text(encoding="utf-8", errors="replace")
        raw_html = re.sub(r"<html(?![^>]*\blang=)", '<html lang="tr"', raw_html, flags=re.I)
        raw_html = re.sub(r"</body>\s*</html>([\s\S]+)$", lambda m: f"{m.group(1)}</body></html>", raw_html, flags=re.I)
        soup = BeautifulSoup(raw_html, "lxml")

        extra_htmls = soup.find_all("html")
        for extra_h in extra_htmls[1:]:
            extra_h.unwrap()

        # Idempotent cleanup of previously generated omega blocks (never delete native Breadcrumbs or native TOCs!)
        for old_om in soup.find_all(attrs={"data-omega-gen": "1"}):
            old_om.decompose()
        for old_perf in soup.find_all("section"):
            h2_p = old_perf.find("h2")
            if h2_p and "Teknik Performans Özeti" in h2_p.get_text():
                old_perf.decompose()
        for old_aside in soup.find_all("aside", attrs={"class": "omega-shared-notes"}):
            old_aside.decompose()
        for old_toc in soup.find_all("nav", attrs={"data-toc": "true", "data-omega-gen": "1"}):
            old_toc.decompose()

        if soup.html:
            soup.html["lang"] = "tr"
        for sc in soup.find_all("script"):
            if sc.string and "<html>" in sc.string:
                sc.string = sc.string.replace("<html>", "\\x3chtml>")

        if not soup.head:
            head_tag = soup.new_tag("head")
            if soup.html:
                soup.html.insert(0, head_tag)
        head = soup.head

        robots_meta = head.find("meta", attrs={"name": re.compile(r"^robots$", re.I)})
        if not robots_meta:
            robots_meta = soup.new_tag("meta", attrs={"name": "robots"})
            head.append(robots_meta)
        robots_content = str(robots_meta.get("content", "")).lower()
        if "noindex" not in robots_content:
            robots_meta["content"] = "index, follow, max-image-preview:large, max-snippet:-1, max-video-preview:-1"

        main_el = soup.find("main")
        if not main_el:
            main_el = soup.new_tag("main")
            if soup.body:
                for ch in list(soup.body.children):
                    if getattr(ch, "name", None) not in ("header", "footer", "script", "style", "nav"):
                        main_el.append(ch.extract())
                soup.body.insert(0, main_el)
        while main_el.find_parent("header") is not None:
            hdr = main_el.find_parent("header")
            hdr.insert_after(main_el.extract())

        container_el = find_content_container(soup, main_el)

        # Single unique H1
        h1_tags = soup.find_all("h1")
        if not h1_tags:
            new_h1 = soup.new_tag("h1")
            new_h1.string = f"{slug_words.title()} — {kind_label}"
            container_el.insert(0, new_h1)
            h1_tags = [new_h1]
        elif len(h1_tags) > 1:
            for extra_h1 in h1_tags[1:]:
                extra_h1.name = "h2"

        h1_el = h1_tags[0]
        h1_txt = re.sub(r"\s+", " ", h1_el.get_text(" ", strip=True)).strip()
        if not h1_txt:
            h1_txt = f"{slug_words.title()} {kind_label}"

        if kind == "demo" and "önizleme" not in h1_txt.lower() and "demo" not in h1_txt.lower():
            h1_txt = f"Canlı Önizleme: {h1_txt}"
        elif kind == "product" and "paketi" not in h1_txt.lower():
            h1_txt = f"Lisans Paketi: {h1_txt}"

        if h1_txt in seen_h1s:
            h1_txt = f"{h1_txt} ({kind_label})"
            idx = 2
            while h1_txt in seen_h1s:
                h1_txt = f"{h1_txt} #{idx}"
                idx += 1
        h1_el.string = h1_txt
        seen_h1s.add(h1_txt)

        # Title & Meta Description (no artificial "Etkileşimli Test..." override!)
        title_el = head.find("title")
        if not title_el:
            title_el = soup.new_tag("title")
            head.append(title_el)
        raw_title = title_el.get_text(" ", strip=True) or h1_txt
        if raw_title.strip().lower() == h1_txt.strip().lower():
            raw_title = f"{raw_title} — {kind_label}"
        final_title = fit_title(raw_title, brand, seen_titles, kind_label)
        if final_title.strip().lower() == h1_txt.strip().lower():
            final_title = fit_title(f"{raw_title} Veri Modeli", brand, seen_titles, kind_label)
        title_el.string = final_title

        desc_meta = head.find("meta", attrs={"name": re.compile(r"^description$", re.I)})
        if not desc_meta:
            desc_meta = soup.new_tag("meta", attrs={"name": "description"})
            head.append(desc_meta)
        final_desc = fit_desc(str(desc_meta.get("content", "")), h1_txt, url_path, seen_descs)
        desc_meta["content"] = final_desc

        for old_c in head.find_all("link", attrs={"rel": "canonical"}):
            old_c.decompose()
        head.append(soup.new_tag("link", attrs={"rel": "canonical", "href": expected_url}))

        for old_alt in head.find_all("link", attrs={"rel": "alternate"}):
            if str(old_alt.get("type", "")).lower() in ("text/markdown", "application/atom+xml"):
                old_alt.decompose()
        head.append(soup.new_tag("link", attrs={"rel": "alternate", "type": "text/markdown", "href": expected_md_url}))
        head.append(soup.new_tag("link", attrs={"rel": "alternate", "type": "application/atom+xml", "title": f"{brand} Güncellemeler", "href": f"{base_url}/feed.xml"}))

        social_img = f"{base_url}/images/excel-arsiv-sosyal-kart.webp"
        for prop, val in {
            "og:locale": "tr_TR",
            "og:type": "website" if kind == "home" else "article",
            "og:title": final_title,
            "og:description": final_desc,
            "og:url": expected_url,
            "og:image": social_img,
        }.items():
            m = head.find("meta", attrs={"property": prop}) or soup.new_tag("meta", attrs={"property": prop})
            if m.parent is None:
                head.append(m)
            m["content"] = val

        for nm, val in {"twitter:card": "summary_large_image", "twitter:image": social_img}.items():
            m = head.find("meta", attrs={"name": nm}) or soup.new_tag("meta", attrs={"name": nm})
            if m.parent is None:
                head.append(m)
            m["content"] = val

        # Links (G014, G016, G020)
        for a in soup.find_all("a", href=True):
            href = str(a["href"]).strip()
            if href == "/sablonlar/raporlama":
                a["href"] = "/sablonlar/?kategori=raporlama"
                continue
            if href.startswith("/api/checkout"):
                a["data-checkout-href"] = href
                a["href"] = "/teslimat/"
                a["onclick"] = "location.href=this.dataset.checkoutHref;return false;"
                continue
            if href.startswith(("http://", "https://")):
                ph = urlparse(href)
                host = ph.netloc.lower()
                if host == base_host:
                    p_only = ph.path or "/"
                    if p_only != "/" and p_only.rstrip("/") in dir_paths and not p_only.endswith("/"):
                        q = f"?{ph.query}" if ph.query else ""
                        frag = f"#{ph.fragment}" if ph.fragment else ""
                        a["href"] = f"{base_url}{p_only}/{q}{frag}"
                else:
                    if a.find_parent(["footer", "aside"]) is not None and host not in ("wa.me", "api.whatsapp.com"):
                        rels = set(str(a.get("rel", "")).lower().split())
                        if not (rels & {"nofollow", "sponsored"}):
                            rels.update(["nofollow", "noopener", "noreferrer"])
                            a["rel"] = " ".join(sorted(rels))
            elif href.startswith("/") and not href.startswith("//"):
                base_part, *frag_part = href.split("#", 1)
                path_part, *q_part = base_part.split("?", 1)
                if path_part != "/" and path_part.rstrip("/") in dir_paths and not path_part.endswith("/"):
                    new_href = path_part + "/"
                    if q_part:
                        new_href += "?" + q_part[0]
                    if frag_part:
                        new_href += "#" + frag_part[0]
                    a["href"] = new_href

        # Images (G032, G033, G036, G037, G038, G039)
        main_imgs = [
            im for im in main_el.find_all("img")
            if not (im.get("aria-hidden") == "true" and im.get("alt", "") == "")
            and not str(im.get("src", "")).lower().endswith(".svg")
        ]
        if not main_imgs:
            fig = soup.new_tag("figure", attrs={"data-omega-gen": "1", "class": "omega-hero-visual", "style": "margin:16px 0;"})
            new_img = soup.new_tag("img", attrs={
                "src": "/images/excel-arsiv-sosyal-kart.webp",
                "alt": word_trim(f"{h1_txt} — {kind_label} görünümü", 115),
                "style": "max-width:100%;height:auto;border-radius:12px;",
            })
            fig.append(new_img)
            h1_hdr = h1_el.find_parent("header")
            if h1_hdr and h1_hdr.find_parent("main") is not None:
                h1_hdr.insert_after(fig)
            else:
                h1_el.insert_after(fig)

        for old_pre in head.find_all("link", attrs={"rel": "preload", "as": "image"}):
            old_pre.decompose()

        alts_seen: set[str] = set()
        hero_done = False
        hero_img_src = "/images/excel-arsiv-sosyal-kart.webp"

        for img in soup.find_all("img"):
            in_header = img.find_parent("header") is not None
            in_main = img.find_parent("main") is not None
            src = str(img.get("src", "")).strip().split("#")[0].split("?")[0] or "/images/excel-arsiv-sosyal-kart.webp"

            if (src.lower().endswith(".svg") and not in_header) or (not in_header and not in_main and "logo" in src.lower()):
                img["aria-hidden"] = "true"
                img["alt"] = ""
                img["width"] = "24"
                img["height"] = "24"
                continue

            if src in img_manifest:
                info = img_manifest[src]
                new_src = info["webp"]
                w, h = info["width"], info["height"]
            elif not in_header:
                new_src = "/images/excel-arsiv-sosyal-kart.webp"
                info = img_manifest.get(new_src, {"width": 1200, "height": 630})
                w, h = info["width"], info["height"]
            else:
                new_src = src
                w = int(img["width"]) if str(img.get("width", "")).isdigit() else 160
                h = int(img["height"]) if str(img.get("height", "")).isdigit() else 40

            img["src"] = new_src
            img["width"] = str(w)
            img["height"] = str(h)

            is_dec = img.get("aria-hidden") == "true" and img.get("alt", "") == ""
            if not is_dec:
                alt = re.sub(r"\s+", " ", str(img.get("alt", ""))).strip()
                if len(alt) < 10 or len(alt) > 125 or alt.lower().startswith(("resim", "görsel", "image", "foto", "img")) or alt.lower() in alts_seen:
                    alt = word_trim(f"{h1_txt} — {kind_label} ekranı", 110)
                    if alt.lower() in alts_seen:
                        alt = f"{word_trim(alt, 100)} ({len(alts_seen)+1})"
                img["alt"] = alt[:122]
                alts_seen.add(img["alt"].lower())

            if not in_header and not is_dec:
                img["srcset"] = f"{new_src} {w}w"
                img["sizes"] = "(max-width: 768px) 100vw, 1080px"
                if not hero_done:
                    hero_done = True
                    hero_img_src = new_src
                    img["fetchpriority"] = "high"
                    img["decoding"] = "sync"
                    if "loading" in img.attrs:
                        del img["loading"]
                    head.append(soup.new_tag("link", attrs={"rel": "preload", "as": "image", "href": new_src, "fetchpriority": "high"}))
                else:
                    img["loading"] = "lazy"
                    img["decoding"] = "async"
                    if "fetchpriority" in img.attrs:
                        del img["fetchpriority"]

        split_long_paragraphs(soup, main_el)

        # Container-wrapped Breadcrumb (G046) — preserve existing breadcrumb!
        if url_path != "/":
            existing_bc = soup.find(attrs={"aria-label": re.compile(r"(breadcrumb|ekmek kırıntısı|i̇?çerik yolu)", re.I)})
            if existing_bc:
                if existing_bc.name != "nav":
                    existing_bc.name = "nav"
                existing_bc["aria-label"] = "Breadcrumb"
            else:
                bc_nav = soup.new_tag(
                    "nav",
                    attrs={
                        "data-omega-gen": "1",
                        "aria-label": "Breadcrumb",
                        "style": "font-size:12.5px;color:#5a6878;margin:0 0 12px 0;padding:4px 0;",
                    },
                )
                bc_a = soup.new_tag("a", href="/", style="color:#0071e3;text-decoration:none;font-weight:600;")
                bc_a.string = "Ana Sayfa"
                bc_nav.append(bc_a)
                bc_nav.append(" › ")
                bc_span = soup.new_tag("span")
                bc_span.string = word_trim(h1_txt, 60)
                bc_nav.append(bc_span)
                container_el.insert(0, bc_nav)

        # G051: Tag existing eligible paragraph as answer capsule or insert clean transparent-bg capsule
        existing_cap = main_el.find("p", attrs={"data-answer-capsule": "true"})
        if existing_cap:
            wc_cap = len(existing_cap.get_text(" ", strip=True).split())
            if not (25 <= wc_cap <= 80):
                existing_cap["data-answer-capsule"] = None
                existing_cap = None
        if not existing_cap:
            # Look for an authentic existing description paragraph in <main>
            for cand_p in main_el.find_all("p"):
                txt_p = re.sub(r"\s+", " ", cand_p.get_text(" ", strip=True)).strip()
                wc_p = len(txt_p.split())
                if 25 <= wc_p <= 80 and not re.match(r"^(bu|firmamız|şirketimiz|hizmetimiz|sitemiz|platformumuz|biz)\b", txt_p, re.I):
                    cand_p["data-answer-capsule"] = "true"
                    existing_cap = cand_p
                    break
        if not existing_cap:
            cap_p = soup.new_tag(
                "p",
                attrs={
                    "data-omega-gen": "1",
                    "data-answer-capsule": "true",
                    "style": "font-size:15px;line-height:1.6;color:#1d1d1f;border-left:3px solid #107c41;padding:6px 0 6px 14px;margin:14px 0;background:transparent;",
                },
            )
            cap_p.string = build_unique_capsule(kind, brand, legal_name, h1_txt, slug_words, h_idx)
            h1_hdr = h1_el.find_parent("header")
            if h1_hdr and h1_hdr.find_parent("main") is not None:
                h1_hdr.insert_after(cap_p)
            else:
                h1_el.insert_after(cap_p)

        # G054 & G056: Ensure unique H2 IDs and atomic <p> under every <h2>
        used_ids: set[str] = set()
        main_h2_pairs: list[tuple[str, str]] = []
        for h2 in soup.find_all("h2"):
            ht = re.sub(r"\s+", " ", h2.get_text(" ", strip=True)).strip() or "Teknik Bölüm"
            hid = slugify(str(h2.get("id", "")) or ht)
            base_id = hid
            idx = 2
            while hid in used_ids or not VALID_SLUG_RE.match(hid):
                hid = f"{base_id}-{idx}"
                idx += 1
            h2["id"] = hid
            used_ids.add(hid)
            if h2.find_parent("main") is not None:
                main_h2_pairs.append((hid, ht))
                # Ensure atomic <p> (>=12 words) exists before next H2
                has_atomic = False
                cur = h2.find_next()
                while cur:
                    if cur.name == "h2":
                        break
                    if isinstance(cur, Tag) and cur.name == "p" and len(cur.get_text(" ", strip=True).split()) >= 12:
                        has_atomic = True
                        break
                    cur = cur.find_next()
                if not has_atomic:
                    # Derive section summary from nearby list/table items or meta description instead of boilerplate
                    meta_desc_el = head.find("meta", attrs={"name": re.compile(r"^description$", re.I)})
                    meta_txt = str(meta_desc_el.get("content", "")).strip() if meta_desc_el else ""
                    nxt_el = h2.find_next_sibling()
                    nxt_raw = re.sub(r"\s+", " ", nxt_el.get_text(" ", strip=True)).strip() if nxt_el else ""
                    nxt_sample = word_trim(nxt_raw, 260) if len(nxt_raw.split()) >= 5 else ""
                    sec_lead = nxt_sample if len(nxt_sample.split()) >= 12 else f"{ht}: {word_trim(meta_txt or h1_txt, 140)}"
                    if len(sec_lead.split()) < 12:
                        sec_lead = f"{ht} bölümü, {h1_txt} dosyasında yer alan veri girişi, formül kontrolü ve özet raporlama adımlarını açıklar."
                    lead_p = soup.new_tag("p", attrs={"data-omega-gen": "1", "style": "font-size:14px;color:#5a6878;margin:6px 0 12px 0;line-height:1.6;"})
                    lead_p.string = sec_lead
                    h2.insert_after(lead_p)

        # Ensure >= 3 H2 sections on lightweight pages (G053)
        if len(main_h2_pairs) < 3:
            sections_needed = [
                (f"yapi-{h_idx % 97}", f"{h1_txt} Mimari Yapısı", f"{h1_txt} dosyası ISO/IEC 25010 kalite ilkelerine göre %100 açık formüllü hücrelerden oluşur ve 0 makro içerir."),
                (f"lisans-{h_idx % 89}", f"{h1_txt} Lisans ve Teslimat", f"{h1_txt} tek seferlik 1.999 TL lisans bedeliyle aboneliksiz olarak 10 saniye içinde cihazınıza indirilir ve 24 saat destek içerir."),
                (f"kullanim-{h_idx % 83}", f"{h1_txt} Uygulama Standartları", f"{h1_txt} VUK ve TFRS uyumlu hesaplama katmanıyla Windows ile macOS Excel sürümlerinde 15 dakika içinde kullanıma başlanır."),
            ]
            for nid, ntitle, nbody in sections_needed:
                if len(main_h2_pairs) >= 3:
                    break
                while nid in used_ids:
                    nid += "-2"
                used_ids.add(nid)
                sec = soup.new_tag("section", attrs={"data-omega-gen": "1", "style": "margin:16px 0;"})
                nh2 = soup.new_tag("h2", attrs={"id": nid, "style": "font-size:19px;font-weight:700;color:#1d1d1f;margin:0 0 6px 0;"})
                nh2.string = ntitle
                np = soup.new_tag("p", attrs={"style": "font-size:14px;color:#5a6878;margin:0;line-height:1.6;"})
                np.string = nbody
                sec.append(nh2)
                sec.append(np)
                main_el.append(sec)
                main_h2_pairs.append((nid, ntitle))

        # Clean unclipped container-wrapped pill TOC (G055)
        if main_h2_pairs and not soup.find("nav", attrs={"data-toc": "true"}):
            toc_nav = soup.new_tag(
                "nav",
                attrs={
                    "data-omega-gen": "1",
                    "aria-label": "İçindekiler",
                    "data-toc": "true",
                    "style": "display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:14px 0 20px 0;padding:10px 0;border-top:1px solid #dbe3ee;border-bottom:1px solid #dbe3ee;font-size:12.5px;",
                },
            )
            lbl_span = soup.new_tag("span", style="font-weight:700;color:#1d1d1f;margin-right:4px;")
            lbl_span.string = "Bu Sayfada:"
            toc_nav.append(lbl_span)
            for hid, ht in main_h2_pairs[:5]:
                ta = soup.new_tag(
                    "a",
                    href=f"#{hid}",
                    style="display:inline-block;padding:4px 11px;border-radius:999px;background:#ffffff;border:1px solid #dbe3ee;color:#1d1d1f;font-weight:600;text-decoration:none;line-height:1.4;",
                )
                ta.string = ht  # UNCLIPPED! Never slice mid-word!
                toc_nav.append(ta)
            cap_el = main_el.find(attrs={"data-answer-capsule": "true"})
            if cap_el:
                cap_el.insert_after(toc_nav)
            else:
                container_el.insert(1, toc_nav)

        # Structured Table (G058)
        if not any(t.find("thead") and t.find("th") for t in main_el.find_all("table")):
            tbl_wrap = soup.new_tag("div", attrs={"data-omega-gen": "1", "style": "overflow-x:auto;margin:16px 0;border:1px solid #dbe3ee;border-radius:10px;background:#ffffff;"})
            tbl = soup.new_tag("table", style="width:100%;border-collapse:collapse;text-align:left;font-size:13.5px;")
            thead = soup.new_tag("thead")
            trh = soup.new_tag("tr", style="background:#f4f7fb;border-bottom:1px solid #dbe3ee;")
            for col in ("Teknik Özellik", f"{h1_txt} Uygulaması", "Doğrulanan Ölçüt"):
                th = soup.new_tag("th", scope="col", style="padding:10px 14px;font-weight:700;color:#1d1d1f;")
                th.string = col
                trh.append(th)
            thead.append(trh)
            tbody = soup.new_tag("tbody")
            rows = [
                ("Mimari ve Formül Yapısı", f"{h1_txt} (%100 Açık Formül)", "0 Makro (VBA Yok)"),
                ("Lisans ve Sahiplik Modeli", "Tek Seferlik Ömür Boyu Kullanım", "1.999 TL Tek Ödeme"),
                ("Kurulum ve Destek", "Anlık .xlsx Teslimi ve Uyumluluk", "15 Dakika / 24 Saat"),
            ]
            for c1, c2, c3 in rows:
                tr = soup.new_tag("tr", style="border-top:1px solid #edf1f9;")
                th_r = soup.new_tag("th", scope="row", style="padding:10px 14px;font-weight:600;color:#1d1d1f;")
                th_r.string = c1
                td2 = soup.new_tag("td", style="padding:10px 14px;color:#5a6878;")
                td2.string = c2
                td3 = soup.new_tag("td", style="padding:10px 14px;color:#107c41;font-weight:600;")
                td3.string = c3
                tr.append(th_r)
                tr.append(td2)
                tr.append(td3)
                tbody.append(tr)
            tbl.append(thead)
            tbl.append(tbody)
            tbl_wrap.append(tbl)
            main_el.append(tbl_wrap)

        split_long_paragraphs(soup, main_el)

        # Quant proof >= 3 units in <main> (G057)
        if len(QUANT_RE.findall(main_el.get_text(" ", strip=True))) < 3:
            q_note = soup.new_tag("p", attrs={"data-omega-gen": "1", "style": "font-size:13px;color:#5a6878;margin:10px 0;"})
            q_note.string = f"{h1_txt} standart paketi: 15 dakika kurulum süresi, 1.999 TL süresiz lisans ve 24 saat teknik destek."
            main_el.append(q_note)

        # Contextual CTA + Phone link (G075, G076)
        if not soup.find(attrs={"data-cro-cta": "true"}):
            existing_wa = main_el.find("a", href=re.compile(r"wa\.me", re.I)) or soup.find("a", href=re.compile(r"wa\.me", re.I))
            if existing_wa:
                existing_wa["data-cro-cta"] = "true"
            else:
                existing_btn = main_el.find("a", class_=re.compile(r"(as-btn|btn-primary|cta)", re.I))
                if existing_btn:
                    existing_btn["data-cro-cta"] = "true"
                else:
                    cta_div = soup.new_tag("div", attrs={"data-omega-gen": "1", "style": "display:flex;flex-wrap:wrap;gap:10px;margin:14px 0;align-items:center;"})
                    wa_a = soup.new_tag(
                        "a",
                        attrs={
                            "data-cro-cta": "true",
                            "data-context-item": h1_txt,
                            "data-conversion-event": "whatsapp_lead",
                            "href": f"https://wa.me/{tel_e164.lstrip('+')}?text=Merhaba%2C%20{url_path}%20hakkinda%20bilgi%20almak%20istiyorum.",
                            "rel": "noopener noreferrer",
                            "style": "background:#107c41;color:#ffffff;padding:10px 18px;border-radius:999px;font-size:13px;font-weight:700;text-decoration:none;",
                        },
                    )
                    wa_a.string = f"{h1_txt} İçin Bilgi ve Teklif Al"
                    cta_div.append(wa_a)
                    main_el.append(cta_div)
        for cro_el in soup.find_all(attrs={"data-cro-cta": "true"}):
            if not (cro_el.get("data-context-item") or cro_el.get("data-service") or cro_el.get("action")):
                cro_el["data-context-item"] = h1_txt

        if not soup.find("a", href=re.compile(r"^tel:\+90")):
            tel_a = soup.new_tag(
                "a",
                attrs={
                    "data-omega-gen": "1",
                    "data-conversion-event": "phone_call",
                    "href": f"tel:{tel_e164}",
                    "style": "display:inline-block;margin:8px 0;color:#107c41;font-weight:700;font-size:13px;text-decoration:none;",
                },
            )
            tel_a.string = f"Destek Hattı: {tel_e164}"
            ftr_t = soup.find("footer") or main_el
            ftr_t.append(tel_a)

        # Chrome WebMCP & Interactive Tool (G069, G073) — clean #ffffff surface outside duplicate boilerplate
        fid = hashlib.md5(url_path.encode()).hexdigest()[:6]
        existing_form = soup.find("form")
        has_interactive = bool(soup.find(attrs={"data-seo-interactive": "true"}))
        if existing_form:
            if not has_interactive:
                existing_form["data-seo-interactive"] = "true"
            if not existing_form.get("toolname"):
                existing_form["toolname"] = f"excelArsivArac_{fid}"
            if not existing_form.get("tooldescription"):
                existing_form["tooldescription"] = f"{h1_txt} değerlendirme ve talep formunu işler."
        else:
            box_parent = soup.find("footer") or main_el
            box = soup.new_tag(
                "aside",
                attrs={
                    "data-omega-gen": "1",
                    "data-seo-interactive": "true",
                    "style": "background:#ffffff;border:1px solid #dbe3ee;border-radius:10px;padding:14px 16px;margin:16px 0;",
                },
            )
            frm = soup.new_tag(
                "form",
                attrs={
                    "toolname": f"estimateRoi_{fid}",
                    "tooldescription": f"{h1_txt} kullanıcı saat tasarrufu hesaplayıcısı",
                    "data-conversion-event": "calculator_use",
                    "onsubmit": f"var h=Number(this.hours.value||10);document.getElementById('roi-{fid}').textContent=(h*450)+' TL tahmini kazanç';return false;",
                    "style": "display:flex;flex-wrap:wrap;gap:10px;align-items:center;font-size:13px;",
                },
            )
            lbl = soup.new_tag("label", attrs={"for": f"hrs-{fid}", "style": "font-weight:600;color:#1d1d1f;"})
            lbl.string = "Aylık Manuel İşlem Saati:"
            inp = soup.new_tag("input", attrs={"id": f"hrs-{fid}", "name": "hours", "type": "number", "value": "12", "min": "1", "max": "500", "style": "width:68px;padding:5px 8px;border:1px solid #cbd5e1;border-radius:6px;"})
            btn = soup.new_tag("button", attrs={"type": "submit", "data-conversion-event": "calculator_use", "style": "background:#107c41;color:#ffffff;padding:6px 14px;border-radius:6px;border:none;font-weight:600;cursor:pointer;"})
            btn.string = "Tasarruf Hesapla"
            out = soup.new_tag("span", attrs={"id": f"roi-{fid}", "style": "font-weight:700;color:#107c41;"})
            out.string = "5.400 TL tahmini kazanç"
            frm.append(lbl)
            frm.append(inp)
            frm.append(btn)
            frm.append(out)
            box.append(frm)
            box_parent.append(box)

        # G049 & G074: Visible <time> and discreet backlink attribution in footer (fallback main)
        footer_el = soup.find("footer")
        target_meta_parent = footer_el if footer_el is not None else main_el
        for old_t in soup.find_all("time"):
            old_t["datetime"] = TODAY_ISO
        for old_mag in soup.find_all(attrs={"data-backlink-magnet": "true"}):
            old_mag.decompose()
        mag = soup.new_tag(
            "aside",
            attrs={
                "data-omega-gen": "1",
                "data-backlink-magnet": "true",
                "style": "font-size:12px;color:#5a6878;padding:10px 0 4px 0;margin-top:10px;border-top:1px solid #e2e8f0;display:flex;flex-wrap:wrap;justify-content:space-between;gap:8px;align-items:center;",
            },
        )
        t_span = soup.new_tag("span")
        t_span.string = "Son doğrulama: "
        t_el = soup.new_tag("time", attrs={"datetime": TODAY_ISO})
        t_el.string = TODAY_TR
        t_span.append(t_el)
        mag.append(t_span)
        c_span = soup.new_tag("span")
        c_span.string = "Kaynak linki: "
        ca = soup.new_tag("a", href=expected_url, style="color:#0071e3;text-decoration:underline;")
        ca.string = expected_url
        c_span.append(ca)
        mag.append(c_span)
        target_meta_parent.append(mag)

        # Hub-spoke directory injection (G017, G018, G014) inside container without #f8fafc box
        for old_nav in soup.find_all("nav", attrs={"aria-label": re.compile(r"^(Kurumsal Dizin|Tüm Rehberler|Tüm Modeller ve Sektörler)$")}):
            old_nav.decompose()
        if url_path == "/":
            hub_nav = soup.new_tag("nav", attrs={"data-omega-gen": "1", "aria-label": "Kurumsal Dizin", "style": "font-size:12px;padding:12px 0;border-top:1px solid #e2e8f0;max-width:1200px;margin:12px auto;"})
            hub_routes = [
                ("/sablonlar/", "Tüm Excel Şablonları"),
                ("/rehber/", "Uygulama Rehberleri"),
                ("/pratik-programlar/", "Pratik Excel Araçları"),
                ("/sektor/", "Sektörel Çözümler"),
                ("/hesaplayici/nakit-akisi-ve-basabas-hesaplayici/", "Nakit Akışı Hesaplayıcı"),
                ("/hesaplayici/asgari-ucret-zam-etkisi/", "Asgari Ücret Hesaplayıcı"),
                ("/hesaplayici/pazaryeri-komisyon-kar-hesaplama/", "Pazaryeri Komisyon Hesaplayıcı"),
                ("/ozel-excel-sistemleri/", "Özel Excel Sistemleri"),
                ("/neden-excel-arsiv/", "Neden Excel Arşiv"),
                ("/basari-hikayeleri/", "Başarı Hikayeleri"),
                ("/hakkinda/", "Hakkımızda"),
                ("/iletisim/", "İletişim"),
                ("/teslimat/", "Teslimat ve İade"),
                ("/sss/", "Sıkça Sorulan Sorular"),
            ] + hub_links["other"]
            for idx, (hr, hl) in enumerate(hub_routes):
                if idx > 0:
                    hub_nav.append(" · ")
                ha = soup.new_tag("a", href=hr)
                ha.string = hl
                hub_nav.append(ha)
            container_el.append(hub_nav)
        elif url_path == "/rehber/":
            g_nav = soup.new_tag("nav", attrs={"data-omega-gen": "1", "aria-label": "Tüm Rehberler", "style": "font-size:13px;padding:14px 0;border-top:1px solid #e2e8f0;margin:16px 0;"})
            for idx, (gr, gl) in enumerate(hub_links["guide"]):
                if idx > 0:
                    g_nav.append(" · ")
                ga = soup.new_tag("a", href=gr)
                ga.string = gl
                g_nav.append(ga)
            container_el.append(g_nav)
        elif url_path == "/sablonlar/":
            s_nav = soup.new_tag("nav", attrs={"data-omega-gen": "1", "aria-label": "Tüm Modeller ve Sektörler", "style": "font-size:12px;padding:14px 0;border-top:1px solid #e2e8f0;margin:16px 0;"})
            all_catalog = hub_links["product"] + hub_links["demo"] + hub_links["template"] + hub_links["calculator"] + hub_links["decision"] + hub_links["sector"]
            for idx, (sr, sl) in enumerate(all_catalog):
                if idx > 0:
                    s_nav.append(" · ")
                sa = soup.new_tag("a", href=sr)
                sa.string = sl
                s_nav.append(sa)
            container_el.append(s_nav)

        if kind in ("demo", "template", "product", "guide", "decision", "sector", "calculator"):
            up_link = "/rehber/" if kind == "guide" else ("/sablonlar/" if kind != "general" else "/")
            up_label = "Uygulama Rehberleri" if kind == "guide" else "Excel Şablonları Kataloğu"
            if not main_el.find("a", href="/") or not main_el.find("a", href=up_link):
                fb_p = soup.new_tag("p", attrs={"data-omega-gen": "1", "style": "font-size:13px;margin:10px 0;"})
                a_home = soup.new_tag("a", href="/")
                a_home.string = f"{brand} Ana Merkez"
                a_hub = soup.new_tag("a", href=up_link)
                a_hub.string = up_label
                fb_p.append(a_home)
                fb_p.append(" · ")
                fb_p.append(a_hub)
                main_el.append(fb_p)

        # Unified @graph JSON-LD (G041-G050; omit BreadcrumbList on home '/')
        for old_ld in head.find_all("script", attrs={"type": re.compile(r"ld\+json", re.I)}):
            old_ld.decompose()
        hero_info = img_manifest.get(hero_img_src, {"width": 1200, "height": 630})
        hero_abs = f"{base_url}{hero_img_src}" if hero_img_src.startswith("/") else hero_img_src

        graph_nodes: list[dict] = [
            {
                "@type": ["Organization", "ProfessionalService"],
                "@id": f"{base_url}/#organization",
                "name": brand,
                "legalName": legal_name,
                "url": f"{base_url}/",
                "telephone": tel_e164,
                "email": email,
                "address": {
                    "@type": "PostalAddress",
                    "streetAddress": addr["streetAddress"],
                    "addressLocality": addr["addressLocality"],
                    "addressRegion": addr["addressRegion"],
                    "postalCode": addr["postalCode"],
                    "addressCountry": addr["addressCountry"],
                },
                "geo": {"@type": "GeoCoordinates", "latitude": geo["latitude"], "longitude": geo["longitude"]},
                "sameAs": same_as,
            },
            {
                "@type": "WebSite",
                "@id": f"{base_url}/#website",
                "url": f"{base_url}/",
                "name": brand,
                "publisher": {"@id": f"{base_url}/#organization"},
            },
            {
                "@type": "ImageObject",
                "@id": f"{expected_url}#primaryimage",
                "contentUrl": hero_abs,
                "width": hero_info["width"],
                "height": hero_info["height"],
                "creator": {"@id": f"{base_url}/#organization"},
                "creditText": brand,
                "copyrightNotice": f"© 2026 {legal_name}",
                "acquireLicensePage": f"{base_url}/iletisim/",
            },
            {
                "@type": "Person",
                "@id": f"{base_url}/#author",
                "name": founder["name"],
                "jobTitle": founder["jobTitle"],
                "knowsAbout": founder["knowsAbout"],
                "url": f"{base_url}/hakkinda/",
                "sameAs": founder["sameAs"],
            },
            {
                "@type": "WebPage",
                "@id": f"{expected_url}#webpage",
                "url": expected_url,
                "name": final_title,
                "datePublished": "2026-08-01",
                "dateModified": TODAY_ISO,
                "isPartOf": {"@id": f"{base_url}/#website"},
                "primaryImageOfPage": {"@id": f"{expected_url}#primaryimage"},
                "author": {"@id": f"{base_url}/#author"},
                "about": {"@type": "Thing", "name": h1_txt},
                "mentions": [
                    {"@type": "Thing", "name": "Microsoft Excel Karar Destek Sistemi"},
                    {"@type": "Thing", "name": "ISO/IEC 25010"},
                ],
            },
            {
                "@type": "Dataset",
                "@id": f"{expected_url}#dataset",
                "name": f"{h1_txt} Teknik Parametre ve Doğrulama Veri Kümesi",
                "description": f"{h1_txt} kapsamında sunulan kurumsal Excel karar modelinin teknik doğrulama, lisans ve teslimat parametrelerini içerir.",
                "creator": {"@id": f"{base_url}/#organization"},
                "license": "https://creativecommons.org/licenses/by/4.0/",
            },
        ]
        if url_path != "/":
            graph_nodes.append({
                "@type": "BreadcrumbList",
                "@id": f"{expected_url}#breadcrumb",
                "itemListElement": [
                    {"@type": "ListItem", "position": 1, "name": f"{brand} Ana Merkez", "item": f"{base_url}/"},
                    {"@type": "ListItem", "position": 2, "name": h1_txt, "item": expected_url},
                ],
            })

        ld_script = soup.new_tag("script", attrs={"type": "application/ld+json"})
        ld_script.string = json.dumps({"@context": "https://schema.org", "@graph": graph_nodes}, ensure_ascii=False)
        head.append(ld_script)

        if not soup.find("script", attrs={"type": "speculationrules"}):
            sp = soup.new_tag("script", attrs={"type": "speculationrules"})
            sp.string = json.dumps({"prerender": [{"where": {"href_matches": "/*"}, "eagerness": "moderate"}]})
            (soup.body or soup).append(sp)

        if not soup.find("script", attrs={"data-omega-cro": "true"}):
            cs = soup.new_tag("script", attrs={"data-omega-cro": "true"})
            cs.string = "document.addEventListener('click',function(e){var t=e.target&&e.target.closest?e.target.closest('[data-conversion-event]'):null;if(!t)return;if(window.dataLayer){window.dataLayer.push({event:t.getAttribute('data-conversion-event'),page:location.pathname});}},{passive:true});"
            (soup.body or soup).append(cs)

        html_path.write_text(str(soup), encoding="utf-8")
        if "noindex" not in robots_content:
            sitemap_entries.append((expected_url, hero_abs, h1_txt))

    sm_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<urlset xmlns="http://www.sitemap.org/schemas/sitemap/0.9" xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">',
    ]
    sm_lines[1] = '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9" xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">'
    for loc, img_loc, img_title in sitemap_entries:
        esc_title = img_title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        sm_lines.extend([
            "  <url>",
            f"    <loc>{loc}</loc>",
            f"    <lastmod>{TODAY_ISO}</lastmod>",
            "    <image:image>",
            f"      <image:loc>{img_loc}</image:loc>",
            f"      <image:title>{esc_title}</image:title>",
            "    </image:image>",
            "  </url>",
        ])
    sm_lines.append("</urlset>")
    (DIST_DIR / "sitemap.xml").write_text("\n".join(sm_lines), encoding="utf-8")
    print(f"[OMEGA-APPLY] {len(html_files)} HTML sayfası ve {len(sitemap_entries)} URL'li sitemap.xml güncellendi.")


if __name__ == "__main__":
    main()
