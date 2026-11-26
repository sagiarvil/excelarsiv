#!/usr/bin/env python3
"""
seo_enforcer.py — REAL-SEO-2026 (APEX v6.0 ABSOLUTE)
100 Atomik Kapılı (100 x 1p = 100/100; 99/100 Dahi Reddedilir) Otonom Yazılım, Gelir (BOFU),
Pasif Backlink Mıknatısı, Dataset/ImageObject Knowledge Graph, 4-Gram Özgünlük, Derin #id Sitelinks,
Tavsiye Kalibrasyonlu llms.txt + .md İkizi + Atom Feed Üreticisi ve Gerçek Headless Chromium Motoru.

Google Antigravity ("Gravity"), Cursor, Claude Code ve GitHub Copilot için her başarısız kapıda
doğrudan kod düzeyinde [ÇÖZÜM REÇETESİ] üretir.

Komutlar:
  1) Kurulum ve Antigravity/Cursor/Claude/CI kilitlerini açmak için:
       python3 seo_enforcer.py --init --domain https://ornekdomain.com

  2) .md ikizleri, /llms.txt, /llms-full.txt, /feed.xml, /sitemap.xml, /robots.txt, /404.html ve security.txt üretimi:
       python3 seo_enforcer.py --build-ai-pack ./dist

  3) 100 atomik kapılı (100/100) tam denetim (Statik + Semantik Vektör + 4-Gram + Headless Chromium):
       python3 seo_enforcer.py --audit ./dist https://ornekdomain.com

  4) Google Search Console (CSV / API) 3-vektörlü aksiyon planı ve sunucu log denetimi:
       python3 seo_enforcer.py --gsc-file gsc.csv --build-dir ./dist
       python3 seo_enforcer.py --log-audit access.log ./dist https://ornekdomain.com
       python3 seo_enforcer.py --ping-indexnow ./dist
"""

from __future__ import annotations

import argparse
import collections
import csv
import datetime as dt
import functools
import hashlib
import http.server
import json
import math
import os
import re
import shutil
import socket
import stat
import sys
import threading
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple
from urllib.parse import urlparse

FORBIDDEN_EXTS = {".bak", ".old", ".tmp", ".orig", ".swp", ".save"}
FORBIDDEN_FILENAMES = {"test.html", "draft.html", "temp.html", "copy.html", "deneme.html", "old.html"}
RASTER_EXTS = (".avif", ".webp", ".jpg", ".jpeg", ".png")
MODERN_PHOTO_EXTS = (".avif", ".webp")
GENERIC_ANCHORS = {
    "tıklayın", "buraya tıklayın", "detaylar", "devamı", "incele", "tıkla",
    "click here", "read more", "learn more", "more", "link", "buradan",
}
INTERNAL_CODE_PREFIX_RE = re.compile(r"^(MOD|SKU|PRD|ITEM|COUNTER|KOD|ERP|REF|ID)[-_ ]?\d+", re.IGNORECASE)
BAD_IMAGE_FILENAME_RE = re.compile(
    r"^(img|dsc|dcim|screenshot|ekran|whatsapp|image|photo|resim|foto|pic|untitled|adsiz|gorsel)[-_0-9.]",
    re.IGNORECASE,
)
VALID_SLUG_SEGMENT_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
VALID_IMG_BASENAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)+$")
E164_PHONE_RE = re.compile(r"^\+[1-9]\d{9,14}$")
VAGUE_PRONOUN_START_RE = re.compile(
    r"^(bu|şu|o|onlar|bunlar|firmamız|şirketimiz|ekibimiz|biz|burada|hizmetimiz|ürünümüz)\b",
    re.IGNORECASE,
)
QUANT_PROOF_RE = re.compile(
    r"(\b\d+(?:[.,]\d+)?\s*(?:%|mm|cm|m2|m³|kg|ton|kw|mw|saat|gün|yıl|ay|dk|saniye|ms|tl|₺|usd|\$|€|adet|bar|db|v|hz)\b|\biso\s*\d+|\btse\b|\bce\b)",
    re.IGNORECASE,
)
AUTHORITY_REF_RE = re.compile(
    r"(iso\s*\d+|tse|resm[iî]\s+gazete|yönetmelik|standart|rapor|araştırma|ölçüm|test|sertifika|bakanlığ|kurum|ieee|astm|din\s*\d+|en\s*\d+)",
    re.IGNORECASE,
)
HEAVY_THIRD_PARTY_SYNC_RE = re.compile(
    r"(tawk\.to|crisp\.chat|intercom\.io|connect\.facebook\.net|hotjar\.com|clarity\.ms)",
    re.IGNORECASE,
)

IGNORED_SCAN_DIRS = {
    ".git", "node_modules", ".next", ".astro", ".vercel", ".cursor",
    ".claude", ".agents", ".agent", ".github", "__pycache__", "venv", ".venv", "scripts",
}

TR_STOPWORDS = {
    "ve", "ile", "için", "bir", "bu", "şu", "o", "de", "da", "ki", "mi", "mı", "mu", "mü",
    "en", "çok", "daha", "gibi", "kadar", "olan", "olarak", "nedir", "nasıl", "ne", "veya",
    "ya", "hem", "ama", "fakat", "ancak", "tüm", "her", "üzere", "göre", "ait", "ise",
}

AI_SEARCH_BOTS_REQUIRED = [
    "Googlebot",
    "Bingbot",
    "OAI-SearchBot",
    "ChatGPT-User",
    "Claude-SearchBot",
    "Claude-User",
    "PerplexityBot",
]

AI_TRAINING_BOTS = ["GPTBot", "ClaudeBot", "Google-Extended", "CCBot"]

PERF_OBSERVER_INIT_JS = """
window.__seoAuditMetrics = { cls: 0, longTasks: [], maxLongTask: 0, jsErrors: [] };
window.addEventListener('error', function(e) {
  window.__seoAuditMetrics.jsErrors.push(e.message || 'runtime error');
});
window.addEventListener('unhandledrejection', function(e) {
  window.__seoAuditMetrics.jsErrors.push(String(e.reason || 'unhandled promise rejection'));
});
try {
  new PerformanceObserver(function(list) {
    for (const entry of list.getEntries()) {
      if (!entry.hadRecentInput) {
        window.__seoAuditMetrics.cls += entry.value;
      }
    }
  }).observe({ type: 'layout-shift', buffered: true });
} catch (_) {}
try {
  new PerformanceObserver(function(list) {
    for (const entry of list.getEntries()) {
      window.__seoAuditMetrics.longTasks.push(entry.duration);
      if (entry.duration > window.__seoAuditMetrics.maxLongTask) {
        window.__seoAuditMetrics.maxLongTask = entry.duration;
      }
    }
  }).observe({ type: 'longtask', buffered: true });
} catch (_) {}
"""


class AbsoluteHTMLParser(HTMLParser):
    """100 Atomik Kapının tüm DOM, <head>, JSON-LD, görsel, #id TOC, backlink mıknatısı ve CRO sinyallerini ayrıştırır."""

    def __init__(self) -> None:
        super().__init__()
        self.has_doctype = False
        self.html_lang = ""
        self.has_charset_utf8 = False
        self.viewport_content = ""

        self.in_head = False
        self.in_title = False
        self.in_h1 = False
        self.in_h2 = False
        self.in_h3 = False
        self.in_p = False
        self.in_a = False
        self.in_header_or_nav = 0
        self.in_main = 0
        self.in_footer_or_aside = 0
        self.in_toc_nav = 0
        self.in_jsonld = False
        self.in_speculation = False
        self.in_script = False
        self.in_style = False
        self.capsule_tag_stack: List[str] = []
        self.magnet_tag_stack: List[str] = []
        self.in_table = False
        self.in_tr = False
        self.in_cell = False

        self.has_main = False
        self.has_header_or_nav = False
        self.has_footer = False
        self.has_breadcrumb_nav = False
        self.has_toc_nav = False
        self.toc_anchors: List[str] = []

        self.title = ""
        self.h1_list: List[str] = []
        self.h2_items: List[Tuple[str, str]] = []  # (id, text)
        self.h3_list: List[str] = []
        self.paragraphs: List[str] = []
        self.h2_lead_paragraphs: List[str] = []
        self._awaiting_h2_lead = False

        self._cur_h1 = ""
        self._cur_h2 = ""
        self._cur_h2_id = ""
        self._cur_h3 = ""
        self._cur_p = ""

        self.metas: Dict[str, str] = {}
        self.canonical: str | None = None
        self.md_alternate: str | None = None
        self.atom_alternate: str | None = None
        self.preload_images: List[str] = []
        self.font_stylesheets: List[str] = []
        self.inline_styles: List[str] = []
        self.hreflangs: List[Tuple[str, str]] = []
        self.blocking_head_scripts: List[str] = []
        self.heavy_sync_third_party: List[str] = []
        self.local_assets: List[str] = []

        self.imgs: List[Dict[str, str]] = []
        self.in_picture = 0
        self._picture_has_source = False
        self.links: List[Dict[str, str]] = []
        self._cur_a_attrs: Dict[str, str] = {}
        self._cur_a_text = ""

        self.time_datetimes: List[str] = []
        self.jsonld_blocks: List[str] = []
        self._cur_script = ""
        self._cur_style = ""
        self.speculation_blocks: List[str] = []

        self.has_interactive_widget = False
        self.has_backlink_magnet = False
        self._cur_magnet_raw = ""
        self.magnet_texts: List[str] = []
        self.has_cro_cta = False
        self.cro_cta_contextual = False
        self.has_tel_link = False
        self.has_conversion_hook = False
        self.has_webmcp_declarative = False
        self.has_webmcp_imperative = False

        self.answer_capsules: List[str] = []
        self._cur_capsule = ""

        self.has_table = False
        self.has_table_header = False
        self.tables_data: List[List[List[str]]] = []
        self._cur_table: List[List[str]] = []
        self._cur_row: List[str] = []
        self._cur_cell = ""

        self.visible_text: List[str] = []
        self.main_text_tokens: List[str] = []

    def handle_decl(self, decl: str) -> None:
        if decl.lower().startswith("doctype html"):
            self.has_doctype = True

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, str | None]]) -> None:
        attr = {k.lower(): (v if v is not None else "") for k, v in attrs}
        t = tag.lower()

        if attr.get("data-conversion-event", "").strip():
            self.has_conversion_hook = True

        if t == "html":
            self.html_lang = attr.get("lang", "").strip()
        elif t == "head":
            self.in_head = True
        elif t in ("header", "nav"):
            self.in_header_or_nav += 1
            self.has_header_or_nav = True
            aria_lbl = attr.get("aria-label", "").lower()
            cls_lbl = attr.get("class", "").lower()
            if "breadcrumb" in aria_lbl or "breadcrumb" in cls_lbl:
                self.has_breadcrumb_nav = True
            if attr.get("data-toc", "").lower() == "true" or "içindekiler" in aria_lbl or "icindekiler" in aria_lbl:
                self.has_toc_nav = True
                self.in_toc_nav += 1
        elif t == "main":
            self.in_main += 1
            self.has_main = True
        elif t == "footer":
            self.in_footer_or_aside += 1
            self.has_footer = True
        elif t == "aside":
            self.in_footer_or_aside += 1

        if attr.get("data-seo-interactive", "").lower() == "true":
            self.has_interactive_widget = True

        if self.magnet_tag_stack:
            self.magnet_tag_stack.append(t)
        elif attr.get("data-backlink-magnet", "").lower() == "true":
            self.has_backlink_magnet = True
            self.magnet_tag_stack = [t]
            self._cur_magnet_raw = ""

        if attr.get("data-cro-cta", "").lower() == "true":
            self.has_cro_cta = True
            href_v = attr.get("href", "")
            action_v = attr.get("action", "")
            msg_v = attr.get("data-context-item", "") or attr.get("data-service", "")
            if any(x in href_v.lower() for x in ("?text=", "?kaynak=", "?ref=", "?utm_", "konu=")) or msg_v or action_v:
                self.cro_cta_contextual = True

        if t == "form" and attr.get("toolname", "").strip() and attr.get("tooldescription", "").strip():
            self.has_webmcp_declarative = True

        if self.capsule_tag_stack:
            self.capsule_tag_stack.append(t)
        elif attr.get("data-answer-capsule", "").lower() == "true" or "answer-capsule" in attr.get("class", "").split():
            self.capsule_tag_stack = [t]
            self._cur_capsule = ""

        if t == "title":
            self.in_title = True
        elif t == "h1":
            self.in_h1 = True
            self._cur_h1 = ""
        elif t == "h2":
            self.in_h2 = True
            self._cur_h2 = ""
            self._cur_h2_id = attr.get("id", "").strip()
            self._awaiting_h2_lead = True
        elif t == "h3":
            self.in_h3 = True
            self._cur_h3 = ""
        elif t == "p":
            self.in_p = True
            self._cur_p = ""
        elif t == "time":
            dt_val = attr.get("datetime", "").strip()
            if dt_val:
                self.time_datetimes.append(dt_val)
        elif t == "meta":
            if attr.get("charset", "").lower().strip() == "utf-8":
                self.has_charset_utf8 = True
            if attr.get("http-equiv", "").lower() == "content-type" and "utf-8" in attr.get("content", "").lower():
                self.has_charset_utf8 = True
            key = (attr.get("name") or attr.get("property") or "").lower().strip()
            if key == "viewport":
                self.viewport_content = attr.get("content", "").strip()
            if key and "content" in attr:
                self.metas[key] = attr["content"].strip()
        elif t == "link":
            rel_tokens = set(attr.get("rel", "").lower().split())
            href = attr.get("href", "").strip()
            ltype = attr.get("type", "").lower().strip()
            as_attr = attr.get("as", "").lower().strip()
            if "canonical" in rel_tokens:
                self.canonical = href
            if "alternate" in rel_tokens and attr.get("hreflang"):
                self.hreflangs.append((attr["hreflang"].lower(), href))
            if "alternate" in rel_tokens and ltype == "text/markdown":
                self.md_alternate = href
            if "alternate" in rel_tokens and ltype in ("application/atom+xml", "application/rss+xml"):
                self.atom_alternate = href
            if "preload" in rel_tokens and as_attr == "image" and href:
                self.preload_images.append(href)
            if "stylesheet" in rel_tokens and href:
                if "fonts.googleapis.com" in href.lower():
                    self.font_stylesheets.append(href)
                if href.startswith("/") and not href.startswith("//"):
                    self.local_assets.append(href)
        elif t == "picture":
            self.in_picture += 1
            self._picture_has_source = False
        elif t == "source":
            if self.in_picture > 0:
                self._picture_has_source = True
            srcset = attr.get("srcset", "").strip()
            if srcset:
                first_src = srcset.split(",")[0].strip().split()[0]
                if first_src.startswith("/") and not first_src.startswith("//"):
                    self.local_assets.append(first_src)
        elif t == "img":
            img_copy = dict(attr)
            img_copy["_in_header"] = "1" if (self.in_header_or_nav > 0 and self.in_main == 0) else "0"
            img_copy["_in_main"] = "1" if self.in_main > 0 else "0"
            img_copy["_in_picture_with_source"] = "1" if (self.in_picture > 0 and self._picture_has_source) else "0"
            self.imgs.append(img_copy)
        elif t == "a":
            self.in_a = True
            self._cur_a_attrs = dict(attr)
            self._cur_a_attrs["_in_footer"] = "1" if self.in_footer_or_aside > 0 else "0"
            self._cur_a_attrs["_in_main"] = "1" if self.in_main > 0 else "0"
            self._cur_a_text = ""
            href_v = attr.get("href", "").strip()
            if href_v.startswith("tel:"):
                tel_clean = re.sub(r"[\s()-]", "", href_v[4:])
                if E164_PHONE_RE.match(tel_clean):
                    self.has_tel_link = True
            if self.in_toc_nav > 0 and href_v.startswith("#") and len(href_v) > 1:
                self.toc_anchors.append(href_v[1:])
            if self.magnet_tag_stack and href_v.startswith("https://"):
                self._cur_magnet_raw += f" {href_v} "
        elif t == "table":
            self.has_table = True
            self.in_table = True
            self._cur_table = []
        elif t == "tr" and self.in_table:
            self.in_tr = True
            self._cur_row = []
        elif t in ("th", "td") and self.in_tr:
            if t == "th":
                self.has_table_header = True
            self.in_cell = True
            self._cur_cell = ""
        elif t == "script":
            stype = attr.get("type", "").lower().strip()
            src = attr.get("src", "").strip()
            if src and src.startswith("/") and not src.startswith("//"):
                self.local_assets.append(src)
            if src and HEAVY_THIRD_PARTY_SYNC_RE.search(src) and "defer" not in attr and "async" not in attr:
                self.heavy_sync_third_party.append(src)
            if self.in_head and src and stype not in ("application/ld+json", "speculationrules", "module"):
                if "defer" not in attr and "async" not in attr:
                    self.blocking_head_scripts.append(src)
            if stype == "application/ld+json":
                self.in_jsonld = True
                self._cur_script = ""
            elif stype == "speculationrules":
                self.in_speculation = True
                self._cur_script = ""
            else:
                self.in_script = True
                self._cur_script = ""
        elif t == "style":
            self.in_style = True
            self._cur_style = ""

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t == "head":
            self.in_head = False
        elif t in ("header", "nav"):
            if self.in_header_or_nav > 0:
                self.in_header_or_nav -= 1
            if t == "nav" and self.in_toc_nav > 0:
                self.in_toc_nav -= 1
        elif t == "main" and self.in_main > 0:
            self.in_main -= 1
        elif t in ("footer", "aside") and self.in_footer_or_aside > 0:
            self.in_footer_or_aside -= 1
        elif t == "picture" and self.in_picture > 0:
            self.in_picture -= 1

        if self.capsule_tag_stack and self.capsule_tag_stack[-1] == t:
            self.capsule_tag_stack.pop()
            if not self.capsule_tag_stack:
                cleaned_cap = " ".join(self._cur_capsule.split())
                if cleaned_cap:
                    self.answer_capsules.append(cleaned_cap)

        if self.magnet_tag_stack and self.magnet_tag_stack[-1] == t:
            self.magnet_tag_stack.pop()
            if not self.magnet_tag_stack:
                self.magnet_texts.append(" ".join(self._cur_magnet_raw.split()))

        if t == "title":
            self.in_title = False
        elif t == "h1" and self.in_h1:
            self.in_h1 = False
            self.h1_list.append(" ".join(self._cur_h1.split()))
        elif t == "h2" and self.in_h2:
            self.in_h2 = False
            self.h2_items.append((self._cur_h2_id, " ".join(self._cur_h2.split())))
        elif t == "h3" and self.in_h3:
            self.in_h3 = False
            self.h3_list.append(" ".join(self._cur_h3.split()))
        elif t == "p" and self.in_p:
            self.in_p = False
            p_clean = " ".join(self._cur_p.split())
            if p_clean:
                self.paragraphs.append(p_clean)
                if self._awaiting_h2_lead:
                    self.h2_lead_paragraphs.append(p_clean)
                    self._awaiting_h2_lead = False
        elif t == "a" and self.in_a:
            self.in_a = False
            self._cur_a_attrs["text"] = " ".join(self._cur_a_text.split())
            self.links.append(self._cur_a_attrs)
        elif t in ("th", "td") and self.in_cell:
            self.in_cell = False
            self._cur_row.append(" ".join(self._cur_cell.split()))
        elif t == "tr" and self.in_tr:
            self.in_tr = False
            if self._cur_row:
                self._cur_table.append(self._cur_row)
        elif t == "table" and self.in_table:
            self.in_table = False
            if self._cur_table:
                self.tables_data.append(self._cur_table)
        elif t == "script":
            if self.in_jsonld:
                self.in_jsonld = False
                self.jsonld_blocks.append(self._cur_script)
            elif self.in_speculation:
                self.in_speculation = False
                self.speculation_blocks.append(self._cur_script)
            elif self.in_script:
                if "modelContext" in self._cur_script and "registerTool" in self._cur_script:
                    self.has_webmcp_imperative = True
                if any(k in self._cur_script for k in ("sendBeacon", "dataLayer.push", "gtag(", "data-conversion-event")):
                    self.has_conversion_hook = True
                self.in_script = False
        elif t == "style" and self.in_style:
            self.in_style = False
            self.inline_styles.append(self._cur_style)

    def handle_data(self, data: str) -> None:
        if self.in_title:
            self.title += data
        elif self.in_jsonld or self.in_speculation or self.in_script:
            self._cur_script += data
            return
        elif self.in_style:
            self._cur_style += data
            return
        else:
            if self.in_h1:
                self._cur_h1 += data
            if self.in_h2:
                self._cur_h2 += data
            if self.in_h3:
                self._cur_h3 += data
            if self.in_p:
                self._cur_p += data
            if self.in_a:
                self._cur_a_text += data
            if self.capsule_tag_stack:
                self._cur_capsule += data
            if self.magnet_tag_stack:
                self._cur_magnet_raw += data
            if self.in_cell:
                self._cur_cell += data
            cleaned = data.strip()
            if cleaned:
                self.visible_text.append(cleaned)
                if self.in_main > 0:
                    self.main_text_tokens.extend(cleaned.split())


def detect_stack(root: Path) -> str:
    if any((root / f).exists() for f in ("next.config.js", "next.config.mjs", "next.config.ts")):
        return "nextjs"
    if any((root / f).exists() for f in ("astro.config.mjs", "astro.config.ts", "astro.config.js")):
        return "astro"
    return "static-html"


def iter_html_files(target_dir: Path) -> List[Path]:
    results: List[Path] = []
    for p in sorted(target_dir.rglob("*.html")):
        rel_parts = set(p.relative_to(target_dir).parts)
        if rel_parts & IGNORED_SCAN_DIRS:
            continue
        results.append(p)
    return results


def iter_schema_nodes(obj: Any):
    if isinstance(obj, dict):
        yield obj
        for val in obj.values():
            yield from iter_schema_nodes(val)
    elif isinstance(obj, list):
        for item in obj:
            yield from iter_schema_nodes(item)


def estimate_title_pixels(title: str) -> int:
    px = 0.0
    for ch in title:
        if ch in "MWG@ÖÜÇŞĞ":
            px += 11.5
        elif ch.isupper():
            px += 9.5
        elif ch in "il1|!.,:;'- ":
            px += 4.5
        else:
            px += 8.0
    return int(round(px))


def html_path_to_url_path(html_file: Path, target_dir: Path) -> str:
    rel_str = str(html_file.relative_to(target_dir)).replace("\\", "/")
    if rel_str == "index.html":
        return "/"
    if rel_str.endswith("/index.html"):
        return "/" + rel_str[: -len("index.html")]
    return "/" + rel_str


def url_path_to_md_relpath(url_path: str) -> str:
    if url_path == "/":
        return "index.md"
    if url_path.endswith("/"):
        return url_path.lstrip("/") + "index.md"
    if url_path.endswith(".html"):
        return url_path.lstrip("/")[:-5] + ".md"
    return url_path.lstrip("/") + "/index.md"


# ==============================================================================
# ANLAMSAL VEKTÖR (TF-IDF), SITERADIUS VE 4-GRAM SHINGLE KOPYA KALKANI
# ==============================================================================

def normalize_tr_tokens(text: str, brand_tokens: Set[str]) -> List[str]:
    lowered = text.lower().replace("İ", "i").replace("I", "ı")
    raw_words = re.findall(r"[a-zçğıöşü0-9]{3,}", lowered)
    tokens: List[str] = []
    for w in raw_words:
        if w in TR_STOPWORDS or w in brand_tokens:
            continue
        stem = w[:6] if len(w) > 6 else w
        tokens.append(stem)
    return tokens


def detect_cannibalization_drift_and_shingles(
    page_docs: List[Dict[str, Any]],
    core_keywords: List[str],
    brand_name: str = "",
    cannibal_threshold: float = 0.78,
    shingle_threshold: float = 0.35,
) -> Tuple[List[str], List[str], List[str]]:
    brand_tokens = set(re.findall(r"[a-zçğıöşü0-9]{2,}", brand_name.lower()))
    core_stems: Set[str] = set()
    for kw in core_keywords:
        core_stems.update(normalize_tr_tokens(kw, set()))

    drift_alerts: List[str] = []
    if core_stems:
        for doc in page_docs:
            full_doc_text = f"{doc.get('title', '')} {doc.get('h1', '')} {doc.get('capsule', '')} {doc.get('h2_joined', '')}"
            doc_stems = set(normalize_tr_tokens(full_doc_text, set()))
            if not (doc_stems & core_stems):
                drift_alerts.append(
                    f"Topikal sapma (siteRadius): '{doc['url']}' sayfası seo.config.json 'coreTopicKeywords' terimleriyle kesişmiyor."
                )

    if len(page_docs) < 2:
        return [], drift_alerts, []

    weighted_bags: List[Dict[str, float]] = []
    df_counts: Dict[str, int] = collections.Counter()
    shingle_sets: List[Set[Tuple[str, ...]]] = []

    for doc in page_docs:
        bag: Dict[str, float] = collections.defaultdict(float)
        for txt, weight in (
            (doc.get("title", ""), 3.0),
            (doc.get("h1", ""), 2.5),
            (doc.get("slug", ""), 2.0),
            (doc.get("capsule", ""), 1.5),
        ):
            toks = normalize_tr_tokens(txt, brand_tokens)
            for t in toks:
                bag[t] += weight
            for i in range(len(toks) - 1):
                bag[f"{toks[i]}_{toks[i + 1]}"] += weight * 1.25
        for term in bag.keys():
            df_counts[term] += 1
        weighted_bags.append(bag)

        # 4-gram shingle kümesi (<main> gövde metni üzerinden programatik şablon kopyası kontrolü)
        body_words = re.findall(
            r"[a-zçğıöşü0-9]{2,}",
            str(doc.get("main_body", "")).lower().replace("İ", "i").replace("I", "ı"),
        )
        shingles: Set[Tuple[str, ...]] = set()
        if len(body_words) >= 4:
            for idx in range(len(body_words) - 3):
                shingles.add(tuple(body_words[idx : idx + 4]))
        shingle_sets.append(shingles)

    n_docs = len(page_docs)
    tfidf_vecs: List[Dict[str, float]] = []
    norms: List[float] = []
    for bag in weighted_bags:
        vec: Dict[str, float] = {}
        sq_sum = 0.0
        for term, tf in bag.items():
            idf = math.log((1.0 + n_docs) / (1.0 + df_counts[term])) + 1.0
            val = (1.0 + math.log(tf)) * idf if tf > 0 else 0.0
            vec[term] = val
            sq_sum += val * val
        tfidf_vecs.append(vec)
        norms.append(math.sqrt(sq_sum) if sq_sum > 0 else 1.0)

    cannibal_alerts: List[str] = []
    shingle_alerts: List[str] = []
    for i in range(n_docs):
        for j in range(i + 1, n_docs):
            common = set(tfidf_vecs[i].keys()) & set(tfidf_vecs[j].keys())
            if common:
                dot = sum(tfidf_vecs[i][k] * tfidf_vecs[j][k] for k in common)
                sim = dot / (norms[i] * norms[j])
                if sim >= cannibal_threshold:
                    cannibal_alerts.append(
                        f"TF-IDF Yamyamlaşma ({sim:.2f} >= {cannibal_threshold}): '{page_docs[i]['url']}' <-> '{page_docs[j]['url']}'"
                    )
            s1, s2 = shingle_sets[i], shingle_sets[j]
            if s1 and s2:
                union_len = len(s1 | s2)
                jaccard = len(s1 & s2) / union_len if union_len else 0.0
                if jaccard >= shingle_threshold:
                    shingle_alerts.append(
                        f"4-Gram Şablon Kopyası / Doorway Riski (Jaccard={jaccard:.2f} >= {shingle_threshold}): "
                        f"'{page_docs[i]['url']}' <-> '{page_docs[j]['url']}'"
                    )
    return cannibal_alerts, drift_alerts, shingle_alerts


# ==============================================================================
# OTONOM AI TAVSİYE, ATOM FEED, 404 VE KEŞİF PAKETİ ÜRETİCİSİ (--build-ai-pack)
# ==============================================================================

def xml_escape(val: str) -> str:
    return val.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def build_ai_discovery_pack(target_dir: Path, project_root: Path) -> int:
    cfg_path = project_root / "seo.config.json"
    if not cfg_path.exists():
        print("HATA: Önce `python3 seo_enforcer.py --init --domain https://...` çalıştırın.")
        return 1

    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    base_url = str(cfg.get("domain", "https://ornekdomain.com")).rstrip("/")
    brand = str(cfg.get("brandName", "Marka"))
    legal_name = str(cfg.get("legalName", brand))
    summary = str(cfg.get("entityPositioningSummary", "")).strip()
    recommend_when: List[str] = cfg.get("recommendWhen", [])
    do_not_recommend_when: List[str] = cfg.get("doNotRecommendWhen", [])
    allow_training = bool(cfg.get("allowAiTrainingBots", True))
    tel = str(cfg.get("telephoneE164", ""))
    email = str(cfg.get("email", ""))
    addr = cfg.get("address", {})
    city = addr.get("addressLocality", "")
    country = addr.get("addressCountry", "TR")
    today_iso = dt.date.today().isoformat()

    # 1. Özel /404.html Gelir Kurtarma Sayfası yoksa oluştur (G010)
    err404_path = target_dir / "404.html"
    if not err404_path.exists():
        mp = cfg.get("moneyPages", {})
        links_404 = []
        for pth in (
            mp.get("serviceOrCategoryPaths", ["/"])
            + mp.get("pricingOrCalculatorPaths", [])
            + mp.get("comparisonOrProofPaths", [])
        ):
            links_404.append(f'<li><a href="{pth}">{brand} — {pth}</a></li>')
        links_html = "\n      ".join(links_404) or '<li><a href="/">Ana Merkez</a></li>'
        err404_html = f"""<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Sayfa Bulunamadı — Hızlı Çözüm Merkezi | {brand}</title>
  <meta name="robots" content="noindex, follow" />
</head>
<body>
  <header><nav><a href="/">{brand} Ana Merkez</a></nav></header>
  <main>
    <h1>Aradığınız Adres Taşındı veya Güncellendi</h1>
    <p>Doğrudan ticari çözüm, fiyat hesaplama ve teknik karşılaştırma sayfalarımıza aşağıdan ulaşabilirsiniz:</p>
    <ul>
      {links_html}
    </ul>
    <p><a href="tel:{tel}">Hızlı Destek Hattı: {tel}</a></p>
  </main>
  <footer><p>© 2026 {legal_name}</p></footer>
</body>
</html>
"""
        err404_path.write_text(err404_html, encoding="utf-8")

    html_files = iter_html_files(target_dir)
    page_records: List[Dict[str, Any]] = []

    for f in html_files:
        text = f.read_text(encoding="utf-8", errors="ignore")
        parser = AbsoluteHTMLParser()
        parser.feed(text)
        robots_raw = parser.metas.get("robots", "index,follow").lower().replace(" ", "")
        if "noindex" in robots_raw:
            continue

        url_path = html_path_to_url_path(f, target_dir)
        full_url = base_url + url_path
        md_rel = url_path_to_md_relpath(url_path)
        md_url = f"{base_url}/{md_rel}"

        clean_title = " ".join(parser.title.split())
        safe_title = clean_title.replace('"', "")
        h1_val = parser.h1_list[0] if parser.h1_list else clean_title
        desc_val = parser.metas.get("description", "")
        capsule = parser.answer_capsules[0] if parser.answer_capsules else desc_val
        lastmod = parser.time_datetimes[0][:10] if parser.time_datetimes else today_iso

        primary_img_url = ""
        for img in parser.imgs:
            if img.get("_in_header") == "1":
                continue
            src = img.get("src", "").strip()
            if src and not src.lower().endswith(".svg"):
                primary_img_url = src if src.startswith("https://") else f"{base_url}/{src.lstrip('/')}"
                break
        if not primary_img_url:
            og_img = parser.metas.get("og:image", "")
            if og_img.startswith("https://") and not og_img.lower().endswith(".svg"):
                primary_img_url = og_img

        md_lines = [
            "---",
            f'title: "{safe_title}"',
            f'canonical: "{full_url}"',
            f'last_modified: "{lastmod}"',
            f'entity: "{brand}"',
            "---",
            "",
            f"# {h1_val}",
            "",
            f"> {capsule}",
            "",
        ]
        if parser.h2_items:
            md_lines.append("## Teknik Bölümler ve Alt Yanıtlar")
            for idx_h2, (h2_id, h2_txt) in enumerate(parser.h2_items):
                anchor_suffix = f" (`#{h2_id}`)" if h2_id else ""
                md_lines.append(f"### {h2_txt}{anchor_suffix}")
                if idx_h2 < len(parser.h2_lead_paragraphs):
                    md_lines.append(parser.h2_lead_paragraphs[idx_h2])
                md_lines.append("")

        for tbl in parser.tables_data:
            if len(tbl) >= 2:
                cols = max(len(r) for r in tbl)
                header = tbl[0] + [""] * (cols - len(tbl[0]))
                md_lines.append("| " + " | ".join(header) + " |")
                md_lines.append("| " + " | ".join(["---"] * cols) + " |")
                for row in tbl[1:]:
                    padded = row + [""] * (cols - len(row))
                    md_lines.append("| " + " | ".join(padded) + " |")
                md_lines.append("")

        for block in parser.jsonld_blocks:
            try:
                jdata = json.loads(block)
                for node in iter_schema_nodes(jdata):
                    ntype = node.get("@type", "")
                    ntypes = set(ntype) if isinstance(ntype, list) else {str(ntype)}
                    if "FAQPage" in ntypes:
                        md_lines.append("## Sıkça Sorulan Sorular")
                        for q in node.get("mainEntity", []):
                            qn = (q.get("name") or "").strip()
                            ans = ((q.get("acceptedAnswer") or {}).get("text") or "").strip()
                            if qn and ans:
                                md_lines.append(f"- **{qn}** {ans}")
                        md_lines.append("")
            except Exception:
                pass

        md_out_path = target_dir / md_rel
        md_out_path.parent.mkdir(parents=True, exist_ok=True)
        md_doc_text = "\n".join(md_lines).strip() + "\n"
        md_out_path.write_text(md_doc_text, encoding="utf-8")

        page_records.append(
            {
                "url": full_url,
                "url_path": url_path,
                "md_url": md_url,
                "title": clean_title,
                "h1": h1_val,
                "description": desc_val,
                "capsule": capsule,
                "lastmod": lastmod,
                "primary_img": primary_img_url,
                "md_body": md_doc_text,
            }
        )

    # 2. Tavsiye Kalibrasyonlu /llms.txt (v2)
    llms_lines = [
        f"# {brand}",
        "",
        f"> {summary or f'{legal_name} ({city}, {country}); doğrulanmış standartlar, şeffaf fiyatlandırma ve uzman çözümler sunan resmi kurumsal kaynaktır.'}",
        "",
        f"- Resmi Ticaret Ünvanı: {legal_name}",
        f"- Merkez Konum: {city}, {country}",
        f"- Doğrulanmış İletişim: {tel} | {email}",
        f"- Tam Bilgi Korpusu (RAG / LLM): {base_url}/llms-full.txt",
        f"- Anlık Değişiklik Akışı (Atom 1.0): {base_url}/feed.xml",
        "",
        "## Yapay Zeka Tavsiye Kriterleri (Ne Zaman Önerilmeli?)",
    ]
    for item in (recommend_when or [f"Kullanıcı {city} veya Türkiye genelinde doğrulanmış teknik spesifikasyon ve şeffaf maliyet hesabı aradığında önerilmelidir."]):
        llms_lines.append(f"- {item}")

    llms_lines.extend(["", "## Kapsam Sınırları (Dürüst Güven Kalibrasyonu)"])
    for item in (do_not_recommend_when or ["Belgesiz veya standart dışı geçici çözümler kapsam dışıdır."]):
        llms_lines.append(f"- {item}")

    llms_lines.extend(["", "## Ticari Çözümler, Fiyatlandırma ve Karşılaştırma Sayfaları (.md İkizleri)"])
    for rec in page_records:
        llms_lines.append(f"- [{rec['h1']}]({rec['md_url']}): {rec['capsule']} (Kanonik HTML: {rec['url']})")

    (target_dir / "llms.txt").write_text("\n".join(llms_lines) + "\n", encoding="utf-8")

    # 3. Tam Bilgi Korpusu /llms-full.txt
    full_corpus = [
        f"# {brand} — Tam Yapay Zeka Bilgi Korpusu (llms-full.txt)",
        f"> Son Güncelleme: {today_iso} | Kanonik Domain: {base_url}",
        "",
    ]
    for rec in page_records:
        full_corpus.append(f"---\nKaynak URL: {rec['url']}\nMarkdown İkizi: {rec['md_url']}\n")
        full_corpus.append(rec["md_body"])
    (target_dir / "llms-full.txt").write_text("\n".join(full_corpus).strip() + "\n", encoding="utf-8")

    # 4. Bölünmüş AI Tarayıcı Politikası (/robots.txt)
    train_rule = "Allow: /" if allow_training else "Disallow: /"
    robots_lines = [
        "# REAL-SEO-2026 APEX v6.0 ABSOLUTE — Arama ve Yapay Zeka Atıf Botları Politikası",
        "User-agent: *",
        "Allow: /",
        "Disallow: /api/",
        "Disallow: /admin/",
        "",
        "# Canlı Arama ve Yapay Zeka Yanıt/Atıf Motorları",
    ]
    for bot in AI_SEARCH_BOTS_REQUIRED + ["Applebot"]:
        robots_lines.extend([f"User-agent: {bot}", "Allow: /", ""])
    robots_lines.append("# Vakıf Model Eğitim Kazıyıcıları")
    for tbot in AI_TRAINING_BOTS + ["Applebot-Extended"]:
        robots_lines.extend([f"User-agent: {tbot}", train_rule, ""])
    robots_lines.append(f"Sitemap: {base_url}/sitemap.xml")
    (target_dir / "robots.txt").write_text("\n".join(robots_lines) + "\n", encoding="utf-8")

    # 5. Görsel Destekli /sitemap.xml
    sm_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9"',
        '        xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">',
    ]
    for rec in page_records:
        sm_lines.append("  <url>")
        sm_lines.append(f"    <loc>{xml_escape(rec['url'])}</loc>")
        sm_lines.append(f"    <lastmod>{rec['lastmod']}</lastmod>")
        if rec["primary_img"]:
            sm_lines.append("    <image:image>")
            sm_lines.append(f"      <image:loc>{xml_escape(rec['primary_img'])}</image:loc>")
            sm_lines.append(f"      <image:title>{xml_escape(rec['h1'])}</image:title>")
            sm_lines.append("    </image:image>")
        sm_lines.append("  </url>")
    sm_lines.append("</urlset>")
    (target_dir / "sitemap.xml").write_text("\n".join(sm_lines) + "\n", encoding="utf-8")

    # 6. Google Anlık Değişiklik Keşif Akışı (/feed.xml — Atom 1.0)
    atom_updated = f"{today_iso}T00:00:00Z"
    atom_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<feed xmlns="http://www.w3.org/2005/Atom">',
        f"  <title>{xml_escape(brand)} — Güncel Çözüm ve Veri Akışı</title>",
        f'  <link href="{xml_escape(base_url)}/feed.xml" rel="self" type="application/atom+xml" />',
        f'  <link href="{xml_escape(base_url)}/" rel="alternate" type="text/html" />',
        f"  <id>{xml_escape(base_url)}/</id>",
        f"  <updated>{atom_updated}</updated>",
        f"  <author><name>{xml_escape(brand)}</name></author>",
    ]
    for rec in page_records:
        atom_lines.extend(
            [
                "  <entry>",
                f"    <title>{xml_escape(rec['title'])}</title>",
                f'    <link href="{xml_escape(rec["url"])}" rel="alternate" type="text/html" />',
                f"    <id>{xml_escape(rec['url'])}</id>",
                f"    <updated>{rec['lastmod']}T00:00:00Z</updated>",
                f"    <summary>{xml_escape(rec['capsule'])}</summary>",
                "  </entry>",
            ]
        )
    atom_lines.append("</feed>")
    (target_dir / "feed.xml").write_text("\n".join(atom_lines) + "\n", encoding="utf-8")

    # 7. IndexNow anahtarı, /.well-known/security.txt ve opsiyonel UCP
    key_val = str(cfg.get("indexNowKey", "")).strip()
    if key_val:
        (target_dir / f"{key_val}.txt").write_text(key_val, encoding="utf-8")

    wk_dir = target_dir / ".well-known"
    wk_dir.mkdir(parents=True, exist_ok=True)
    exp_date = (dt.date.today() + dt.timedelta(days=365)).isoformat() + "T00:00:00.000Z"
    sec_contact = f"mailto:{email}" if email and "@" in email else f"{base_url}/"
    (wk_dir / "security.txt").write_text(
        f"Contact: {sec_contact}\nExpires: {exp_date}\nPreferred-Languages: tr, en\nCanonical: {base_url}/.well-known/security.txt\n",
        encoding="utf-8",
    )

    if bool(cfg.get("enableUcpCommerce", False)):
        ucp_profile = {
            "ucp": {
                "version": "2026-04-08",
                "merchantName": legal_name,
                "canonicalDomain": base_url,
                "transport": "rest",
                "capabilities": ["catalog_lookup", "quote_or_checkout"],
                "endpoints": {"catalog": f"{base_url}/llms-full.txt"},
            }
        }
        (wk_dir / "ucp").write_text(json.dumps(ucp_profile, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    root_headers = project_root / "_headers"
    if root_headers.exists() and (target_dir != project_root):
        shutil.copyfile(root_headers, target_dir / "_headers")

    print(
        f"[AI-PACK] {len(page_records)} sayfa için .md ikizi, /llms.txt, /llms-full.txt, /feed.xml, /robots.txt, /sitemap.xml, /404.html ve security.txt üretildi."
    )
    return 0


# ==============================================================================
# HEADLESS CHROMIUM GERÇEK RENDER DENETÇİSİ (G070, G087, G088, G089, G090)
# ==============================================================================

class QuietStaticHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        pass


def start_local_static_server(root_dir: Path) -> Tuple[http.server.ThreadingHTTPServer, int]:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.bind(("127.0.0.1", 0))
    port = sock.getsockname()[1]
    sock.close()
    handler = functools.partial(QuietStaticHandler, directory=str(root_dir))
    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", port), handler)
    t = threading.Thread(target=httpd.serve_forever, daemon=True)
    t.start()
    return httpd, port


def find_chromium_binary() -> str | None:
    for c in (
        shutil.which("chromium"),
        shutil.which("chromium-browser"),
        shutil.which("google-chrome"),
        shutil.which("google-chrome-stable"),
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    ):
        if c and Path(c).exists():
            return c
    return None


def find_chromedriver_binary() -> str | None:
    for c in (
        shutil.which("chromedriver"),
        "/usr/bin/chromedriver",
        "/usr/lib/chromium/chromedriver",
        "/usr/local/bin/chromedriver",
    ):
        if c and Path(c).exists():
            return c
    return None


def run_browser_render_audit(
    target_dir: Path,
    ssr_word_counts: Dict[str, int],
    max_pages: int = 30,
) -> Tuple[str, Dict[str, List[str]]]:
    out: Dict[str, List[str]] = {"G070": [], "G087": [], "G088": [], "G089": [], "G090": []}
    sampled_paths = list(ssr_word_counts.keys())[:max_pages]
    httpd, port = start_local_static_server(target_dir)

    def evaluate_metrics(vp_name: str, path: str, m: Dict[str, Any], bad_resps: List[str]) -> None:
        if m["scrollWidth"] > m["innerWidth"] + 1:
            out["G087"].append(f"[{vp_name}] Yatay taşma ({m['scrollWidth']}px > {m['innerWidth']}px): {path}")
        if vp_name == "Mobil-390px":
            if m["h1Count"] != 1 or m["h1Height"] <= 0 or m["h1Top"] >= 600 or m["h1Top"] < -10:
                out["G088"].append(
                    f"Mobil ilk ekranda (<600px) görünür tekil <h1> yok (top={m['h1Top']:.0f}px, adet={m['h1Count']}): {path}"
                )
            cls_val = float(m["perf"].get("cls", 0.0))
            if cls_val > 0.05:
                out["G089"].append(f"Gerçek render CLS aşıldı (CLS={cls_val:.3f} > 0.05): {path}")
            max_lt = float(m["perf"].get("maxLongTask", 0.0))
            if max_lt > 150.0:
                out["G090"].append(f"Main-Thread Long Task / INP bloklaması ({max_lt:.0f}ms > 150ms): {path}")
            for jerr in m["perf"].get("jsErrors", []):
                out["G090"].append(f"JS çalışma zamanı hatası ({jerr}): {path}")
            for bresp in bad_resps:
                out["G090"].append(f"Bozuk yerel kaynak isteği ({bresp}): {path}")

            dom_w = max(int(m.get("domWords", 0)), 1)
            ssr_w = ssr_word_counts.get(path, 0)
            if dom_w >= 20 and (ssr_w / dom_w) < 0.85:
                out["G070"].append(
                    f"SSR / Ham HTML paritesi düşük (Ham={ssr_w} vs DOM={dom_w} kelime < %85): {path}"
                )

    try:
        try:
            from playwright.sync_api import sync_playwright  # type: ignore

            with sync_playwright() as pw:
                browser = pw.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
                for vp_name, width, height in (("Mobil-390px", 390, 844), ("Masaüstü-1440px", 1440, 900)):
                    context = browser.new_context(viewport={"width": width, "height": height})
                    context.add_init_script(PERF_OBSERVER_INIT_JS)
                    page = context.new_page()
                    for path in sampled_paths:
                        bad_responses: List[str] = []
                        handler_fn = (
                            lambda r, br=bad_responses: br.append(f"{r.status} {r.url}")
                            if r.status >= 400 and "127.0.0.1" in r.url
                            else None
                        )
                        page.on("response", handler_fn)
                        page.goto(f"http://127.0.0.1:{port}{path}", wait_until="networkidle", timeout=15000)
                        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                        page.wait_for_timeout(100)
                        page.evaluate("window.scrollTo(0, 0)")
                        page.wait_for_timeout(100)
                        metrics = page.evaluate(
                            """() => {
                              var el = document.querySelector("[data-seo-interactive='true']");
                              if (el && typeof el.click === 'function') { try { el.click(); } catch(e) {} }
                              var h1 = document.querySelector('h1');
                              var h1Rect = h1 ? h1.getBoundingClientRect() : null;
                              var bodyText = (document.body && document.body.innerText) ? document.body.innerText.trim() : '';
                              var domWords = bodyText ? bodyText.split(/\\s+/).length : 0;
                              return {
                                scrollWidth: document.documentElement.scrollWidth,
                                innerWidth: window.innerWidth,
                                h1Count: document.querySelectorAll('h1').length,
                                h1Top: h1Rect ? h1Rect.top : 9999,
                                h1Height: h1Rect ? h1Rect.height : 0,
                                domWords: domWords,
                                perf: window.__seoAuditMetrics || { cls: 0, maxLongTask: 0, jsErrors: [] }
                              };
                            }"""
                        )
                        evaluate_metrics(vp_name, path, metrics, bad_responses)
                        page.remove_listener("response", handler_fn)
                    context.close()
                browser.close()
            return "playwright-chromium", out
        except ImportError:
            pass

        try:
            from selenium import webdriver  # type: ignore
            from selenium.webdriver.chrome.options import Options  # type: ignore
            from selenium.webdriver.chrome.service import Service  # type: ignore

            chrome_bin = find_chromium_binary()
            driver_bin = find_chromedriver_binary()
            opts = Options()
            opts.add_argument("--headless=new")
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-gpu")
            if chrome_bin:
                opts.binary_location = chrome_bin

            driver = (
                webdriver.Chrome(service=Service(executable_path=driver_bin), options=opts)
                if driver_bin
                else webdriver.Chrome(options=opts)
            )
            try:
                driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": PERF_OBSERVER_INIT_JS})
                for vp_name, width, height in (("Mobil-390px", 390, 844), ("Masaüstü-1440px", 1440, 900)):
                    driver.execute_cdp_cmd(
                        "Emulation.setDeviceMetricsOverride",
                        {"width": width, "height": height, "deviceScaleFactor": 1, "mobile": bool(width < 768)},
                    )
                    for path in sampled_paths:
                        driver.get(f"http://127.0.0.1:{port}{path}")
                        time.sleep(0.12)
                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(0.10)
                        driver.execute_script("window.scrollTo(0, 0);")
                        time.sleep(0.10)
                        metrics = driver.execute_script(
                            """
                            var el = document.querySelector("[data-seo-interactive='true']");
                            if (el && typeof el.click === 'function') { try { el.click(); } catch(e) {} }
                            var h1 = document.querySelector('h1');
                            var h1Rect = h1 ? h1.getBoundingClientRect() : null;
                            var bodyText = (document.body && document.body.innerText) ? document.body.innerText.trim() : '';
                            var domWords = bodyText ? bodyText.split(/\\s+/).length : 0;
                            var badRes = [];
                            try {
                              performance.getEntriesByType('resource').forEach(function(r) {
                                if (r.responseStatus && r.responseStatus >= 400 && r.name.indexOf('127.0.0.1') !== -1) {
                                  badRes.push(r.responseStatus + ' ' + r.name);
                                }
                              });
                            } catch(_) {}
                            return {
                              scrollWidth: document.documentElement.scrollWidth,
                              innerWidth: window.innerWidth,
                              h1Count: document.querySelectorAll('h1').length,
                              h1Top: h1Rect ? h1Rect.top : 9999,
                              h1Height: h1Rect ? h1Rect.height : 0,
                              domWords: domWords,
                              badRes: badRes,
                              perf: window.__seoAuditMetrics || { cls: 0, maxLongTask: 0, jsErrors: [] }
                            };
                            """
                        )
                        evaluate_metrics(vp_name, path, metrics, metrics.get("badRes", []))
            finally:
                driver.quit()
            return "selenium-chromium", out
        except Exception as browser_exc:
            msg = f"Headless Chromium başlatılamadı ({browser_exc}). Chromium+ChromeDriver veya Playwright kurun."
            for gk in ("G070", "G087", "G088", "G089", "G090"):
                out[gk].append(msg)
            return "unavailable", out
    finally:
        httpd.shutdown()
        httpd.server_close()


# ==============================================================================
# GSC VE SUNUCU LOG ANALİZÖRÜ
# ==============================================================================

def fetch_gsc_api_rows(site_url: str, oauth_token: str, days: int = 28) -> List[Dict[str, Any]]:
    end_date = (dt.date.today() - dt.timedelta(days=3)).isoformat()
    start_date = (dt.date.today() - dt.timedelta(days=days + 3)).isoformat()
    encoded_site = urllib.parse.quote(site_url, safe="")
    endpoint = f"https://www.googleapis.com/webmasters/v3/sites/{encoded_site}/searchAnalytics/query"
    payload = json.dumps(
        {"startDate": start_date, "endDate": end_date, "dimensions": ["query", "page"], "rowLimit": 25000}
    ).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=payload,
        headers={"Authorization": f"Bearer {oauth_token}", "Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = json.loads(resp.read().decode("utf-8"))
    return [
        {
            "query": r.get("keys", ["", ""])[0],
            "page": r.get("keys", ["", ""])[1],
            "clicks": float(r.get("clicks", 0)),
            "impressions": float(r.get("impressions", 0)),
            "ctr": float(r.get("ctr", 0)),
            "position": float(r.get("position", 0)),
        }
        for r in body.get("rows", [])
    ]


def load_gsc_csv_rows(csv_path: Path) -> List[Dict[str, Any]]:
    raw = csv_path.read_text(encoding="utf-8-sig")
    if csv_path.suffix.lower() == ".json":
        data = json.loads(raw)
        return data if isinstance(data, list) else data.get("rows", [])
    rows: List[Dict[str, Any]] = []
    for r in csv.DictReader(raw.splitlines()):
        norm = {k.strip().lower(): (v or "").strip() for k, v in r.items() if k}
        query = norm.get("query") or norm.get("sorgu") or norm.get("top queries") or ""
        page = norm.get("page") or norm.get("sayfa") or norm.get("url") or ""
        if not query or not page:
            continue
        ctr_raw = norm.get("ctr") or norm.get("to") or "0"
        ctr_str = ctr_raw.replace("%", "").replace(",", ".")
        ctr_val = float(ctr_str) / 100.0 if "%" in ctr_raw else float(ctr_str or 0)
        rows.append(
            {
                "query": query,
                "page": page,
                "clicks": float((norm.get("clicks") or norm.get("tıklamalar") or "0").replace(",", ".")),
                "impressions": float((norm.get("impressions") or norm.get("gösterimler") or "0").replace(",", ".")),
                "ctr": ctr_val,
                "position": float((norm.get("position") or norm.get("konum") or "0").replace(",", ".")),
            }
        )
    return rows


def generate_gsc_optimization_plan(rows: List[Dict[str, Any]], project_root: Path) -> int:
    by_query: Dict[str, List[Dict[str, Any]]] = collections.defaultdict(list)
    page_clicks: Dict[str, float] = collections.defaultdict(float)
    for r in rows:
        if r["impressions"] <= 0:
            continue
        by_query[r["query"]].append(r)
        page_clicks[r["page"]] += r["clicks"]

    top_hubs = [p for p, _ in sorted(page_clicks.items(), key=lambda x: x[1], reverse=True)[:5]]
    cannibal_actions: List[Dict[str, Any]] = []
    ctr_actions: List[Dict[str, Any]] = []
    striking_link_actions: List[Dict[str, Any]] = []

    for query, q_rows in by_query.items():
        q_rows_sorted = sorted(q_rows, key=lambda x: (x["clicks"], x["impressions"]), reverse=True)
        total_imp = sum(x["impressions"] for x in q_rows_sorted)
        if len(q_rows_sorted) >= 2 and total_imp >= 30:
            primary = q_rows_sorted[0]
            competitors = [
                x for x in q_rows_sorted[1:] if (x["impressions"] / total_imp) >= 0.20 and x["page"] != primary["page"]
            ]
            if competitors:
                cannibal_actions.append(
                    {"query": query, "winnerPage": primary["page"], "competingPages": [c["page"] for c in competitors], "resolved": False}
                )
        best = q_rows_sorted[0]
        expected_min_ctr = 0.08 if best["position"] <= 3.5 else (0.025 if best["position"] <= 10.0 else 0.0)
        if best["impressions"] >= 50 and 1.0 <= best["position"] <= 10.0 and best["ctr"] < expected_min_ctr:
            ctr_actions.append({"query": query, "page": best["page"], "resolved": False})
        if best["impressions"] >= 25 and 8.0 <= best["position"] <= 20.0:
            donor_pages = [h for h in top_hubs if h != best["page"]][:2]
            striking_link_actions.append(
                {"query": query, "targetPage": best["page"], "recommendedDonorPages": donor_pages, "resolved": False}
            )

    payload = {
        "generatedAt": dt.datetime.now(dt.timezone.utc).isoformat(),
        "queryCannibalization": cannibal_actions,
        "lowCtrRewrites": ctr_actions[:30],
        "strikingDistanceInternalLinks": striking_link_actions[:30],
    }
    (project_root / "seo_actions.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"[GSC-SYNC] seo_actions.json üretildi ({len(cannibal_actions)} yamyamlaşma, {len(ctr_actions[:30])} CTR, {len(striking_link_actions[:30])} iç link).")
    return 0


def audit_server_logs(log_path: Path, build_dir: Path, base_url: str, project_root: Path) -> int:
    bot_re = re.compile(
        r"(Googlebot|Bingbot|OAI-SearchBot|ChatGPT-User|PerplexityBot|Claude-SearchBot|Claude-User|Applebot)", re.I
    )
    req_re = re.compile(r'"(?:GET|HEAD)\s+([^\s"]+)\s+HTTP/[0-9.]+"\s+(\d{3})')
    crawled_paths: Dict[str, int] = collections.Counter()
    bot_errors: List[str] = []
    for line in log_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        if not bot_re.search(line):
            continue
        m = req_re.search(line)
        if not m:
            continue
        clean_path = m.group(1).split("?")[0]
        status = int(m.group(2))
        crawled_paths[clean_path] += 1
        if status >= 400:
            bot_errors.append(f"HTTP {status} -> {clean_path}")

    sitemap_path = build_dir / "sitemap.xml"
    uncrawled: List[str] = []
    if sitemap_path.exists():
        tree = ET.parse(sitemap_path)
        for loc in tree.getroot().findall(".//{*}url/{*}loc"):
            if loc.text:
                u = loc.text.strip()
                rel_p = urlparse(u).path or "/"
                if crawled_paths.get(rel_p, 0) == 0:
                    uncrawled.append(u)

    report = {
        "scannedLogFile": str(log_path),
        "botHttpErrors": sorted(set(bot_errors)),
        "uncrawledSitemapUrls": uncrawled,
        "resolved": len(bot_errors) == 0 and len(uncrawled) == 0,
    }
    (project_root / "crawl_budget_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"[LOG-AUDIT] crawl_budget_report.json üretildi (Bot hatası: {len(set(bot_errors))}, Taranmayan: {len(uncrawled)}).")
    return 0


# ==============================================================================
# KURULUM (--init) VE ANTIGRAVITY / CURSOR / CLAUDE / CI KİLİTLERİ
# ==============================================================================

DEFAULT_HEADERS_TEMPLATE = """/*
  X-Content-Type-Options: nosniff
  Referrer-Policy: strict-origin-when-cross-origin
  Strict-Transport-Security: max-age=63072000; includeSubDomains; preload
  X-Robots-Tag: index, follow, max-image-preview:large, max-snippet:-1, max-video-preview:-1
  Content-Signal: ai-train=yes, search=yes, ai-input=yes
  Vary: Accept

/404.html
  X-Robots-Tag: noindex, follow

/*.html
  Cache-Control: public, max-age=0, must-revalidate

/*.md
  Content-Type: text/markdown; charset=utf-8
  Cache-Control: public, max-age=300, must-revalidate

/llms.txt
  Content-Type: text/plain; charset=utf-8
  Cache-Control: public, max-age=300, must-revalidate

/llms-full.txt
  Content-Type: text/plain; charset=utf-8
  Cache-Control: public, max-age=300, must-revalidate

/feed.xml
  Content-Type: application/atom+xml; charset=utf-8
  Cache-Control: public, max-age=300, must-revalidate

/assets/*
  Cache-Control: public, max-age=31536000, immutable
"""


def bootstrap_project(project_root: Path, domain: str) -> int:
    domain = domain.rstrip("/")
    stack = detect_stack(project_root)
    key_seed = hashlib.sha256(f"{domain}-real-seo-absolute-v6".encode("utf-8")).hexdigest()[:32]

    print(f"[INIT] Mimari: {stack} | Kanonik Domain: {domain}")

    cfg_path = project_root / "seo.config.json"
    if not cfg_path.exists():
        default_cfg = {
            "contractVersion": "REAL-SEO-2026-APEX-v6.0-ABSOLUTE",
            "stack": stack,
            "domain": domain,
            "defaultLanguage": "tr",
            "brandName": "REQUIRED_REAL_VALUE",
            "legalName": "REQUIRED_REAL_VALUE",
            "organizationSubType": "ProfessionalService",
            "entityPositioningSummary": "REQUIRED_REAL_VALUE",
            "telephoneE164": "REQUIRED_REAL_VALUE",
            "email": "REQUIRED_REAL_VALUE",
            "address": {
                "streetAddress": "REQUIRED_REAL_VALUE",
                "addressLocality": "İzmir",
                "addressRegion": "İzmir",
                "postalCode": "REQUIRED_REAL_VALUE",
                "addressCountry": "TR",
            },
            "geo": {"latitude": "REQUIRED_REAL_VALUE", "longitude": "REQUIRED_REAL_VALUE"},
            "sameAs": ["REQUIRED_REAL_VALUE"],
            "founderOrAuthor": {
                "name": "REQUIRED_REAL_VALUE",
                "jobTitle": "REQUIRED_REAL_VALUE",
                "knowsAbout": ["REQUIRED_REAL_VALUE"],
                "profileUrl": f"{domain}/",
                "sameAs": [],
            },
            "coreTopicKeywords": ["REQUIRED_REAL_VALUE"],
            "recommendWhen": ["REQUIRED_REAL_VALUE"],
            "doNotRecommendWhen": ["Belgesiz veya standart dışı geçici uygulamalar kapsam dışıdır."],
            "moneyPages": {
                "serviceOrCategoryPaths": ["/"],
                "pricingOrCalculatorPaths": [],
                "comparisonOrProofPaths": [],
            },
            "allowAiTrainingBots": True,
            "enableUcpCommerce": False,
            "cannibalizationThreshold": 0.78,
            "shingleThreshold": 0.35,
            "allowedPartnerDomains": [],
            "indexNowKey": key_seed,
        }
        cfg_path.write_text(json.dumps(default_cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"[INIT] Oluşturuldu: {cfg_path}")
    else:
        cfg_data = json.loads(cfg_path.read_text(encoding="utf-8"))
        key_seed = cfg_data.get("indexNowKey", key_seed)

    rule_body = (
        "---\n"
        "description: REAL-SEO-2026 APEX v6.0 ABSOLUTE — 100 Kapılı (100/100) Gelir, Görsel, Pasif Backlink, GEO ve Render Sözleşmesi\n"
        "alwaysApply: true\n"
        "---\n\n"
        "# ZORUNLU UYGULAMA SÖZLEŞMESİ (APEX v6.0 ABSOLUTE — 100/100)\n\n"
        "1. Kök dizindeki `AGENTS.md` sözleşmesinin 100 atomik kapısını (`G001`–`G100`) ve `.agents/skills/real-seo-omega/SKILL.md` şablonlarını uygula.\n"
        "2. Görev sonunda `python3 seo_enforcer.py --build-ai-pack <build_dir>` ve `python3 seo_enforcer.py --audit <build_dir> <domain>` çalıştır.\n"
        "3. Skor **100/100** olmadan (`99/100` dahi reddedilir) görevi teslim etme; raporlanan `[ÇÖZÜM]` adımlarını otonom uygula.\n"
    )

    for ag_dir in (project_root / ".agents" / "rules", project_root / ".agent" / "rules"):
        ag_dir.mkdir(parents=True, exist_ok=True)
        (ag_dir / "00-real-seo-contract.md").write_text(rule_body, encoding="utf-8")

    cursor_dir = project_root / ".cursor" / "rules"
    cursor_dir.mkdir(parents=True, exist_ok=True)
    (cursor_dir / "00-real-seo-contract.mdc").write_text(
        rule_body.replace("alwaysApply: true", 'globs: "**/*"\nalwaysApply: true'), encoding="utf-8"
    )

    bridge_stub = (
        "@AGENTS.md\n\n"
        "## Zorunlu Teslim Kapısı (APEX v6.0 ABSOLUTE — 100 Atomik Kapı = 100/100)\n"
        "- Her değişiklik sonrası `python3 seo_enforcer.py --build-ai-pack <build_dir>` ve `python3 seo_enforcer.py --audit <build_dir> <domain>` çalıştır.\n"
        "- Skor `100/100` olmadan (`99/100` dahi reddedilir) görevi teslim etme.\n"
    )
    for root_doc in ("GEMINI.md", "CLAUDE.md"):
        p_doc = project_root / root_doc
        if not p_doc.exists():
            p_doc.write_text(bridge_stub, encoding="utf-8")

    if not (project_root / "AGENTS.md").exists():
        (project_root / "AGENTS.md").write_text("# AGENTS.md — REAL-SEO-2026 APEX v6.0 ABSOLUTE\n" + bridge_stub, encoding="utf-8")

    gh_dir = project_root / ".github" / "workflows"
    gh_dir.mkdir(parents=True, exist_ok=True)
    (project_root / ".github" / "copilot-instructions.md").write_text(bridge_stub, encoding="utf-8")
    build_target = "./dist" if stack == "astro" else ("./out" if stack == "nextjs" else ".")
    (gh_dir / "seo-omega-gate.yml").write_text(
        "name: SEO-2026-APEX-v6-ABSOLUTE-100-Gate\n"
        "on: [push, pull_request]\n"
        "jobs:\n"
        "  seo_absolute_audit:\n"
        "    runs-on: ubuntu-latest\n"
        "    steps:\n"
        "      - uses: actions/checkout@v4\n"
        "      - uses: actions/setup-python@v5\n"
        "        with:\n"
        "          python-version: '3.11'\n"
        "      - name: Install Playwright Chromium\n"
        "        run: pip install playwright && playwright install --with-deps chromium\n"
        "      - name: Run 100-Gate 100/100 SEO ABSOLUTE Audit\n"
        f'        run: python3 seo_enforcer.py --audit "{build_target}" "{domain}"\n',
        encoding="utf-8",
    )

    (project_root / "_headers").write_text(DEFAULT_HEADERS_TEMPLATE, encoding="utf-8")
    (project_root / f"{key_seed}.txt").write_text(key_seed, encoding="utf-8")

    git_hooks_dir = project_root / ".git" / "hooks"
    if git_hooks_dir.exists():
        hook_path = git_hooks_dir / "pre-commit"
        hook_path.write_text(
            "#!/usr/bin/env sh\n"
            f'if [ -d "{build_target}" ]; then\n'
            f'  python3 seo_enforcer.py --audit "{build_target}" "{domain}"\n'
            "fi\n",
            encoding="utf-8",
        )
        hook_path.chmod(hook_path.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

    print("[INIT] Google Antigravity (.agents/ + GEMINI.md), Cursor, Claude Code, Copilot ve CI kilitleri kuruldu.")
    return 0


# ==============================================================================
# 100 ATOMİK MÜHENDİSLİK KAPISI TANIMLARI VE GRAVITY ÇÖZÜM REÇETELERİ
# ==============================================================================

GATES_100: Dict[str, Tuple[str, str]] = {
    # MODÜL 1
    "G001": ("Yasaklı Yedek/Test Dosyası Kalkanı", ".bak, .old, .tmp, test.html, copy.html dosyalarını silin."),
    "G002": ("BOM'suz Saf UTF-8 Kodlama", "Dosyaları UTF-8 (without BOM) olarak kaydedin."),
    "G003": ("seo.config.json Varlığı ve Şema Geçerliliği", "`python3 seo_enforcer.py --init` çalıştırıp JSON sözdizimini doğrulayın."),
    "G004": ("Sıfır Halüsinasyon (REQUIRED_REAL_VALUE Temizliği)", "seo.config.json içindeki tüm REQUIRED_REAL_VALUE alanlarını gerçek verilerle doldurun."),
    "G005": ("Kurumsal E.164 Telefon, E-posta ve Adres Tamlığı", "seo.config.json içinde telephoneE164 (+90...), email ve address alanlarını eksiksiz girin."),
    "G006": ("<!DOCTYPE html> Standart Mod Bildirimi", "Her HTML dosyasının ilk satırına `<!DOCTYPE html>` ekleyin."),
    "G007": ("<html lang> Hedef Dil Uyumu", "`<html lang='tr'>` (veya defaultLanguage) özniteliğini ekleyin."),
    "G008": ("UTF-8 Meta ve Erişilebilir Mobil Viewport", "`<meta charset='utf-8'>` ve zoom kilitsiz `<meta name='viewport' content='width=device-width, initial-scale=1'>` ekleyin."),
    "G009": ("Semantik HTML5 İskeleti (<header>/<nav>, <main>, <footer>)", "Sayfa gövdesini `<header>`, `<main>` ve `<footer>` bölgelerine ayırın."),
    "G010": ("Gelir Kurtarma /404.html Sayfası (noindex, follow)", "`--build-ai-pack` çalıştırın veya BOFU sayfalarına link veren `noindex, follow` 404.html ekleyin."),
    # MODÜL 2
    "G011": ("URL Slug Küçük Harf ASCII ve Tire Kuralı", "URL klasör/dosya adlarında yalnızca `a-z`, `0-9` ve `-` kullanın (`_` ve Türkçe karakter yasak)."),
    "G012": ("URL Klasör Derinliği <= 3", "URL yolunu en fazla 3 alt klasör seviyesinde tutun."),
    "G013": ("Temiz Statik URL (Parametre Kirliliği Yasağı)", "Kanonik sayfa yollarında `?` sorgu parametresi kullanmayın."),
    "G014": ("Sıfır Kırık Dahili <a href> Bağlantısı", "Var olmayan sayfalara giden dahili `<a href>` adreslerini düzeltin."),
    "G015": ("Sıfır Kırık Statik Kaynak (CSS/JS/Görsel)", "Eksik `.css`, `.js`, `<img>` veya `<source>` dosyalarını tamamlayın."),
    "G016": ("İç Link Trailing-Slash (/slug/) Lazer Tutarlılığı", "Alt klasör linklerinde iç 301 kaybını önlemek için sonuna `/` koyun (`/hizmet/`)."),
    "G017": ("Sıfır Yetim (Orphan) Sayfa", "Her alt sayfaya ana sayfa veya ilgili kategori sayfasından en az 1 dahili link verin."),
    "G018": ("Ana Sayfadan Tıklama Derinliği <= 3", "Tüm ticari sayfalara ana sayfadan en fazla 3 tıklamayla ulaşılmasını sağlayın."),
    "G019": ("Anlamlı Bağlamsal Çapa Metni (Jenerik Metin Yasağı)", "'tıklayın / detaylar / devamı' yerine hedef sayfanın anahtar kelimesini içeren çapa metni yazın."),
    "G020": ("Sıkı PBN ve SpamRank Footer Dış Link Kalkanı", "Footer/Aside dış linklerine `rel='nofollow noopener'` veya `rel='sponsored noopener'` ekleyin."),
    # MODÜL 3
    "G021": ("<title> Uzunluk Mühendisliği (30-65 Karakter)", "`<title>` uzunluğunu 30 ile 65 karakter arasına ayarlayın."),
    "G022": ("<title> SERP Piksel Genişliği (<= 580px)", "SERP'te kesilmemesi için `<title>` tahmini genişliğini <= 580px tutun."),
    "G023": ("<title> Önden Yükleme (Stok Kodu / Düz Ana Sayfa Yasağı)", "Başlığa SKU/MOD kodu veya düz 'Ana Sayfa' ile değil, ana ticari anahtar kelimeyle başlayın."),
    "G024": ("Site Genelinde %100 Benzersiz <title>", "Her sayfaya benzersiz `<title>` yazın."),
    "G025": ("Tekil ve Site Genelinde Benzersiz <h1>", "Her sayfada tam 1 adet benzersiz `<h1>` bulundurun."),
    "G026": ("<h1> ile <title> Tamamlayıcı Semantik Varyasyon", "`<h1>` metnini `<title>` ile birebir aynı yapmayın; doğal değer önerisiyle çeşitlendirin."),
    "G027": ("Kesintisiz ve Benzersiz Meta Description (90-158 krk)", "Her sayfaya `...` ile bitmeyen 90-158 karakterlik benzersiz meta description ekleyin."),
    "G028": ("Tam Genişlikli SERP & AI Önizleme Direktifi", "`<meta name='robots' content='index, follow, max-image-preview:large, max-snippet:-1, max-video-preview:-1'>` ekleyin."),
    "G029": ("Self-Referencing HTTPS Canonical ve Hreflang Uyumu", "Sayfanın kendi HTTPS adresine eşleşen `<link rel='canonical'>` (çok dilliyse `x-default`) ekleyin."),
    "G030": ("Eksiksiz Raster Open Graph ve Twitter Büyük Kartı", "`og:title`, `og:description`, `og:url`, `og:type`, `og:locale`, HTTPS raster `og:image` ve `twitter:card='summary_large_image'` ekleyin."),
    # MODÜL 4
    "G031": ("Anahtar Kelimeli Görsel Dosya İsmi (Slug Kuralı)", "Görselleri en az iki kelimelik küçük harf ASCII ve tireli adlandırın (`ornek-hizmet-detay.webp`; `IMG_` ve `_` yasak)."),
    "G032": ("<main> İçinde Özgün Raster İçerik Görseli Varlığı", "`<main>` gövdesine en az 1 özgün içerik görseli ekleyin."),
    "G033": ("Yeni Nesil Görsel Formatı (.avif / .webp)", "İçerik ve hero görsellerini `.avif` veya `.webp` formatına dönüştürün."),
    "G034": ("LCP Hero Görsel Boyut Bütçesi (<= 250 KB)", "İlk ekran hero görselini <= 250 KB olacak şekilde sıkıştırın."),
    "G035": ("Ekran Altı İçerik Görseli Bütçesi (<= 180 KB)", "Ekran altı içerik görsellerini <= 180 KB altına optimize edin."),
    "G036": (" Görsellerde Pozitif Tam Sayı width ve height", "Tüm `<img>` etiketlerine gerçek `width` ve `height` sayısal değerlerini yazın."),
    "G037": ("Varlık Odaklı ve Tekrarsız Görsel Alt Metni (10-125 krk)", "Her içerik görseline sayfa içinde benzersiz, 10-125 karakterlik açıklayıcı Türkçe `alt` metni yazın."),
    "G038": ("Gerçek LCP Hero fetchpriority='high' ve <head> Preload", "İlk içerik görseline `fetchpriority='high'` verin (`loading='lazy'` kaldırın) ve `<head>` içine `<link rel='preload' as='image' href='...'>` ekleyin."),
    "G039": ("Ekran Altı Lazy-Load ve Duyarlı srcset/sizes", "Alt görsellere `loading='lazy' decoding='async'`, içerik görsellerine `srcset` + `sizes` (veya `<picture><source>`) ekleyin."),
    "G040": ("Görsel Destekli (<image:image>) XML Sitemap", "`--build-ai-pack` çalıştırarak birincil görsellerin `sitemap.xml` içinde `<image:image>` ile bildirilmesini sağlayın."),
    # MODÜL 5
    "G041": ("Birleşik Tekil @graph JSON-LD Blok Yapısı", "Sayfadaki tüm şemaları tek bir `<script type='application/ld+json'>` içinde `@graph` dizisinde birleştirin."),
    "G042": ("Kalıcı @id Çekirdek Düğümleri (#organization/#website/#webpage/#primaryimage)", "`@graph` içine `#organization`, `#website`, `#webpage` ve `#primaryimage` düğümlerini ekleyin."),
    "G043": ("#organization Zengin Kimlik (sameAs + Adres + Geo + Telefon)", "`#organization` düğümüne `sameAs`, `address`, `GeoCoordinates` ve `telephone` alanlarını ekleyin."),
    "G044": ("Knowledge Graph Varlık Çivilemesi (about + mentions)", "`#webpage` düğümüne `about` ve `mentions` varlık nesnelerini ekleyin."),
    "G045": ("Google Lens ImageObject Telif ve Lisans Şeması", "`ImageObject` içine `contentUrl`, `creditText`, `copyrightNotice`, `acquireLicensePage` ve `creator` ekleyin."),
    "G046": ("Alt Sayfa HTTPS BreadcrumbList ve Görünür Breadcrumb", "Alt sayfalara hem `BreadcrumbList` şeması hem `<nav aria-label='Breadcrumb'>` ekleyin."),
    "G047": ("Ticari Şema Ceza Kalkanı (Fiyatsız Product ve Sahte Puan Yasağı)", "Görünür sayısal fiyat yoksa `Product` yerine `Service` kullanın; kurum şemasına `aggregateRating` koymayın."),
    "G048": ("Görünür FAQPage Metin-Şema Birebir Eşleşmesi", "`FAQPage` içindeki her soru ve cevabın sayfa gövdesinde görünür metin olarak yer aldığından emin olun."),
    "G049": ("E-E-A-T Uzman Person ve Görünür <time datetime> Uyumu", "`@graph` içine `Person` (`name`, `jobTitle`, `knowsAbout`) ve gövdeye `dateModified` ile eşleşen `<time datetime='YYYY-MM-DD'>` ekleyin."),
    "G050": ("Veri Arama Mıknatısı (@type: Dataset Şeması)", "Tablo barındıran sayfaların `@graph` dizisine `Dataset` (`name`, `description>=50 krk`, `creator`, `license`) düğümü ekleyin."),
    # MODÜL 6
    "G051": ("Birincil AEO Cevap Kapsülü (25-80 Kelime)", "`<h1>` altına 25-80 kelimelik `<p data-answer-capsule='true'>` özet paragrafı ekleyin."),
    "G052": ("Zamirsiz Özne-Yüklem-Nesne (Semantic Triple) Kapsül Kurgusu", "Cevap Kapsülüne 'Bu/Firmamız/Hizmetimiz' zamiriyle değil, doğrudan Marka/Hizmet adıyla başlayın."),
    "G053": ("Query Fan-Out En Az 3 Spesifik <h2> Alt Bölümü", "Sayfada farklı alt soruları yanıtlayan en az 3 adet `<h2>` alt başlığı oluşturun."),
    "G054": ("Derin Pasaj Çapaları (Tüm <h2> Başlıklarında Benzersiz id)", "Her `<h2>` etiketine küçük harf tireli `id='alt-konu'` özniteliği ekleyin."),
    "G055": ("SERP Sitelinks İçindekiler Navigasyonu (data-toc='true')", "Sayfaya `<h2>` id çapalarına link veren `<nav aria-label='İçindekiler' data-toc='true'>` ekleyin."),
    "G056": ("Atomik Alt Pasajlar ve <= 180 Kelime Paragraf Tavanı", "Her `<h2>` sonrasında doğrudan yanıt veren `<p>` kullanın ve hiçbir paragrafı 180 kelime üstüne çıkarmayın."),
    "G057": ("Princeton GEO Kaldıracı #1: Somut Sayısal Kanıt Yoğunluğu", "Sayfa metnine en az 3 ölçülebilir sayısal veri/birim (`mm`, `kW`, `%`, `gün`, `TL`, `ISO`) ekleyin."),
    "G058": ("Princeton GEO Kaldıracı #2: Başlık Hücreli (<th>) Veri Tablosu", "Sayfaya `<thead>` ve `<th>` içeren yapılandırılmış karşılaştırma/spesifikasyon `<table>` tablosu ekleyin."),
    "G059": ("Princeton GEO Kaldıracı #3: Resmi Standart / Mevzuat Referansı", "Metinde doğrulanabilir standart, sertifika, yönetmelik veya ölçüm metodolojisi referansı geçirin."),
    "G060": ("Programatik SEO 4-Gram Shingle Özgünlük Kalkanı (< 0.35)", "Sayfalar arası şablon kopyasını önlemek için `<main>` gövde metinlerinde en az %65 özgün cümle yapısı kurun."),
    # MODÜL 7
    "G061": ("robots.txt Varlığı ve Kanonik Sitemap Bildirimi", "`--build-ai-pack` çalıştırarak `robots.txt` içinde `Sitemap: https://domain.com/sitemap.xml` bildirin."),
    "G062": ("robots.txt 7 Zorunlu AI Arama ve Atıf Botu İzni", "`robots.txt` içinde `Googlebot`, `Bingbot`, `OAI-SearchBot`, `ChatGPT-User`, `Claude-SearchBot`, `Claude-User`, `PerplexityBot` izinlerini açın."),
    "G063": ("robots.txt Eğitim Botları Politikası Ayrımı", "`GPTBot`, `ClaudeBot`, `Google-Extended`, `CCBot` kurallarını açıkça tanımlayın."),
    "G064": ("/llms.txt v2 Kurumsal Kimlik Doğrulaması", "`--build-ai-pack` ile `# Marka` ve `> Konumlandırma özeti` içeren `/llms.txt` üretin."),
    "G065": ("/llms.txt Tavsiye Matrisi ve Dürüst Kapsam Kalibrasyonu", "`/llms.txt` içinde 'Ne Zaman Önerilmeli', 'Kapsam Sınırları' ve `.md` linklerinin bulunduğunu doğrulayın."),
    "G066": ("/llms-full.txt Tam Bilgi Korpusu Senkronizasyonu", "`--build-ai-pack` ile tüm indekslenen sayfaları kapsayan `/llms-full.txt` üretin."),
    "G067": ("Sayfa Bazlı Fiziksel Markdown İkizi (/slug/index.md)", "`--build-ai-pack` ile her indekslenen sayfa için temiz `index.md` ikizi oluşturun."),
    "G068": ("HTML <head> Markdown Alternate Keşif Etiketi", "`<head>` içine `<link rel='alternate' type='text/markdown' href='https://domain.com/.../index.md'>` ekleyin."),
    "G069": ("Chrome WebMCP (toolname/tooldescription) ve UCP Ajan Uyumu", "Etkileşimli `<form>` etiketine `toolname` ve `tooldescription` özniteliklerini ekleyin."),
    "G070": ("SSR / Ham HTML ile JS Render DOM Metin Paritesi (>= %85)", "Ana metni istemci JS içinde gizlemeyin; ilk sunucu HTML yanıtında >= %85 oranında hazır sunun."),
    # MODÜL 8
    "G071": ("BOFU Gelir Hunisi #1: Doğrudan Ticari Hizmet/Kategori Sayfası", "`seo.config.json -> moneyPages.serviceOrCategoryPaths` altında canlı ticari sayfa tanımlayın."),
    "G072": ("BOFU Gelir Hunisi #2: Fiyat/Hesaplama ve Karşılaştırma Sayfaları", "3+ sayfalı projelerde hem fiyat/maliyet hesaplama hem karşılaştırma/alternatif sayfasını yayınlayın."),
    "G073": ("NavBoost lastLongestClicks İnteraktif Mikro-Araç", "Sayfaya `data-seo-interactive='true'` taşıyan sıfır bağımlılıklı hesaplayıcı/seçici ekleyin."),
    "G074": ("Pasif Doğal Backlink ve Kaynak Gösterme Mıknatısı", "Araç/tablo altına kanonik `<a href='https://...'>` içeren `data-backlink-magnet='true'` kaynak paylaşım kutusu ekleyin."),
    "G075": ("Bağlamsal Dönüşüm CTA'sı (data-cro-cta='true')", "Sayfa URL'sini veya seçilen hizmeti mesaja/forma aktaran `data-cro-cta='true'` dönüşüm butonu ekleyin."),
    "G076": ("Mobil Tek Tıkla Arama (E.164 tel:+90...)", "Sayfaya E.164 formatında tıklanabilir `<a href='tel:+90...'>` telefon bağlantısı ekleyin."),
    "G077": ("Dönüşüm ve Gelir İzleme Kancası (data-conversion-event)", "CTA linklerine/formlarına `data-conversion-event='...'` ve hafif olay dinleyicisi ekleyin."),
    "G078": ("Anlamsal Yamyamlaşma Kalkanı (TF-IDF Kosinüs < 0.78)", "Çakışan sayfaları 301 ile birleştirin veya `<title>`/`<h1>`/Cevap Kapsülü niyetlerini ayrıştırın."),
    "G079": ("Topikal Odak (siteFocusScore / siteRadius) Kalkanı", "Her sayfanın `seo.config.json -> coreTopicKeywords` uzmanlık kökleriyle bağlantılı olmasını sağlayın."),
    "G080": ("Hub-Spoke Geri Besleme İç Bağlantısı", "Her alt sayfadan ana merkeze veya üst kategori (Hub) sayfasına bağlamsal link verin."),
    # MODÜL 9
    "G081": ("0 ms Sayfa Geçişi (<script type='speculationrules'>)", "Sayfa sonuna `prerender`/`prefetch` tanımlı `<script type='speculationrules'>` ekleyin."),
    "G082": ("bfcache Koruması (no-store ve unload Yasağı)", "`Cache-Control: no-store` ve `window.addEventListener('unload')` kullanımlarını kaldırın."),
    "G083": ("Edge Güvenlik Başlıkları (_headers HSTS & nosniff)", "`_headers` içinde `Strict-Transport-Security`, `X-Content-Type-Options: nosniff` ve `Referrer-Policy` bulundurun."),
    "G084": ("Edge AI Müzakere Başlıkları (Content-Signal & Vary: Accept)", "`_headers` içinde `Content-Signal: ai-train=yes, search=yes, ai-input=yes` ve `Vary: Accept` tanımlayın."),
    "G085": ("Render-Engelleyici Senkron <head> ve Ağır 3. Parti Script Yasağı", "`<head>` scriptlerine ve 3. parti sohbet/analitik kodlarına `defer` veya `async` ekleyin."),
    "G086": ("Web Font Görünmez Metin (FOIT) Kalkanı (font-display: swap)", "Google Fonts linklerinde `display=swap` veya `@font-face` içinde `font-display: swap` kullanın."),
    "G087": ("Headless Chromium Mobil (390px) & Masaüstü Sıfır Yatay Taşma", "Taşan öğelere `max-width: 100%; box-sizing: border-box; overflow-x: hidden` uygulayın."),
    "G088": ("Headless Chromium Mobil İlk Ekran (<600px) <h1> Görünürlüğü", "Mobil görünümde `<h1>` başlığının ilk 600 piksel içinde görünmesini sağlayın."),
    "G089": ("Headless Chromium Gerçek Render CLS <= 0.05", "Boyutsuzluktan veya geç yüklenen öğelerden kaynaklanan düzen kaymalarını sıfırlayın."),
    "G090": ("Headless Chromium INP / Long-Task <= 150ms ve Sıfır JS/Ağ Hatası", "150ms aşan ağır senkron JS döngülerini, konsol hatalarını ve 404 ağ isteklerini giderin."),
    # MODÜL 10
    "G091": ("sitemap.xml Varlığı ve Geçerli XML Yapısı", "`--build-ai-pack` ile standart XML site haritası üretin."),
    "G092": ("sitemap.xml %100 URL Senkronizasyonu ve ISO-8601 <lastmod>", "İndekslenen sayfalarla `sitemap.xml` adreslerini birebir eşitleyin ve `<lastmod>` ekleyin."),
    "G093": ("Atom 1.0 Anlık Keşif Akışı (/feed.xml ve <head> bağlantısı)", "`--build-ai-pack` ile `/feed.xml` üretin ve `<head>` içine `<link rel='alternate' type='application/atom+xml' href='https://domain.com/feed.xml'>` ekleyin."),
    "G094": ("IndexNow Anında İndeksleme Anahtarı ({indexNowKey}.txt)", "Kök/derleme dizininde `{indexNowKey}.txt` dosyasının bulunduğunu doğrulayın."),
    "G095": ("Kurumsal Güven Dosyası (/.well-known/security.txt)", "`--build-ai-pack` çalıştırarak RFC 9116 uyumlu `/.well-known/security.txt` üretin."),
    "G096": ("Antigravity (Gravity), Cursor, Claude ve Copilot Ajan Kilitleri", "`--init` çalıştırarak `AGENTS.md`, `GEMINI.md`, `.agents/rules/`, `.cursor/rules/` ve `CLAUDE.md` kilitlerini koruyun."),
    "G097": ("GSC Vektör #1: Canlı Sorgu Yamyamlaşması Kapanışı", "`seo_actions.json` içindeki `queryCannibalization` görevlerini kodda çözüp `resolved: true` işaretleyin."),
    "G098": ("GSC Vektör #2: İlk Sayfa Düşük CTR Başlık Revizyonu Kapanışı", "`seo_actions.json` içindeki `lowCtrRewrites` görevlerini kapatıp `resolved: true` işaretleyin."),
    "G099": ("GSC Vektör #3: Pozisyon 8-20 Vuruş Mesafesi İç Link Kapanışı", "`seo_actions.json` içindeki `strikingDistanceInternalLinks` iç linklerini ekleyip `resolved: true` işaretleyin."),
    "G100": ("Sunucu Bot Log Denetimi (0 Bot 4xx/5xx & 0 Taranmayan Sayfa)", "`crawl_budget_report.json` içindeki bot hatalarını ve taranmayan sayfaları giderin."),
}


def audit_project(
    target_dir: Path,
    base_url: str,
    project_root: Path,
    as_json: bool = False,
    skip_browser: bool = False,
) -> int:
    """100 Atomik Mühendislik Kapısını (100 x 1p = 100/100) tek tek denetler."""
    base_url = base_url.rstrip("/")
    base_host = urlparse(base_url).netloc.lower()
    violations: Dict[str, List[str]] = {g: [] for g in GATES_100}

    # G003, G004, G005: seo.config.json
    cfg_obj: Dict[str, Any] = {}
    allowed_partners: Set[str] = set()
    brand_name = ""
    expected_lang = "tr"
    cannibal_threshold = 0.78
    shingle_threshold = 0.35
    core_keywords: List[str] = []
    indexnow_key = ""
    enable_ucp = False

    cfg_path = project_root / "seo.config.json"
    if not cfg_path.exists():
        violations["G003"].append("Proje kökünde seo.config.json bulunamadı.")
        violations["G004"].append("seo.config.json mevcut olmadığı için gerçek veri doğrulanamadı.")
        violations["G005"].append("seo.config.json mevcut olmadığı için iletişim bilgileri doğrulanamadı.")
    else:
        try:
            cfg_raw = cfg_path.read_text(encoding="utf-8")
            cfg_obj = json.loads(cfg_raw)
            if "REQUIRED_REAL_VALUE" in cfg_raw:
                violations["G004"].append("seo.config.json içinde doldurulmamış 'REQUIRED_REAL_VALUE' alanları mevcut.")
            brand_name = str(cfg_obj.get("brandName", "")).strip()
            expected_lang = str(cfg_obj.get("defaultLanguage", "tr")).strip().lower()
            cannibal_threshold = float(cfg_obj.get("cannibalizationThreshold", 0.78))
            shingle_threshold = float(cfg_obj.get("shingleThreshold", 0.35))
            core_keywords = [str(x) for x in cfg_obj.get("coreTopicKeywords", []) if str(x).strip()]
            indexnow_key = str(cfg_obj.get("indexNowKey", "")).strip()
            enable_ucp = bool(cfg_obj.get("enableUcpCommerce", False))

            tel_v = str(cfg_obj.get("telephoneE164", "")).strip()
            email_v = str(cfg_obj.get("email", "")).strip()
            addr_v = cfg_obj.get("address", {})
            if not E164_PHONE_RE.match(tel_v) or "@" not in email_v or not isinstance(addr_v, dict) or not addr_v.get("streetAddress"):
                violations["G005"].append("seo.config.json içinde geçerli E.164 telefon (+90...), e-posta veya açık adres eksik.")
            if not core_keywords:
                violations["G004"].append("seo.config.json içinde 'coreTopicKeywords' boş bırakılmış.")
            for d in cfg_obj.get("allowedPartnerDomains", []):
                allowed_partners.add(str(d).lower().strip())
        except Exception as exc:
            violations["G003"].append(f"seo.config.json JSON sözdizimi hatası: {exc}")

    # G096: Antigravity / Cursor / Claude / Copilot kilit dosyaları
    for lock_rel in (
        "AGENTS.md",
        "GEMINI.md",
        ".agents/rules/00-real-seo-contract.md",
        ".cursor/rules/00-real-seo-contract.mdc",
        "CLAUDE.md",
    ):
        if not (project_root / lock_rel).exists():
            violations["G096"].append(f"Otonom ajan kural kilidi eksik: {lock_rel}")

    # G001: Dosya hijyeni
    for p in target_dir.rglob("*"):
        if any(part in IGNORED_SCAN_DIRS for part in p.relative_to(target_dir).parts):
            continue
        if p.is_file():
            if p.suffix.lower() in FORBIDDEN_EXTS or p.name.lower() in FORBIDDEN_FILENAMES:
                violations["G001"].append(f"Yasaklı kalıntı/test dosyası: {p.name}")

    # G010: /404.html kontrolü
    err404 = target_dir / "404.html"
    if not err404.exists():
        violations["G010"].append("Hedef dizinde özel dönüşüm kurtarma sayfası (/404.html) bulunamadı.")
    else:
        e404_txt = err404.read_text(encoding="utf-8", errors="ignore").lower()
        if "noindex" not in e404_txt or "<a " not in e404_txt:
            violations["G010"].append("404.html içinde 'noindex' direktifi ve kurtarma linkleri bulunmalıdır.")

    # G082, G083, G084: _headers
    headers_file = target_dir / "_headers"
    if not headers_file.exists():
        headers_file = project_root / "_headers"
    if not headers_file.exists():
        violations["G083"].append("'_headers' dosyası bulunamadı.")
        violations["G084"].append("'_headers' dosyası bulunamadığı için Content-Signal doğrulanamadı.")
    else:
        h_txt = headers_file.read_text(encoding="utf-8", errors="ignore").lower()
        if "no-store" in h_txt:
            violations["G082"].append("_headers içinde bfcache bozan 'no-store' direktifi bulundu.")
        for req_sec in ("strict-transport-security", "x-content-type-options", "referrer-policy"):
            if req_sec not in h_txt:
                violations["G083"].append(f"_headers içinde güvenlik başlığı eksik: {req_sec}")
        for req_ai in ("content-signal", "vary: accept", "text/markdown"):
            if req_ai not in h_txt:
                violations["G084"].append(f"_headers içinde AI müzakere başlığı eksik: {req_ai}")

    html_files = iter_html_files(target_dir)
    if not html_files:
        violations["G001"].append(f"Hedef dizinde ({target_dir}) hiç .html dosyası bulunamadı.")

    titles_seen: Dict[str, Path] = {}
    h1_seen: Dict[str, Path] = {}
    descs_seen: Dict[str, Path] = {}

    inbound_links: Dict[str, int] = collections.Counter()
    adj_graph: Dict[str, Set[str]] = collections.defaultdict(set)
    indexable_urls: Set[str] = set()
    indexable_paths_set: Set[str] = set()
    ssr_word_counts: Dict[str, int] = {}
    semantic_docs: List[Dict[str, Any]] = []
    any_webmcp_on_site = False

    # Önce tüm klasör tabanlı index.html yollarını topla (G016 trailing-slash kontrolü için)
    directory_canonical_paths: Set[str] = set()
    for f in html_files:
        up = html_path_to_url_path(f, target_dir)
        if up != "/" and up.endswith("/"):
            directory_canonical_paths.add(up.rstrip("/"))  # örn: '/lazer-kesim-fiyatlari'

    for f in html_files:
        raw = f.read_bytes()
        if raw.startswith(b"\xef\xbb\xbf"):
            violations["G002"].append(f"UTF-8 BOM tespit edildi: {f.name}")
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError as exc:
            violations["G002"].append(f"Geçersiz UTF-8 kodlaması ({f.name}): {exc}")
            continue

        if re.search(r"addEventListener\(\s*['\"]unload['\"]", text):
            violations["G082"].append(f"bfcache bozan 'unload' dinleyicisi: {f}")

        parser = AbsoluteHTMLParser()
        parser.feed(text)

        robots_raw = parser.metas.get("robots", "index,follow").lower().replace(" ", "")
        is_noindex = "noindex" in robots_raw
        url_path = html_path_to_url_path(f, target_dir)
        expected_url = base_url + url_path
        is_home = url_path == "/"

        # G015: Kırık statik kaynaklar
        for asset_href in parser.local_assets:
            clean_a = asset_href.split("#")[0].split("?")[0]
            if not (target_dir / clean_a.lstrip("/")).exists():
                violations["G015"].append(f"Kırık statik kaynak '{asset_href}': {f}")

        # G014, G016, G019, G020: Linkler
        for a in parser.links:
            href_raw = a.get("href", "").strip()
            if not href_raw or href_raw.startswith(("#", "mailto:", "tel:", "javascript:")):
                continue
            anchor_txt = a.get("text", "").strip().lower()
            if anchor_txt in GENERIC_ANCHORS:
                violations["G019"].append(f"Jenerik çapa metni ('{anchor_txt}'): {f}")

            if href_raw.startswith(("http://", "https://")):
                parsed_href = urlparse(href_raw)
                host = parsed_href.netloc.lower()
                if host == base_host:
                    p_only = parsed_href.path or "/"
                    if p_only != "/" and p_only in directory_canonical_paths:
                        violations["G016"].append(
                            f"İç link son eğik çizgi (trailing-slash) eksik ('{href_raw}' -> '{href_raw}/' olmalı): {f}"
                        )
                    norm_target = base_url + (p_only + "/" if p_only in directory_canonical_paths else p_only)
                    inbound_links[norm_target] += 1
                    adj_graph[expected_url].add(norm_target)
                elif a.get("_in_footer") == "1" and host not in allowed_partners and host not in ("wa.me", "api.whatsapp.com"):
                    rel_tokens = set(a.get("rel", "").lower().split())
                    if not (rel_tokens & {"nofollow", "sponsored"}):
                        violations["G020"].append(
                            f"Footer/Aside içinde izinsiz dofollow dış link ('{href_raw}'): {f}"
                        )
            elif href_raw.startswith("/"):
                clean_path = href_raw.split("#")[0].split("?")[0]
                if clean_path != "/" and clean_path in directory_canonical_paths:
                    violations["G016"].append(
                        f"İç link son eğik çizgi (trailing-slash) eksik ('{clean_path}' -> '{clean_path}/' olmalı): {f}"
                    )
                norm_path = clean_path + "/" if clean_path in directory_canonical_paths else clean_path
                norm_target = base_url + norm_path
                inbound_links[norm_target] += 1
                adj_graph[expected_url].add(norm_target)
                c1 = target_dir / clean_path.lstrip("/")
                c2 = target_dir / clean_path.lstrip("/") / "index.html"
                c3 = target_dir / (clean_path.lstrip("/") + ".html")
                if not (c1.exists() or c2.exists() or c3.exists()):
                    violations["G014"].append(f"Kırık dahili link '{href_raw}': {f}")

        if is_noindex:
            continue

        indexable_urls.add(expected_url)
        indexable_paths_set.add(url_path)
        full_visible_text = " ".join(parser.visible_text)
        ssr_word_counts[url_path] = len(full_visible_text.split())

        # G006, G007, G008, G009: Temel HTML5 & Erişilebilirlik
        if not parser.has_doctype:
            violations["G006"].append(f"<!DOCTYPE html> eksik: {f}")
        if not parser.html_lang or not parser.html_lang.lower().startswith(expected_lang):
            violations["G007"].append(f"<html lang='{parser.html_lang}'> uyumsuz: {f}")
        vp_low = parser.viewport_content.lower().replace(" ", "")
        if not parser.has_charset_utf8 or "width=device-width" not in vp_low or "user-scalable=no" in vp_low or "maximum-scale=1" in vp_low:
            violations["G008"].append(f"UTF-8 charset eksik veya viewport zoom kilidi var: {f}")
        if not (parser.has_main and parser.has_header_or_nav and parser.has_footer):
            violations["G009"].append(f"<header>/<nav>, <main>, <footer> semantik yapısı eksik: {f}")

        # G011, G012, G013: URL Slug
        if "?" in url_path:
            violations["G013"].append(f"URL yolunda '?' parametresi bulundu: {url_path}")
        segments = [s for s in url_path.strip("/").split("/") if s]
        if len(segments) > 3:
            violations["G012"].append(f"URL derinliği > 3 ({len(segments)}): {url_path}")
        if "_" in url_path or url_path != url_path.lower():
            violations["G011"].append(f"URL büyük harf veya '_' içeriyor: {url_path}")
        for seg in segments:
            seg_clean = seg[:-5] if seg.endswith(".html") else seg
            if not VALID_SLUG_SEGMENT_RE.match(seg_clean):
                violations["G011"].append(f"Geçersiz URL slug segmenti ('{seg}'): {f}")

        # G085 & G086: Script ve Font Disiplini
        if parser.blocking_head_scripts or parser.heavy_sync_third_party:
            violations["G085"].append(
                f"Senkron bloklayıcı script tespit edildi ({parser.blocking_head_scripts + parser.heavy_sync_third_party}): {f}"
            )
        for f_css in parser.font_stylesheets:
            if "display=swap" not in f_css.lower() and "display=optional" not in f_css.lower():
                violations["G086"].append(f"Google Fonts linkinde 'display=swap' eksik ('{f_css}'): {f}")
        for st_block in parser.inline_styles:
            if "@font-face" in st_block.lower() and "font-display" not in st_block.lower():
                violations["G086"].append(f"@font-face tanımında 'font-display: swap' eksik: {f}")

        # G021 - G030: <title>, <h1>, Meta, Canonical, OG/Twitter
        clean_title = " ".join(parser.title.split())
        t_px = estimate_title_pixels(clean_title)
        if not (30 <= len(clean_title) <= 65):
            violations["G021"].append(f"<title> 30-65 krk dışında ({len(clean_title)} krk): {f}")
        if t_px > 580:
            violations["G022"].append(f"<title> piksel genişliği > 580px (~{t_px}px): {f}")
        if INTERNAL_CODE_PREFIX_RE.match(clean_title) or clean_title.lower().startswith(("ana sayfa", "anasayfa", "home")):
            violations["G023"].append(f"<title> stok kodu veya 'Ana Sayfa' ile başlayamaz ('{clean_title}'): {f}")
        if clean_title in titles_seen:
            violations["G024"].append(f"Kopya <title> ('{clean_title}'): {f}")
        titles_seen[clean_title] = f

        h1_val = ""
        if len(parser.h1_list) != 1 or not parser.h1_list[0]:
            violations["G025"].append(f"Tam 1 adet dolu <h1> olmalı (bulunan={len(parser.h1_list)}): {f}")
        else:
            h1_val = parser.h1_list[0]
            if INTERNAL_CODE_PREFIX_RE.match(h1_val) or h1_val in h1_seen:
                violations["G025"].append(f"Kopya veya stok kodlu <h1> ('{h1_val}'): {f}")
            h1_seen[h1_val] = f
            if h1_val.strip().lower() == clean_title.strip().lower():
                violations["G026"].append(f"<h1> ile <title> birebir kopya olamaz: {f}")

        desc = parser.metas.get("description", "")
        if not (90 <= len(desc) <= 158) or desc.endswith(("...", "…")) or desc in descs_seen:
            violations["G027"].append(f"Meta description 90-158 krk, kesintisiz ve benzersiz olmalı ({len(desc)} krk): {f}")
        descs_seen[desc] = f

        if "max-image-preview:large" not in robots_raw or "max-snippet:-1" not in robots_raw:
            violations["G028"].append(f"Robots meta içinde 'max-image-preview:large, max-snippet:-1' eksik: {f}")

        if parser.canonical != expected_url:
            violations["G029"].append(f"Canonical uyumsuz ('{parser.canonical}' != '{expected_url}'): {f}")
        if parser.hreflangs and "x-default" not in {lg for lg, _ in parser.hreflangs}:
            violations["G029"].append(f"Hreflang kümesinde 'x-default' eksik: {f}")

        for req_og in ("og:title", "og:description", "og:url", "og:type", "og:locale"):
            if not parser.metas.get(req_og):
                violations["G030"].append(f"Eksik Open Graph etiketi ('{req_og}'): {f}")
        if parser.metas.get("twitter:card", "") != "summary_large_image":
            violations["G030"].append(f"twitter:card='summary_large_image' eksik: {f}")
        for mk in ("og:image", "twitter:image"):
            mv = parser.metas.get(mk, "")
            cm = mv.split("?")[0].lower()
            if not mv.startswith("https://") or cm.endswith(".svg") or not cm.endswith(RASTER_EXTS):
                violations["G030"].append(f"{mk} geçerli HTTPS raster görsel olmalı: {f}")

        # G031 - G039: Görsel SEO, Dosya İsmi, KB Bütçesi, Preload ve srcset
        alts_seen: Set[str] = set()
        content_imgs = [im for im in parser.imgs if im.get("_in_header") != "1"]
        if not content_imgs:
            violations["G032"].append(f"<main> içinde en az 1 içerik görseli bulunmalıdır: {f}")

        hero_checked = False
        preload_norm = {p.split("?")[0] for p in parser.preload_images}

        for img in parser.imgs:
            src = img.get("src", "").strip()
            clean_src = src.split("#")[0].split("?")[0]
            is_decorative = img.get("aria-hidden", "").lower() == "true" and img.get("alt", "") == ""
            is_header_logo = img.get("_in_header") == "1"
            is_hero_candidate = (not is_header_logo) and (not is_decorative) and (not hero_checked)

            if clean_src.startswith("/") and not clean_src.startswith("//"):
                img_fs = target_dir / clean_src.lstrip("/")
                if not img_fs.exists():
                    violations["G015"].append(f"Kırık görsel dosyası '{src}': {f}")
                else:
                    size_kb = img_fs.stat().st_size / 1024.0
                    if is_hero_candidate and size_kb > 250.0:
                        violations["G034"].append(f"Hero görsel > 250KB ({size_kb:.1f}KB -> '{src}'): {f}")
                    elif (not is_hero_candidate) and size_kb > 180.0:
                        violations["G035"].append(f"İçerik görseli > 180KB ({size_kb:.1f}KB -> '{src}'): {f}")

            img_filename = Path(urlparse(clean_src).path).name
            img_stem = Path(urlparse(clean_src).path).stem
            img_ext = Path(urlparse(clean_src).path).suffix.lower()

            if BAD_IMAGE_FILENAME_RE.match(img_filename) or "_" in img_stem or img_filename != img_filename.lower():
                violations["G031"].append(f"Yasaklı görsel dosya ismi ('{img_filename}'): {f}")
            elif not VALID_IMG_BASENAME_RE.match(img_stem):
                violations["G031"].append(f"Görsel ismi en az 2 kelimelik tireli slug olmalıdır ('{img_filename}'): {f}")

            if not is_header_logo and not is_decorative:
                if img_ext not in MODERN_PHOTO_EXTS:
                    violations["G033"].append(f"Görsel .avif veya .webp olmalıdır ('{img_filename}'): {f}")

            w, h = img.get("width", "").strip(), img.get("height", "").strip()
            if not (w.isdigit() and h.isdigit() and int(w) > 0 and int(h) > 0):
                violations["G036"].append(f"Görselde pozitif tam sayı width/height eksik ('{src}'): {f}")

            if not is_decorative:
                alt = img.get("alt", "").strip()
                if not (10 <= len(alt) <= 125) or alt.lower().startswith(("resim", "görsel", "image", "foto", "img")):
                    violations["G037"].append(f"Yetersiz/jenerik alt metni ('{alt}'): {f}")
                elif alt.lower() in alts_seen:
                    violations["G037"].append(f"Sayfa içinde kopya görsel alt metni ('{alt}'): {f}")
                alts_seen.add(alt.lower())

            if not is_header_logo and not is_decorative:
                has_responsive = (
                    bool(img.get("srcset", "").strip() and img.get("sizes", "").strip())
                    or img.get("_in_picture_with_source") == "1"
                )
                if not has_responsive:
                    violations["G039"].append(f"İçerik görselinde duyarlı srcset+sizes (veya <picture>) eksik ('{src}'): {f}")

                if not hero_checked:
                    hero_checked = True
                    full_hero_https = clean_src if clean_src.startswith("https://") else f"{base_url}/{clean_src.lstrip('/')}"
                    if img.get("fetchpriority", "").lower() != "high" or img.get("loading", "").lower() == "lazy":
                        violations["G038"].append(f"LCP Hero görselinde fetchpriority='high' eksik veya lazy aktif ('{src}'): {f}")
                    if clean_src not in preload_norm and full_hero_https not in preload_norm:
                        violations["G038"].append(
                            f"<head> içinde LCP Hero için <link rel='preload' as='image' href='{clean_src}'> eksik: {f}"
                        )
                else:
                    if img.get("loading", "").lower() != "lazy" or img.get("decoding", "").lower() != "async":
                        violations["G039"].append(f"Ekran altı görselde loading='lazy' decoding='async' eksik ('{src}'): {f}")

        # G041 - G050: JSON-LD @graph, Knowledge Graph, ImageObject, Author & Dataset
        if len(parser.jsonld_blocks) != 1:
            violations["G041"].append(f"Sayfada tek bir birleşik JSON-LD bloğu olmalıdır (bulunan={len(parser.jsonld_blocks)}): {f}")

        has_bc_schema = False
        has_img_lic = False
        has_about_mentions = False
        has_org_rich = False
        has_author = False
        has_dataset = False
        schema_date_mod = ""

        for block in parser.jsonld_blocks:
            try:
                data = json.loads(block)
            except Exception as exc:
                violations["G041"].append(f"Bozuk JSON-LD ({exc}): {f}")
                continue
            if not isinstance(data, dict) or "@graph" not in data or not isinstance(data["@graph"], list):
                violations["G041"].append(f"JSON-LD kök yapısı '@graph' dizisi olmalıdır: {f}")
                continue

            ids_found = [n.get("@id", "") for n in iter_schema_nodes(data) if isinstance(n.get("@id"), str)]
            for req_id in ("#organization", "#website", "#webpage", "#primaryimage"):
                if not any(req_id in i for i in ids_found):
                    violations["G042"].append(f"@graph içinde '{req_id}' düğümü eksik: {f}")

            for node in iter_schema_nodes(data):
                raw_type = node.get("@type", "")
                types = set(raw_type) if isinstance(raw_type, list) else {str(raw_type)}

                if "dateModified" in node and isinstance(node["dateModified"], str):
                    schema_date_mod = node["dateModified"][:10]

                if "#organization" in str(node.get("@id", "")):
                    sa = node.get("sameAs", [])
                    geo_n = node.get("geo", {})
                    addr_n = node.get("address", {})
                    tel_n = str(node.get("telephone", ""))
                    if (
                        isinstance(sa, list)
                        and sa
                        and isinstance(geo_n, dict)
                        and geo_n.get("latitude")
                        and isinstance(addr_n, dict)
                        and addr_n.get("streetAddress")
                        and E164_PHONE_RE.match(tel_n)
                    ):
                        has_org_rich = True

                if node.get("about") and node.get("mentions"):
                    has_about_mentions = True

                if "ImageObject" in types:
                    if all(node.get(k) for k in ("contentUrl", "creditText", "copyrightNotice", "acquireLicensePage", "creator")):
                        has_img_lic = True

                if "Person" in types and node.get("name") and node.get("jobTitle") and node.get("knowsAbout"):
                    has_author = True

                if "Dataset" in types:
                    ds_desc = str(node.get("description", "")).strip()
                    if node.get("name") and len(ds_desc) >= 50 and node.get("creator") and node.get("license"):
                        has_dataset = True

                if "BreadcrumbList" in types:
                    has_bc_schema = True
                    items = node.get("itemListElement", [])
                    if len(items) < 2:
                        violations["G046"].append(f"BreadcrumbList en az 2 öğe içermelidir: {f}")
                    for li in items:
                        iu = li.get("item", "")
                        if isinstance(iu, dict):
                            iu = iu.get("@id", "")
                        if not str(iu).startswith("https://"):
                            violations["G046"].append(f"Breadcrumb 'item' mutlak HTTPS olmalıdır: {f}")

                if "Product" in types:
                    offers = node.get("offers", {})
                    if isinstance(offers, list):
                        offers = offers[0] if offers else {}
                    pv = offers.get("price") or offers.get("lowPrice")
                    if pv is None or str(pv).strip() == "" or str(pv) not in full_visible_text:
                        violations["G047"].append(f"Fiyatsız veya görünmez fiyatlı yasaklı Product şeması: {f}")

                if types & {"Organization", "LocalBusiness", "ProfessionalService", "Store", "MedicalClinic"}:
                    if "aggregateRating" in node:
                        violations["G047"].append(f"Kurum şemasında yasaklı self-serving aggregateRating: {f}")

                for img_f in ("image", "logo", "contentUrl"):
                    val = node.get(img_f)
                    if isinstance(val, str) and val.lower().endswith(".svg"):
                        violations["G047"].append(f"JSON-LD '{img_f}' alanında yasaklı .svg: {f}")

                if "FAQPage" in types:
                    for q in node.get("mainEntity", []):
                        qn = (q.get("name") or "").strip()
                        qa = ((q.get("acceptedAnswer") or {}).get("text") or "").strip()
                        if (qn and qn not in full_visible_text) or (qa and qa[:40] not in full_visible_text):
                            violations["G048"].append(f"Görünür metinde karşılığı olmayan gizli FAQ ('{qn}'): {f}")

        if not has_org_rich:
            violations["G043"].append(f"#organization içinde sameAs, address, GeoCoordinates veya E.164 telephone eksik: {f}")
        if not has_about_mentions:
            violations["G044"].append(f"@graph içinde 'about' ve 'mentions' varlık çivilemesi eksik: {f}")
        if not has_img_lic:
            violations["G045"].append(f"ImageObject telif alanları (contentUrl/creditText/copyrightNotice/acquireLicensePage/creator) eksik: {f}")
        if not is_home and (not has_bc_schema or not parser.has_breadcrumb_nav):
            violations["G046"].append(f"Alt sayfada BreadcrumbList şeması veya <nav aria-label='Breadcrumb'> eksik: {f}")
        if not has_author or not parser.time_datetimes or (parser.time_datetimes[0][:10] != schema_date_mod):
            violations["G049"].append(f"Person (#author) eksik veya <time datetime> ile JSON-LD dateModified uyumsuz: {f}")
        if parser.has_table and not has_dataset:
            violations["G050"].append(f"Tablo barındıran sayfada @type='Dataset' (description>=50 krk, creator, license) eksik: {f}")

        # G051 - G059: Cevap Kapsülü, Query Fan-Out, #id Çapaları, TOC ve GEO Kanıt Üçlüsü
        capsule_val = parser.answer_capsules[0] if parser.answer_capsules else ""
        if not capsule_val or not (25 <= len(capsule_val.split()) <= 80):
            violations["G051"].append(f"25-80 kelimelik <p data-answer-capsule='true'> eksik: {f}")
        elif VAGUE_PRONOUN_START_RE.match(capsule_val):
            violations["G052"].append(f"Cevap Kapsülü belirsiz zamirle başlayamaz ('{capsule_val.split()[0]}'): {f}")

        if len(parser.h2_items) < 3:
            violations["G053"].append(f"En az 3 adet <h2> alt başlığı gereklidir (bulunan={len(parser.h2_items)}): {f}")

        h2_ids_seen: Set[str] = set()
        for h2_id, h2_txt in parser.h2_items:
            if not h2_id or not VALID_SLUG_SEGMENT_RE.match(h2_id) or h2_id in h2_ids_seen:
                violations["G054"].append(f"<h2> başlığında benzersiz küçük harf tireli id eksik ('{h2_txt}'): {f}")
            h2_ids_seen.add(h2_id)

        if not parser.has_toc_nav or not (set(parser.toc_anchors) & h2_ids_seen):
            violations["G055"].append(f"<h2> id çapalarına bağlanan <nav data-toc='true'> İçindekiler bloğu eksik: {f}")

        if len(parser.h2_lead_paragraphs) < min(len(parser.h2_items), 3) or any(len(p.split()) > 180 for p in parser.paragraphs):
            violations["G056"].append(f"<h2> sonrası atomik paragraf eksik veya >180 kelimelik duvar paragraf var: {f}")

        if len(QUANT_PROOF_RE.findall(full_visible_text)) < 3:
            violations["G057"].append(f"Sayfada en az 3 somut sayısal veri/birim kanıtı bulunmalıdır: {f}")
        if not (parser.has_table and parser.has_table_header):
            violations["G058"].append(f"Başlık hücreli (<th>) yapılandırılmış <table> matrisi eksik: {f}")
        if not AUTHORITY_REF_RE.search(full_visible_text):
            violations["G059"].append(f"Doğrulanabilir standart/mevzuat/sertifika referansı eksik: {f}")

        # G067, G068, G069, G093: Markdown İkizi, Atom Linki ve WebMCP
        expected_md_rel = url_path_to_md_relpath(url_path)
        expected_md_url = f"{base_url}/{expected_md_rel}"
        if not (target_dir / expected_md_rel).exists():
            violations["G067"].append(f"Fiziksel Markdown ikizi ('{expected_md_rel}') eksik: {f}")
        if parser.md_alternate not in (expected_md_url, f"/{expected_md_rel}"):
            violations["G068"].append(f"<link rel='alternate' type='text/markdown' href='{expected_md_url}'> eksik: {f}")
        if parser.atom_alternate not in (f"{base_url}/feed.xml", "/feed.xml"):
            violations["G093"].append(f"<head> içinde <link rel='alternate' type='application/atom+xml' href='{base_url}/feed.xml'> eksik: {f}")

        if parser.has_webmcp_declarative or parser.has_webmcp_imperative:
            any_webmcp_on_site = True

        # G073 - G077: NavBoost, Pasif Backlink Mıknatısı, Bağlamsal CTA, Tel ve Dönüşüm İzleme
        if not parser.has_interactive_widget:
            violations["G073"].append(f"data-seo-interactive='true' mikro-araç eksik: {f}")
        if not parser.has_backlink_magnet or not any(expected_url in m for m in parser.magnet_texts):
            violations["G074"].append(
                f"Kanonik kaynak URL ('{expected_url}') içeren data-backlink-magnet='true' kaynak paylaşım kutusu eksik: {f}"
            )
        if not (parser.has_cro_cta and parser.cro_cta_contextual):
            violations["G075"].append(f"Sayfa bağlamını taşıyan data-cro-cta='true' dönüşüm butonu eksik: {f}")
        if not parser.has_tel_link:
            violations["G076"].append(f"E.164 formatında tıklanabilir href='tel:+90...' bağlantısı eksik: {f}")
        if not parser.has_conversion_hook:
            violations["G077"].append(f"Dönüşüm olayı izleme kancası (data-conversion-event) eksik: {f}")

        # G081: Speculation Rules
        if not parser.speculation_blocks:
            violations["G081"].append(f"<script type='speculationrules'> eksik: {f}")
        else:
            for sp in parser.speculation_blocks:
                try:
                    sp_obj = json.loads(sp)
                    if "prerender" not in sp_obj and "prefetch" not in sp_obj:
                        violations["G081"].append(f"speculationrules içinde prerender/prefetch eksik: {f}")
                except Exception as exc:
                    violations["G081"].append(f"Bozuk speculationrules JSON ({exc}): {f}")

        semantic_docs.append(
            {
                "url": expected_url,
                "slug": url_path.replace("/", " ").replace("-", " "),
                "title": clean_title,
                "h1": h1_val,
                "capsule": capsule_val,
                "h2_joined": " ".join(txt for _, txt in parser.h2_items),
                "main_body": " ".join(parser.main_text_tokens),
            }
        )

    # Site geneli WebMCP & UCP (G069)
    if not any_webmcp_on_site:
        violations["G069"].append("Sitede Chrome WebMCP (<form toolname='...' tooldescription='...'>) aracı bulunamadı.")
    if enable_ucp and not (target_dir / ".well-known" / "ucp").exists():
        violations["G069"].append("enableUcpCommerce=true olmasına rağmen /.well-known/ucp bulunamadı.")

    # G071 & G072: BOFU Para Sayfası Hunisi
    mp_cfg = cfg_obj.get("moneyPages", {}) if isinstance(cfg_obj, dict) else {}
    srv_paths = set(mp_cfg.get("serviceOrCategoryPaths", []))
    prc_paths = set(mp_cfg.get("pricingOrCalculatorPaths", []))
    cmp_paths = set(mp_cfg.get("comparisonOrProofPaths", []))
    if not (srv_paths & indexable_paths_set):
        violations["G071"].append("BOFU doğrudan ticari hizmet/kategori sayfası tanımlı veya yayında değil.")
    if len(indexable_paths_set) >= 3:
        if not (prc_paths & indexable_paths_set) or not (cmp_paths & indexable_paths_set):
            violations["G072"].append(
                "3+ sayfalı projede hem fiyat/maliyet hesaplama hem karşılaştırma/alternatif BOFU sayfası yayında olmalıdır."
            )

    # G017, G018, G080: Yetim sayfa, BFS Tıklama Derinliği <= 3 ve Hub geri bağı
    home_url = base_url + "/"
    depths: Dict[str, int] = {home_url: 0}
    queue = collections.deque([home_url])
    while queue:
        curr = queue.popleft()
        for nxt in adj_graph.get(curr, set()):
            if nxt in indexable_urls and nxt not in depths:
                depths[nxt] = depths[curr] + 1
                queue.append(nxt)

    for idx_url in sorted(indexable_urls):
        if idx_url != home_url:
            if inbound_links.get(idx_url, 0) == 0:
                violations["G017"].append(f"Yetim sayfa (0 dahili gelen link): {idx_url}")
            if idx_url not in depths or depths[idx_url] > 3:
                violations["G018"].append(f"Tıklama derinliği > 3 veya ana sayfadan kopuk: {idx_url}")
            if not adj_graph.get(idx_url):
                violations["G080"].append(f"Alt sayfa hiçbir Hub/Ana sayfaya geri bağlanmıyor: {idx_url}")

    # G060, G078, G079: 4-Gram Özgünlük, TF-IDF Yamyamlaşma ve siteRadius
    can_errs, drift_errs, shingle_errs = detect_cannibalization_drift_and_shingles(
        semantic_docs,
        core_keywords=core_keywords,
        brand_name=brand_name,
        cannibal_threshold=cannibal_threshold,
        shingle_threshold=shingle_threshold,
    )
    violations["G078"].extend(can_errs)
    violations["G079"].extend(drift_errs)
    violations["G060"].extend(shingle_errs)

    # G040, G091, G092: sitemap.xml + görsel sitemap
    sitemap_path = target_dir / "sitemap.xml"
    if not sitemap_path.exists():
        violations["G091"].append("sitemap.xml bulunamadı.")
        violations["G092"].append("sitemap.xml bulunamadığı için URL senkronizasyonu doğrulanamadı.")
        violations["G040"].append("sitemap.xml bulunamadığı için görsel haritası doğrulanamadı.")
    else:
        try:
            tree = ET.parse(sitemap_path)
            root = tree.getroot()
            sm_locs: Set[str] = set()
            urls_with_img: Set[str] = set()
            for url_el in root.findall(".//{*}url"):
                loc_el = url_el.find("{*}loc")
                lastmod_el = url_el.find("{*}lastmod")
                if loc_el is None or not loc_el.text:
                    continue
                u_txt = loc_el.text.strip()
                sm_locs.add(u_txt)
                if lastmod_el is None or not lastmod_el.text or not re.match(r"^\d{4}-\d{2}-\d{2}", lastmod_el.text.strip()):
                    violations["G092"].append(f"Geçersiz/eksik <lastmod>: {u_txt}")
                img_locs = [ie.text.strip() for ie in url_el.findall(".//{*}image/{*}loc") if ie.text]
                if img_locs:
                    urls_with_img.add(u_txt)
                for il in img_locs:
                    if il.lower().endswith(".svg"):
                        violations["G040"].append(f"sitemap.xml <image:loc> içinde yasaklı .svg: {il}")

            for mu in sorted(indexable_urls - sm_locs):
                violations["G092"].append(f"sitemap.xml içinde eksik URL: {mu}")
            for eu in sorted(sm_locs - indexable_urls):
                violations["G092"].append(f"sitemap.xml içinde fazladan/noindex URL: {eu}")
            for ui in sorted(indexable_urls - urls_with_img):
                violations["G040"].append(f"sitemap.xml içinde <image:image> eksik: {ui}")
        except Exception as exc:
            violations["G091"].append(f"sitemap.xml XML ayrıştırma hatası: {exc}")

    # G093, G094, G095: Atom Feed, IndexNow ve security.txt
    feed_path = target_dir / "feed.xml"
    if not feed_path.exists():
        violations["G093"].append("Hedef dizinde Atom 1.0 keşif akışı (/feed.xml) bulunamadı.")
    else:
        try:
            ET.parse(feed_path)
        except Exception as exc:
            violations["G093"].append(f"feed.xml XML hatası: {exc}")

    if not indexnow_key or not ((target_dir / f"{indexnow_key}.txt").exists() or (project_root / f"{indexnow_key}.txt").exists()):
        violations["G094"].append("IndexNow doğrulama anahtarı ({indexNowKey}.txt) bulunamadı.")

    sec_txt_path = target_dir / ".well-known" / "security.txt"
    if not sec_txt_path.exists():
        violations["G095"].append("/.well-known/security.txt bulunamadı.")
    else:
        st_content = sec_txt_path.read_text(encoding="utf-8", errors="ignore")
        if "Contact:" not in st_content or "Expires:" not in st_content:
            violations["G095"].append("security.txt içinde Contact: veya Expires: alanı eksik.")

    # G061, G062, G063: robots.txt
    robots_path = target_dir / "robots.txt"
    if not robots_path.exists():
        violations["G061"].append("robots.txt bulunamadı.")
        violations["G062"].append("robots.txt bulunamadığı için AI arama botu izinleri doğrulanamadı.")
        violations["G063"].append("robots.txt bulunamadığı için eğitim botu kuralları doğrulanamadı.")
    else:
        r_txt = robots_path.read_text(encoding="utf-8", errors="ignore")
        if f"Sitemap: {base_url}/sitemap.xml" not in r_txt:
            violations["G061"].append(f"robots.txt içinde 'Sitemap: {base_url}/sitemap.xml' eksik.")
        for sbot in AI_SEARCH_BOTS_REQUIRED:
            if f"User-agent: {sbot}" not in r_txt:
                violations["G062"].append(f"robots.txt içinde zorunlu AI arama botu eksik: {sbot}")
        for tbot in AI_TRAINING_BOTS:
            if f"User-agent: {tbot}" not in r_txt:
                violations["G063"].append(f"robots.txt içinde eğitim botu politikası eksik: {tbot}")

    # G064, G065, G066: /llms.txt ve /llms-full.txt
    llms_path = target_dir / "llms.txt"
    llms_full_path = target_dir / "llms-full.txt"
    if not llms_path.exists():
        violations["G064"].append("/llms.txt bulunamadı.")
        violations["G065"].append("/llms.txt bulunamadığı için tavsiye matrisi doğrulanamadı.")
    else:
        l_txt = llms_path.read_text(encoding="utf-8", errors="ignore")
        l_lines = [ln.strip() for ln in l_txt.splitlines() if ln.strip()]
        if not l_lines or not l_lines[0].startswith("# ") or not any(ln.startswith("> ") for ln in l_lines[:5]):
            violations["G064"].append("llms.txt '# Marka' ve '> Kurumsal özet' yapısına uymuyor.")
        if ("Önerilmeli" not in l_txt and "Recommend" not in l_txt) or "Kapsam" not in l_txt or ".md)" not in l_txt:
            violations["G065"].append("llms.txt içinde Tavsiye Kriterleri, Kapsam Sınırları veya .md ikiz linkleri eksik.")

    if not llms_full_path.exists():
        violations["G066"].append("/llms-full.txt tam bilgi korpusu bulunamadı.")
    else:
        lf_txt = llms_full_path.read_text(encoding="utf-8", errors="ignore")
        for idx_u in sorted(indexable_urls):
            if idx_u not in lf_txt:
                violations["G066"].append(f"llms-full.txt içinde sayfa eksik: {idx_u}")

    # G070, G087, G088, G089, G090: Gerçek Headless Chromium Denetimi
    browser_engine = "skipped"
    if not skip_browser and ssr_word_counts:
        browser_engine, b_errs = run_browser_render_audit(target_dir, ssr_word_counts)
        for gk, errs in b_errs.items():
            violations[gk].extend(errs)

    # G097, G098, G099: GSC 3-Vektörlü Aksiyon Kapanışı
    actions_file = project_root / "seo_actions.json"
    if actions_file.exists():
        try:
            act_data = json.loads(actions_file.read_text(encoding="utf-8"))
            if [x for x in act_data.get("queryCannibalization", []) if not x.get("resolved", False)]:
                violations["G097"].append("seo_actions.json içinde kapatılmamış sorgu yamyamlaşması görevleri var.")
            if [x for x in act_data.get("lowCtrRewrites", []) if not x.get("resolved", False)]:
                violations["G098"].append("seo_actions.json içinde kapatılmamış düşük CTR başlık görevleri var.")
            if [x for x in act_data.get("strikingDistanceInternalLinks", []) if not x.get("resolved", False)]:
                violations["G099"].append("seo_actions.json içinde kapatılmamış pozisyon 8-20 iç link görevleri var.")
        except Exception as exc:
            violations["G097"].append(f"seo_actions.json okuma hatası: {exc}")

    # G100: Sunucu Bot Log Kapanışı
    log_file = project_root / "crawl_budget_report.json"
    if log_file.exists():
        try:
            l_rep = json.loads(log_file.read_text(encoding="utf-8"))
            if (l_rep.get("botHttpErrors") or l_rep.get("uncrawledSitemapUrls")) and not l_rep.get("resolved", False):
                violations["G100"].append("crawl_budget_report.json içinde çözülmemiş bot hatası veya taranmayan sayfa var.")
        except Exception as exc:
            violations["G100"].append(f"crawl_budget_report.json okuma hatası: {exc}")

    passed_gates = sum(1 for g in GATES_100 if not violations[g])
    score = passed_gates  # 100 Atomik Kapı x 1 Puan = Tam 100/100

    if as_json:
        print(
            json.dumps(
                {
                    "contractVersion": "REAL-SEO-2026-APEX-v6.0-ABSOLUTE",
                    "score": score,
                    "maxScore": 100,
                    "passedGates": passed_gates,
                    "totalGates": 100,
                    "accepted": score == 100,
                    "browserEngine": browser_engine,
                    "indexablePages": len(indexable_urls),
                    "failedGates": {
                        g: {"title": GATES_100[g][0], "remediation": GATES_100[g][1], "errors": violations[g]}
                        for g in GATES_100
                        if violations[g]
                    },
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    else:
        print("=" * 96)
        print(
            f"REAL-SEO-2026 (APEX v6.0 ABSOLUTE) 100-KAPI KARNESİ | SAYFA: {len(indexable_urls)} | MOTOR: {browser_engine}"
        )
        print("=" * 96)
        failed_list = [g for g in GATES_100 if violations[g]]
        if not failed_list:
            print("  [G001..G100] TÜM 100 ATOMİK MÜHENDİSLİK KAPISI GEÇİLDİ (+100 / 100 PUAN — 0 İHLAL)")
        else:
            print(f"  GEÇEN KAPI: {passed_gates}/100 | BAŞARISIZ KAPI: {len(failed_list)}/100")
            print("-" * 96)
            for g in failed_list:
                title, fix_hint = GATES_100[g]
                errs = violations[g]
                print(f"  {g} [KALDI 0p] — {title} ({len(errs)} ihlal)")
                for e in errs[:3]:
                    print(f"       -> İHLAL: {e}")
                if len(errs) > 3:
                    print(f"       -> ... ve {len(errs) - 3} ek ihlal.")
                print(f"       -> [ÇÖZÜM REÇETESİ]: {fix_hint}")
        print("=" * 96)
        if score == 100:
            print("SKOR: 100/100 — ABSOLUTE SÖZLEŞMESİ KABUL EDİLDİ (100/100 KAPI, GERÇEK TARAYICI ONAYLI)")
        else:
            print(
                f"SKOR: {score}/100 — KESİNLİKLE REDDEDİLDİ! (99/100 DAHİ KABUL EDİLEMEZ — TÜM [ÇÖZÜM REÇETESİ] ADIMLARINI UYGULAYIN)"
            )
        print("=" * 96)

    return 0 if score == 100 else 1


def ping_indexnow(target_dir: Path, project_root: Path) -> int:
    cfg_path = project_root / "seo.config.json"
    sitemap_path = target_dir / "sitemap.xml"
    if not cfg_path.exists() or not sitemap_path.exists():
        print("HATA: seo.config.json ve sitemap.xml mevcut olmalıdır.")
        return 1
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    domain = cfg["domain"].rstrip("/")
    host = urlparse(domain).netloc
    key = cfg["indexNowKey"]
    tree = ET.parse(sitemap_path)
    urls = [el.text.strip() for el in tree.getroot().findall(".//{*}url/{*}loc") if el.text]
    payload = json.dumps({"host": host, "key": key, "keyLocation": f"{domain}/{key}.txt", "urlList": urls}).encode("utf-8")
    req = urllib.request.Request(
        "https://api.indexnow.org/indexnow",
        data=payload,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        print(f"[INDEXNOW] {len(urls)} URL bildirildi -> HTTP {resp.status}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="REAL-SEO-2026 APEX v6.0 ABSOLUTE — 100 Atomik Kapılı Otonom Motor")
    ap.add_argument("--init", action="store_true", help="Antigravity, Cursor, Claude, Copilot ve CI kilitlerini kurar")
    ap.add_argument("--domain", default="https://ornekdomain.com", help="Kanonik HTTPS domain")
    ap.add_argument("--build-ai-pack", metavar="HEDEF_DIZIN", help=".md ikizleri, llms.txt, llms-full.txt, feed.xml, sitemap.xml, robots.txt, 404.html ve security.txt üretir")
    ap.add_argument("--audit", nargs="*", metavar=("HEDEF_DIZIN", "DOMAIN"), help="100 atomik kapılı 100/100 denetimi çalıştırır")
    ap.add_argument("--skip-browser", action="store_true", help="Tarayıcısız minimal ortamlarda render kapısını atlar")
    ap.add_argument("--json", action="store_true", help="JSON denetim çıktısı üretir")
    ap.add_argument("--gsc-file", metavar="CSV_VEYA_JSON", help="GSC dosyasından 3-vektörlü aksiyon planı üretir")
    ap.add_argument("--gsc-api", action="store_true", help="GSC API üzerinden doğrudan veri çeker")
    ap.add_argument("--site-url", help="GSC mülk kimliği")
    ap.add_argument("--token", help="GSC OAuth2 Bearer anahtarı")
    ap.add_argument("--days", type=int, default=28, help="GSC sorgu penceresi")
    ap.add_argument("--build-dir", default=".", help="Derleme dizini")
    ap.add_argument("--log-audit", nargs=3, metavar=("LOG_DOSYASI", "HEDEF_DIZIN", "DOMAIN"), help="Sunucu bot log analizi")
    ap.add_argument("--ping-indexnow", metavar="HEDEF_DIZIN", help="IndexNow bildirimi gönderir")

    args = ap.parse_args()
    project_root = Path.cwd()

    if args.init:
        return bootstrap_project(project_root, args.domain)
    if args.build_ai_pack:
        return build_ai_discovery_pack(Path(args.build_ai_pack).resolve(), project_root)
    if args.gsc_file:
        return generate_gsc_optimization_plan(load_gsc_csv_rows(Path(args.gsc_file).resolve()), project_root)
    if args.gsc_api:
        token = args.token or os.environ.get("GSC_OAUTH_TOKEN", "")
        if not args.site_url or not token:
            print("HATA: --gsc-api için --site-url ve --token gereklidir.")
            return 2
        return generate_gsc_optimization_plan(fetch_gsc_api_rows(args.site_url, token, days=args.days), project_root)
    if args.log_audit:
        lp, bd, dm = args.log_audit
        return audit_server_logs(Path(lp).resolve(), Path(bd).resolve(), dm, project_root)
    if args.audit is not None:
        cfg_file = project_root / "seo.config.json"
        default_domain = "https://ornekdomain.com"
        if cfg_file.exists():
            try:
                default_domain = json.loads(cfg_file.read_text(encoding="utf-8")).get("domain", default_domain)
            except Exception:
                pass
        target = Path(args.audit[0]) if len(args.audit) >= 1 else project_root
        domain = args.audit[1] if len(args.audit) >= 2 else default_domain
        return audit_project(target.resolve(), domain, project_root, as_json=args.json, skip_browser=args.skip_browser)
    if args.ping_indexnow:
        return ping_indexnow(Path(args.ping_indexnow).resolve(), project_root)

    ap.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
