---
name: ag-html-export-expert
description: HTML & Bootstrap export specialist. Generates W3C-valid, sanitized, 100% SEO-aligned markup compatible with Engine V3, schema_validator, and heading_semantics_check.
---

# HTML & BOOTSTRAP SEO EXPORT MASTER SKILL

> [!IMPORTANT]
> **CRITICAL USER OUTPUT MANDATE:** You MUST output all user-facing chat responses, explanations, progress reports, `toolAction`, and `toolSummary` strictly in **100% Turkish**. Never output English to the user.

This skill governs the production of semantic HTML5, Bootstrap 5 templates, and standalone exports strictly aligned with automated SEO analysis tools.

---

## 🎯 1. SEO & ENGINE V3.0 EXPORT REQUIREMENTS
1. **Heading Semantics (`heading_semantics_check.js`):**
   - Exactly ONE `<h1>` tag per document.
   - Never skip heading levels (`<h2>` must precede `<h3>`).
2. **14KB AST Budget:**
   - Place metadata, OpenGraph, JSON-LD schema, and `<main>` content inside the initial 14KB crawl chunk.
3. **Schema.org Integration (`schema_validator.js`):**
   - Embed valid `<script type="application/ld+json">` with `WebSite`, `BreadcrumbList`, or `Organization`.
4. **Image Attributes:**
   - Every `<img>` must feature descriptive `alt`, `width`, `height`, `loading="lazy"`, and `decoding="async"`.

---

## 🎨 2. BOOTSTRAP 5 FRAMEWORK STANDARDS
1. **Semantic Landmark Pairing:** Combine native `<header>`, `<nav>`, `<main>`, `<section>`, `<article>`, `<aside>`, `<footer>` with Bootstrap grid classes.
2. **Direct Grid Hierarchy:** `.col-*` must be a direct child of `.row`. Never inject intermediary wrapper `<div>`s.
3. **Modern Data Attributes:** Use `data-bs-*` attributes. jQuery and external CDN links are strictly forbidden.

---

## 🧪 3. VERIFICATION PROTOCOL
After every HTML/Bootstrap export, execute:
- `node .../heading_semantics_check.js <file>`
- `node .../schema_validator.js <file>`
- `node .../bom_utf8_scan.js --fix <file>`

---

## 🇹🇷 4. TÜRKÇE KARAKTER & JSON VERİ BÜTÜNLÜĞÜ (MANDATORY P0)
When generating HTML5 and Bootstrap 5 templates:
1. **Document Header:** Always start with `<html lang="tr">` and `<meta charset="UTF-8">` as the first head tag.
2. **Native UTF-8 Characters:** All text, buttons, modals, and navigation items must feature 100% accurate Turkish letters (`ç, Ç, ğ, Ğ, ı, I, i, İ, ö, Ö, ş, Ş, ü, Ü`). Never downgrade to ASCII and never use HTML entities like `&ccedil;`.
3. **Form Submissions & AJAX:** Any JS form handler must send `Content-Type: application/json; charset=utf-8` and `Accept: application/json; charset=utf-8`.
4. **Form Regex Patterns:** Never use English-only `[a-zA-Z]` regex patterns on Turkish name/address fields. Always include `[a-zA-ZçÇğĞıİöÖşŞüÜ\s]`.
5. **No Mojibake:** Zero tolerance for legacy encoding artifacts (`Ý, Ð, Þ, ý, ð, þ`).
