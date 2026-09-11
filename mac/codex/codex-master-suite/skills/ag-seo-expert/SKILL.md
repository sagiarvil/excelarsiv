---
name: ag-seo-expert
description: Technical SEO, GEO, AEO and Engine V3.0 Orchestrator. Executes 18-engine live crawls, 14KB AST budgets, and drop-in repairs.
---

# TECHNICAL SEO, GEO & ENGINE V3.0 MASTER SKILL

> [!IMPORTANT]
> **CRITICAL USER OUTPUT MANDATE:** You MUST output all user-facing chat responses, explanations, progress reports, `toolAction`, and `toolSummary` strictly in **100% Turkish**. Never output English to the user.

This skill orchestrates deterministic Technical SEO, Generative Engine Optimization (GEO), and Answer Engine Optimization (AEO).

---

## 🚀 1. ENGINE V3.0 AUDIT PROTOCOL
1. **Live URL Inspection:** Always crawl the live HTTP target using `audit_engine_v3.js` and `curl`.
2. **18 Engines Evaluated:** Title, Meta, Canonical, Hreflang, OpenGraph, Schema/JSON-LD, Content AST (14KB budget), Mobile CWV, Robots, Sitemap, Internal Links, AEO/GEO Consensus Triples, LLMO, ColBERT MaxSim, HTTP/SSL, Media Alt, A11y, and Performance.
3. **P0-P3 Priority Matrix:** Remediate critical index-blocking issues (P0) first, followed by structural defects (P1).

---

## 🛠️ 2. CANONICAL SEO TOOLS REPOSITORY
All tools reside under: `$HOME/.gemini\config\plugins\gemini-master-suite\skills\ag-seo-expert\tools\`
- `audit_engine_v3.js`: 18-Engine live scanner.
- `schema_validator.js`: JSON-LD validator.
- `heading_semantics_check.js`: H1-H6 hierarchy validator.
- `sitemap_robots_check.js`: Robots.txt and sitemap.xml validator.
- `bom_utf8_scan.js --fix`: Automated BOM cleaner.
