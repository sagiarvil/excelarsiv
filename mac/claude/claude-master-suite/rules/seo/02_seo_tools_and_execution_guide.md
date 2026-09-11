# OFFICIAL SEO TOOLSET & EXECUTION GUIDE (TOOLS & ENGINES)

This guide documents the canonical suite of SEO analysis and validation tools available in the system.

---

## 🛠️ 1. CANONICAL SEO TOOLS REPOSITORY
**Tools Root Directory:** `$HOME/.gemini\config\plugins\gemini-master-suite\skills\ag-seo-expert\tools\`

| Tool Script | Domain & Responsibility | Execution Command |
|---|---|---|
| `audit_engine_v3.js` | **18-Engine Engine V3:** Comprehensive technical, GEO, and LLMO live scanner. | `node .../audit_engine_v3.js [TARGET_URL]` |
| `schema_validator.js` | **Structured Data Validator:** Tests Organization, Product, Breadcrumb, FAQPage JSON-LD. | `node .../schema_validator.js <path_or_url>` |
| `heading_semantics_check.js` | **Heading Hierarchy Audit:** Validates H1-H6 levels, single H1 rule, and semantic landmarks. | `node .../heading_semantics_check.js <path_or_url>` |
| `sitemap_robots_check.js` | **Robots & Sitemap Validator:** Tests robots.txt directives and sitemap.xml indexing rules. | `node .../sitemap_robots_check.js [TARGET_URL]` |
| `opengraph_check.js` | **Social Graph Validator:** Validates og:title, og:image, twitter:card, and dimensions. | `node .../opengraph_check.js <path_or_url>` |
| `internal_link_check.js` | **Link Integrity Checker:** Detects 404 dead links, orphan pages, and anchor text hygiene. | `node .../internal_link_check.js [TARGET_URL]` |
| `web_vitals_hints.js` | **Core Web Vitals Checker:** Analyzes LCP, CLS, INP asset budgets and blocking resources. | `node .../web_vitals_hints.js <path_or_url>` |
| `dom_utf8_full_test.js` | **DOM & Encoding Verifier:** Detects unclosed HTML tags and UTF-8 encoding anomalies. | `node .../dom_utf8_full_test.js <directory>` |
| `bom_utf8_scan.js` | **BOM Sanitizer:** Removes destructive Byte Order Marks from template/PHP files (`--fix`). | `node .../bom_utf8_scan.js --fix <filePath>` |
| `curl` | **HTTP & SSL Inspector:** Tests HTTP status codes (200, 301, 404), redirect hops, and SSL certs. | `curl -I -L [TARGET_URL]` |

---

## ⚙️ 2. AUDIT & REMEDIATION PIPELINE
1. **Live Discovery:** Execute `audit_engine_v3.js` and `curl` on the live domain.
2. **Prioritization Matrix (P0 to P3):**
   - **P0 (Critical):** 404s, indexation blocking, canonical conflicts, syntax-broken schema.
   - **P1 (High):** Missing H1, invalid OpenGraph, payload exceeding 14KB AST budget.
   - **P2 (Medium):** Missing image alt tags, incomplete breadcrumb trees.
   - **P3 (Optimization):** Internal link graph distribution and anchor text refinement.
3. **Local Remediation:** Patch codebase files locally and mirror to the `guncelleme/` drop-in folder.
4. **Physical Verification:** Verify repaired files using `bom_utf8_scan.js --fix` and `dom_utf8_full_test.js`.
