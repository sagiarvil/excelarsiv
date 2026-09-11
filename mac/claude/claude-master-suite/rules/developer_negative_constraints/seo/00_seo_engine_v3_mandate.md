# ENGINE V3.0 18-ENGINE TECHNICAL SEO/GEO/AEO MANDATE

This mandate defines the execution protocol for search engine, generative engine (GEO), and answer engine (AEO) audits.

## 1. CORE OPERATIONAL INVARIANTS
1. **Live URL Auditing:** All scans MUST execute against the target live website URL. If the live URL is unknown, query the user directly.
2. **18-Engine V3 Standard:** The audit must evaluate all 18 engines: Title, Meta, Canonical, Hreflang, OpenGraph, Schema/JSON-LD, Content AST (14KB budget), Mobile CWV, Robots, Sitemap, Internal Links, AEO/GEO Consensus Triples, LLMO, ColBERT MaxSim, HTTP/SSL, Media Alt, A11y, and Performance.
3. **Local Remediation & Drop-in Packaging:** Discoveries must be resolved in local codebase files and packaged into the `guncelleme/` drop-in tree matching original directory paths.
