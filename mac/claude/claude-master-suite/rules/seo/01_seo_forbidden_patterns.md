# SEO EXPERT — STRICTLY FORBIDDEN ANTI-PATTERNS

This standard outlines non-negotiable prohibitions for Technical SEO, GEO, and AEO engineering.

## 1. STRICTLY FORBIDDEN SEO PATTERNS
1. **Auditing Without Live URL:** Performing speculative local audits without testing the live HTTP target is FORBIDDEN.
2. **Multiple H1 Tags:** Using more than one `<h1>` tag per page is STRICTLY FORBIDDEN.
3. **Keyword Stuffing & Cloaking:** Artificially cramming keywords into titles or content, or serving different DOM trees to bots vs users, is BANNED.
4. **Missing or Broken Canonical Links:** Generating non-self-referencing or conflicting `<link rel="canonical">` tags is FORBIDDEN.
5. **Leftover Development Noindex Tags:** Deploying `<meta name="robots" content="noindex, nofollow">` to production is a critical failure.
6. **14KB AST Budget Violation:** Structuring the HTML such that primary content and schema require more than 14KB in the initial payload is BANNED.
7. **Missing Alt Attributes & Broken JSON-LD:** Omitting descriptive image `alt` tags or emitting invalid JSON-LD syntax is FORBIDDEN.
