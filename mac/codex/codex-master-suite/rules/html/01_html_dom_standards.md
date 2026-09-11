# SEMANTIC HTML5 & DOM INTEGRITY STANDARDS

This standard enforces valid DOM structures, modern semantics, accessibility, and visual stability.

## 1. STRICT DESIGN LOCK & TEMPLATE INTEGRITY
1. **Zero Unsolicited Restructuring:** Never rewrite, modernize, or alter working HTML templates, layouts, or stylesheets unless explicitly requested.
2. **DOM Hierarchy Preservation:** Maintain all existing container classes, element IDs, and semantic nesting.

## 2. SEMANTIC STRUCTURE & NEGATIVE CONSTRAINTS
1. **Unclosed Tags Strictly Forbidden:** Every opening tag (`<div>`, `<section>`, `<p>`, `<ul>`, `<table>`) MUST have an exactly matching closing tag.
2. **Anti-Div Soup:** Avoid generic `<div>` nesting. Use structural HTML5 elements (`<header>`, `<nav>`, `<main>`, `<article>`, `<section>`, `<aside>`, `<footer>`).
3. **Image Attributes:** Every `<img>` tag MUST have an informative `alt` attribute and explicit `width` and `height` (or CSS `aspect-ratio`) to prevent Cumulative Layout Shift (CLS).
4. **Heading Hierarchy:** Only one `<h1>` tag per page. Never skip levels (e.g., `<h2>` directly to `<h4>`).
5. **Form Controls:** Every input, select, and textarea MUST have an associated `<label for="inputId">` or `aria-label`.
6. **No Inline Styles or SVG Bloat:** Inline `style="..."` and inline `<svg>` are banned; use external stylesheets and asset files.
