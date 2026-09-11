# HTML & DOM ERROR RESOLUTIONS & ACCESSIBILITY CHECKLISTS

This document outlines deterministic procedures for repairing broken markup, layout shift, and accessibility violations.

## 1. UNCLOSED / MISMATCHED DOM TAG REPAIR
- **Symptom:** Broken grid alignments, collapsed footers, or leaked container bounds.
- **Resolution Procedure:**
  1. Inspect indentation levels hierarchically from root to leaf.
  2. Map each opening element to its closing pair.
  3. Ensure template loop directives (`@foreach`, `while`) encapsulate balanced opening and closing tags.

## 2. PREVENTING CUMULATIVE LAYOUT SHIFT (CLS)
```html
<!-- CORRECT RESPONSIVE IMAGE -->
<img src="hero.webp" alt="Cloud Infrastructure Overview" width="1280" height="720" loading="lazy" decoding="async" class="w-100 h-auto">
```

## 3. ACCESSIBILITY (WCAG 2.2 AA) PROTOCOL
- All interactive controls must be native `<button>` or `<a>` elements (never `<div onclick>`).
- Ensure full keyboard focusability and minimum 4.5:1 color contrast for standard text.
