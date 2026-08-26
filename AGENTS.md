## Mandatory reading order

Before making any user-facing UI, CSS, layout, responsive or component change:

1. Read `DESIGN.md` completely.
2. Read the target page/component and the relevant tokens in `src/styles/global.css`.
3. Check `scripts/ci/protected-surfaces.mjs` before touching a protected surface.
4. Preserve route, schema, SEO, commerce and product-data contracts unless the task explicitly changes them.

`DESIGN.md` is the canonical visual contract. `DNA.txt` contains historical project context but must not override current v5 runtime tokens or `DESIGN.md` when they differ.

## UI execution rules

- Reuse existing components and design tokens before creating new ones.
- Do not invent random colors, fonts, spacing, radius or component styles inside page markup.
- 21st.dev may be used as composition/interaction inspiration only; do not import a React/Next runtime or add a visual dependency solely to copy a reference.
- New/edited surfaces must be responsive from 320px through large desktop and must not rely on page-level overflow masking to hide layout defects.
- Dense tables/data may scroll only inside an intentional local wrapper; the page itself must not require horizontal scrolling.
- Preserve Turkish visible copy, accessibility, reduced motion, real-product proof and commercial truth.
- Do not weaken protected-surface, SEO, commerce or source-language guards to make a build pass.

## Required validation

For UI work run, at minimum:

```bash
npm run guard:design
npm run build
npm test
```

A visual change is incomplete if it introduces clipping, horizontal overflow, broken links, invalid product/commerce behavior, SEO regressions or protected-surface drift.

## Development

When starting the dev server, use background mode:

```bash
astro dev --background
```

Manage the background server with `astro dev stop`, `astro dev status`, and `astro dev logs`.

## Documentation

Full documentation: https://docs.astro.build

Consult these guides before working on related tasks:

- [Adding pages, dynamic routes, or middleware](https://docs.astro.build/en/guides/routing/)
- [Working with Astro components](https://docs.astro.build/en/basics/astro-components/)
- [Using React, Vue, Svelte, or other framework components](https://docs.astro.build/en/guides/framework-components/)
- [Adding or managing content](https://docs.astro.build/en/guides/content-collections/)
- [Adding styles or using Tailwind](https://docs.astro.build/en/guides/styling/)
- [Supporting multiple languages](https://docs.astro.build/en/guides/internationalization/)
