## Canonical SEO / GEO / LLMS mandate — mandatory

Before any SEO, GEO, AEO, LLMS, sitemap, robots, canonical, structured-data, search-content, internal-link, redirect or search-measurement change, read `SAGIARVIL_SEARCH_REVENUE_OS_MANDATE.md` completely.

For those scopes, `SAGIARVIL_SEARCH_REVENUE_OS_MANDATE.md` is the single canonical Search Revenue mandate and supersedes older SEO/GEO/LLMS instruction documents. Runtime SEO data, `robots.txt`, sitemap generators/files, `llms.txt`, `llms-full.txt`, `/llms/**`, schema code and `data/seo/**` remain operational/source assets governed by that mandate; they are not obsolete instruction documents.

User's latest explicit instruction remains highest authority. `DESIGN.md` remains the canonical visual contract and product/commerce/source-of-truth rules remain authoritative in their own domains.

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

---

## MANDATE-SEO-GEO-2026-V6 MİMARİ STANDARDI

> **ZORUNLU VE BAĞLAYICI MİMARİ STANDART:** Bu projede yapılacak tüm sayfa, şema, sitemap, routing ve içerik değişiklikleri MANDATE-SEO-GEO-2026-V6 şartnamesine %100 uymak zorundadır. Registry kaydı (`src/seo/registry.ts`) olmadan sayfa üretilemez, derin alt-graf (`/llms/*.md`) olmadan amiral gemisi rota açılamaz.

### 8 Adımlı Deterministik Uygulama Protokolü (Playbook)

1. **SSOT REGISTRY TANIMI:** `src/seo/registry.ts` dosyasına gidilir; rota, canonicalRoute, primaryIntent, primaryEntity ve modifiedAt alanları tanımlanır.
2. **HERO ANSWER ENGINE KODLAMASI:** Sayfa HTML şablonunun en üstüne (ilk 100 piksel) 80-120 kelimelik `div.hero-answer-engine` özet kutusu eklenir (sayısal veri, kesin tanım, teknik parametreler).
3. **@graph JSON-LD ŞEMA ENJEKSİYONU:** `src/seo/schema-builder.ts` çağrılarak Organization, WebSite, WebPage, BreadcrumbList ve sektörel şemalar ham HTML içine gömülür.
4. **ÇOK KATMANLI LLM BİLGİ DOKÜMANI ÜRETİMİ:** `public/llms/pages/[slug].md` oluşturulur (Saf token-yoğun Markdown, RDF üçlüleri, FAQ) ve kök `public/llms.txt` manifestosuna bağlanır.
5. **SITEMAP & ROBOTS.TXT DOĞRULAMASI:** Sayfanın sitemap.xml içerisine yalnız canonical adresiyle girdiği doğrulanır; robots.txt içinde GPTBot, ClaudeBot, PerplexityBot izinleri teyit edilir.
6. **SUNUCU / EDGE MIME-TYPE YAPILANDIRMASI:** `/llms.txt` ve `/llms/**` yollarının `text/markdown; charset=utf-8` başlığıyla döndüğü doğrulanır.
7. **MULTI-HUB INDEXNOW DAĞITIMI:** `node scripts/notify-indexnow.js https://excelarsiv.com/[slug]` komutuyla Bing, Yandex ve IndexNow API'ye anlık PUSH yapılır.
8. **CANLI PROD HEALTH CHECK & KORUMA:** `npm run seo:ci-gate` çalıştırılarak tüm G0-G9 kapılarından sıfır hatayla geçildiği onaylanır.

