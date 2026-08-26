# Excel Arşiv Design System Contract

Version: 1.1.0  
Scope: `excelarsiv.com` source UI, visual system, responsive behavior, commercial UX and future agent work  
Status: mandatory design contract

## 0. Purpose

This file is the canonical design instruction for every user-facing change in this repository. It exists to prevent page-by-page visual drift, random component invention, broken responsive behavior and generic AI/SaaS styling.

The site is not a generic template marketplace. It presents ready-to-use Excel business systems for finance, accounting and operations. The design must therefore communicate three things at the same time:

1. Excel familiarity and operational credibility.
2. Premium product quality and commercial trust.
3. Decision support: input -> calculation -> control -> management output.

`DESIGN.md` governs design decisions. Runtime values must still be implemented through the existing token and component system in `src/styles/global.css` and `src/components/`.

If an older document conflicts with the current source or this file, use this priority order:

1. User-approved commercial/product truth and protected-surface contracts.
2. This `DESIGN.md`.
3. Current runtime tokens/components in `src/styles/global.css` and `src/components/`.
4. Existing page implementation when it does not conflict with 1-3.
5. External inspiration such as 21st.dev or high-trust financial product sites.

`DNA.txt` contains valuable project history but some visual values are historical. It must not override current v5 runtime tokens or this file.

---

## 1. Design thesis

### 1.1 Core visual idea

Excel Arşiv uses a **Living Workbook / Decision Console** visual language.

The site should feel like a highly refined financial workbook translated into a modern web interface, not like a dark startup landing page and not like a generic rounded SaaS kit.

Required characteristics:

- light theme only;
- strong information hierarchy;
- workbook/grid references used with restraint;
- real product screens prioritized over decorative mockups;
- data, status and result visualization used as evidence;
- compact premium spacing rather than oversized empty SaaS sections;
- typography optimized for Turkish finance and accounting content;
- commercial CTAs visible but never visually noisy;
- no invented testimonials, fabricated KPIs or fake live indicators.

### 1.2 Brand behavior

The visual system should make the user understand the product before asking them to buy it.

Preferred narrative:

`problem -> input -> calculation -> control/risk -> management output -> proof/demo -> purchase`

Avoid:

`hero slogan -> decorative cards -> generic benefits -> buy now`

---

## 2. 21st.dev reference policy

External reference: `https://21st.dev/community/components/s/comparison`

21st.dev is an **interaction and composition reference**, not a dependency and not a source to copy blindly.

Useful reference categories include:

- comparison tables and before/after comparisons;
- pricing sections;
- stats/KPI blocks;
- dashboards;
- feature sections;
- calls to action;
- FAQs;
- cards and bento/grids;
- navigation patterns;
- scroll areas when dense data requires them.

Rules:

1. Never import a React/Next.js component only because a 21st.dev example uses it.
2. Rebuild the useful interaction pattern in Astro + existing Tailwind/CSS architecture.
3. Do not add a dependency only for appearance.
4. Do not copy colors, fonts, radius or generic SaaS styling from a reference.
5. The Excel Arşiv token system remains authoritative.
6. Comparison components must remain readable on mobile. Dense tables may scroll inside their own bounded wrapper; the page itself must never require horizontal scrolling.
7. Animation from references may be adopted only when it uses opacity/transform, respects reduced motion and does not harm LCP/CLS.

The reference is successful only when the final component looks native to Excel Arşiv.

---

## 3. Runtime design tokens

Do not invent new values inside page/component markup when an existing token can express the decision.

The current v5 token family lives in `src/styles/global.css`.

### 3.1 Primary color family

Use the existing Excel-oriented family:

- `--xl-green: #107c41` — primary Excel action/accent.
- `--xl-green-dark: #0c5a30` — hover/strong accent.
- `--xl-green-tint: #e9f5ee` — soft accent surface.
- `--xl-select: #217346` — selection/workbook reference.
- `--xl-blue: #0f6cbd` — secondary informational/data accent.

Do not introduce a second brand-green family in a component.

### 3.2 Surfaces

Use:

- `--xl-canvas` / `--surface-elevated` for raised content and workbook windows.
- `--xl-paper` / `--surface-0` for the page plane.
- `--xl-header` / `--surface-1` for secondary workbook/header surfaces.
- `--xl-head-hov` / `--surface-2` for subtle hover/selected surfaces.

A page should normally have no more than three simultaneous surface levels.

### 3.3 Text

Use the existing ink hierarchy:

- `--xl-ink` / `--ink-900`: primary headings and high-importance values.
- `--xl-ink-2` / `--ink-700`: body text.
- `--xl-ink-3` / `--ink-500`: secondary copy.
- softer aliases only where contrast remains WCAG-safe.

Do not create gray text with arbitrary opacity when an ink token exists.

### 3.4 Status colors

Status color must carry meaning, not decoration.

- positive: conditional-format green tokens;
- warning: conditional-format amber tokens;
- negative/critical: conditional-format red tokens;
- neutral: ink tokens.

A status must never rely on color alone. Pair it with text, icon, shape or state label.

### 3.5 Typography

Current runtime families:

- primary/display/body: `Manrope` through `--font-sans`, `--font-display`, `--font-body`;
- labels/data/technical cells: `IBM Plex Mono` through `--font-mono`.

Use the existing type scale:

- `--t-display` / `--type-display`;
- `--t-h1` / `--type-h1`;
- `--t-h2` / `--type-h2`;
- `--t-h3` / `--type-h3`;
- `--t-body`;
- `--t-small`;
- `--t-xs`;
- `--t-micro`.

Rules:

- one H1 per page;
- finance/data labels may use mono, body paragraphs should not;
- headings use tight tracking only through existing tokens;
- paragraphs should remain readable at 16px equivalent or greater on mobile;
- no decorative font addition.

### 3.6 Radius

Use the existing micro-radius system only:

- `--r-none: 0`;
- `--r-sm: 2px`;
- `--r-base: 4px`;
- `--r-lg: 6px`;
- `--r-xl: 6px`.

Large 16-32px SaaS pills/cards are not part of the default system. A component that needs a visual capsule must have a functional reason.

### 3.7 Spacing

Use the existing rhythm and spacing tokens before adding any value:

- `--space-1: 4px`;
- `--space-2: 8px`;
- `--space-3: 12px`;
- `--space-4: 16px`;
- `--space-6: 24px`;
- `--space-8: 32px`;
- `--space-12: 48px`;
- `--space-16: 64px`;
- `--space-24: 96px`;
- `--space-32: 128px`;
- `--space-40: 160px`.

Commercial page rhythm:

- normal section gap: `--section-gap`;
- large section gap: `--section-gap-large`;
- mobile section gap: `--section-gap-mobile`.

Avoid one-off values unless they solve a measured layout issue and cannot be represented by the token scale.

### 3.8 Containers

Use the existing container system:

- primary content: `--container-main: 1180px`;
- wide visual/data section: `--container-wide: 1320px`;
- desktop page padding: `--page-pad-desktop: 24px`;
- tablet page padding: `--page-pad-tablet: 20px`;
- mobile page padding: `--page-pad-mobile: 16px`.

Never use a fixed content width without a responsive max-width/calc boundary.

---

## 4. Layout system

### 4.1 Grid

The default structural grid is 12 columns with the existing gutter/max-width tokens.

Use asymmetry where it improves information hierarchy. Preferred compositions include:

- 4/8 or 5/7 explanatory splits;
- 3/9 technical index + content;
- 7/5 product proof + purchase decision;
- full-width proof/data scene followed by compact explanatory copy.

Do not force every section into the same three-card row.

### 4.2 Density

Excel Arşiv should be information-rich without becoming cramped.

Use compact density for:

- KPI rows;
- workbook chrome;
- comparison tables;
- technical specifications;
- pricing facts;
- input/output lists.

Use more breathing room for:

- hero headline;
- decision narrative transitions;
- product proof imagery;
- primary CTA area.

### 4.3 Visual hierarchy

One viewport should normally have one dominant decision.

Allowed emphasis order:

1. page/product name or user problem;
2. proof/result;
3. primary action;
4. technical detail;
5. supporting navigation.

Do not make every card, badge and button visually loud.

---

## 5. Component contract

Before creating a component, search `src/components/` for an existing primitive or commercial component that can be extended.

Preferred reusable vocabulary includes the existing components and these conceptual roles:

- `SiteHeader` / navigation;
- `AnnouncementBar`;
- `Button`;
- `Badge`;
- `Breadcrumbs`;
- `ExcelFrame` / workbook scene;
- product/template card;
- proof/demo box;
- eligibility list;
- spec table;
- sheet-map table;
- FAQ accordion;
- checkout/purchase panel;
- mobile commerce bar;
- footer.

New components should represent a stable product concept, not a single page decoration.

### 5.1 Product hero

A product hero must answer, above the fold where practical:

- what business problem it solves;
- what the user receives;
- one high-value result or proof;
- price/demo/purchase path;
- compatibility or trust fact when relevant.

Do not fill the hero with multiple competing CTAs.

### 5.2 Product preview / proof

Real product screenshots are preferred over mockups.

Rules:

- preserve screenshot aspect ratio;
- never stretch an Excel image;
- readable crop over decorative crop;
- include descriptive alt text;
- allow zoom/detail only when it materially helps evaluation;
- do not label a fake UI as a real product screen.

### 5.3 KPI / status card

A KPI card must include context. A naked number is not proof.

Recommended structure:

`label -> value -> direction/status -> short interpretation`

### 5.4 Comparison

Use comparison when the user must choose or understand transformation.

Preferred modes inspired by 21st.dev:

- before vs after;
- existing manual process vs Excel Arşiv system;
- product A vs product B when differences matter;
- package vs single product;
- feature table only when rows are decision-relevant.

Mobile rules:

- prefer stacked comparison cards for short comparisons;
- for dense tables, use a local scroll container with visible scroll affordance;
- keep the first label column understandable;
- do not compress text below a readable size to fit columns;
- no page-level horizontal overflow.

### 5.5 Pricing / purchase block

The purchase area must clearly communicate:

- current price;
- KDV state;
- one-time/subscription state;
- payment provider;
- demo availability when applicable;
- delivery behavior.

The primary CTA should be visually unambiguous. Secondary links must not look stronger than purchase/demo decisions.

### 5.6 FAQ

Use native semantic `details/summary` behavior unless there is a measured reason not to.

No JS accordion dependency for visual novelty.

---

## 6. Page archetypes

### 6.1 Home page

The home page should operate as a commercial command center.

Recommended narrative order:

1. search/problem-led hero;
2. immediate trust facts (demo, KDV, one-time payment, Shopier);
3. flagship systems;
4. input -> calculation -> risk -> management output explanation;
5. real product screens;
6. category/problem discovery;
7. before/after comparison;
8. purchase/delivery flow;
9. package decision;
10. footer/trust close.

Do not redesign protected home surfaces casually. Respect `scripts/ci/protected-surfaces.mjs`.

### 6.2 Catalog `/sablonlar`

Search and problem discovery have priority over decorative filtering.

Product cards should expose enough information to decide whether opening the detail page is worth it:

- name;
- problem/result;
- category;
- price;
- demo/proof availability;
- detail action.

### 6.3 Product detail

Preferred sequence:

1. breadcrumbs + product hero;
2. real visual proof;
3. problem and result;
4. inputs;
5. calculations/controls;
6. outputs/management decisions;
7. workbook/sheet map;
8. suitability/requirements;
9. demo;
10. price/purchase;
11. comparison or alternatives when useful;
12. FAQ;
13. related systems.

The page should answer "what will this change in my work?" before drowning the user in technical fields.

### 6.4 Special/custom Excel systems

Custom-system pages may have a distinct sales narrative, but they still inherit typography, color, spacing, accessibility and responsive rules from this contract.

Any explicit isolation rule in `protected-surfaces.mjs` remains binding.

---

## 7. Responsive contract — non-negotiable

Responsive quality is a release criterion, not a polishing step.

Required verification widths:

- 320px;
- 360px;
- 375px;
- 390px;
- 430px;
- 768px;
- 1024px;
- 1280px;
- 1440px.

### 7.1 Zero page-level horizontal overflow

No edited surface may require the user to scroll the page horizontally.

Rules:

- grid/flex children that contain long content must be able to shrink (`min-width: 0` / equivalent);
- use `minmax(0, 1fr)` for flexible grid tracks;
- long URLs, codes and technical labels must wrap or scroll only inside a deliberate local code/data container;
- images/video/canvas must never exceed their content box;
- a dense table may have its own horizontal scroll wrapper;
- do not use `100vw` inside a padded container; prefer `100%` or a container token;
- fixed pixel widths need a max-width/responsive fallback;
- transformed/decorative elements may not enlarge the document scroll width;
- sticky/fixed CTAs must respect safe-area insets and viewport edges.

Important legacy note: the current global stylesheet contains a historical `body { overflow-x: hidden; }` mask. Do not treat that as proof that a surface is overflow-safe. New or edited surfaces must be locally safe without relying on the mask. Removing that legacy rule requires a dedicated cross-page audit because `global.css` is protected.

### 7.2 Text

- no unreadably small mobile text;
- no heading that forces a single unbreakable line;
- preserve Turkish diacritics and words;
- use `text-wrap: balance` for headings when suitable;
- use safe wrapping for long technical strings.

### 7.3 Touch and controls

Interactive mobile targets should normally provide at least 44x44px effective hit area.

Do not place two primary actions so close that accidental taps are likely.

### 7.4 Tables and data visualization

Desktop table density must not be achieved by shrinking mobile typography.

At narrow widths choose one:

1. transform rows into stacked label/value cards;
2. reduce non-essential columns;
3. local horizontal scroll with clear edge/scroll affordance.

Do not hide decision-critical data merely to preserve a desktop layout.

---

## 8. Accessibility contract

Target WCAG 2.2 AA or better.

Required:

- semantic landmarks;
- one logical H1 and ordered heading hierarchy;
- visible `:focus-visible` state;
- keyboard-accessible controls;
- descriptive image alt text;
- form fields with programmatic labels;
- status not communicated only by color;
- reduced-motion behavior;
- sufficient text/background contrast;
- no hover-only critical information;
- skip link retained.

A premium visual that breaks keyboard or mobile use is a failed component.

---

## 9. Motion contract

Motion exists only to clarify hierarchy or state.

Allowed properties by default:

- opacity;
- transform.

Preferred duration range: approximately 120-220ms for direct UI feedback. Longer section reveal effects must be subtle and must not delay content visibility.

Forbidden as default behavior:

- layout-shifting entrance animations;
- perpetual decorative motion;
- scroll-jacking;
- large parallax effects;
- animation that makes a financial data value harder to read.

`prefers-reduced-motion` must remain respected.

---

## 10. Commercial UX contract

The design must support trust and purchase comprehension.

Every commercial surface should preserve these facts when relevant:

- demo can be examined before purchase;
- displayed price/KDV state is explicit;
- subscription vs one-time payment is explicit;
- Shopier is the payment step;
- delivery/return behavior is clear;
- compatibility and macro requirements are clear;
- real screens are labeled accurately.

Do not add fake scarcity, fake countdowns, fake sales counters or fabricated customer claims.

---

## 11. Content density and writing in UI

UI copy should be concise but specific.

Prefer:

- "6. haftada nakit açığını görün"
- "Fiili kasa ile kayıt farkını aynı gün yakalayın"
- "Formül hücreleri korunur"

Avoid:

- "işinizi bir üst seviyeye taşıyın"
- "devrim niteliğinde"
- "benzersiz çözüm"
- generic AI/SaaS marketing filler.

The interface should demonstrate the outcome instead of claiming quality.

---

## 12. Anti-patterns

Do not introduce the following without an explicit, documented reason:

- new brand color families;
- random hex values in page/component markup;
- random fonts;
- large pill-shaped SaaS cards as the default visual language;
- excessive glassmorphism;
- dark theme sections that break the light system;
- gradients used only to signal "premium";
- decorative charts with no business meaning;
- fake Excel UI labeled as real;
- infinite carousels for essential information;
- page-level horizontal scroll;
- `overflow-x: hidden` as the primary fix for a broken component;
- fixed-width mobile content;
- tiny text to force desktop tables into mobile;
- more than one visually dominant CTA in the same decision block;
- one-off component styles when a shared primitive can solve the case.

---

## 13. Protected surfaces and change discipline

The repository contains explicit protected-surface hashes in `scripts/ci/protected-surfaces.mjs`.

Before editing a protected file:

1. confirm the visual change is actually required;
2. understand all pages that consume the surface;
3. update the surface intentionally, not as collateral cleanup;
4. run the full build/test suite;
5. update the protected baseline only when the redesign is deliberate and validated.

Do not weaken the protected-surface guard to make a build pass.

---

## 14. QA checklist for every visual change

A UI change is incomplete until all applicable checks pass.

### Structure

- [ ] Existing component reused or extension justified.
- [ ] No duplicate design token invented.
- [ ] No unintended protected-surface change.
- [ ] Semantic HTML remains valid.

### Responsive

- [ ] 320px verified.
- [ ] 375/390px verified.
- [ ] 430px verified.
- [ ] 768px verified.
- [ ] 1024px verified.
- [ ] 1280/1440px verified.
- [ ] No page-level horizontal overflow.
- [ ] No clipped CTA, image, badge, table or sticky element.
- [ ] Long Turkish copy and long product names tested.

### Accessibility

- [ ] Keyboard flow works.
- [ ] Focus is visible.
- [ ] Contrast is safe.
- [ ] Touch targets are usable.
- [ ] Reduced motion is safe.

### Commercial

- [ ] Price/KDV/payment facts remain clear.
- [ ] Demo/proof is not hidden by decoration.
- [ ] Primary CTA remains obvious.
- [ ] No fake proof or misleading claim was added.

### Technical

- [ ] `npm run guard:design` passes.
- [ ] `npm run build` passes.
- [ ] `npm test` passes before merge.
- [ ] Existing SEO and protected-surface guards pass.

---

## 15. Failure modes and mitigation

### Failure mode A — design drift

Symptom: each new page introduces new colors, radius, card style or spacing.

Mitigation:

- reuse current tokens;
- extend a shared component;
- reject one-off styling unless the use case is unique and durable;
- keep `DESIGN.md` in the mandatory agent reading order.

### Failure mode B — responsive layout looks correct only because overflow is hidden

Symptom: mobile content is clipped, but the page does not visibly scroll sideways because a parent masks overflow.

Mitigation:

- local shrinkability (`min-width: 0`);
- `minmax(0, 1fr)`;
- bounded media;
- deliberate local table/code scroll;
- viewport verification at 320-430px;
- never use another overflow mask as the fix.

### Failure mode C — 21st.dev reference turns into a generic SaaS clone

Symptom: Excel visual language disappears and the page becomes rounded cards, gradients and generic feature grids.

Mitigation:

- take composition/interaction only;
- map every visual decision back to Excel Arşiv tokens;
- keep real workbook proof dominant;
- do not import the reference component/runtime.

### Failure mode D — visual upgrade breaks SEO/commercial behavior

Symptom: markup, links, product truth or checkout path changes while redesigning.

Mitigation:

- preserve route/schema/data contracts;
- keep the product/commercial source of truth separate from presentation;
- run existing smoke, SEO and commerce validation after every redesign.

### Failure mode E — fintech inspiration becomes visual imitation

Symptom: Excel Arşiv starts copying another financial brand's palette, layout proportions, typography, illustrations or branded components.

Mitigation:

- borrow navigation logic, trust sequencing and conversion mechanics only;
- keep Excel Arşiv tokens, workbook evidence and component vocabulary authoritative;
- never reproduce another site's branded artwork, exact copy, logo treatment or signature visual composition.

---

## 16. Agent execution prompt

For any future UI task, use this execution model:

> Read `DESIGN.md`, `AGENTS.md`, the target page/component and the relevant tokens in `src/styles/global.css` before changing code. Treat `DESIGN.md` as the visual contract and existing product/commerce data as factual truth. Reuse existing components and tokens. Use 21st.dev and high-trust financial product sites only as composition, information-architecture and interaction inspiration; do not import a new runtime or copy another brand's visual identity. Preserve protected surfaces, SEO, routes, schema and commerce behavior. Implement the requested change with responsive behavior designed from 320px through 1440px. Do not use page-level overflow masking to hide layout defects. Validate long Turkish copy, tables, screenshots, CTAs, focus states and reduced motion. Run `npm run guard:design`, `npm run build` and `npm test`. A change is complete only when the requested visual result is achieved without overflow, clipping, broken navigation, schema/SEO regressions or commercial-flow regressions.

---

## 17. Definition of done

A design change is done only when:

- it looks native to Excel Arşiv rather than like an imported template;
- it uses the current design tokens;
- it adds no unnecessary dependency;
- it preserves factual product/commercial behavior;
- it remains usable from 320px to large desktop;
- no viewport-level horizontal overflow is introduced;
- real product proof remains readable;
- accessibility fundamentals pass;
- repository guards/build/tests pass.

---

## 18. Fintech Trust & Conversion Architecture

This section captures the transferable strengths of high-trust financial product interfaces such as iyzico without copying their brand identity, artwork, copy, palette or proprietary visual treatment.

The objective is not to make Excel Arşiv look like a payment company. The objective is to make a large and technically complex product portfolio feel easy to understand, safe to evaluate and easy to buy.

### 18.1 Intent-first navigation

Users are not required to know the product name before they can find the right system.

Primary discovery should begin with the business outcome or pain point. Preferred top-level decision families are:

- Nakit Akışı;
- Tahsilat;
- Banka & Kredi;
- Kârlılık;
- Muhasebe & Vergi;
- Stok & Üretim;
- Personel / Operasyon when the catalog justifies it.

Rules:

1. A user should reach a relevant product family in no more than two meaningful decisions.
2. Do not expose the full catalog in the primary navigation.
3. Product-name search remains available for users who already know what they need.
4. Search and intent routing must converge on the same canonical product pages; do not create duplicate SEO routes for UI convenience.

### 18.2 Persona routing

Persona is a secondary routing aid, not the first taxonomy by default.

Use personas only when they materially change the recommended content or product set. Supported examples include:

- Patron / Yönetici;
- Finans;
- Muhasebe;
- Mali Müşavir;
- Operasyon.

Do not duplicate the same product catalog five times under persona labels. Persona routing should narrow or reorder, not fragment the source of truth.

### 18.3 Outcome taxonomy

Every product must have a clear user-facing outcome classification in addition to its internal category.

A product card or recommendation should answer at least one of these questions:

- What will I know after using it?
- What risk will I detect earlier?
- What repetitive work will it reduce?
- What management decision will it support?

Avoid category labels that explain the file format but not the business outcome.

### 18.4 Evidence ladder

Every high-intent commercial page should move from claim to evidence in a deliberate order:

`problem -> solution -> real product screen -> how it works -> input/output -> controls/risk -> demo/proof -> compatibility/trust -> price -> purchase`

Rules:

1. A significant claim should be followed by nearby proof where practical.
2. Real workbook screenshots outrank decorative illustrations as product evidence.
3. Product specifications support evidence; they do not replace evidence.
4. A demo must be presented as evaluation proof, not as a low-value freebie.
5. If evidence does not exist, do not fabricate it. Use accurate product facts instead.

### 18.5 Trust density

Trust must be distributed through the journey rather than isolated in one logo wall or footer.

High-value trust facts include only verifiable information such as:

- real Excel screenshots;
- demo availability;
- compatible Excel versions;
- macro requirement or macro-free status;
- protected formula behavior where factual;
- sheet/workbook structure;
- delivery method;
- one-time/subscription status;
- payment provider;
- KDV state;
- support/contact path;
- documented product controls and validation behavior.

Forbidden:

- invented customer logos;
- fabricated transaction/sales counts;
- fake review scores;
- fake live-user indicators;
- unsupported "most preferred" or market-leader claims.

### 18.6 Contextual CTA hierarchy

A CTA should match the user's current level of certainty.

Default hierarchy:

- discovery/hero: `Sistemi Bul` or equivalent problem-led action;
- catalog/product card: `İncele`;
- proof/demo area: `Demoyu Gör` / `Demoyu İncele` when available;
- final purchase area: `Satın Al`.

Rules:

1. Do not repeat the same visually dominant CTA in every section.
2. One decision block normally has one dominant action.
3. A secondary CTA must support the primary decision, not compete with it.
4. Sticky mobile purchase UI may repeat the purchase action because it solves navigation friction, but it must not cover content or violate safe-area rules.

### 18.7 Solution ecosystem and cross-sell

Related products are not random recommendations.

A related-system block should be generated from the user's business workflow:

`previous step -> current decision -> next step`

Examples:

- tahsilat takibi -> nakit akışı -> ödeme planlama;
- banka limitleri -> kredi taksit takibi -> nakit etkisi;
- satış/kârlılık -> fiyat ayarlama -> yönetim paneli.

Cross-sell must help the user complete a workflow. Do not use unrelated products only to increase card count.

### 18.8 Progressive disclosure

The first viewport explains business value. Technical detail appears when the user needs it to reduce risk or confirm suitability.

Preferred disclosure order:

1. outcome/problem;
2. proof;
3. how it works;
4. suitability and requirements;
5. technical workbook detail;
6. purchase terms.

Do not put sheet counts, formula terminology or implementation trivia ahead of the business result unless that technical fact is itself a purchase blocker.

### 18.9 Home-page conversion sequence

When a future deliberate redesign of the protected home surface is approved, prefer this sequence:

`INTENT -> RIGHT SOLUTION -> REAL PRODUCT -> HOW IT WORKS -> WHAT IT CONTROLS -> EVIDENCE -> DEMO -> PURCHASE`

Implementation guidance:

- Hero starts with the user's control problem, not a generic quality claim.
- Intent choices use a compact decision interface, not a wall of dozens of cards.
- Flagship products appear after the user understands the catalog logic.
- Real workbook proof appears before long marketing copy.
- Trust facts are attached to the relevant decision stage.
- Purchase is the final confidence step, not the first thing shouted at the user.

Protected-surface rules remain binding. This sequence is a design target, not permission to bypass `scripts/ci/protected-surfaces.mjs`.

### 18.10 Conversion anti-patterns

Reject the following:

- placing all catalog products in a mega-menu;
- organizing the site only by internal category names;
- generic SaaS card walls with equal visual weight;
- three or more competing primary CTAs in one viewport;
- a hero that says "premium" without showing proof;
- a trust section built from unverifiable logos or numbers;
- pushing technical Excel detail before the business outcome;
- unrelated cross-sell blocks;
- copying another financial brand's palette, illustrations, navigation chrome, wording or signature component styling.

### 18.11 Measurable conversion-design acceptance gates

For future UI changes that touch discovery, product evaluation or purchase paths, verify:

- [ ] Product family reachable within <= 2 meaningful user decisions where intent routing applies.
- [ ] One dominant CTA per decision block.
- [ ] No fabricated trust evidence.
- [ ] Product proof is real or explicitly labeled as illustrative.
- [ ] Technical detail does not precede the core business outcome without a documented reason.
- [ ] Related products share a real workflow relationship.
- [ ] Mobile intent controls and CTA targets meet the touch-size contract.
- [ ] No page-level horizontal overflow at required widths.
- [ ] Discovery and search resolve to canonical product routes.
- [ ] Protected-surface, SEO, commerce and accessibility guards still pass.

A conversion-oriented design is successful only if it reduces decision friction without weakening Excel Arşiv's native workbook identity, factual accuracy or technical safeguards.
