# SEO-ALIGNED HTML & BOOTSTRAP 5 EXPORT SPECIFICATION
# Compatibility with Engine V3.0, schema_validator.js, and heading_semantics_check.js

This standard mandates that any agent generating or exporting HTML and Bootstrap 5 markup MUST produce code that passes all canonical SEO, GEO, and structured data validation tools with a 100% score.

---

## 🚀 1. THE 14KB AST CRITICAL WINDOW BUDGET
Search engine crawlers and generative AI bots prioritize the first 14KB of the HTML document.
1. **Critical Head & Body Order:** The top 14KB of the generated markup MUST contain:
   - Complete technical meta tags (`title`, `description`, `canonical`, `viewport`, `robots`).
   - Social metadata (OpenGraph `og:title`, `og:image`, `og:description` and Twitter cards).
   - Valid JSON-LD structured data block (`<script type="application/ld+json">`).
   - Primary content element (`<main id="main-content">`) with the single `<h1>` heading.
2. **Zero Head Bloat:** Never inject massive inline JavaScript or render-blocking monolithic CSS in the document head before primary semantic markup.

---

## 🏷️ 2. HEADING HIERARCHY & SEMANTIC TREE (HEADING CHECKER COMPLIANT)
The generated markup MUST pass `heading_semantics_check.js` with ZERO errors:
1. **Single H1 Rule:** Exactly ONE `<h1>` tag per document. It must encapsulate the primary target topic and brand.
2. **Strict Hierarchy Order:** Never skip levels:
   - `<h1>` -> `<h2>` -> `<h3>` -> `<h4>`
   - Transitioning directly from `<h2>` to `<h4>` is STRICTLY FORBIDDEN.
3. **Bootstrap Visual Headings:** When a smaller heading requires larger visual styling, use Bootstrap typography utility classes on the semantic tag (e.g., `<h2 class="h1 fw-bold">`), NEVER downgrade the semantic tag to a `<div>` or misalign the hierarchy.

---

## 📊 3. EMBEDDED STRUCTURED DATA (SCHEMA_VALIDATOR COMPLIANT)
Every exported HTML page MUST embed an error-free, W3C/Schema.org compliant JSON-LD block validated by `schema_validator.js`:
```html
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@graph": [
    {
      "@type": "WebSite",
      "@id": "https://example.com/#website",
      "url": "https://example.com",
      "name": "Site Name",
      "publisher": {
        "@type": "Organization",
        "name": "Brand Name",
        "logo": "https://example.com/assets/logo.png"
      }
    },
    {
      "@type": "BreadcrumbList",
      "@id": "https://example.com/page/#breadcrumb",
      "itemListElement": [
        {
          "@type": "ListItem",
          "position": 1,
          "name": "Ana Sayfa",
          "item": "https://example.com"
        },
        {
          "@type": "ListItem",
          "position": 2,
          "name": "Kategori",
          "item": "https://example.com/kategori"
        }
      ]
    }
  ]
}
</script>
```

---

## 🎨 4. BOOTSTRAP 5 SEMANTIC LANDMARK FUSION
Exported markup must combine native HTML5 semantic landmark elements with Bootstrap responsive layout classes:
```html
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SEO & Bootstrap Optimized Page</title>
  <meta name="description" content="Concise, keyword-focused description under 160 characters.">
  <link rel="canonical" href="https://example.com/current-page">
  <!-- OpenGraph & JSON-LD -->
  <link rel="stylesheet" href="/assets/css/bootstrap.min.css">
</head>
<body>
  <!-- Header & Navigation Landmark -->
  <header class="site-header border-bottom">
    <nav class="navbar navbar-expand-lg navbar-light bg-light" aria-label="Ana Navigasyon">
      <div class="container">
        <a class="navbar-brand fw-bold" href="/">LOGO</a>
        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navMenu" aria-controls="navMenu" aria-expanded="false" aria-label="Menüyü Aç">
          <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navMenu">
          <ul class="navbar-nav ms-auto">
            <li class="nav-item"><a class="nav-link active" aria-current="page" href="/">Ana Sayfa</a></li>
          </ul>
        </div>
      </div>
    </nav>
  </header>

  <!-- Main Content Landmark -->
  <main id="main-content" class="container py-5">
    <div class="row g-4">
      <article class="col-12 col-lg-8">
        <h1 class="display-5 fw-bold mb-3">Ana Başlık (Single H1)</h1>
        <p class="lead text-muted">Özet ve giriş paragrafı.</p>
        
        <section class="mt-4">
          <h2 class="h3 fw-semibold">Alt Bölüm Başlığı (H2)</h2>
          <p>Açıklayıcı içerik metni.</p>
          <img src="/assets/img/example.webp" alt="Açıklayıcı görsel metni" width="800" height="450" loading="lazy" decoding="async" class="img-fluid rounded shadow-sm">
        </section>
      </article>

      <!-- Sidebar Landmark -->
      <aside class="col-12 col-lg-4" aria-label="Yan Menü">
        <div class="card p-3 shadow-sm">
          <h2 class="h5 fw-bold">İlgili Bağlantılar</h2>
          <ul class="list-unstyled mb-0">
            <li><a href="/detay" class="text-decoration-none">Ayrıntılı Rehberi İnceleyin</a></li>
          </ul>
        </div>
      </aside>
    </div>
  </main>

  <!-- Footer Landmark -->
  <footer class="site-footer bg-dark text-white py-4 mt-auto">
    <div class="container text-center">
      <p class="mb-0 text-muted">&copy; 2026 Tüm Hakları Saklıdır.</p>
    </div>
  </footer>

  <script src="/assets/js/bootstrap.bundle.min.js"></script>
</body>
</html>
```

---

## 🔍 5. PHYSICAL AUDIT CHECKLIST BEFORE SIGN-OFF
After exporting or modifying any HTML/Bootstrap markup, the agent MUST run:
1. `node $HOME/.gemini\config\plugins\gemini-master-suite\skills\ag-seo-expert\tools\heading_semantics_check.js <file>`
2. `node $HOME/.gemini\config\plugins\gemini-master-suite\skills\ag-seo-expert\tools\schema_validator.js <file>`
3. `node $HOME/.gemini\config\plugins\gemini-master-suite\skills\ag-seo-expert\tools\dom_utf8_full_test.js <file>`
All three checks MUST exit with 0 errors.

---

## 🇹🇷 6. TÜRKÇE KARAKTER & JSON VERİ BÜTÜNLÜĞÜ STANDARDI (MANDATORY P0)
Any agent generating HTML, Bootstrap 5 templates, forms, or frontend components MUST strictly obey these rules:

1. **Mandatory `<html lang="tr">` and `<meta charset="UTF-8">`:**
   - The root tag MUST declare `lang="tr"` so browsers properly handle Turkish I-casing and screen readers pronounce correctly.
   - The very first tag inside `<head>` MUST be `<meta charset="UTF-8">`.
   ```html
   <!DOCTYPE html>
   <html lang="tr">
   <head>
     <meta charset="UTF-8">
     <meta name="viewport" content="width=device-width, initial-scale=1.0">
   ```

2. **Native UTF-8 Turkish Characters (Zero ASCII Degradation & Zero Entity Bloat):**
   - In all Bootstrap components (Navbar, Cards, Modals, Buttons, Alerts, Tables, Forms, Placeholders):
     ALWAYS use proper Turkish characters: `ç, Ç, ğ, Ğ, ı, I, i, İ, ö, Ö, ş, Ş, ü, Ü`.
   - **BANNED:** Never downgrade words to English ASCII (e.g. NEVER write "Giris Yap", "Kaydetmek Icin Tiklayiniz" -> ALWAYS write "Giriş Yap", "Kaydetmek İçin Tıklayınız").
   - **BANNED:** Never use obsolete HTML entities for Turkish letters (e.g. do not write `&ccedil;`, `&#305;` -> write native `ç`, `ı`).

3. **Bootstrap Typography & Turkish "I" Case Sensitivity:**
   - When using Bootstrap `.text-uppercase` or `.text-lowercase`, ensure `<html lang="tr">` is present so `i` correctly uppercases to `İ` (and not English `I`).
   - For critical CTAs and headings, author the text directly in correct uppercase (e.g. `GİRİŞ YAP`, `İLETİŞİM`, `IŞIK`).

4. **Bootstrap AJAX / Fetch Form Submissions (Mandatory Headers):**
   - Every JavaScript fetch/AJAX request submitting Bootstrap form data to backend APIs MUST include:
   ```javascript
   fetch('/api/endpoint', {
     method: 'POST',
     headers: {
       'Content-Type': 'application/json; charset=utf-8',
       'Accept': 'application/json; charset=utf-8'
     },
     body: JSON.stringify(formData)
   });
   ```

5. **HTML5 Form Validation Patterns:**
   - In form fields accepting Turkish text (Names, Cities, Messages), HTML5 `pattern` attributes MUST support Turkish characters:
     `pattern="[a-zA-ZçÇğĞıİöÖşŞüÜ\s]+"` (never restrict solely to `[a-zA-Z]`).

6. **Mojibake Auto-Sanitization:**
   - If incoming template strings contain corrupted characters (`Ý, Ð, Þ, ý, ð, þ`), they must immediately be mapped to (`İ, Ğ, Ş, ı, ğ, ş`).
