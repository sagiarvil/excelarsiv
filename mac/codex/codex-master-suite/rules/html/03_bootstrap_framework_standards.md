# BOOTSTRAP 5 FRAMEWORK STANDARDS & INVARIANTS

This standard governs responsive design, grid hierarchy, utility classes, and components when using Bootstrap 5.

## 1. MODERN BOOTSTRAP 5.3+ & JQUERY BAN
1. **Native Vanilla JavaScript:** Bootstrap 5 relies entirely on vanilla JS. Including jQuery (`jquery.min.js`) is STRICTLY FORBIDDEN.
2. **Local Assets Only (CDN Ban):** Bootstrap CSS and JS bundles must be loaded from local project directories (`/assets/css/bootstrap.min.css`, `/assets/js/bootstrap.bundle.min.js`).

## 2. GRID SYSTEM HIERARCHY (.CONTAINER -> .ROW -> .COL)
1. **Direct Child Invariant:**
   - A `.row` must always reside within a `.container` or `.container-fluid`.
   - `.col-*` classes MUST be direct children of a `.row`.
   - NEVER inject intermediary wrapper `<div>` elements between a `.row` and its `.col` children.
2. **12-Column Grid Rule:** Grid columns in a single horizontal group must sum to 12. Structure breakpoints mobile-first:
   ```html
   <div class="container py-4">
     <div class="row g-3">
       <div class="col-12 col-md-6 col-lg-4">
         <div class="card h-100">...</div>
       </div>
       <div class="col-12 col-md-6 col-lg-4">
         <div class="card h-100">...</div>
       </div>
       <div class="col-12 col-md-12 col-lg-4">
         <div class="card h-100">...</div>
       </div>
     </div>
   </div>
   ```

## 3. BOOTSTRAP 5 DATA ATTRIBUTES SPECIFICATION
- **Bootstrap 4 Syntax Banned:** Legacy `data-toggle` and `data-target` attributes are invalid in Bootstrap 5.
- **Mandatory `data-bs-*` Syntax:**
  - Modals: `data-bs-toggle="modal" data-bs-target="#exampleModal"`
  - Dismiss: `data-bs-dismiss="modal"` or `data-bs-dismiss="alert"`
  - Dropdowns: `data-bs-toggle="dropdown"`
  - Collapses: `data-bs-toggle="collapse" data-bs-target="#collapseId"`

## 4. UTILITY CLASSES OVER INLINE STYLES
- Replace inline styles with Bootstrap utility classes:
  - Spacing: `mt-3`, `mb-4`, `p-3`, `gap-3`
  - Flexbox: `d-flex`, `align-items-center`, `justify-content-between`
  - Typography: `fw-bold`, `text-muted`, `text-truncate`

## 5. FORMS & ACCESSIBILITY
- Form controls must use `.form-control`, `.form-select`, and `.form-check-input`.
- Icon-only buttons must provide an explicit `aria-label` or `.visually-hidden` descriptor:
  ```html
  <button type="button" class="btn btn-outline-danger btn-sm" aria-label="Delete Record">
    <i class="bi bi-trash" aria-hidden="true"></i>
  </button>
  ```

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
