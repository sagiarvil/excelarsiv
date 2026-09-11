# DETERMINISTIC SEO RESOLUTION PLAYBOOKS

This guide details root-cause recipes for resolving search and generative engine indexing failures.

## 1. VALID JSON-LD ORGANIZATION & BREADCRUMB SCHEMA
```html
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Organization",
  "name": "Enterprise Brand",
  "url": "https://example.com",
  "logo": "https://example.com/assets/logo.png",
  "contactPoint": {
    "@type": "ContactPoint",
    "telephone": "+90-212-000-0000",
    "contactType": "customer service",
    "areaServed": "TR",
    "availableLanguage": "Turkish"
  }
}
</script>
```

## 2. DYNAMIC TITLE & META DESCRIPTION ISOLATION
```php
<title><?= htmlspecialchars($pageTitle . ' | ' . $siteName, ENT_QUOTES, 'UTF-8'); ?></title>
<meta name="description" content="<?= htmlspecialchars($metaDescription, ENT_QUOTES, 'UTF-8'); ?>">
<link rel="canonical" href="<?= htmlspecialchars($canonicalUrl, ENT_QUOTES, 'UTF-8'); ?>">
```

## 3. OPENGRAPH & TWITTER CARD SPECIFICATION
- og:image MUST be an absolute HTTPS URL and meet minimum 1200x630px dimensions:
```html
<meta property="og:image" content="https://example.com/assets/og-image.jpg">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta name="twitter:card" content="summary_large_image">
```
