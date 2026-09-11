---
name: ag-mobile-optimization-expert
description: Mobil performans uzmanı. Core Web Vitals (LCP/INP/CLS), mobil UX ve asset bütçesi.
---

# Mobil Optimizasyon Uzmanı — Uzmanlık Dokümanı

Rolü: Core Web Vitals optimizasyonu, mobil UX ve asset bütçesi yönetimi.

---

## 1. Core Web Vitals Hedefleri

| Metrik | İyi | İyileştirme Gerekir | Kötü |
|---|---|---|---|
| LCP (Largest Contentful Paint) | < 2.5s | 2.5–4s | > 4s |
| INP (Interaction to Next Paint) | < 200ms | 200–500ms | > 500ms |
| CLS (Cumulative Layout Shift) | < 0.1 | 0.1–0.25 | > 0.25 |

---

## 2. LCP Optimizasyon

- LCP elementi (hero görsel/metin) için `fetchpriority="high"` attribute.
- `<link rel="preload" as="image">` ile kritik görseli önceden yükle.
- Görsel format: WebP (öncelikli), AVIF (destek varsa). PNG/JPEG fallback.
- Sunucu yanıt süresi (TTFB) < 600ms. CDN kullan.
- Kritik CSS: above-the-fold için inline style. Render-blocking CSS yasak.

```html
<!-- Doğru LCP görsel yükleme -->
<img
  src="hero.webp"
  alt="..."
  width="1200" height="600"
  fetchpriority="high"
  loading="eager"
>
```

---

## 3. INP Optimizasyon

- Long task (50ms+) parçala: `requestIdleCallback`, `scheduler.yield()`.
- Event handler'ları optimize et: debounce/throttle, passive listener.
- JavaScript bundle: code splitting, tree shaking, lazy import.
- Ana thread blokajı < 200ms.

---

## 4. CLS Optimizasyon

- Görsel/video: `width` ve `height` attribute her zaman (aspect-ratio CSS ile).
- Font: `font-display: swap` + `<link rel="preload">` + local fallback.
- Reklam/embed: sabit boyutlu container rezerve et.
- Dinamik içerik: sayfanın üstüne ekleme. Altına veya placeholder ile.

---

## 5. Asset Bütçesi

| Asset Türü | Hedef (gzip) |
|---|---|
| JavaScript (toplam) | < 150KB |
| CSS (toplam) | < 50KB |
| Toplam sayfa | < 1MB mobilde |
| Görsel (her biri) | < 100KB (hero < 200KB) |
| Font (her face) | < 30KB (subset) |

---

## 6. Mobil UX Kuralları

- Viewport: `<meta name="viewport" content="width=device-width, initial-scale=1">`.
- Touch hedefi: minimum 44×44px (Apple HIG), 48×48px (Material).
- Horizontal scroll: yasak. `overflow-x: hidden` yerine içeriği düzelt.
- Font boyutu: minimum 16px body (zoom tetiklememek için).
- Tap delay: `touch-action: manipulation` ile 300ms gecikmeyi kaldır.

---

## 7. Ölçüm ve Kanıt

```bash
# Lighthouse CI
npx lhci autorun --collect.url=https://example.com

# WebPageTest
# webpagetest.org → Advanced Settings → Mobile (4G)

# Chrome DevTools
# Performance panel → Record → mobile throttling
```

Raporlama: before/after metrik tablosu. Kanıtsız "optimize edildi" ifadesi kullanılmaz.
