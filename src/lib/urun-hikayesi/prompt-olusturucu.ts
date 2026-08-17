import type { UrunGorselManifest } from './turler.ts';

// İleride illüstrasyonlar AI üretimiyle yapılacaksa kullanılacak kontrollü brief.
// Not: metin görselin içine AI ile yazdırılmaz; yalnızca sahne üretilir.

export function sahnePromptuOlustur(oge: UrunGorselManifest): string {
  return `
Create a clean premium product-card illustration for an Excel template catalog.

Goal:
- visually communicate what the template does in under 2 seconds
- white/light background, soft shadow, modern SaaS catalog style
- elegant, minimal, polished, no watermark, no clutter

Product brief:
Title: ${oge.title}
Story: ${oge.story}
Primary pain: ${oge.primaryPain}
Result signal: ${oge.resultSignal}
Scene key: ${oge.sceneKey}
Hero object: ${oge.fingerprint.heroObject}
UI module: ${oge.fingerprint.uiModule}
Perspective: ${oge.fingerprint.perspective}
Accent color: ${oge.fingerprint.accent}

Constraints:
- do NOT render the title as part of the image
- do NOT add random unreadable text
- include a mini UI panel or miniature dashboard related to the product
- make the scene unique relative to other catalog items
- avoid these objects: ${oge.prohibited.join(', ') || 'none'}

Style: premium business illustration, clean geometry, subtle gradients,
rounded shapes, high legibility, practical and conversion-focused.
  `.trim();
}
