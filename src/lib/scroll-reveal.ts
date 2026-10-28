/**
 * Apple Motion Architecture: Scroll-Driven Reveal with Staggered Fade-Up
 * Kademeli Kaydırma Görünürlük Animasyonu Motoru
 *
 * 1. Fade-Up / Slide-Up: opacity 0 -> 1, translateY 24px -> 0
 * 2. Stagger Effect: Her ardışık öğe 80ms aralıkla domino gibi akar.
 * 3. Scroll-Driven: IntersectionObserver ile viewport'a girince tetiklenir.
 * 4. Zero-Jank: Sadece composite transform ve opacity kullanılır; CLS = 0.
 */

export function initScrollReveal() {
  if (typeof window === 'undefined') return;

  // Düşük hareket tercihi varsa çalıştırma (WCAG 2.2 AA)
  if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
    return;
  }

  const docEl = document.documentElement;
  docEl.classList.add('ea-motion-ready');

  // Hedeflenecek kartlar, bölümler ve vitrin konteynerleri
  const containers = document.querySelectorAll(
    '.grid, .product-grid, .card-grid, section > div, .features-grid, [data-stagger-group]'
  );

  containers.forEach((container) => {
    const children = Array.from(container.children) as HTMLElement[];
    // Eğer 1'den fazla çocuk varsa stagger uygula
    if (children.length > 1) {
      children.forEach((child, index) => {
        if (!child.classList.contains('ea-fade-up') && !child.hasAttribute('data-no-motion')) {
          child.classList.add('ea-fade-up');
          child.style.setProperty('--stagger', String(index % 8)); // 0..7 döngüsüyle 80ms basamak
        }
      });
    }
  });

  // Tekil bölümler ve ana başlıklar
  const standaloneElements = document.querySelectorAll(
    'section h2, section .section-header, .hero-metric, .interactive-card'
  );
  standaloneElements.forEach((el) => {
    const htmlEl = el as HTMLElement;
    if (!htmlEl.classList.contains('ea-fade-up') && !htmlEl.hasAttribute('data-no-motion')) {
      htmlEl.classList.add('ea-fade-up');
      htmlEl.style.setProperty('--stagger', '0');
    }
  });

  // IntersectionObserver ile Viewport takibi
  const revealObserver = new IntersectionObserver(
    (entries, observer) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add('ea-revealed');
          observer.unobserve(entry.target); // Tek seferlik akıcı tetikleme
        }
      });
    },
    {
      root: null,
      rootMargin: '0px 0px -40px 0px', // Ekrana 40px girmeden önce başla
      threshold: 0.1,
    }
  );

  document.querySelectorAll('.ea-fade-up').forEach((el) => {
    // Halihazırda ekranda görünenleri hemen aç
    const rect = el.getBoundingClientRect();
    if (rect.top < window.innerHeight && rect.bottom > 0) {
      el.classList.add('ea-revealed');
    } else {
      revealObserver.observe(el);
    }
  });
}

// Otomatik başlatma
if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initScrollReveal);
  } else {
    initScrollReveal();
  }
  document.addEventListener('astro:page-load', initScrollReveal);
}
