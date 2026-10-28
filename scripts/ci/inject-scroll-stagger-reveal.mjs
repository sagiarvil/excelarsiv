#!/usr/bin/env node
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'dist';
const MARKER = 'data-apple-stagger-reveal';

function walk(dir) {
  const out = [];
  try {
    for (const name of readdirSync(dir)) {
      const full = join(dir, name);
      const stat = statSync(full);
      if (stat.isDirectory()) out.push(...walk(full));
      else if (name.endsWith('.html')) out.push(full);
    }
  } catch (err) {
    console.error(`Walk error in ${dir}:`, err.message);
  }
  return out;
}

const MOTION_PAYLOAD = `
<style ${MARKER}>
/* Apple Motion Architecture: Scroll-Driven Reveal with Staggered Fade-Up */
html.ea-motion-ready .ea-fade-up {
  opacity: 0 !important;
  transform: translate3d(0, 24px, 0) !important;
  transition: opacity 0.65s cubic-bezier(0.16, 1, 0.3, 1),
              transform 0.65s cubic-bezier(0.16, 1, 0.3, 1) !important;
  transition-delay: calc(var(--stagger, 0) * 80ms) !important;
  will-change: opacity, transform;
}
html.ea-motion-ready .ea-fade-up.ea-revealed {
  opacity: 1 !important;
  transform: translate3d(0, 0, 0) !important;
}
@media (prefers-reduced-motion: reduce) {
  html.ea-motion-ready .ea-fade-up {
    opacity: 1 !important;
    transform: none !important;
    transition: none !important;
  }
}
</style>
<script ${MARKER}>
(() => {
  if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const init = () => {
    document.documentElement.classList.add('ea-motion-ready');
    const selectors = [
      '.grid > *',
      '.product-grid > *',
      '.card-grid > *',
      'section .grid > *',
      '.features-grid > *',
      '[data-product-card]',
      '.product-card',
      '.metric-card',
      '.pricing-card'
    ];
    const targets = document.querySelectorAll(selectors.join(','));
    let groupMap = new Map();
    targets.forEach((el) => {
      const parent = el.parentElement;
      if (!groupMap.has(parent)) groupMap.set(parent, []);
      groupMap.get(parent).push(el);
    });

    groupMap.forEach((children) => {
      children.forEach((child, idx) => {
        child.classList.add('ea-fade-up');
        child.style.setProperty('--stagger', String(idx % 8));
      });
    });

    const observer = new IntersectionObserver((entries, obs) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add('ea-revealed');
          obs.unobserve(entry.target);
        }
      });
    }, { rootMargin: '0px 0px -40px 0px', threshold: 0.1 });

    document.querySelectorAll('.ea-fade-up').forEach(el => {
      const rect = el.getBoundingClientRect();
      if (rect.top < window.innerHeight && rect.bottom > 0) {
        el.classList.add('ea-revealed');
      } else {
        observer.observe(el);
      }
    });
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
  document.addEventListener('astro:page-load', init);
})();
</script>
`;

function processHtmlFiles() {
  const files = walk(DIST);
  let modifiedCount = 0;

  for (const file of files) {
    let content = readFileSync(file, 'utf8');
    if (content.includes(MARKER)) continue;

    // Head kapanışından veya body sonundan önce yerleştir
    if (content.includes('</head>')) {
      content = content.replace('</head>', `${MOTION_PAYLOAD}\n</head>`);
      writeFileSync(file, content, 'utf8');
      modifiedCount++;
    } else if (content.includes('</body>')) {
      content = content.replace('</body>', `${MOTION_PAYLOAD}\n</body>`);
      writeFileSync(file, content, 'utf8');
      modifiedCount++;
    }
  }

  console.log(`[Apple Motion Gate] Scroll-Driven Reveal with Staggered Fade-Up: ${modifiedCount} HTML sayfasına enjekte edildi.`);
}

processHtmlFiles();
