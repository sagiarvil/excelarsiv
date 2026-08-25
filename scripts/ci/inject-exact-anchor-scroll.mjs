#!/usr/bin/env node
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'dist';
const MARKER = 'data-exact-anchor-scroll';

function walk(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const full = join(dir, name);
    const stat = statSync(full);
    if (stat.isDirectory()) out.push(...walk(full));
    else if (name.endsWith('.html')) out.push(full);
  }
  return out;
}

const payload = `
<style ${MARKER}>
  html{scroll-behavior:auto!important}
  [id]{scroll-margin-top:88px}
  @media(max-width:760px){[id]{scroll-margin-top:72px}}
</style>
<script ${MARKER}>
(() => {
  const OFFSET_GAP = 12;

  const visibleTopObstruction = () => {
    const selectors = ['.site-header', '.nav', '[data-sticky-header]', 'header[style*="sticky"]', 'header[style*="fixed"]'];
    let bottom = 0;
    for (const selector of selectors) {
      for (const el of document.querySelectorAll(selector)) {
        const style = getComputedStyle(el);
        if (style.display === 'none' || style.visibility === 'hidden') continue;
        const rect = el.getBoundingClientRect();
        if (rect.height <= 0 || rect.bottom <= 0) continue;
        const positioned = style.position === 'sticky' || style.position === 'fixed' || selector === '.site-header' || selector === '.nav';
        if (positioned && rect.top <= 2) bottom = Math.max(bottom, rect.bottom);
      }
    }
    return Math.max(0, bottom);
  };

  const resolveTarget = (raw) => {
    if (!raw || raw === '#') return null;
    let id = raw.startsWith('#') ? raw.slice(1) : raw;
    try { id = decodeURIComponent(id); } catch {}
    return document.getElementById(id);
  };

  const jump = (target, hash, updateHistory = true) => {
    if (!target) return;
    const y = Math.max(0, window.scrollY + target.getBoundingClientRect().top - visibleTopObstruction() - OFFSET_GAP);
    window.scrollTo({ top: y, left: 0, behavior: 'auto' });
    if (updateHistory && hash) history.pushState(null, '', hash);
    if (!target.hasAttribute('tabindex')) target.setAttribute('tabindex', '-1');
    try { target.focus({ preventScroll: true }); } catch {}
  };

  document.addEventListener('click', (event) => {
    if (event.defaultPrevented || event.button !== 0 || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
    const trigger = event.target.closest('a[href^="#"], [data-scroll-target]');
    if (!trigger) return;
    const raw = trigger.getAttribute('data-scroll-target') || trigger.getAttribute('href');
    const target = resolveTarget(raw);
    if (!target) return;
    event.preventDefault();
    const hash = raw.startsWith('#') ? raw : '#' + raw.replace(/^#/, '');
    jump(target, hash, true);
  }, { passive: false });

  document.addEventListener('change', (event) => {
    const control = event.target.closest('select[data-scroll-target], input[data-scroll-target]');
    if (!control) return;
    const raw = control.getAttribute('data-scroll-target');
    const target = resolveTarget(raw);
    if (target) jump(target, raw.startsWith('#') ? raw : '#' + raw, true);
  });

  const correctInitialHash = () => {
    if (!location.hash) return;
    const target = resolveTarget(location.hash);
    if (!target) return;
    requestAnimationFrame(() => requestAnimationFrame(() => jump(target, location.hash, false)));
  };

  window.addEventListener('hashchange', correctInitialHash);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', correctInitialHash, { once: true });
  else correctInitialHash();
})();
</script>`;

let touched = 0;
for (const file of walk(DIST)) {
  let html = readFileSync(file, 'utf8');
  if (html.includes(MARKER) || !html.includes('</body>')) continue;
  html = html.replace('</body>', `${payload}\n</body>`);
  writeFileSync(file, html);
  touched++;
}

if (!touched) {
  console.error('Exact anchor scroll injector: no HTML files updated.');
  process.exit(1);
}

console.log(`Exact anchor scroll injector: ${touched} HTML files updated.`);
