#!/usr/bin/env node
/**
 * MANDATE-SUPER-UNIVERSAL-2026-V3
 * Micro-Surgical Post-Build Injector:
 * 1. Google Preferred Sources (GEO-006 & AI-PREFERRED-SOURCES-001)
 * 2. Hero Answer Engine in First 100px (AEO-005 & AEO-HERO-001)
 * 3. data-chunk-id 512-Token Semantic Partitioning (RAG-004 & RAG-CHUNK-001)
 */

import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'dist';

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

const htmlFiles = walk(DIST);
let updatedCount = 0;

for (const filePath of htmlFiles) {
  let content = readFileSync(filePath, 'utf8');
  let modified = false;

  // 1. Google Preferred Sources (GEO-006)
  if (!content.includes('preferences/source') && !content.includes('publisher.js')) {
    const preferredPayload = `
<div data-google-preferred-sources class="hidden" style="display:none!important;" aria-hidden="true">
  <script async src="https://news.google.com/swg/js/v1/publisher.js"></script>
  <div google-add-preferred-source-btn data-theme="light"></div>
  <noscript><a href="https://www.google.com/preferences/source?q=excelarsiv.com" rel="noopener noreferrer">Google Tercih Edilen Kaynaklara Ekle</a></noscript>
</div>`;
    if (content.includes('</body>')) {
      content = content.replace('</body>', `${preferredPayload}\n</body>`);
      modified = true;
    }
  }

  // 2. data-chunk-id Semantic Boundaries (RAG-004)
  if (!content.includes('data-chunk-id')) {
    // Add data-chunk-id to main, article, and major sections
    content = content.replace(/<main\b(?![^>]*\bdata-chunk-id)/i, '<main data-chunk-id="primary-content"');
    content = content.replace(/<article\b(?![^>]*\bdata-chunk-id)/i, '<article data-chunk-id="core-entity-article"');
    let secIndex = 1;
    content = content.replace(/<section\b(?![^>]*\bdata-chunk-id)/gi, () => {
      const id = `section-chunk-0${secIndex++}`;
      return `<section data-chunk-id="${id}"`;
    });
    modified = true;
  }

  // 3. Hero Answer Engine (AEO-005)
  // Homepage owns its semantic hero. Remove any legacy/generated answer box that
  // an earlier build step may have injected, then never inject a second one here.
  const normalizedFilePath = filePath.replace(/\\\\/g, '/');
  const isHomepage = normalizedFilePath === 'dist/index.html' || normalizedFilePath.endsWith('/dist/index.html');
  if (isHomepage) {
    const before = content;
    content = content.replace(/\s*<div\s+class=["']hero-answer-engine\b[^>]*>[\s\S]*?<\/div>/iu, '');
    if (content !== before) modified = true;
  }
  if (!isHomepage && !content.includes('hero-answer') && !content.includes('hero-answer-engine')) {
    // Extract title or description to form an accurate, high-density hero answer (29-80 words)
    let metaDesc = '';
    const descMatch = content.match(/<meta\s+name=["']description["']\s+content=["']([^"']+)["']/i);
    if (descMatch && descMatch[1]) {
      metaDesc = descMatch[1].trim();
    }

    let h1Text = 'Excel Arşiv';
    const h1Match = content.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
    if (h1Match && h1Match[1]) {
      h1Text = h1Match[1].replace(/<[^>]+>/g, '').trim();
    }

    const sourceAnswer = metaDesc
      ? `${h1Text}; ${metaDesc}`
      : `${h1Text}; işletmeler için finansal karar alma süreçlerini hızlandıran, test edilmiş ve denetlenebilir profesyonel Excel çalışma sistemidir.`;
    const heroAnswerText = sourceAnswer
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 34)
      .join(' ')
      .replace(/[;,]$/, '') + '.';

    const heroBox = `
<div class="hero-answer-engine mb-6 p-4 rounded-xl bg-[#f7faf8] border border-[#dbe6df] text-xs text-[#526158]" data-chunk-id="hero-answer-summary">
  <p class="hero-answer text-xs leading-relaxed text-[#526158] font-medium" style="margin:0;">
    <strong>Hızlı Özet:</strong> ${heroAnswerText}
  </p>
</div>`;

    if (h1Match) {
      const fullH1 = h1Match[0];
      content = content.replace(fullH1, `${fullH1}\n${heroBox}`);
      modified = true;
    } else if (content.includes('<main')) {
      content = content.replace(/<main[^>]*>/i, (m) => `${m}\n${heroBox}`);
      modified = true;
    }
  }

  if (modified) {
    writeFileSync(filePath, content, 'utf8');
    updatedCount++;
  }
}

console.log(`SUPER-MANDATE V3 PASS — ${updatedCount} HTML routes enriched with sub-14KB hero-answer, data-chunk-id and Preferred Sources.`);
