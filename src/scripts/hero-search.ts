export interface SearchItem {
  name: string;
  summary: string;
  category: string;
  url: string;
  kind?: 'product' | 'category' | 'guide' | 'problem';
  keywords?: string;
  priceTL?: number;
  sheetCount?: number;
}

const escapeHtml = (value: string): string =>
  value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');

const normalize = (text: string): string => text.toLocaleLowerCase('tr-TR');

function highlightMatch(text: string, query: string): string {
  if (!query) return escapeHtml(text);
  const q = normalize(query.trim());
  const normText = normalize(text);
  const idx = normText.indexOf(q);
  if (idx === -1) return escapeHtml(text);

  const before = text.slice(0, idx);
  const match = text.slice(idx, idx + query.trim().length);
  const after = text.slice(idx + query.trim().length);

  return `${escapeHtml(before)}<mark class="hero-search__mark">${escapeHtml(match)}</mark>${escapeHtml(after)}`;
}

export function mountHeroSearch(root: HTMLElement, items: SearchItem[]): void {
  const form = root.querySelector<HTMLFormElement>('form');
  const input = root.querySelector<HTMLInputElement>('input[type="search"]');
  const list = root.querySelector<HTMLUListElement>('ul');
  const kbdEl = root.querySelector<HTMLElement>('.hero-search__kbd');
  if (!form || !input || !list) return;

  // Platform shortcut detection (⌘K on Mac, Ctrl+K on Windows/Linux)
  if (kbdEl && typeof navigator !== 'undefined') {
    const isMac = /(Mac|iPhone|iPod|iPad)/i.test(navigator.platform || navigator.userAgent);
    kbdEl.textContent = isMac ? '⌘K' : 'Ctrl+K';
  }

  const close = (): void => {
    list.hidden = true;
    list.innerHTML = '';
    input.setAttribute('aria-expanded', 'false');
  };

  const open = (): void => {
    list.hidden = false;
    input.setAttribute('aria-expanded', 'true');
  };

  const render = (query: string): void => {
    const rawQ = query.trim();
    const q = normalize(rawQ);
    if (q.length < 2) {
      close();
      return;
    }

    const results = items
      .filter((item) => {
        const haystack = `${item.name} ${item.summary} ${item.category} ${item.keywords ?? ''}`;
        return normalize(haystack).includes(q);
      })
      .slice(0, 6);

    if (results.length === 0) {
      list.innerHTML = `
        <li class="hero-search__empty">
          <svg class="hero-search__empty-icon" viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="11" cy="11" r="8"></circle>
            <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
          </svg>
          <div>
            <strong class="block text-sm text-neutral-800">Sonuç bulunamadı</strong>
            <span class="text-xs text-neutral-500">“${escapeHtml(rawQ)}” için eşleşen hazır sistem bulunamadı.</span>
          </div>
          <a href="/sablonlar" class="hero-search__empty-link">Tüm Şablonları Listele →</a>
        </li>`;
      open();
      return;
    }

    const kindConfig: Record<string, { label: string; badgeClass: string }> = {
      product: { label: 'SİSTEM', badgeClass: 'hero-search__badge--product' },
      category: { label: 'KATEGORİ', badgeClass: 'hero-search__badge--category' },
      guide: { label: 'REHBER', badgeClass: 'hero-search__badge--guide' },
      problem: { label: 'ÇÖZÜM', badgeClass: 'hero-search__badge--problem' },
    };

    const itemsHtml = results
      .map((result, index) => {
        const conf = kindConfig[result.kind ?? 'product'] ?? kindConfig.product;
        const pricePill = result.priceTL
          ? `<span class="hero-search__price font-mono font-bold">${result.priceTL.toLocaleString('tr-TR')} TL</span>`
          : '';
        const sheetPill = result.sheetCount
          ? `<span class="hero-search__sheets font-mono">${result.sheetCount} Sayfa</span>`
          : '';

        return `
          <li class="hero-search__item">
            <a href="${result.url}" data-result-index="${index}" class="hero-search__link">
              <div class="hero-search__item-main">
                <div class="hero-search__item-head">
                  <span class="hero-search__badge ${conf.badgeClass}">${conf.label}</span>
                  <span class="hero-search__item-cat">${escapeHtml(result.category)}</span>
                </div>
                <span class="hero-search__name">${highlightMatch(result.name, rawQ)}</span>
                ${result.summary ? `<p class="hero-search__item-summary">${escapeHtml(result.summary)}</p>` : ''}
              </div>
              <div class="hero-search__item-meta">
                ${sheetPill}
                ${pricePill}
                <span class="hero-search__item-arrow">→</span>
              </div>
            </a>
          </li>`;
      })
      .join('');

    const footerHtml = `
      <li class="hero-search__footer">
        <span class="hero-search__footer-hint">↑↓ Gezin · ↵ Seç · ESC Kapat</span>
        <a href="/sablonlar" class="hero-search__footer-all">51 Sistemin Tamamını Gör →</a>
      </li>
    `;

    list.innerHTML = itemsHtml + footerHtml;
    open();
  };

  input.addEventListener('input', () => render(input.value));
  input.addEventListener('focus', () => {
    if (normalize(input.value).length >= 2) render(input.value);
  });

  input.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      close();
      input.blur();
      return;
    }
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    const links = Array.from(list.querySelectorAll<HTMLAnchorElement>('a[data-result-index]'));
    if (links.length === 0) return;
    event.preventDefault();
    const active = list.querySelector<HTMLAnchorElement>('a[data-result-index].is-highlighted');
    const activeIndex = active ? links.indexOf(active) : -1;
    const nextIndex =
      event.key === 'ArrowDown'
        ? (activeIndex + 1) % links.length
        : (activeIndex - 1 + links.length) % links.length;

    links.forEach((link, index) => {
      const isActive = index === nextIndex;
      link.classList.toggle('is-highlighted', isActive);
      if (isActive) link.scrollIntoView({ block: 'nearest' });
    });
  });

  document.addEventListener('click', (event) => {
    if (!root.contains(event.target as Node)) close();
  });

  // Global hotkey: ⌘K, Ctrl+K or / to focus search
  window.addEventListener('keydown', (event) => {
    const isHotkey = (event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'k';
    const isSlash = event.key === '/' && document.activeElement !== input && !['INPUT', 'TEXTAREA', 'SELECT'].includes((document.activeElement?.tagName || ''));
    if (isHotkey || isSlash) {
      event.preventDefault();
      root.scrollIntoView({ behavior: 'smooth', block: 'center' });
      input.focus();
      input.select();
    }
  });

  form.addEventListener('submit', (event) => {
    if (normalize(input.value).length < 2) {
      event.preventDefault();
      return;
    }
    const highlighted = list.querySelector<HTMLAnchorElement>('a[data-result-index].is-highlighted');
    if (highlighted) {
      event.preventDefault();
      window.location.href = highlighted.href;
    }
  });
}
