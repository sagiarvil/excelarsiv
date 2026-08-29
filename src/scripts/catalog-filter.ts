import {
  searchEngine,
  normalizeTurkish,
  findDidYouMeanSuggestion,
  buildSearchCorpus,
  type SearchItem,
} from '../lib/search-engine';

const ensureCommandStackingLayer = (): void => {
  if (document.getElementById('catalog-command-stacking-layer')) return;
  const style = document.createElement('style');
  style.id = 'catalog-command-stacking-layer';
  style.textContent = `
    .catalog-section .site-container { position: relative !important; isolation: isolate !important; }
    .catalog-command { position: relative !important; z-index: 120 !important; overflow: visible !important; }
    .catalog-command__head { position: relative !important; z-index: 121 !important; }
    .catalog-command__search { position: relative !important; z-index: 160 !important; overflow: visible !important; }
    [data-catalog-command] { position: relative !important; z-index: 180 !important; overflow: visible !important; isolation: auto !important; }
    [data-catalog-command-panel] { z-index: 10000 !important; }
    .catalog-command__filters { position: relative !important; z-index: 20 !important; }
    [data-template-grid-wrap], [data-template-grid], [data-template-item], [data-template-card] { position: relative !important; z-index: 1 !important; }
    .empty-state__dym-btn { display: inline-flex; align-items: center; gap: 8px; margin-top: 14px; padding: 10px 18px; border: 1px solid #b7e0c7; border-radius: 12px; background: #f0faf4; color: #166534; font-family: var(--font-body); font-size: 13px; font-weight: 700; cursor: pointer; transition: all .15s ease; }
    .empty-state__dym-btn:hover { background: #e2f6ea; border-color: #7ece9d; transform: translateY(-1px); }
    .empty-state__dym-btn strong { text-decoration: underline; color: #0e4c25; }
  `;
  document.head.append(style);
};

export function mountCatalogFilter(): void {
  const grid = document.querySelector<HTMLElement>('[data-template-grid]');
  if (!grid) return;

  ensureCommandStackingLayer();

  const cards = Array.from(grid.querySelectorAll<HTMLElement>('[data-template-card]'));
  const search = document.querySelector<HTMLInputElement>('[data-catalog-search]');
  const searchClear = document.querySelector<HTMLButtonElement>('[data-catalog-search-clear]');
  const chips = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-category-filter]'));
  const resultCount = document.querySelector<HTMLElement>('[data-result-count]');
  const clearAll = document.querySelector<HTMLButtonElement>('[data-catalog-clear]');

  // Build items representation for search engine
  const catalogItems: SearchItem[] = cards.map((card) => ({
    name: card.dataset.name ?? '',
    summary: card.dataset.summary ?? '',
    category: card.dataset.categoryLabel ?? '',
    categorySlug: card.dataset.category ?? '',
    url: card.querySelector<HTMLAnchorElement>('a')?.getAttribute('href') ?? '/sablonlar',
  }));
  const corpus = buildSearchCorpus(catalogItems);

  const ensureEmptyState = (): HTMLElement => {
    const wrap = grid.closest<HTMLElement>('[data-template-grid-wrap]');
    let empty = grid.parentElement?.querySelector<HTMLElement>('[data-empty-state]') ?? null;
    if (!empty) {
      empty = document.createElement('div');
      empty.className = 'empty-state';
      empty.dataset.emptyState = '';
      empty.hidden = true;
      const icon = document.createElement('span');
      icon.className = 'empty-state__icon';
      icon.setAttribute('aria-hidden', 'true');
      icon.textContent = 'Ø';
      const title = document.createElement('h2');
      title.className = 'empty-state__title';
      title.textContent = 'Sonuç bulunamadı';
      const text = document.createElement('p');
      text.className = 'empty-state__text';
      text.dataset.emptyStateText = '';
      empty.append(icon, title, text);
      wrap?.append(empty);
    }
    return empty;
  };

  const params = new URLSearchParams(window.location.search);
  let activeCategory = params.get('kategori') ?? '';

  if (!chips.some((chip) => chip.dataset.category === activeCategory)) activeCategory = '';

  const setChipState = (): void => {
    chips.forEach((chip) => {
      const isActive = chip.dataset.category === activeCategory;
      chip.classList.toggle('is-active', isActive);
      chip.setAttribute('aria-pressed', String(isActive));
    });
  };

  const apply = (): void => {
    const rawQuery = (search?.value ?? '').trim();
    let visible = 0;

    let matchedItemNames = new Set<string>();
    let didYouMeanSuggestion: string | null = null;

    if (rawQuery.length > 0) {
      const response = searchEngine(catalogItems, rawQuery, {
        limit: 100,
        enableFuzzy: true,
        enableIntentParsing: true,
      });
      response.results.forEach((item) => matchedItemNames.add(item.name));
      didYouMeanSuggestion = response.didYouMean;
    }

    for (const card of cards) {
      const name = card.dataset.name ?? '';
      const category = card.dataset.category ?? '';
      const matchCategory = activeCategory === '' || category === activeCategory;
      const matchQuery = rawQuery.length === 0 || matchedItemNames.has(name);
      const show = matchCategory && matchQuery;
      const item = card.closest<HTMLElement>('[data-template-item]') ?? card;
      item.hidden = !show;
      if (show) visible += 1;
    }

    const empty = ensureEmptyState();
    if (empty) {
      empty.hidden = visible > 0;
      const text = empty.querySelector<HTMLElement>('[data-empty-state-text]');
      const existingDym = empty.querySelector<HTMLButtonElement>('.empty-state__dym-btn');
      if (existingDym) existingDym.remove();

      if (text) {
        if (rawQuery.length > 0) {
          text.textContent = `“${rawQuery}” için eşleşen şablon bulunamadı.`;
          if (didYouMeanSuggestion) {
            const dymBtn = document.createElement('button');
            dymBtn.type = 'button';
            dymBtn.className = 'empty-state__dym-btn';
            dymBtn.innerHTML = `Bunu mu demek istediniz: <strong>“${didYouMeanSuggestion}”</strong> (Sonuçları göster)`;
            dymBtn.addEventListener('click', () => {
              if (search && didYouMeanSuggestion) {
                search.value = didYouMeanSuggestion;
                apply();
                search.focus();
              }
            });
            empty.append(dymBtn);
          }
        } else {
          text.textContent = 'Bu kategoriye uygun şablon bulunamadı.';
        }
      }
    }
    if (resultCount) resultCount.textContent = String(visible);
    if (searchClear) searchClear.hidden = rawQuery.length === 0;
    if (clearAll) clearAll.hidden = rawQuery.length === 0 && activeCategory === '';

    const url = new URL(window.location.href);
    if (rawQuery.length > 0) url.searchParams.set('q', rawQuery);
    else url.searchParams.delete('q');
    if (activeCategory) url.searchParams.set('kategori', activeCategory);
    else url.searchParams.delete('kategori');
    history.replaceState(null, '', url);
  };

  setChipState();

  if (search) search.value = params.get('q') ?? '';

  search?.addEventListener('input', apply);
  searchClear?.addEventListener('click', () => {
    if (!search) return;
    search.value = '';
    search.focus();
    apply();
  });

  chips.forEach((chip) => {
    chip.addEventListener('click', () => {
      activeCategory = chip.dataset.category ?? '';
      setChipState();
      apply();
    });
  });

  clearAll?.addEventListener('click', () => {
    activeCategory = '';
    if (search) search.value = '';
    setChipState();
    apply();
    search?.focus();
  });

  apply();
}
