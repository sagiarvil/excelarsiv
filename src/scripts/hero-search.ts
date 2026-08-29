import { searchEngine, type SearchItem } from '../lib/search-engine';

export type { SearchItem };

const escapeHtml = (value: string): string =>
  value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');

const iconForKind = (kind: SearchItem['kind']): string => {
  if (kind === 'category') return '<path d="M4 5h16v4H4V5Zm0 6h7v8H4v-8Zm9 0h7v8h-7v-8Z" fill="currentColor"/>';
  if (kind === 'guide') return '<path d="M5 4h14a1 1 0 0 1 1 1v14H7a3 3 0 0 0-2 .76V4Zm2 3v2h9V7H7Zm0 4v2h9v-2H7Zm0 4v2h6v-2H7Z" fill="currentColor"/>';
  if (kind === 'problem') return '<path d="M12 3 2.8 19h18.4L12 3Zm1 12h-2v-2h2v2Zm0-4h-2V7h2v4Z" fill="currentColor"/>';
  return '<path d="M4 4h16v16H4V4Zm3 3v4h4V7H7Zm6 0v4h4V7h-4Zm-6 6v4h4v-4H7Zm6 0v4h4v-4h-4Z" fill="currentColor"/>';
};

export function mountHeroSearch(root: HTMLElement, items: SearchItem[]): void {
  const form = root.querySelector<HTMLFormElement>('form');
  const input = root.querySelector<HTMLInputElement>('input[type="search"]');
  const panel = root.querySelector<HTMLElement>('[data-search-panel]');
  const list = root.querySelector<HTMLUListElement>('ul');
  const count = root.querySelector<HTMLElement>('[data-result-count]');
  if (!form || !input || !panel || !list) return;

  const kindLabel: Record<string, string> = {
    product: 'Ürün',
    category: 'Kategori',
    guide: 'Rehber',
    problem: 'Problem',
    template: 'Şablon',
  };

  const close = (): void => {
    panel.hidden = true;
    list.innerHTML = '';
    input.setAttribute('aria-expanded', 'false');
    if (count) count.textContent = '0 sonuç';
  };

  const open = (): void => {
    panel.hidden = false;
    input.setAttribute('aria-expanded', 'true');
  };

  const renderCardList = (displayItems: SearchItem[]): string =>
    displayItems
      .map((result, index) => {
        const kind = result.kind ?? 'product';
        const kindText = kindLabel[kind] ?? 'Ürün';
        return `
          <li>
            <a href="${escapeHtml(result.url)}" data-result-index="${index}">
              <span class="hero-search__result-icon" aria-hidden="true">
                <svg viewBox="0 0 24 24" width="18" height="18">${iconForKind(kind)}</svg>
              </span>
              <span class="hero-search__result-copy">
                <span class="hero-search__name">${escapeHtml(result.name)}</span>
                <span class="hero-search__summary">${escapeHtml(result.summary)}</span>
              </span>
              <span class="hero-search__result-side">
                <span class="hero-search__kind">${escapeHtml(kindText)}</span>
                <span class="hero-search__category">${escapeHtml(result.category)}</span>
                <span class="hero-search__arrow" aria-hidden="true">→</span>
              </span>
            </a>
          </li>`;
      })
      .join('');

  const render = (query: string): void => {
    const rawQuery = (query || '').trim();
    root.classList.toggle('has-query', rawQuery.length > 0);

    if (!rawQuery) {
      if (count) count.textContent = `${Math.min(5, items.length)} öneri`;
      list.innerHTML = renderCardList(items.slice(0, 5));
      open();
      return;
    }

    const response = searchEngine(items, rawQuery, { limit: 5 });
    const { results, didYouMean, suggestedItems } = response;

    let html = '';

    // "Bunu mu demek istediniz?" banner
    if (didYouMean) {
      html += `
        <li class="hero-search__suggestion-item">
          <button type="button" class="hero-search__did-you-mean" data-suggest-term="${escapeHtml(didYouMean)}">
            <span class="hero-search__dym-label">Bunu mu demek istediniz?</span>
            <strong class="hero-search__dym-term">“${escapeHtml(didYouMean)}”</strong>
            <span class="hero-search__dym-action">Uygula ↵</span>
          </button>
        </li>
      `;
    }

    if (results.length > 0) {
      if (count) count.textContent = `${results.length} sonuç`;
      html += renderCardList(results);
    } else if (suggestedItems.length > 0) {
      if (count) count.textContent = `${suggestedItems.length} yakın eşleşme`;
      html += `
        <li class="hero-search__info-banner">
          <span>“<strong>${escapeHtml(rawQuery)}</strong>” için doğrudan sonuç bulunamadı. İlgili en yakın sistemler:</span>
        </li>
      `;
      html += renderCardList(suggestedItems);
    } else {
      if (count) count.textContent = '0 sonuç';
      html += `
        <li class="hero-search__empty">
          “${escapeHtml(rawQuery)}” için eşleşme bulunamadı. Nakit akışı, cari, kasa, stok veya kârlılık gibi bir terim deneyin.
        </li>
      `;
    }

    list.innerHTML = html;

    // Attach click listeners to suggestion buttons
    const suggestBtn = list.querySelector<HTMLButtonElement>('[data-suggest-term]');
    if (suggestBtn) {
      suggestBtn.addEventListener('click', (e) => {
        e.preventDefault();
        const term = suggestBtn.dataset.suggestTerm;
        if (term) {
          input.value = term;
          render(term);
          input.focus();
        }
      });
    }

    open();
  };

  input.addEventListener('input', () => render(input.value));
  input.addEventListener('focus', () => render(input.value));

  input.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      close();
      input.blur();
      return;
    }

    if (event.key === 'Enter') {
      const highlighted = list.querySelector<HTMLAnchorElement>('a[data-result-index].is-highlighted');
      if (highlighted) {
        event.preventDefault();
        window.location.href = highlighted.href;
        return;
      }
    }

    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    const links = Array.from(list.querySelectorAll<HTMLAnchorElement>('a[data-result-index]'));
    if (links.length === 0) return;
    event.preventDefault();
    const active = list.querySelector<HTMLAnchorElement>('a[data-result-index].is-highlighted');
    const activeIndex = active ? links.indexOf(active) : -1;
    const nextIndex = event.key === 'ArrowDown'
      ? (activeIndex + 1) % links.length
      : (activeIndex - 1 + links.length) % links.length;

    links.forEach((link, index) => {
      const isActive = index === nextIndex;
      link.classList.toggle('is-highlighted', isActive);
      link.setAttribute('aria-selected', String(isActive));
      if (isActive) link.scrollIntoView({ block: 'nearest' });
    });
  });

  document.addEventListener('keydown', (event) => {
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'k') {
      event.preventDefault();
      input.focus();
      input.select();
    }
  });

  document.addEventListener('click', (event) => {
    if (!root.contains(event.target as Node)) close();
  });

  form.addEventListener('submit', (event) => {
    const highlighted = list.querySelector<HTMLAnchorElement>('a[data-result-index].is-highlighted');
    if (highlighted) {
      event.preventDefault();
      window.location.href = highlighted.href;
      return;
    }
    if ((input.value || '').trim().length < 2) {
      event.preventDefault();
      render(input.value);
    }
  });
}
