export interface SearchItem {
  name: string;
  summary: string;
  category: string;
  url: string;
  kind?: 'product' | 'category' | 'guide' | 'problem';
  keywords?: string;
}

const escapeHtml = (value: string): string =>
  value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');

const normalize = (text: string): string => text.toLocaleLowerCase('tr-TR');

export function mountHeroSearch(root: HTMLElement, items: SearchItem[]): void {
  const form = root.querySelector<HTMLFormElement>('form');
  const input = root.querySelector<HTMLInputElement>('input[type="search"]');
  const panel = root.querySelector<HTMLElement>('[data-search-panel]');
  const list = root.querySelector<HTMLUListElement>('ul');
  if (!form || !input || !panel || !list) return;

  const kindLabel: Record<string, string> = {
    product: 'Ürün',
    category: 'Kategori',
    guide: 'Rehber',
    problem: 'Problem',
  };

  const close = (): void => {
    panel.hidden = true;
    list.innerHTML = '';
    input.setAttribute('aria-expanded', 'false');
  };

  const open = (): void => {
    panel.hidden = false;
    input.setAttribute('aria-expanded', 'true');
  };

  const renderItems = (results: SearchItem[], query = ''): void => {
    if (results.length === 0) {
      list.innerHTML = `<li class="hero-search__empty">“${escapeHtml(query.trim())}” için sonuç bulunamadı. Kasa, cari, nakit veya stok gibi bir terim deneyin.</li>`;
      open();
      return;
    }

    list.innerHTML = results
      .map(
        (result, index) => `
          <li>
            <a href="${result.url}" data-result-index="${index}">
              <span class="hero-search__result-main">
                <span class="hero-search__result-dot" aria-hidden="true"></span>
                <span>
                  <span class="hero-search__name">${escapeHtml(result.name)}</span>
                  <span class="hero-search__summary">${escapeHtml(result.summary)}</span>
                </span>
              </span>
              <span class="hero-search__meta">${escapeHtml(kindLabel[result.kind ?? 'product'] ?? 'Ürün')} · ${escapeHtml(result.category)}</span>
            </a>
          </li>`
      )
      .join('');
    open();
  };

  const render = (query: string): void => {
    const q = normalize(query.trim());
    root.classList.toggle('has-query', q.length > 0);

    if (!q) {
      renderItems(items.slice(0, 5));
      return;
    }

    const results = items
      .filter((item) => {
        const haystack = `${item.name} ${item.summary} ${item.category} ${item.keywords ?? ''}`;
        return normalize(haystack).includes(q);
      })
      .slice(0, 6);

    renderItems(results, query);
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
    if (normalize(input.value).length < 2) {
      event.preventDefault();
      render(input.value);
    }
  });
}
