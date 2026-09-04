import { growthEvents, trackGrowthEvent } from '../lib/growth.ts';

export interface FinderProduct {
  slug: string;
  name: string;
  summary: string;
  category: string;
  categoryName: string;
  priceTL: number;
  url: string;
  demoUrl: string;
}

export interface FinderAnswers {
  alan: string;
  tip: string;
  hacim: string;
  donem: string;
  problem: string;
}

const HERO: Record<string, string[]> = {
  'nakit-akisi': ['13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi', 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi'],
  'muhasebe-ve-vergi': ['cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi', 'kdv-tevkifat-mahsup-iade-listesi'],
  'stok-ve-uretim': ['stok-satis-ve-nakit-baglanma-sistemi', 'uretim-recetesi-ve-zam-yansitma-hesaplayici'],
  'satis-ve-fiyatlama': ['pos-komisyon-ve-net-tahsilat-kontrol-sistemi', 'trendyol-komisyon-sonrasi-net-kar'],
  'personel-ve-bordro': ['kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici', 'kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi'],
  'finansal-analiz': ['aylik-patron-finans-paneli', 'proje-ve-is-bazinda-gercek-karlilik-sistemi'],
  'butce-ve-planlama': ['kobi-finans-yonetim-paketi', '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi'],
};

const PROBLEM_SLUG: Record<string, string> = {
  nakit: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
  tahsilat: 'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi',
  kar: 'proje-ve-is-bazinda-gercek-karlilik-sistemi',
  vergi: 'kdv-tevkifat-mahsup-iade-listesi',
  stok: 'stok-satis-ve-nakit-baglanma-sistemi',
  personel: 'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici',
};

export function secUrunler(products: FinderProduct[], answers: FinderAnswers): FinderProduct[] {
  const bySlug = new Map(products.map((item) => [item.slug, item]));
  const ranked: string[] = [];
  const problemSlug = PROBLEM_SLUG[answers.problem];
  if (problemSlug) ranked.push(problemSlug);
  for (const slug of HERO[answers.alan] ?? []) ranked.push(slug);
  if (answers.donem === 'gun') ranked.unshift('akilli-kasa-defteri-ve-nakit-kontrol-sistemi');
  if (answers.donem === 'hafta') ranked.unshift('13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi');
  if (answers.tip === 'eticaret') ranked.unshift('trendyol-komisyon-sonrasi-net-kar');
  if (answers.tip === 'uretim') ranked.unshift('uretim-recetesi-ve-zam-yansitma-hesaplayici');
  if (answers.tip === 'saha') ranked.unshift('insaat-hakedis-santiye-maliyet');
  const unique: FinderProduct[] = [];
  for (const slug of ranked) {
    const product = bySlug.get(slug);
    if (product && !unique.some((item) => item.slug === product.slug)) unique.push(product);
    if (unique.length === 3) break;
  }
  if (unique.length === 0) return products.filter((item) => item.category === answers.alan).slice(0, 3);
  return unique.slice(0, 3);
}

export function mountUrunBulucu(root: HTMLElement, products: FinderProduct[]): void {
  const form = root.querySelector<HTMLFormElement>('[data-finder-form]');
  const output = root.querySelector<HTMLElement>('[data-finder-sonuc]');
  if (!form || !output) return;

  trackGrowthEvent(growthEvents.toolView, { tool_id: 'akilli-excel-urun-bulucu', page_type: 'hero_tool' });
  let started = false;
  form.addEventListener('change', () => {
    if (started) return;
    started = true;
    trackGrowthEvent(growthEvents.toolStart, { tool_id: 'akilli-excel-urun-bulucu', page_type: 'hero_tool' });
  }, { once: true });

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const data = new FormData(form);
    const answers: FinderAnswers = {
      alan: String(data.get('alan') ?? ''),
      tip: String(data.get('tip') ?? ''),
      hacim: String(data.get('hacim') ?? ''),
      donem: String(data.get('donem') ?? ''),
      problem: String(data.get('problem') ?? ''),
    };
    const selected = secUrunler(products, answers);
    const primary = selected[0];

    trackGrowthEvent(growthEvents.toolComplete, {
      tool_id: 'akilli-excel-urun-bulucu',
      intent: answers.problem || answers.alan,
      product: primary?.slug || 'none',
      lead_score: selected.length ? 70 : 20,
    });
    trackGrowthEvent(growthEvents.toolResult, {
      tool_id: 'akilli-excel-urun-bulucu',
      tool_result: primary?.slug || 'no_match',
      product: primary?.slug || 'none',
      intent: answers.problem || answers.alan,
    });

    output.hidden = false;
    output.innerHTML = selected
      .map((item, index) => `
        <a class="${index === 0 ? 'is-ana' : ''}" href="${item.url}" data-growth-result-product="${item.slug}">
          <small>${index === 0 ? 'Ana sistem' : 'İlgili sistem'} · ${item.categoryName}</small>
          <strong>${item.name}</strong>
          <span>${item.summary}</span>
        </a>`)
      .join('');
  });

  output.addEventListener('click', (event) => {
    const anchor = (event.target as Element | null)?.closest<HTMLAnchorElement>('[data-growth-result-product]');
    if (!anchor) return;
    trackGrowthEvent(growthEvents.ctaClick, {
      cta_id: 'finder-product-result',
      tool_id: 'akilli-excel-urun-bulucu',
      product: anchor.dataset.growthResultProduct || 'unknown',
    });
  });
}
