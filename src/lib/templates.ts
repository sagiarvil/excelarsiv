import { getCollection, type CollectionEntry } from 'astro:content';
import { pickCatalogScreenshot } from './catalog-screenshot';
import { kapakYolu } from './kapak';
import { categories, getCategoryName, type CategorySlug } from './categories';
import { shopierUrlForPrice } from './shopier';
import type { SearchItem } from './search';

export interface TemplateSheetMap {
  name: string;
  kind: 'input' | 'calculation' | 'output';
}

export interface TemplatePreview {
  src: string;
  alt: string;
}

export type ScreenshotFocus = 'result' | 'top' | 'center';

export interface TemplateViewModel {
  slug: string;
  name: string;
  summary: string;
  categorySlug: CategorySlug;
  categoryName: string;
  priceTL: number;
  sheetCount: number;
  version?: string;
  fileFormat?: 'xlsx' | 'xlsm';
  sheetMap?: TemplateSheetMap[];
  outputs?: string[];
  preview?: TemplatePreview;
  kapak?: string;
  screenshotFocus: ScreenshotFocus;
  url: string;
  shopierUrl: string;
}

export type TemplateEntry = CollectionEntry<'templates'>;

export function toTemplateViewModel(entry: TemplateEntry): TemplateViewModel {
  const data = entry.data;
  return {
    slug: entry.id,
    name: data.name,
    summary: data.summary,
    categorySlug: data.category,
    categoryName: getCategoryName(data.category),
    priceTL: data.priceTL,
    sheetCount: data.sheetCount,
    version: data.version,
    fileFormat: data.fileFormat,
    sheetMap: data.sheetMap,
    outputs: data.outputs,
    preview: pickCatalogScreenshot(data.screenshots),
    kapak: kapakYolu(entry.id),
    screenshotFocus: 'result',
    url: `/sablon/${entry.id}`,
    shopierUrl: shopierUrlForPrice(data.priceTL),
  };
}

export async function getAllTemplates(): Promise<TemplateViewModel[]> {
  const entries = await getCollection('templates');
  return entries.map(toTemplateViewModel);
}

export async function getFeaturedTemplates(count = 6): Promise<TemplateViewModel[]> {
  const all = await getAllTemplates();
  return all.slice(0, count);
}

export async function getTemplatesByCategory(categorySlug: string): Promise<TemplateViewModel[]> {
  const all = await getAllTemplates();
  return all.filter((t) => t.categorySlug === categorySlug);
}

export async function getTemplateSearchIndex(): Promise<SearchItem[]> {
  const all = await getAllTemplates();
  const guides = await getCollection('guides');
  const products: SearchItem[] = all.map((t) => ({
    name: t.name,
    summary: t.summary,
    category: t.categoryName,
    url: t.url,
    kind: 'product',
    keywords: `${t.categorySlug} ${t.outputs?.join(' ') ?? ''}`,
  }));
  const categoryItems: SearchItem[] = categories.map((category) => ({
    name: `${category.name} Excel şablonları`,
    summary: category.description,
    category: category.name,
    url: `/sablonlar/${category.slug}`,
    kind: 'category',
    keywords: category.primaryQuery,
  }));
  const guideItems: SearchItem[] = guides.map((guide) => ({
    name: guide.data.title,
    summary: guide.data.description,
    category: getCategoryName(guide.data.category),
    url: `/rehber/${guide.id}`,
    kind: 'guide',
    keywords: guide.data.primaryQuery,
  }));
  const problems: SearchItem[] = [
    { name: 'Tahsilat ve müşteri riski', summary: 'Cari hesap, vade ve tahsilat kontrolü', category: 'problem', url: '/sablon/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi', kind: 'problem', keywords: 'tahsilat cari risk vade' },
    { name: 'Nakit akışı planı', summary: '13 haftalık ödeme ve nakit açığı', category: 'problem', url: '/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi', kind: 'problem', keywords: 'nakit ödeme planı kasa' },
    { name: 'POS net tahsilat', summary: 'Komisyon sonrası gerçek tahsilat', category: 'problem', url: '/sablon/pos-komisyon-ve-net-tahsilat-kontrol-sistemi', kind: 'problem', keywords: 'pos komisyon tahsilat' },
  ];
  return [...products, ...categoryItems, ...guideItems, ...problems];
}
