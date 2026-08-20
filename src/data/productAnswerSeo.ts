import { productSeo, type ProductSeoEntry } from './productSeo';

export type ProductIntent = 'commercial-investigation';

export interface ProductAnswerSeoEntry extends ProductSeoEntry {
  intent: ProductIntent;
  answerQuestion: string;
  answerSummary: string;
}

function buildAnswerQuestion(primaryQuery: string): string {
  const normalized = primaryQuery.trim();
  return `${normalized.charAt(0).toLocaleUpperCase('tr-TR')}${normalized.slice(1)} hangi ihtiyacı çözer?`;
}

export const productAnswerSeo: Record<string, ProductAnswerSeoEntry> = Object.fromEntries(
  Object.entries(productSeo).map(([slug, entry]) => [
    slug,
    {
      ...entry,
      intent: 'commercial-investigation' as const,
      answerQuestion: buildAnswerQuestion(entry.primaryQuery),
      answerSummary: entry.description,
    },
  ]),
);

export function getProductAnswerSeo(slug: string): ProductAnswerSeoEntry | null {
  return productAnswerSeo[slug] ?? null;
}
