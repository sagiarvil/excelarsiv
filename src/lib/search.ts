export interface SearchItem {
  name: string;
  summary: string;
  category: string;
  url: string;
  kind?: 'product' | 'category' | 'guide' | 'problem';
  keywords?: string;
}

const normalize = (text: string): string => text.toLocaleLowerCase('tr-TR');

export function searchTemplates(items: SearchItem[], query: string, limit = 6): SearchItem[] {
  const q = normalize(query.trim());
  if (q.length < 2) return [];
  return items
    .filter((item) => {
      const haystack = `${item.name} ${item.summary} ${item.category} ${item.keywords ?? ''}`;
      return normalize(haystack).includes(q);
    })
    .slice(0, limit);
}
