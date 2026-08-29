export interface SearchItem {
  name: string;
  summary: string;
  category: string;
  url: string;
  kind?: 'product' | 'category' | 'guide' | 'problem' | 'template';
  keywords?: string | string[];
  priceTL?: number;
  categorySlug?: string;
  [key: string]: unknown;
}

export interface SearchMatch<T extends SearchItem = SearchItem> {
  item: T;
  score: number;
  matchType: 'exact' | 'prefix' | 'contains' | 'token' | 'fuzzy' | 'intent';
  highlightedName?: string;
}

export interface SearchResponse<T extends SearchItem = SearchItem> {
  results: T[];
  matches: SearchMatch<T>[];
  totalMatches: number;
  didYouMean: string | null;
  suggestedItems: T[];
  originalQuery: string;
  cleanQuery: string;
  intentTokens: string[];
}

export interface SearchOptions {
  limit?: number;
  minScore?: number;
  enableFuzzy?: boolean;
  enableIntentParsing?: boolean;
  dictionary?: string[];
}

const TURKISH_STOP_WORDS = new Set([
  'nasil', 'nedir', 'ne', 'icin', 'yapmak', 'istiyorum', 'hangi', 'nerede',
  'hesaplama', 'hesaplamasi', 'hesaplanir', 'yapilir', 'kullanabilirim',
  'var', 'mi', 'mu', 'mu', 'sablonu', 'sablonlari', 'tablosu', 'formulu',
  'ornegi', 'programi', 'sistemi', 'dosyasi', 'araci', 'listesi', 'takibi',
  'takip', 'etmek', 'icin', 've', 'veya', 'ile', 'de', 'da', 'miyim', 'midir',
  'olan', 'olarak', 'gibi', 'kadar', 'cok', 'az', 'en', 'ise', 'bu', 'su', 'o'
]);

export function normalizeTurkish(text: string, options: { deAccent?: boolean; removePunctuation?: boolean } = {}): string {
  if (!text) return '';
  let res = text
    .toLocaleLowerCase('tr-TR')
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '');

  if (options.deAccent ?? true) {
    res = res
      .replace(/[ıİ]/g, 'i')
      .replace(/[ğĞ]/g, 'g')
      .replace(/[şŞ]/g, 's')
      .replace(/[çÇ]/g, 'c')
      .replace(/[öÖ]/g, 'o')
      .replace(/[üÜ]/g, 'u')
      .replace(/[âÂ]/g, 'a')
      .replace(/[îÎ]/g, 'i')
      .replace(/[ûÛ]/g, 'u');
  }

  if (options.removePunctuation ?? true) {
    res = res.replace(/[^a-z0-9\s-]/gi, ' ').replace(/\s+/g, ' ').trim();
  }

  return res;
}

export function levenshteinDistance(a: string, b: string): number {
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  const matrix: number[][] = [];

  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }

  return matrix[b.length][a.length];
}

export function stringSimilarity(a: string, b: string): number {
  const normA = normalizeTurkish(a);
  const normB = normalizeTurkish(b);
  if (normA === normB) return 1.0;
  if (!normA || !normB) return 0.0;

  const maxLen = Math.max(normA.length, normB.length);
  const dist = levenshteinDistance(normA, normB);
  return Math.max(0, 1 - dist / maxLen);
}

export function extractSearchIntent(query: string): {
  normalizedQuery: string;
  cleanTokens: string[];
  intentTokens: string[];
} {
  const normalizedQuery = normalizeTurkish(query);
  const rawTokens = normalizedQuery.split(/\s+/).filter(Boolean);
  const intentTokens = rawTokens.filter((token) => !TURKISH_STOP_WORDS.has(token) && token.length > 1);
  const cleanTokens = intentTokens.length > 0 ? intentTokens : rawTokens;

  return {
    normalizedQuery,
    cleanTokens,
    intentTokens,
  };
}

export const COMMON_DOMAIN_KEYWORDS: string[] = [
  'nakit akisi', 'kasa defteri', 'cari hesap', 'stok takibi', 'karlilik',
  'amortisman', 'bordro', 'kredi covenant', 'asgari ucret', 'doviz riski',
  'cek senet', 'mutabakat', 'faiz hesaplama', 'hakedis', 'vergi', 'sgk',
  'butce planlama', 'e-fatura', 'pos komisyon', 'trendyol', 'restoran recete',
  'personel maliyeti', 'kidem ihbar', 'defter beyan', 'avans takip', 'uretim maliyeti'
];

export function findDidYouMeanSuggestion(
  query: string,
  corpus: string[]
): { suggestion: string; score: number } | null {
  const normQuery = normalizeTurkish(query);
  if (normQuery.length < 2) return null;

  let bestSuggestion = '';
  let bestScore = 0;

  for (const candidate of corpus) {
    const normCandidate = normalizeTurkish(candidate);
    if (!normCandidate || normCandidate === normQuery) continue;

    if (normCandidate.startsWith(normQuery) && normQuery.length >= 3) {
      const score = 0.85 + (normQuery.length / normCandidate.length) * 0.1;
      if (score > bestScore) {
        bestScore = score;
        bestSuggestion = candidate;
      }
      continue;
    }

    const maxLen = Math.max(normQuery.length, normCandidate.length);
    const dist = levenshteinDistance(normQuery, normCandidate);
    const maxAllowedDist = normQuery.length >= 7 ? 3 : (normQuery.length >= 4 ? 2 : 1);

    if (dist <= maxAllowedDist) {
      const score = 1 - dist / maxLen;
      if (score > bestScore && score >= 0.6) {
        bestScore = score;
        bestSuggestion = candidate;
      }
    }
  }

  if (bestScore >= 0.6 && bestSuggestion) {
    return { suggestion: bestSuggestion, score: bestScore };
  }

  return null;
}

export function buildSearchCorpus<T extends SearchItem>(items: T[]): string[] {
  const corpusSet = new Set<string>(COMMON_DOMAIN_KEYWORDS);

  for (const item of items) {
    if (item.name) corpusSet.add(item.name);
    if (item.category) corpusSet.add(item.category);
    if (item.keywords) {
      const kw = Array.isArray(item.keywords) ? item.keywords : item.keywords.split(/[,;\s]+/);
      for (const k of kw) {
        const trimmed = k.trim();
        if (trimmed.length > 2) corpusSet.add(trimmed);
      }
    }
  }

  return Array.from(corpusSet);
}

export function searchEngine<T extends SearchItem>(
  items: T[],
  query: string,
  options: SearchOptions = {}
): SearchResponse<T> {
  const limit = options.limit ?? 6;
  const enableFuzzy = options.enableFuzzy ?? true;
  const rawQuery = (query || '').trim();

  if (!rawQuery) {
    return {
      results: items.slice(0, limit),
      matches: items.slice(0, limit).map((item) => ({ item, score: 1, matchType: 'exact' })),
      totalMatches: items.length,
      didYouMean: null,
      suggestedItems: [],
      originalQuery: rawQuery,
      cleanQuery: '',
      intentTokens: [],
    };
  }

  const { normalizedQuery, cleanTokens, intentTokens } = extractSearchIntent(rawQuery);
  const corpus = options.dictionary && options.dictionary.length > 0
    ? options.dictionary
    : buildSearchCorpus(items);

  const scoredMatches: SearchMatch<T>[] = [];

  for (const item of items) {
    const normName = normalizeTurkish(item.name);
    const normCategory = normalizeTurkish(item.category || '');
    const normSummary = normalizeTurkish(item.summary || '');
    const normKeywords = normalizeTurkish(
      Array.isArray(item.keywords) ? item.keywords.join(' ') : (item.keywords || '')
    );

    let score = 0;
    let matchType: SearchMatch['matchType'] = 'contains';

    if (normName === normalizedQuery) {
      score += 150;
      matchType = 'exact';
    } else if (normName.startsWith(normalizedQuery)) {
      score += 90;
      matchType = 'prefix';
    } else if (normName.includes(normalizedQuery)) {
      score += 70;
      matchType = 'contains';
    } else if (normCategory.includes(normalizedQuery)) {
      score += 45;
      matchType = 'contains';
    } else if (normKeywords.includes(normalizedQuery)) {
      score += 40;
      matchType = 'contains';
    } else if (normSummary.includes(normalizedQuery)) {
      score += 25;
      matchType = 'contains';
    }

    let tokenMatches = 0;
    for (const token of cleanTokens) {
      if (token.length < 2) continue;

      if (normName.includes(token)) {
        score += 20;
        tokenMatches++;
        matchType = matchType === 'contains' ? 'token' : matchType;
      } else if (normCategory.includes(token)) {
        score += 12;
        tokenMatches++;
      } else if (normKeywords.includes(token)) {
        score += 10;
        tokenMatches++;
      } else if (normSummary.includes(token)) {
        score += 6;
        tokenMatches++;
      } else if (enableFuzzy && token.length >= 3) {
        const nameTokens = normName.split(/\s+/);
        for (const nameWord of nameTokens) {
          if (nameWord.length >= 3) {
            const dist = levenshteinDistance(token, nameWord);
            const maxL = Math.max(token.length, nameWord.length);
            if (dist <= 2 && dist / maxL <= 0.35) {
              score += 35 - dist * 10;
              tokenMatches++;
              matchType = 'fuzzy';
              break;
            }
          }
        }
      }
    }

    if (cleanTokens.length > 1 && tokenMatches >= cleanTokens.length) {
      score += 30;
      matchType = 'intent';
    }

    if (score === 0 && enableFuzzy && normalizedQuery.length >= 4) {
      const dist = levenshteinDistance(normalizedQuery, normName);
      const maxAllowed = normalizedQuery.length >= 8 ? 3 : 2;
      if (dist <= maxAllowed) {
        score = Math.max(10, 40 - dist * 8);
        matchType = 'fuzzy';
      }
    }

    if (score > (options.minScore ?? 0)) {
      scoredMatches.push({ item, score, matchType });
    }
  }

  scoredMatches.sort((a, b) => b.score - a.score || a.item.name.localeCompare(b.item.name, 'tr-TR'));

  const directResults = scoredMatches.slice(0, limit).map((m) => m.item);

  let didYouMean: string | null = null;
  let suggestedItems: T[] = [];

  const isPoorMatch = scoredMatches.length === 0 || (scoredMatches[0] && scoredMatches[0].score < 50);
  if (isPoorMatch && normalizedQuery.length >= 3) {
    const suggestionObj = findDidYouMeanSuggestion(normalizedQuery, corpus);
    if (suggestionObj && normalizeTurkish(suggestionObj.suggestion) !== normalizedQuery) {
      didYouMean = suggestionObj.suggestion;

      const rerun = searchEngine(items, didYouMean, {
        limit: limit,
        enableFuzzy: false,
        enableIntentParsing: true,
      });
      suggestedItems = rerun.results;
    }
  }

  if (directResults.length === 0 && suggestedItems.length === 0 && cleanTokens.length > 0) {
    const fallbackMatches: SearchMatch<T>[] = [];
    for (const item of items) {
      const normName = normalizeTurkish(item.name);
      let bestSim = 0;
      for (const token of cleanTokens) {
        for (const nameWord of normName.split(/\s+/)) {
          const sim = stringSimilarity(token, nameWord);
          if (sim > bestSim) bestSim = sim;
        }
      }
      if (bestSim >= 0.5) {
        fallbackMatches.push({
          item,
          score: Math.round(bestSim * 100),
          matchType: 'fuzzy',
        });
      }
    }
    fallbackMatches.sort((a, b) => b.score - a.score);
    suggestedItems = fallbackMatches.slice(0, limit).map((m) => m.item);
  }

  return {
    results: directResults,
    matches: scoredMatches.slice(0, limit),
    totalMatches: scoredMatches.length,
    didYouMean,
    suggestedItems,
    originalQuery: rawQuery,
    cleanQuery: cleanTokens.join(' '),
    intentTokens,
  };
}

export class UniversalSearchEngine<T extends SearchItem> {
  private items: T[];
  private corpus: string[];

  constructor(items: T[], customDictionary?: string[]) {
    this.items = items;
    this.corpus = customDictionary && customDictionary.length > 0
      ? customDictionary
      : buildSearchCorpus(items);
  }

  public updateItems(newItems: T[]): void {
    this.items = newItems;
    this.corpus = buildSearchCorpus(newItems);
  }

  public search(query: string, options?: SearchOptions): SearchResponse<T> {
    return searchEngine(this.items, query, {
      dictionary: this.corpus,
      ...options,
    });
  }

  public suggest(query: string): string | null {
    const res = findDidYouMeanSuggestion(query, this.corpus);
    return res ? res.suggestion : null;
  }
}
