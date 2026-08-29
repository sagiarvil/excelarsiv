export type {
  SearchItem,
  SearchMatch,
  SearchResponse,
  SearchOptions,
} from './search-engine';

export {
  normalizeTurkish,
  levenshteinDistance,
  stringSimilarity,
  extractSearchIntent,
  findDidYouMeanSuggestion,
  buildSearchCorpus,
  searchEngine,
  UniversalSearchEngine,
  COMMON_DOMAIN_KEYWORDS,
} from './search-engine';

import { searchEngine, type SearchItem } from './search-engine';

export function searchTemplates(items: SearchItem[], query: string, limit = 6): SearchItem[] {
  const res = searchEngine(items, query, { limit });
  return res.results;
}
