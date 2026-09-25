import type { APIRoute } from 'astro';
import { getTemplateSearchIndex } from '../lib/templates';

export const prerender = true;

export const GET: APIRoute = async () => {
  const items = await getTemplateSearchIndex();
  return new Response(JSON.stringify(items), {
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'public, max-age=3600, s-maxage=86400, stale-while-revalidate=604800',
    },
  });
};
