import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));
const EXIT = Object.freeze({ PASS: 0, BLOCK: 1, CONFIG: 4 });

export const FORBIDDEN_ROUTES = ['/excel-araclari', '/paketler'] as const;

const SOURCE_SURFACES = ['src/pages', 'src/components', 'src/layouts'] as const;
const DIST_ARTIFACTS = ['sitemap.xml', 'sitemap-pages.xml', 'sitemap-products.xml', 'llms.txt', 'llms-full.txt', 'ai.txt'] as const;

const DEMO_JARGON = [
  /indexlenebilir\s+demo\s+kapısı/i,
  /SEO\s+açılış/i,
  /SEO\s+kapısı/i,
  /E-E-A-T/,
  /(?:^|[\s"'(>])intent(?:[\s"'<.,]|$)/i,
  /(?:^|[\s"'(>])SERP(?:[\s"'<.,]|$)/,
  /Proof Demo/,
  /premium sistem/i,
  /(?:^|[\s"'(>])moat(?:[\s"'<.,]|$)/i,
];

type Redirect = { from: string; to: string; type: number };
type Ledger = { redirects: Redirect[] };
type FirebaseRedirect = { source?: string; destination?: string; type?: number };
type FirebaseHosting = { target?: string; redirects?: FirebaseRedirect[] };
type FirebaseConfig = { hosting?: FirebaseHosting | FirebaseHosting[] };
type Registry = { records: Array<{ route?: string; status?: string; redirectTarget?: string | null }> };

function hostingTargets(config: FirebaseConfig): FirebaseHosting[] {
  const hosting = config.hosting;
  if (Array.isArray(hosting)) return hosting;
  return hosting ? [hosting] : [];
}

function walkFiles(dir: string, acc: string[] = []): string[] {
  if (!existsSync(dir)) return acc;
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    if (statSync(full).isDirectory()) walkFiles(full, acc);
    else acc.push(full);
  }
  return acc;
}

function posix(path: string): string {
  return relative(ROOT, path).split(sep).join('/');
}

function hrefPattern(route: string): RegExp {
  const escaped = route.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`href=["']${escaped}(?:/)?["']`);
}

export function validateForbiddenRoutes(root = ROOT, distDir = join(root, 'dist'), options: { checkDist?: boolean } = {}): string[] {
  const errors: string[] = [];
  const ledger = JSON.parse(readFileSync(join(root, 'data/seo/redirects.json'), 'utf8')) as Ledger;
  const firebase = JSON.parse(readFileSync(join(root, 'firebase.json'), 'utf8')) as FirebaseConfig;
  const registry = JSON.parse(readFileSync(join(root, 'data/seo/registry/excelarsiv_seo_registry.json'), 'utf8')) as Registry;
  const site = hostingTargets(firebase).find((item) => item.target === 'excelarsiv') ?? hostingTargets(firebase)[0];
  const firebaseRedirects = site?.redirects ?? [];

  for (const route of FORBIDDEN_ROUTES) {
    const pagePath = join(root, 'src/pages', `${route.slice(1)}.astro`);
    if (existsSync(pagePath)) errors.push(`YASAKLI_ROUTE_KAYNAK: ${posix(pagePath)}`);

    const rule = ledger.redirects.find((item) => item.from === route);
    if (!rule) errors.push(`YASAKLI_ROUTE_LEDGER_EKSIK: ${route}`);
    else if (rule.type !== 301) errors.push(`YASAKLI_ROUTE_LEDGER_301_DEGIL: ${route}`);

    const hosted = firebaseRedirects.find((item) => item.source === route);
    if (!hosted) errors.push(`YASAKLI_ROUTE_FIREBASE_EKSIK: ${route}`);
    else if (hosted.type !== 301) errors.push(`YASAKLI_ROUTE_FIREBASE_301_DEGIL: ${route}`);
    else if (rule && hosted.destination !== rule.to) {
      errors.push(`YASAKLI_ROUTE_HEDEF_UYUMSUZ: ${route} ledger=${rule.to} firebase=${hosted.destination}`);
    }

    const live = registry.records.find((record) => record.route === route && record.status === 'live');
    if (live) errors.push(`YASAKLI_ROUTE_REGISTRY_LIVE: ${route}`);
  }

  for (const surface of SOURCE_SURFACES) {
    for (const file of walkFiles(join(root, surface))) {
      if (!/\.(astro|ts|js|mjs|cjs)$/.test(file)) continue;
      const source = readFileSync(file, 'utf8');
      const rel = posix(file);
      for (const route of FORBIDDEN_ROUTES) {
        if (hrefPattern(route).test(source)) errors.push(`YASAKLI_IC_LINK: ${rel} -> ${route}`);
      }
      if (rel !== 'src/pages/excel-araclari.astro' && /Excel Araçları/.test(source)) {
        errors.push(`YASAKLI_NAV_METIN: ${rel} -> Excel Araçları`);
      }
      if (/>\s*Paketler\s*</.test(source) || /label:\s*'Paketler'/.test(source)) {
        errors.push(`YASAKLI_NAV_METIN: ${rel} -> Paketler`);
      }
    }
  }

  const demoSurfaces = [
    ...walkFiles(join(root, 'src/pages/demo')),
    join(root, 'src/pages/hakkinda.astro'),
    join(root, 'src/components/home/MoatStrip.astro'),
    join(root, 'src/pages/hesaplayici/asgari-ucret-zam-etkisi.astro'),
    join(root, 'src/pages/sektor/[slug].astro'),
  ];
  for (const file of demoSurfaces) {
    if (!existsSync(file)) continue;
    const source = readFileSync(file, 'utf8');
    const rel = posix(file);
    for (const pattern of DEMO_JARGON) {
      if (pattern.test(source)) errors.push(`DEMO_JARGON: ${rel} -> ${pattern}`);
    }
  }

  if (options.checkDist !== false && existsSync(distDir)) {
    for (const name of DIST_ARTIFACTS) {
      const path = join(distDir, name);
      if (!existsSync(path)) continue;
      const text = readFileSync(path, 'utf8');
      for (const route of FORBIDDEN_ROUTES) {
        if (text.includes(`${route}<`) || text.includes(`${route}"`) || text.includes(`${route}\n`) || text.includes(`excelarsiv.com${route}`)) {
          errors.push(`YASAKLI_ARTIFACT: ${name} -> ${route}`);
        }
      }
    }
    for (const route of FORBIDDEN_ROUTES) {
      const html = join(distDir, `${route.slice(1)}.html`);
      const indexed = join(distDir, route.slice(1), 'index.html');
      if (existsSync(html) || existsSync(indexed)) errors.push(`YASAKLI_DIST_SAYFA: ${route}`);
    }
  }

  return errors;
}

function main(): void {
  try {
    const sourceOnly = process.argv.includes('--source-only');
    const errors = validateForbiddenRoutes(ROOT, join(ROOT, 'dist'), { checkDist: !sourceOnly });
    if (errors.length) {
      console.error(errors.join('\n'));
      process.exit(EXIT.BLOCK);
    }
    console.log(`FORBIDDEN ROUTES PASS — ${FORBIDDEN_ROUTES.join(', ')}`);
    process.exit(EXIT.PASS);
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(EXIT.CONFIG);
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
