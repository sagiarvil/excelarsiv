import { KATALOG_EKRAN_ODAK } from '../data/katalog-ekran-odak.ts';

export interface CatalogShot {
  src: string;
  alt: string;
}

export interface CatalogShotOdak {
  x: number;
  y: number;
  s: number;
}

const ODAK = KATALOG_EKRAN_ODAK;
const DEFAULT_ODAK: CatalogShotOdak = { x: 50, y: 28, s: 1.62 };

export function screenshotSlug(src: string): string | undefined {
  const match = src.match(/\/screenshots\/(.+)-\d\.(?:png|webp)$/i);
  return match?.[1];
}

export function catalogShotOdak(src: string): CatalogShotOdak {
  const slug = screenshotSlug(src);
  const kayit = slug ? ODAK[slug] : undefined;
  if (!kayit) return DEFAULT_ODAK;
  return { x: kayit.x, y: kayit.y, s: kayit.s };
}

/** Katalog kartı: piksel doluluğu en yüksek kare (ızgara), boş PANO değil. */
export function pickCatalogScreenshot(shots: CatalogShot[]): CatalogShot | undefined {
  if (shots.length === 0) return undefined;
  const slug = screenshotSlug(shots[0].src);
  const n = slug ? ODAK[slug]?.n : undefined;
  if (n) {
    const hit = shots.find((shot) => new RegExp(`-${n}\\.(png|webp)$`, 'i').test(shot.src));
    if (hit) return hit;
  }
  return (
    shots.find((shot) => /-2\.(png|webp)$/i.test(shot.src)) ??
    shots.find((shot) => /-1\.(png|webp)$/i.test(shot.src)) ??
    shots[shots.length - 1]
  );
}

export function catalogScreenshotAlt(name: string): string {
  return `${name} gerçek Excel karar ekranı`;
}
