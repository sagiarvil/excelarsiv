import { analytics } from '../../config/analytics.ts';
import { trackAnalyticsEvent } from '../analytics.ts';

// Kart tıklamasını mevcut GA4 sözleşmesi üzerinden, onay kapısından geçirerek gönderir.
// Varyant parametresi A/B ölçümü için kullanılır.

export function izleUrunKartiTiklamasi(slug: string): boolean {
  return trackAnalyticsEvent(analytics.events.templateCardClick, {
    templateSlug: slug,
    variant: 'product-story',
  });
}
