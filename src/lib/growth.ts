import { hasAnalyticsConsent } from './consent.ts';

export const growthEvents = Object.freeze({
  sessionStart: 'session_start',
  pageView: 'page_view',
  contentEngaged: 'content_engaged',
  toolView: 'tool_view',
  toolStart: 'tool_start',
  toolStep: 'tool_step',
  toolAbandon: 'tool_abandon',
  toolComplete: 'tool_complete',
  toolResult: 'tool_result',
  ctaView: 'cta_view',
  ctaClick: 'cta_click',
  leadStart: 'lead_start',
  leadSubmit: 'lead_submit',
  leadQualified: 'lead_qualified',
  leadDisqualified: 'lead_disqualified',
  whatsappClick: 'whatsapp_click',
  phoneClick: 'phone_click',
  appointmentStart: 'appointment_start',
  appointmentBooked: 'appointment_booked',
  productView: 'product_view',
  comparisonView: 'comparison_view',
  checkoutStart: 'checkout_start',
  purchase: 'purchase',
  upsell: 'upsell',
  repeatPurchase: 'repeat_purchase',
  revenueRecorded: 'revenue_recorded',
} as const);

export type GrowthEventName = typeof growthEvents[keyof typeof growthEvents];
export type GrowthPayload = Record<string, string | number | boolean | null | undefined>;

declare global {
  interface Window {
    gtag?: (...args: unknown[]) => void;
  }
}

const FIRST_TOUCH_KEY = 'excelarsiv:growth:first-touch:v1';
const USER_KEY = 'excelarsiv:growth:anonymous-user:v1';
const SESSION_KEY = 'excelarsiv:growth:session:v1';

function randomId(prefix: string): string {
  const uuid = typeof crypto !== 'undefined' && 'randomUUID' in crypto ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  return `${prefix}_${uuid}`;
}

function readStorage(storage: Storage, key: string): string | null {
  try { return storage.getItem(key); } catch { return null; }
}

function writeStorage(storage: Storage, key: string, value: string): void {
  try { storage.setItem(key, value); } catch { /* privacy mode */ }
}

function parseTouch(url = window.location.href, referrer = document.referrer) {
  const current = new URL(url);
  const source = current.searchParams.get('utm_source') || inferSource(referrer);
  return {
    traffic_source: source,
    medium: current.searchParams.get('utm_medium') || inferMedium(source),
    campaign: current.searchParams.get('utm_campaign') || '',
    referrer: referrer || '',
  };
}

function inferSource(referrer: string): string {
  if (!referrer) return 'direct';
  try {
    const host = new URL(referrer).hostname.toLowerCase();
    if (host.includes('google.')) return 'google';
    if (host.includes('bing.')) return 'bing';
    if (host.includes('chatgpt.com')) return 'chatgpt.com';
    if (host.includes('perplexity.ai')) return 'perplexity.ai';
    if (host.includes('gemini.google.com')) return 'gemini';
    return host.replace(/^www\./, '');
  } catch { return 'referral'; }
}

function inferMedium(source: string): string {
  if (source === 'direct') return 'direct';
  if (['google', 'bing'].includes(source)) return 'organic';
  if (['chatgpt.com', 'perplexity.ai', 'gemini'].includes(source)) return 'ai_referral';
  return 'referral';
}

function getIdentity() {
  let anonymousUserId = readStorage(window.localStorage, USER_KEY);
  if (!anonymousUserId) {
    anonymousUserId = randomId('au');
    writeStorage(window.localStorage, USER_KEY, anonymousUserId);
  }
  let sessionId = readStorage(window.sessionStorage, SESSION_KEY);
  if (!sessionId) {
    sessionId = randomId('ss');
    writeStorage(window.sessionStorage, SESSION_KEY, sessionId);
  }
  return { anonymous_user_id: anonymousUserId, session_id: sessionId };
}

function getAttribution() {
  const current = parseTouch();
  let firstTouch = readStorage(window.localStorage, FIRST_TOUCH_KEY);
  if (!firstTouch) {
    firstTouch = JSON.stringify(current);
    writeStorage(window.localStorage, FIRST_TOUCH_KEY, firstTouch);
  }
  let first: ReturnType<typeof parseTouch> = current;
  try { first = JSON.parse(firstTouch) as ReturnType<typeof parseTouch>; } catch { /* use current */ }
  return {
    traffic_source: current.traffic_source,
    medium: current.medium,
    campaign: current.campaign,
    referrer: current.referrer,
    first_touch: `${first.traffic_source}/${first.medium}${first.campaign ? `/${first.campaign}` : ''}`,
    last_touch: `${current.traffic_source}/${current.medium}${current.campaign ? `/${current.campaign}` : ''}`,
  };
}

export function trackGrowthEvent(eventName: GrowthEventName, payload: GrowthPayload = {}): boolean {
  if (typeof window === 'undefined' || !hasAnalyticsConsent()) return false;
  const context = {
    event_id: randomId('ev'),
    site: 'excelarsiv.com',
    page: window.location.pathname,
    timestamp: new Date().toISOString(),
    ...getIdentity(),
    ...getAttribution(),
    ...payload,
  };
  window.dispatchEvent(new CustomEvent('sagiarvil:growth', { detail: { event: eventName, ...context } }));
  if (typeof window.gtag === 'function') window.gtag('event', eventName, context);
  return true;
}
