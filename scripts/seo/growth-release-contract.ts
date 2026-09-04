import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));
const growth = readFileSync(resolve(ROOT, 'src/lib/growth.ts'), 'utf8');
const finder = readFileSync(resolve(ROOT, 'src/scripts/urun-bulucu.ts'), 'utf8');
const analytics = readFileSync(resolve(ROOT, 'src/lib/analytics.ts'), 'utf8');

const required = [
  'tool_view','tool_start','tool_complete','tool_result','cta_click','lead_submit',
  'product_view','checkout_start','purchase','revenue_recorded'
];
const errors: string[] = [];
for (const event of required) if (!growth.includes(`'${event}'`)) errors.push(`EVENT_MISSING:${event}`);
for (const token of ['anonymous_user_id','session_id','traffic_source','first_touch','last_touch','event_id','timestamp']) {
  if (!growth.includes(token)) errors.push(`CONTEXT_MISSING:${token}`);
}
for (const token of ['growthEvents.toolView','growthEvents.toolStart','growthEvents.toolComplete','growthEvents.toolResult']) {
  if (!finder.includes(token)) errors.push(`FINDER_INSTRUMENTATION_MISSING:${token}`);
}
if (!analytics.includes('GROWTH_ALIAS')) errors.push('ANALYTICS_BRIDGE_MISSING');
for (const forbidden of ['email:', 'phone:', 'tax_id:', 'tc_kimlik:']) {
  if (growth.includes(forbidden)) errors.push(`PII_FIELD_FORBIDDEN:${forbidden}`);
}
if (errors.length) {
  errors.forEach((error) => console.error(`FAIL ${error}`));
  process.exit(1);
}
console.log('GROWTH_RELEASE_CONTRACT_PASS — consent-aware attribution + hero tool + checkout bridge');
