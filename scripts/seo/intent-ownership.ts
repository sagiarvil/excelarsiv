import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));
const EXIT = Object.freeze({ PASS: 0, BLOCK: 1, CONFIG: 4 });

type RecordRow = {
  route?: string;
  status?: string;
  primaryQueryClusterId?: string | null;
  ownerRoute?: string | null;
  type?: string;
};

export function validateIntentOwnership(records: RecordRow[]): string[] {
  const errors: string[] = [];
  const owners = new Map<string, string[]>();
  for (const record of records) {
    if (record.status !== 'live') continue;
    const cluster = record.primaryQueryClusterId;
    if (!cluster) continue;
    if (record.ownerRoute !== record.route) continue;
    const list = owners.get(cluster) ?? [];
    list.push(String(record.route));
    owners.set(cluster, list);
  }
  for (const [cluster, routes] of owners) {
    if (routes.length > 1) errors.push(`INTENT_COLLISION: ${cluster} -> ${routes.join(', ')}`);
  }
  return errors;
}

function main(): void {
  try {
    const registry = JSON.parse(readFileSync(resolve(ROOT, 'data/seo/registry/excelarsiv_seo_registry.json'), 'utf8')) as { records: RecordRow[] };
    const errors = validateIntentOwnership(registry.records);
    if (errors.length) {
      console.error(errors.join('\n'));
      process.exit(EXIT.BLOCK);
    }
    console.log('INTENT OWNERSHIP PASS');
    process.exit(EXIT.PASS);
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(EXIT.CONFIG);
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
