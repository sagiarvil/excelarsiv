#!/usr/bin/env node
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const root = process.cwd();

function mergeCatalog(basePath, extraPath) {
  if (!existsSync(extraPath)) return;
  const base = JSON.parse(readFileSync(basePath, 'utf8'));
  const extra = JSON.parse(readFileSync(extraPath, 'utf8'));
  const merged = {
    tiers: base.tiers,
    products: { ...base.products, ...(extra.products ?? {}) },
  };
  writeFileSync(basePath, `${JSON.stringify(merged, null, 2)}\n`);
}

mergeCatalog(join(root, 'commerce', 'catalog.json'), join(root, 'commerce', 'catalog-extra.json'));
mergeCatalog(join(root, 'functions', 'catalog.json'), join(root, 'functions', 'catalog-extra.json'));
console.log('Extra product catalogs materialized.');
