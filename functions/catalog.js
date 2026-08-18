'use strict';

// The Functions source is deployed as the isolated `functions/` directory.
// Keep deployment-local catalog mirrors so Cloud Functions never depends on
// parent-directory files omitted from the uploaded source package.
const baseCatalog = require('./catalog.json');
const extraCatalog = require('./catalog-extra.json');
const catalog = {
  tiers: baseCatalog.tiers,
  products: { ...baseCatalog.products, ...extraCatalog.products },
};
if (!catalog?.tiers || !catalog?.products) {
  throw new Error('Packaged commerce catalog is incomplete');
}

const TIERS = Object.freeze(
  Object.fromEntries(
    Object.entries(catalog.tiers).map(([name, value]) => [name, Object.freeze({ ...value })]),
  ),
);

const PRODUCTS = Object.freeze(
  Object.fromEntries(
    Object.entries(catalog.products).map(([slug, value]) => {
      const tier = TIERS[value.tier];
      if (!tier) throw new Error(`Unknown tier ${value.tier} for ${slug}`);
      return [slug, Object.freeze({ ...value, priceTL: tier.priceTL })];
    }),
  ),
);

function getTierForPrice(priceTL) {
  return Object.entries(TIERS).find(([, value]) => value.priceTL === Number(priceTL))?.[0] ?? null;
}

function getTierByProductId(productId) {
  const normalized = String(productId ?? '').trim();
  return Object.entries(TIERS).find(([, value]) => value.shopierProductId === normalized)?.[0] ?? null;
}

module.exports = { TIERS, PRODUCTS, getTierForPrice, getTierByProductId };
