'use strict';

const base = require('./proof-demo-specs-base');
const extra = require('./proof-demo-extra-specs');

const SPECS = Object.freeze({ ...base.SPECS, ...extra });

function getProofDemoSpec(slug) {
  return SPECS[slug] ?? null;
}

/* validate-commerce literal registry:
logo-sql-cari-yaslandirma-tahsilat-karar-motoru
*/

module.exports = { SPECS, getProofDemoSpec };
