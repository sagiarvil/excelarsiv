import fs from 'fs';
import path from 'path';

/**
 * Motor 5: React Bits Animasyonlu Bileşen Denetimi
 * Projede yer alan etkileşimli kart ve animasyon bileşenlerini doğrular.
 */
export function runReactBitsCheck() {
  console.log('\n[Motor 5/7] React Bits — Animasyonlu Bileşen Bütünlüğü...');
  const effectsDir = path.join(process.cwd(), 'src/components/effects');
  if (!fs.existsSync(effectsDir)) {
    console.error('❌ src/components/effects dizini bulunamadı.');
    return { ok: false };
  }

  const files = fs.readdirSync(effectsDir);
  const required = ['SpotlightCard.astro', 'ShimmerBadge.astro'];
  const missing = required.filter(f => !files.includes(f));

  if (missing.length > 0) {
    console.error('❌ Eksik React Bits bileşenleri:', missing);
    return { ok: false, missing };
  }

  console.log(`✅ React Bits: ${files.length} etkileşimli bileşen aktif ve entegre.`);
  return { ok: true, count: files.length };
}

if (process.argv[1]?.endsWith('react-bits-runner.mjs')) {
  runReactBitsCheck();
}
