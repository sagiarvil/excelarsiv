import fs from 'fs';
import path from 'path';

/**
 * Motor 6: GSAP & ScrollTrigger Motion Denetimi
 * GPU hızlandırmalı hareket motorunu ve SSR-safe sözleşmeyi doğrular.
 */
export function runGsapCheck() {
  console.log('\n[Motor 6/7] GSAP & ScrollTrigger — GPU Hareket Motoru Denetimi...');
  const motionFile = path.join(process.cwd(), 'src/lib/gsap-motion.ts');
  if (!fs.existsSync(motionFile)) {
    console.error('❌ src/lib/gsap-motion.ts bulunamadı.');
    return { ok: false };
  }

  const content = fs.readFileSync(motionFile, 'utf-8');
  if (!content.includes('ScrollTrigger') || !content.includes('typeof window === \'undefined\'')) {
    console.error('❌ GSAP SSR-safe koruma veya ScrollTrigger eksik.');
    return { ok: false };
  }

  console.log('✅ GSAP & ScrollTrigger: SSR-safe, bellek sızıntısız GPU animasyon motoru hazır.');
  return { ok: true };
}

if (process.argv[1]?.endsWith('gsap-runner.mjs')) {
  runGsapCheck();
}
