import { runDesignLang } from './designlang-runner.mjs';
import { runTasteAudit } from './taste-engine.mjs';
import { runImpeccable } from './impeccable-runner.mjs';
import { runTransitionsCheck } from './transitions-runner.mjs';
import { runReactBitsCheck } from './react-bits-runner.mjs';
import { runGsapCheck } from './gsap-runner.mjs';
import { runPlaywrightCheck } from './playwright-runner.mjs';

/**
 * 7 MOTORLU MASTER TASARIM MÜHENDİSLİĞİ ZİNCİRİ (DESIGN ENGINE CHAIN)
 * Deterministic Pipeline & Verification Gate
 */
async function main() {
  console.log('================================================================');
  console.log('🚀 7 MOTORLU TASARIM MÜHENDİSLİĞİ ZİNCİRİ BAŞLATILIYOR (V1.0)');
  console.log('   Tüm motorlar sırayla ve birbirine bağlı olarak doğrulanıyor.');
  console.log('================================================================');

  const results = [];

  // Motor 1: designlang
  const m1 = runDesignLang();
  results.push({ name: 'designlang (v13.3.2)', status: m1.ok ? 'PASS' : 'FAIL' });

  // Motor 2: Taste Skill
  const m2 = runTasteAudit();
  results.push({ name: 'Taste Skill (Anti-Slop Gate)', status: m2.ok ? 'PASS' : 'FAIL' });

  // Motor 3: Impeccable
  const m3 = runImpeccable();
  results.push({ name: 'Impeccable (v4.8.0)', status: m3.ok ? 'PASS' : 'FAIL' });

  // Motor 4: transitions.dev
  const m4 = runTransitionsCheck();
  results.push({ name: 'transitions.dev (v0.3.0)', status: m4.ok ? 'PASS' : 'FAIL' });

  // Motor 5: React Bits
  const m5 = runReactBitsCheck();
  results.push({ name: 'React Bits (Effects)', status: m5.ok ? 'PASS' : 'FAIL' });

  // Motor 6: GSAP
  const m6 = runGsapCheck();
  results.push({ name: 'GSAP Motion & ScrollTrigger', status: m6.ok ? 'PASS' : 'FAIL' });

  // Motor 7: Playwright
  const m7 = await runPlaywrightCheck();
  results.push({ name: 'Playwright CLI (v1.63.0)', status: m7.ok ? 'PASS' : 'FAIL' });

  console.log('\n================================================================');
  console.log('📊 7 MOTORLU ZİNCİR ENTEGRASYON SONUCU:');
  console.log('================================================================');
  let allPass = true;
  for (const r of results) {
    const icon = r.status === 'PASS' ? '✅' : '❌';
    console.log(`${icon} ${r.name.padEnd(35)} : ${r.status}`);
    if (r.status !== 'PASS') allPass = false;
  }
  console.log('================================================================');

  if (!allPass) {
    console.error('❌ ZİNCİRDE EKSİK VEYA BAŞARISIZ MOTOR BULUNDU.');
    process.exit(1);
  } else {
    console.log('🏆 7/7 MOTOR EKSİKSİZ, AKTİF VE ZİNCİRLEME BAĞLI OLARAK DOĞRULANDI.');
  }
}

main().catch((err) => {
  console.error('Fatal Chain Error:', err);
  process.exit(1);
});
