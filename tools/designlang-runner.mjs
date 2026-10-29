import { execSync } from 'child_process';
import fs from 'fs';
import path from 'path';

/**
 * Motor 1: designlang Entegrasyonu
 * Sitenin tasarım token'larını, tipografisini ve renklerini W3C formatında denetler.
 */
export function runDesignLang(targetUrl = 'https://excelarsiv.com') {
  console.log('\n[Motor 1/7] designlang — Tasarım Dili ve Token Analizi...');
  try {
    const output = execSync(`designlang stats ${targetUrl} 2>&1`, { encoding: 'utf-8' });
    console.log(output.trim());
    return { ok: true, output };
  } catch (error) {
    console.warn('designlang canlı tarama uyarısı (fallback modu):', error.message);
    return { ok: true, fallback: true };
  }
}

if (process.argv[1]?.endsWith('designlang-runner.mjs')) {
  runDesignLang();
}
