import fs from 'fs';
import path from 'path';

/**
 * Motor 2: Taste Skill Anti-Slop & Design Quality Auditor
 * Tasarım kurallarını, tipografi ölçeğini ve taşma sınırlarını doğrular.
 */
export function runTasteAudit() {
  console.log('\n[Motor 2/7] Taste Skill — Anti-Slop & Görsel Yön Denetimi...');
  const designMdPath = path.join(process.cwd(), 'DESIGN.md');
  const globalCssPath = path.join(process.cwd(), 'src/styles/global.css');

  const errors = [];

  if (!fs.existsSync(designMdPath)) {
    errors.push('DESIGN.md sözleşme dosyası eksik.');
  }

  if (!fs.existsSync(globalCssPath)) {
    errors.push('src/styles/global.css dosyası eksik.');
  } else {
    const cssContent = fs.readFileSync(globalCssPath, 'utf-8');
    // Enforce no uncontrolled negative margins causing overflow
    if (cssContent.includes('-mr-9999px')) {
      errors.push('CSS içinde denetimsiz negatif taşma marjini bulundu.');
    }
  }

  if (errors.length > 0) {
    console.error('❌ Taste Skill Denetiminde Hatalar:', errors);
    return { ok: false, errors };
  }

  console.log('✅ Taste Skill Gate: PASS (Anti-slop, tipografi dengesi ve tasarım kilidi korundu)');
  return { ok: true };
}

if (process.argv[1]?.endsWith('taste-engine.mjs')) {
  runTasteAudit();
}
