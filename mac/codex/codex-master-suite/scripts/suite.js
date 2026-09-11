/**
 * suite.js — AI Suite Ajan, Beceri, Araç ve Otonom Güncelleme Katalog Raporlayıcı
 * 
 * Tüm modeller ve terminal oturumları için Uzman Ajanları, Çekirdek Becerileri,
 * Deterministik CLI Araçlarını ve Otonom Güncelleme (Self-Evolution) motorunu listeler.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');

const USER_HOME = process.env.USERPROFILE || process.env.HOME || os.homedir();
const SCRIPT_DIR = __dirname;
const SUITE_ROOT = path.resolve(SCRIPT_DIR, '..');
const MANIFEST_PATH = path.join(SUITE_ROOT, 'plugin.json');

let manifest = {};
if (fs.existsSync(MANIFEST_PATH)) {
  try { manifest = JSON.parse(fs.readFileSync(MANIFEST_PATH, 'utf8')); } catch (e) {}
}

const currentVersion = manifest.version || '1.0.0';
const pluginName = manifest.name || path.basename(SUITE_ROOT);

console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log('🎛️ ANTIGRAVITY & AI SUITE — TAM AJAN, BECERİ VE ARAÇ KATALOĞU (/suite)');
console.log('Eklenti Adı   : ' + pluginName);
console.log('Aktif Sürüm   : ' + currentVersion);
console.log('Sistem Konumu : ' + SUITE_ROOT);
console.log('Kullanıcı     : ' + USER_HOME);
console.log('Model Kapsamı : Tüm Modeller (Gemini, Claude, GPT, Cursor vb.)');
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

// 1. Ajanları Tara
const agentsDir = path.join(SUITE_ROOT, 'agents');
const agents = [];
if (fs.existsSync(agentsDir)) {
  fs.readdirSync(agentsDir).forEach(name => {
    const p = path.join(agentsDir, name, 'agent.md');
    if (fs.existsSync(p)) {
      const txt = fs.readFileSync(p, 'utf8');
      const desc = (txt.match(/description:\s*(.+)/) || [])[1] || '';
      const skill = (txt.match(/skills:\s*(.+)/) || [])[1] || '';
      agents.push({ name, desc: desc.trim(), skill: skill.trim() });
    }
  });
}

// 2. Becerileri Tara
const skillsDir = path.join(SUITE_ROOT, 'skills');
const skills = [];
if (fs.existsSync(skillsDir)) {
  fs.readdirSync(skillsDir).forEach(name => {
    const p = path.join(skillsDir, name, 'SKILL.md');
    if (fs.existsSync(p)) {
      const txt = fs.readFileSync(p, 'utf8');
      const desc = (txt.match(/description:\s*>?\s*(.+)/) || [])[1] || '';
      skills.push({ name, desc: desc.trim() });
    }
  });
}

// 3. CLI Betiklerini Tara
const scriptsDir = path.join(SUITE_ROOT, 'scripts');
const scripts = [];
if (fs.existsSync(scriptsDir)) {
  fs.readdirSync(scriptsDir).forEach(f => {
    if (f.endsWith('.js') || f.endsWith('.py') || f.endsWith('.ps1')) {
      scripts.push(f);
    }
  });
}

console.log('🤖 1. UZMAN AJANLAR (' + agents.length + ' Adet):\n');
agents.forEach((a, i) => {
  console.log('  [' + (i + 1).toString().padStart(2, '0') + '] ' + a.name.padEnd(28) + ' | Beceri: ' + a.skill.padEnd(24) + ' | ' + a.desc.substring(0, 65) + '...');
});

console.log('\n⚡ 2. ÇEKİRDEK BECERİLER (' + skills.length + ' Adet):\n');
skills.forEach((s, i) => {
  console.log('  [' + (i + 1).toString().padStart(2, '0') + '] /' + s.name.padEnd(28) + ' | ' + s.desc.substring(0, 75) + '...');
});

console.log('\n🛠️ 3. DETERMINISTIC CLI ARAÇLARI (' + scripts.length + ' Adet):\n');
const keyScripts = [
  { name: 'self_updater.js', desc: 'Otonom eklenti güncelleme, versiyon artırma (SemVer) ve canlı senkronizasyon' },
  { name: 'error_checker.js', desc: 'Tüm JS/JSON/PHP dosyaları için AST sözdizimi ve UTF-8 BOM denetimi' },
  { name: 'bom_utf8_scan.js', desc: 'UTF-8 BOM baytlarını otomatik tarama ve sıfırlama (--fix)' },
  { name: 'audit_engine_v3.js', desc: '18 Motorlu Engine V3.0 SEO & GEO derinlik denetleyici' },
  { name: 'schema_validator.js', desc: 'JSON-LD şema ve yapısal veri doğrulayıcı' },
  { name: 'heading_semantics_check.js', desc: 'H1-H6 başlık hiyerarşisi ve semantik etiket denetleyici' },
  { name: 'install_plugins.js', desc: 'Taşınabilir çoklu AI eklenti kurulum ve tescil aracı' },
  { name: 'suite.js', desc: 'Tam yetenek kataloğu ve çalışma yönergesi raporlayıcı' }
];

keyScripts.forEach((ks, i) => {
  console.log('  [' + (i + 1).toString().padStart(2, '0') + '] ' + ks.name.padEnd(26) + ' -> ' + ks.desc);
});

console.log('\n🔄 4. OTONOM GÜNCELLEME VE KENDİNİ GELİŞTİRME MOTORU (Self-Evolution):');
console.log('  • Proje Sırasında Canlı Güncelleme: Geliştirici veya Ajan eksik gördüğü bir aracı/kuralı doğrudan ekler.');
console.log('  • Patch Artışı (Hata Düzeltme)   : node scripts/self_updater.js --patch "Düzeltme mesajı"');
console.log('  • Minor Artışı (Yeni Yetenek)   : node scripts/self_updater.js --minor "Yeni özellik mesajı"');
console.log('  • Çift Yönlü Senkronizasyon     : Masaüstü (Desktop/plugins) ve Küresel (.gemini/config/plugins) anında eşitlenir.');
console.log('  • Otomatik Tescil               : agy plugin install ile yeni sürüm anında sisteme kaydedilir.');


console.log('\n🧬 5. OTONOM DİL SENTEZLEME VE AZ TOKEN MİMARİSİ (Language Synthesizer):');
console.log('  • Tek Komutla Dil Ekleme : node scripts/language_synthesizer.js <dil_adi>');
console.log('  • Desteklenen Ekosistemler: python, typescript, golang, rust, csharp, ruby, kotlin vb.');
console.log('  • Otomatik 3 Katman      : rules/<lang>-rules.md | skills/ag-<lang>-developer | agents/<lang>-developer');
console.log('  • Çift Dilli Mimari       : Ajana talimatlar yoğun İngilizce (Token tasarrufu & sıfır hata), kullanıcıya %100 Türkçe!');
console.log('  • Otonom Tescil          : self_updater.js tetiklenir, sürüm artırılır ve agy ile sisteme kaydedilir.');

console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log('✅ /suite kataloğu başarıyla raporlandı.');
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
