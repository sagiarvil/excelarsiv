/**
 * connect_projects.js — config/projects ile messages-suite Eşleme ve Kural Bağlantı Aracı
 * 
 * Portatif Standart: os.homedir() ve ortam değişkenleri ile dinamik çözümlenir.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');

const USER_HOME = process.env.USERPROFILE || process.env.HOME || os.homedir();
const GEMINI_ROOT = path.join(USER_HOME, '.gemini');
const PROJECTS_CONFIG_DIR = path.join(GEMINI_ROOT, 'config', 'projects');
const GLOBAL_MESSAGES_DIR = path.join(GEMINI_ROOT, 'config', 'plugins', 'gemini-messages-suite', 'messages');
const DESKTOP_MESSAGES_DIR = path.join(USER_HOME, 'Desktop', 'plugins', 'gemini', 'gemini-messages-suite', 'messages');

function mkdirp(d) {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
}

function writeIfNotExists(filePath, content) {
  if (!fs.existsSync(filePath)) {
    fs.writeFileSync(filePath, content.trim() + '\n', 'utf8');
  }
}

console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log('🔗 CONFIG/PROJECTS ↔ MESSAGES-SUITE KURAL BAĞLANTISI');
console.log('Kullanıcı Ev Dizini: ' + USER_HOME);
console.log('Projeler Dizini    : ' + PROJECTS_CONFIG_DIR);
console.log('Hafıza Dizini      : ' + GLOBAL_MESSAGES_DIR);
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

mkdirp(GLOBAL_MESSAGES_DIR);
mkdirp(DESKTOP_MESSAGES_DIR);
mkdirp(path.join(GLOBAL_MESSAGES_DIR, 'projects'));
mkdirp(path.join(DESKTOP_MESSAGES_DIR, 'projects'));

if (fs.existsSync(PROJECTS_CONFIG_DIR)) {
  const projectFiles = fs.readdirSync(PROJECTS_CONFIG_DIR).filter(f => f.endsWith('.json'));
  console.log(`Toplam ${projectFiles.length} kayıtlı proje bulundu.\n`);

  const connectedProjects = [];

  projectFiles.forEach(file => {
    const filePath = path.join(PROJECTS_CONFIG_DIR, file);
    try {
      const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
      const projectName = (data.name || path.basename(file, '.json')).replace(/[^a-zA-Z0-9_\-+]/g, '_').toLowerCase();
      
      const pGlobalDir = path.join(GLOBAL_MESSAGES_DIR, 'projects', projectName);
      const pDesktopDir = path.join(DESKTOP_MESSAGES_DIR, 'projects', projectName);
      mkdirp(pGlobalDir);
      mkdirp(pDesktopDir);

      // 1. PROJE_TALİMATLARI.md
      const talimatContent = `# Proje Talimatları — ${data.name || projectName}
Son güncelleme: 2026-09-11

Bu proje kurallarını ve yönergelerini messages-suite üzerinden alır.

## 1. PROJE ÖZEL KURALLARI
- Bu projede yapılan tüm geliştirmeler %100 Türkçe açıklanır.
- BOM kesinlikle yasaktır; tüm dosyalar BOM-suz saf UTF-8 olarak kaydedilir.
- Kod değişikliklerinden sonra test ve error_checker ile doğrulama zorunludur.

## 2. KULLANICI TALİMATLARI
(Kullanıcı bu projeye özel talimat verdiğinde buraya otomatik eklenir)
| Tarih | Talimat | Durum |
|---|---|---|
| 2026-09-11 | Proje messages-suite kural sistemine bağlandı | Aktif |
`;
      writeIfNotExists(path.join(pGlobalDir, 'PROJE_TALİMATLARI.md'), talimatContent);
      writeIfNotExists(path.join(pDesktopDir, 'PROJE_TALİMATLARI.md'), talimatContent);

      // 2. _OKU.txt (Taşınabilir format)
      const okuContent = `PROJE ADI : ${data.name || projectName}
DURUM     : 🟢 AKTİF (messages-suite bağlı)
GÜNCELLEME: 2026-09-11
KURAL YOLU: ~/.gemini/config/plugins/gemini-messages-suite/messages/projects/${projectName}/PROJE_TALİMATLARI.md
HAFIZA    : Bu proje oturum özetlerini ve talimatlarını messages-suite üzerinden takip eder.
`;
      fs.writeFileSync(path.join(pGlobalDir, '_OKU.txt'), okuContent, 'utf8');
      fs.writeFileSync(path.join(pDesktopDir, '_OKU.txt'), okuContent, 'utf8');

      // 3. KONUSMA.md
      const konusmaContent = `# Konuşma ve Oturum Geçmişi — ${data.name || projectName}
- 2026-09-11: Proje messages-suite hafıza sistemine bağlandı.
`;
      writeIfNotExists(path.join(pGlobalDir, 'KONUSMA.md'), konusmaContent);
      writeIfNotExists(path.join(pDesktopDir, 'KONUSMA.md'), konusmaContent);

      // 4. PLAN.md
      const planContent = `# PLAN: ${data.name || projectName}
> **Durum:** 🟢 Hazır | **Prensip:** Az Token · Çok İş · Cerrahi Müdahale

## 🎯 1. HEDEF VE KAPSAM
- Hedef: ${data.name || projectName} geliştirmeleri
- Standart: %100 Türkçe · BOM-suz UTF-8 · Sıfır Dolgu

## 🔍 2. MEVCUT ALTYAPIDAN YENİDEN KULLANILACAKLAR (Token Tasarrufu)
- [ ] Mevcut çalışan controller/route kalıbı
- [ ] Mevcut UI bileşeni veya şablon sınıfı

## ✂️ 3. CERRAHİ GÖREV KIRILIMI
- [ ] Adım 1: Görev planlama ve keşif
- [ ] Adım 2: Cerrahi kod müdahalesi
- [ ] Adım 3: 4 Kalite kapısı denetimi (Syntax, BOM, Güvenlik, Kanıt)

## 🛡️ 4. KALİTE KAPILARI (4/4 PASS)
- [ ] Syntax PASS | BOM-suz PASS | Güvenlik PASS | Kanıt PASS
`;
      writeIfNotExists(path.join(pGlobalDir, 'PLAN.md'), planContent);
      writeIfNotExists(path.join(pDesktopDir, 'PLAN.md'), planContent);

      // 5. Proje JSON dosyasına dinamik kural yollarını bağla
      data.messagesSuite = {
        enabled: true,
        plugin: "gemini-messages-suite",
        universalRulesPath: path.join(GLOBAL_MESSAGES_DIR, 'rules', '01_EVRENSEL_KURALLAR.md').replace(/\\/g, '/'),
        userDirectivesPath: path.join(GLOBAL_MESSAGES_DIR, 'rules', 'KULLANICI_TALİMATLARI.md').replace(/\\/g, '/'),
        projectDirectivesPath: path.join(pGlobalDir, 'PROJE_TALİMATLARI.md').replace(/\\/g, '/'),
        projectStatusPath: path.join(pGlobalDir, '_OKU.txt').replace(/\\/g, '/')
      };

      data.customInstructions = `Bu proje kurallarını, kullanıcı direktiflerini ve hafıza kayıtlarını messages-suite eklentisinden alır.
Kurallar: ${data.messagesSuite.universalRulesPath}
Kişisel Talimatlar: ${data.messagesSuite.userDirectivesPath}
Proje Talimatları: ${data.messagesSuite.projectDirectivesPath}`;

      fs.writeFileSync(filePath, JSON.stringify(data, null, 2) + '\n', 'utf8');
      connectedProjects.push({ name: data.name || projectName, file });
      console.log(`✅ Bağlandı: ${(data.name || projectName).padEnd(25)} -> messages/projects/${projectName}`);
    } catch (err) {
      console.error(`❌ Hata (${file}):`, err.message);
    }
  });

  // 6. INDEX.md dosyasını güncelle
  const indexContent = `# 🗂️ MESSAGES SUITE — PROJE VE OTURUM İNDEKSİ

> **Merkezi Hafıza Konumu:** \`~/.gemini/config/plugins/gemini-messages-suite/messages/\`  
> **Son Güncelleme:** 2026-09-11  
> **Toplam Bağlı Proje:** ${connectedProjects.length}

---

## 📋 Kayıtlı Projeler ve Kural Dosyaları Haritası

| # | Proje Adı | Hafıza ve Talimat Klasörü | Durum |
|---|---|---|---|
${connectedProjects.map((p, idx) => `| ${idx + 1} | **${p.name}** | \`projects/${p.name.replace(/[^a-zA-Z0-9_\-+]/g, '_').toLowerCase()}/\` | 🟢 Bağlı |`).join('\n')}

---
*Bu indeks, config/projects ile messages-suite arasındaki canlı köprü tarafından otomatik yönetilir.*
`;

  fs.writeFileSync(path.join(GLOBAL_MESSAGES_DIR, 'INDEX.md'), indexContent, 'utf8');
  fs.writeFileSync(path.join(DESKTOP_MESSAGES_DIR, 'INDEX.md'), indexContent, 'utf8');
  console.log('\n📄 INDEX.md başarıyla güncellendi.');
  console.log(`🎉 BAŞARILI: Toplam ${connectedProjects.length} proje kural sistemine bağlandı!\n`);
}
