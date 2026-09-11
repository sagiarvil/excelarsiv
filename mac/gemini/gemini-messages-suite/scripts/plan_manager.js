/**
 * plan_manager.js — Az Token Çok İş Plan Yönetim ve Çoklu-AI Çözümleme Motoru
 * 
 * Standart: Karpathy Disiplini · Sıfır Dolgu · Cerrahi Müdahale · Çoklu AI Desteği
 */
'use strict';

const fs = require('fs');
const path = require('path');

const os = require("os");
const USER_HOME = process.env.USERPROFILE || process.env.HOME || os.homedir();
const GLOBAL_MESSAGES = path.join(USER_HOME, ".gemini", "config", "plugins", "gemini-messages-suite", "messages");
const DESKTOP_MESSAGES = path.join(USER_HOME, "Desktop", "plugins", "gemini", "gemini-messages-suite", "messages");

/**
 * resolveProjectKey — Farklı Yapay Zekaların (Gemini, Claude vb.) Proje Adlarını Çözer
 */
function resolveProjectKey(inputName) {
  if (!inputName) return 'default';
  const cleanInput = inputName.toLowerCase().replace(/[^a-z0-9]/g, '');
  const mapPath = path.join(GLOBAL_MESSAGES, 'ai_projects_map.json');

  if (fs.existsSync(mapPath)) {
    try {
      const map = JSON.parse(fs.readFileSync(mapPath, 'utf8'));
      
      // 1. Doğrudan anahtar eşleşmesi
      if (map[inputName.toLowerCase()]) return inputName.toLowerCase();

      // 2. Çoklu AI haritasında slug, uuid veya temiz isim eşleşmesi
      for (const [key, val] of Object.entries(map)) {
        const cleanKey = key.replace(/[^a-z0-9]/g, '');
        if (cleanKey === cleanInput || cleanInput.includes(cleanKey) || cleanKey.includes(cleanInput)) {
          return key;
        }

        // Gemini UUID kontrolü
        if (val.supportedAIs && val.supportedAIs.gemini && val.supportedAIs.gemini.projectId === inputName) {
          return key;
        }

        // Claude Slug kontrolü
        if (val.supportedAIs && val.supportedAIs.claude && val.supportedAIs.claude.associatedSlugs) {
          if (val.supportedAIs.claude.associatedSlugs.some(s => s.slug.toLowerCase().includes(cleanInput) || cleanInput.includes(s.slug.toLowerCase().replace(/[^a-z0-9]/g, '')))) {
            return key;
          }
        }
      }
    } catch(e) {}
  }

  return inputName.toLowerCase().replace(/[^a-z0-9_\-+]/g, '_');
}

function mkdirp(d) {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
}

const command = process.argv[2] || 'help';
const arg1 = process.argv[3];
const arg2 = process.argv[4];

console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log('📋 PLAN YÖNETİMİ — AZ TOKEN ÇOK İŞ & ÇOKLU AI ORKESTRASYONU');
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

switch (command) {
  case 'create':
  case 'new': {
    const projectName = resolveProjectKey(arg1);
    const planTitle = arg2 || `${projectName} Geliştirme Görevi`;
    const timestamp = new Date().toISOString().split('T')[0];

    const planTemplate = `# PLAN: ${planTitle}
> **Proje:** ${projectName} | **Tarih:** ${timestamp} | **Prensip:** Az Token · Çok İş · Cerrahi Müdahale

---

## 🎯 1. HEDEF VE KAPSAM (Zero-Preamble)
- **Amaç:** ${planTitle}
- **Yığın Tespiti:** (PHP / .NET / Node / Python / DB / Frontend)
- **Kısıtlamalar:** BOM kesinlikle yasak, mevcut mimari korunacak, harici gereksiz kütüphane eklenmeyecek.

---

## 🔍 2. MEVCUT ALTYAPIDAN YENİDEN KULLANILACAKLAR (Token Tasarrufu)
> *Sıfırdan yazıp token yakma! Projedeki çalışan desenleri kopyala.*
- [ ] Benzer çalışan controller/route:
- [ ] Mevcut UI bileşeni veya şablon sınıfı:
- [ ] Ortak servis / database helper:

---

## ✂️ 3. CERRAHİ GÖREV KIRILIMI (Surgical Checklist)
> *Her adım tek başına test edilebilir ve dar kapsamlı olmalıdır.*
- [ ] **Adım 1:** 
  - Hedef Dosya: 
  - İşlem: [MODIFY / NEW]
  - Sorumlu Ajan / Eklenti: 
- [ ] **Adım 2:** 
  - Hedef Dosya: 
  - İşlem: [MODIFY / NEW]
  - Sorumlu Ajan / Eklenti: 

---

## 🛡️ 4. KALİTE KAPILARI VE DOĞRULAMA (4/4 PASS)
- [ ] **Kapı 1 (Syntax):** Dil sözdizim kontrolü (0 hata)
- [ ] **Kapı 2 (BOM & DOM):** BOM-suz UTF-8 & dengeli HTML/DOM
- [ ] **Kapı 3 (Güvenlik):** OWASP / Sanitization / SQLi koruması
- [ ] **Kapı 4 (İşlev Kanıtı):** Ekran çıktısı / Terminal logu / Test sonucu

---

## 📦 5. TESLİMAT KANITI
*(Görev bittiğinde elde edilen somut kanıt buraya işlenir)*
`;

    const targetPaths = [
      path.join(GLOBAL_MESSAGES, 'projects', projectName, 'PLAN.md'),
      path.join(DESKTOP_MESSAGES, 'projects', projectName, 'PLAN.md')
    ];

    targetPaths.forEach(tp => {
      mkdirp(path.dirname(tp));
      fs.writeFileSync(tp, planTemplate, 'utf8');
      console.log('✅ Plan oluşturuldu: ' + tp);
    });

    console.log(`\n💡 İpucu: Proje anahtarı '${projectName}' olarak çözümlendi.`);
    break;
  }

  case 'status': {
    const projectName = resolveProjectKey(arg1);
    const planFile = path.join(GLOBAL_MESSAGES, 'projects', projectName, 'PLAN.md');
    if (!fs.existsSync(planFile)) {
      console.log('⚠️ Bu proje için henüz bir PLAN.md oluşturulmamış: ' + planFile);
      console.log('Oluşturmak için: node plan_manager.js new ' + projectName + ' "<Plan Başlığı>"');
      break;
    }

    const content = fs.readFileSync(planFile, 'utf8');
    console.log(`📄 Aktif Plan [${projectName}]: ` + planFile + '\n');
    const lines = content.split('\n');
    lines.forEach(l => {
      if (l.startsWith('#') || l.includes('[ ]') || l.includes('[x]') || l.includes('🎯')) {
        console.log('  ' + l);
      }
    });
    break;
  }

  case 'validate': {
    const filePath = arg1;
    if (!filePath || !fs.existsSync(filePath)) {
      console.error('❌ Geçerli bir plan dosya yolu belirtilmedi.');
      process.exit(1);
    }
    const text = fs.readFileSync(filePath, 'utf8');
    const checks = [
      { name: 'Hedef ve Kapsam Bölümü', pass: text.includes('1. HEDEF VE KAPSAM') },
      { name: 'Mevcut Altyapıdan Yeniden Kullanım (Token Tasarrufu)', pass: text.includes('2. MEVCUT ALTYAPIDAN YENİDEN KULLANILACAKLAR') },
      { name: 'Cerrahi Görev Kırılımı (Checklist)', pass: text.includes('3. CERRAHİ GÖREV KIRILIMI') },
      { name: 'Kalite Kapıları (4/4 PASS)', pass: text.includes('4. KALİTE KAPILARI VE DOĞRULAMA') }
    ];

    console.log('🔍 Plan Standart Uyumluluk Denetimi:');
    let allPass = true;
    checks.forEach(c => {
      console.log(`  ${c.pass ? '✅ PASS' : '❌ FAIL'}: ${c.name}`);
      if (!c.pass) allPass = false;
    });

    if (allPass) {
      console.log('\n🎉 Plan "Az Token Çok İş" standartlarına %100 uygundur.');
    } else {
      console.log('\n⚠️ Plan şablonunda eksik standartlar bulunmaktadır.');
      process.exit(1);
    }
    break;
  }

  default:
    console.log('Kullanım:');
    console.log('  node plan_manager.js new <proje-adi-veya-slug-veya-uuid> "<Plan Başlığı>"');
    console.log('  node plan_manager.js status <proje-adi-veya-slug-veya-uuid>');
    console.log('  node plan_manager.js validate <plan-dosyasi>');
    break;
}
