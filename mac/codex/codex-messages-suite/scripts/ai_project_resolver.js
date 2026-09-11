/**
 * ai_project_resolver.js — Çoklu Yapay Zeka Proje Haritalandırma ve Çözümleme Motoru
 * 
 * Portatif Standart: os.homedir() ve ortam değişkenleri ile dinamik çözümlenir.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');

const USER_HOME = process.env.USERPROFILE || process.env.HOME || os.homedir();
const GEMINI_PROJECTS_DIR = path.join(USER_HOME, '.gemini', 'config', 'projects');
const CLAUDE_PROJECTS_DIR = path.join(USER_HOME, '.claude', 'projects');
const CLAUDE_JSON_PATH = path.join(USER_HOME, '.claude.json');
const MESSAGES_BASE_DIR = path.join(USER_HOME, '.gemini', 'config', 'plugins', 'gemini-messages-suite', 'messages');
const DESKTOP_MESSAGES_DIR = path.join(USER_HOME, 'Desktop', 'plugins', 'gemini', 'gemini-messages-suite', 'messages');

function toPortable(p) {
  if (!p) return p;
  const h1 = USER_HOME.replace(/\\/g, '/');
  const h2 = USER_HOME;
  return p.split(h1).join('~').split(h2).join('~').replace(/\\/g, '/');
}

function normalizeName(str) {
  return (str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log('🌐 ÇOKLU YAPAY ZEKA PROJE ÇÖZÜMLEME VE HARİTALANDIRMA');
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

// 1. Gemini Projelerini Topla
const geminiList = [];
if (fs.existsSync(GEMINI_PROJECTS_DIR)) {
  fs.readdirSync(GEMINI_PROJECTS_DIR).filter(f => f.endsWith('.json')).forEach(file => {
    try {
      const data = JSON.parse(fs.readFileSync(path.join(GEMINI_PROJECTS_DIR, file), 'utf8'));
      geminiList.push({
        file,
        id: data.id || path.basename(file, '.json'),
        name: data.name || path.basename(file, '.json'),
        cleanName: normalizeName(data.name || path.basename(file, '.json')),
        fullPath: path.join(GEMINI_PROJECTS_DIR, file)
      });
    } catch(e) {}
  });
}
console.log(`✅ Gemini: ${geminiList.length} proje tespit edildi.`);

// 2. Claude Projelerini Topla
const claudeSlugs = [];
if (fs.existsSync(CLAUDE_PROJECTS_DIR)) {
  fs.readdirSync(CLAUDE_PROJECTS_DIR).forEach(slug => {
    claudeSlugs.push({
      slug,
      cleanName: normalizeName(slug.split('-').pop()),
      fullPath: path.join(CLAUDE_PROJECTS_DIR, slug)
    });
  });
}
console.log(`✅ Claude: ${claudeSlugs.length} proje slug'ı tespit edildi.`);

// 3. Claude.json Dosya Yollarını Topla
const claudePaths = [];
if (fs.existsSync(CLAUDE_JSON_PATH)) {
  try {
    const cData = JSON.parse(fs.readFileSync(CLAUDE_JSON_PATH, 'utf8'));
    if (cData.projects) {
      Object.keys(cData.projects).forEach(pKey => {
        claudePaths.push({
          path: pKey,
          cleanName: normalizeName(path.basename(pKey))
        });
      });
    }
  } catch(e) {}
}
console.log(`✅ Claude Config: ${claudePaths.length} aktif proje yolu tespit edildi.`);

// 4. Çoklu AI Proje Haritasını Oluştur (Taşınabilir Format)
const projectMap = {};

geminiList.forEach(gp => {
  const key = gp.name.toLowerCase().replace(/[^a-z0-9_\-+]/g, '_');
  const matchedClaudeSlugs = claudeSlugs
    .filter(cs => cs.cleanName.includes(gp.cleanName) || gp.cleanName.includes(cs.cleanName))
    .map(cs => ({ slug: cs.slug, path: toPortable(cs.fullPath) }));

  const matchedClaudePaths = claudePaths
    .filter(cp => cp.cleanName === gp.cleanName || cp.path.toLowerCase().includes(gp.cleanName))
    .map(cp => toPortable(cp.path));

  projectMap[key] = {
    displayName: gp.name,
    supportedAIs: {
      gemini: {
        supported: true,
        format: "UUID_JSON",
        projectId: gp.id,
        messagesDir: `projects/${key}`
      },
      claude: {
        supported: matchedClaudeSlugs.length > 0 || matchedClaudePaths.length > 0,
        format: "PATH_SLUG_AND_LOCAL_MESSAGES",
        associatedSlugs: matchedClaudeSlugs,
        associatedPaths: matchedClaudePaths
      },
      genericAI: {
        supported: true,
        format: "IN_PROJECT_ROOT",
        rulesPath: "rules/01_EVRENSEL_KURALLAR.md",
        planPath: `projects/${key}/PLAN.md`,
        directivesPath: `projects/${key}/PROJE_TALİMATLARI.md`
      }
    },
    centralStorage: {
      relativeDir: `projects/${key}`
    }
  };
});

// 5. Harita Dosyalarını Kaydet
const outputGlobal = path.join(MESSAGES_BASE_DIR, 'ai_projects_map.json');
const outputDesktop = path.join(DESKTOP_MESSAGES_DIR, 'ai_projects_map.json');

if (fs.existsSync(MESSAGES_BASE_DIR)) {
  fs.writeFileSync(outputGlobal, JSON.stringify(projectMap, null, 2) + '\n', 'utf8');
}
if (fs.existsSync(DESKTOP_MESSAGES_DIR)) {
  fs.writeFileSync(outputDesktop, JSON.stringify(projectMap, null, 2) + '\n', 'utf8');
}

console.log(`\n🎉 Toplam ${Object.keys(projectMap).length} proje için Taşınabilir Çoklu AI Haritası üretildi!`);
