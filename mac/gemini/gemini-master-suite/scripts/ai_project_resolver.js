/**
 * ai_project_resolver.js — Çoklu Yapay Zeka Proje Haritalandırma ve Çözümleme Motoru
 * 
 * Amaç: Gemini, Claude, Cursor ve yerel projelerin farklı dosya formatlarını (UUID JSON, Path Slug, Local In-Project)
 * merkezi ve modüler bir yapay zeka proje haritasında (ai_projects_map.json) birleştirmek.
 */
'use strict';

const fs = require('fs');
const path = require('path');

const GEMINI_PROJECTS_DIR = '$HOME/.gemini/config/projects';
const CLAUDE_PROJECTS_DIR = '$HOME/.claude/projects';
const CLAUDE_JSON_PATH = '$HOME/.claude.json';
const MESSAGES_BASE_DIR = '$HOME/.gemini/config/plugins/gemini-messages-suite/messages';
const DESKTOP_MESSAGES_DIR = '$HOME/Desktop/gemini/plugins/gemini-messages-suite/messages';

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

// 4. Çoklu AI Proje Haritasını Oluştur
const projectMap = {};

geminiList.forEach(gp => {
  const key = gp.name.toLowerCase().replace(/[^a-z0-9_\-+]/g, '_');
  const matchedClaudeSlugs = claudeSlugs
    .filter(cs => cs.cleanName.includes(gp.cleanName) || gp.cleanName.includes(cs.cleanName))
    .map(cs => ({ slug: cs.slug, path: cs.fullPath }));

  const matchedClaudePaths = claudePaths
    .filter(cp => cp.cleanName === gp.cleanName || cp.path.toLowerCase().includes(gp.cleanName))
    .map(cp => cp.path);

  projectMap[key] = {
    displayName: gp.name,
    supportedAIs: {
      gemini: {
        supported: true,
        format: "UUID_JSON",
        projectId: gp.id,
        configFile: gp.fullPath,
        messagesDir: `${MESSAGES_BASE_DIR}/projects/${key}`
      },
      claude: {
        supported: matchedClaudeSlugs.length > 0 || matchedClaudePaths.length > 0,
        format: "PATH_SLUG_AND_LOCAL_MESSAGES",
        associatedSlugs: matchedClaudeSlugs,
        associatedPaths: matchedClaudePaths,
        localMessagesDirs: matchedClaudePaths.map(p => path.join(p, '.claude', 'messages'))
      },
      genericAI: {
        supported: true,
        format: "IN_PROJECT_ROOT",
        rulesPath: `${MESSAGES_BASE_DIR}/rules/01_EVRENSEL_KURALLAR.md`,
        planPath: `${MESSAGES_BASE_DIR}/projects/${key}/PLAN.md`,
        directivesPath: `${MESSAGES_BASE_DIR}/projects/${key}/PROJE_TALİMATLARI.md`
      }
    },
    centralStorage: {
      globalMessagesDir: `${MESSAGES_BASE_DIR}/projects/${key}`,
      desktopMessagesDir: `${DESKTOP_MESSAGES_DIR}/projects/${key}`
    }
  };
});

// 5. Harita Dosyalarını Kaydet
const outputGlobal = path.join(MESSAGES_BASE_DIR, 'ai_projects_map.json');
const outputDesktop = path.join(DESKTOP_MESSAGES_DIR, 'ai_projects_map.json');

fs.writeFileSync(outputGlobal, JSON.stringify(projectMap, null, 2) + '\n', 'utf8');
fs.writeFileSync(outputDesktop, JSON.stringify(projectMap, null, 2) + '\n', 'utf8');

console.log(`\n🎉 Toplam ${Object.keys(projectMap).length} proje için Çoklu AI Haritası üretildi!`);
console.log(`📄 Kaydedildi: ${outputGlobal}`);
console.log(`📄 Kaydedildi: ${outputDesktop}`);
