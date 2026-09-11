/**
 * install_plugins.js — Antigravity macOS Özel AI Eklenti Kurulum Aracı
 * 
 * Apple Silicon (M1/M2/M3/M4) ve Intel Mac Sistemleri İçin Özel Optimize Edilmiştir.
 * Homebrew, Laravel Herd, Valet ve MAMP geliştirici ortamlarıyla %100 uyumludur.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execSync } = require('child_process');

const BASE_DIR = __dirname;
const USER_HOME = process.env.HOME || os.homedir();
const GEMINI_ROOT = path.join(USER_HOME, '.gemini');
const CONFIG_DIR = path.join(GEMINI_ROOT, 'config');
const GLOBAL_PLUGINS_DIR = path.join(CONFIG_DIR, 'plugins');

function getAgyExecutable() {
  const localBin = path.join(GEMINI_ROOT, 'bin', 'agy');
  if (fs.existsSync(localBin)) return localBin;
  try {
    const out = execSync('which agy', { stdio: ['pipe', 'pipe', 'pipe'] }).toString().trim();
    if (out) return out.split('\n')[0].trim();
  } catch (e) {}
  return 'agy';
}

function mkdirp(d) {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
}

function copyDirRecursive(src, dst) {
  mkdirp(dst);
  const items = fs.readdirSync(src);
  for (const item of items) {
    const s = path.join(src, item);
    const d = path.join(dst, item);
    const stat = fs.statSync(s);
    if (stat.isDirectory()) {
      if (item === 'node_modules' || item === '.git' || item === '__pycache__') continue;
      copyDirRecursive(s, d);
    } else {
      fs.copyFileSync(s, d);
    }
  }
}

console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log('🍎 ANTIGRAVITY MACOS EKLENTİ KURULUMU (Apple Silicon & Intel)');
console.log('Kaynak Dizin: ' + BASE_DIR);
console.log('Hedef Dizin : ' + GLOBAL_PLUGINS_DIR);
console.log('macOS Sürüm : ' + os.type() + ' ' + os.release() + ' (' + os.arch() + ')');
console.log('CLI Yürütücü: ' + getAgyExecutable());
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

mkdirp(GLOBAL_PLUGINS_DIR);

// 0. MACOS GELİŞTİRİCİ ÇALIŞMA ZAMANI DENETİMİ (Homebrew & Herd)
console.log('🔍 0. macOS Geliştirici Ortamı ve Paket Yöneticisi Denetimi...');
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

const toolsSummary = [
  { name: 'Node.js', cmd: 'node -v', brewCmd: 'brew install node' },
  { name: 'Python 3', cmd: 'python3 --version', brewCmd: 'brew install python' },
  { name: 'Git', cmd: 'git --version', brewCmd: 'brew install git' },
  { name: 'PHP (8.3/8.4)', cmd: 'php -v', brewCmd: 'brew install php' },
  { name: 'Composer', cmd: 'composer -V', brewCmd: 'brew install composer' }
];

let hasMissing = false;
toolsSummary.forEach(t => {
  let ok = false;
  let ver = '';
  try {
    ver = execSync(t.cmd, { stdio: ['pipe', 'pipe', 'pipe'] }).toString().trim().split('\n')[0];
    ok = true;
  } catch (e) {
    // macOS alternatif konumlar (Homebrew /opt/homebrew veya Laravel Herd)
    if (t.name.startsWith('PHP')) {
      const macPhpCandidates = [
        '/opt/homebrew/bin/php',
        '/usr/local/bin/php',
        path.join(USER_HOME, 'Library/Application Support/Herd/bin/php')
      ];
      for (const pc of macPhpCandidates) {
        if (fs.existsSync(pc)) {
          try {
            ver = execSync(`"${pc}" -v`, { stdio: ['pipe', 'pipe', 'pipe'] }).toString().trim().split('\n')[0];
            ok = true;
            ver += ' (macOS Keşfedildi)';
            break;
          } catch (err) {}
        }
      }
    } else if (t.name === 'Composer') {
      const macCompCandidates = [
        '/opt/homebrew/bin/composer',
        '/usr/local/bin/composer',
        path.join(USER_HOME, '.composer/vendor/bin/composer')
      ];
      for (const cc of macCompCandidates) {
        if (fs.existsSync(cc)) {
          ok = true;
          ver = 'Composer Hazır (' + cc + ')';
          break;
        }
      }
    }
  }

  if (ok) {
    console.log(`  ✅ ${t.name.padEnd(14)} : ${ver}`);
  } else {
    hasMissing = true;
    console.log(`  ⚠️ ${t.name.padEnd(14)} : SİSTEMDE BULUNAMADI!`);
    console.log(`     🍎 Homebrew Komutu: ${t.brewCmd}`);
    console.log('');
  }
});

if (hasMissing) {
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('💡 MACOS İÇİN TAVSİYE EDİLEN ÇÖZÜMLER:');
  console.log('1. Homebrew ile tek satırda tüm araçları kurun:');
  console.log('   brew install node python git php composer');
  console.log('2. Hepsi bir arada modern macOS stack: Laravel Herd');
  console.log('   Web: https://herd.laravel.com (PHP, Nginx, SSL arayüzlü yönetim)');
  console.log('3. macOS Veritabanı Yöneticisi: DBngin');
  console.log('   Web: https://dbngin.com (MySQL, PostgreSQL, Redis tek tıkla)');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
} else {
  console.log('🎉 Tüm temel macOS geliştirici araçları hazır ve çalışır durumda!\n');
}

// 1. YAPAY ZEKA EKLENTİLERİNİ TARA (Auto-Discovery)
const aiFolders = ['gemini', 'claude', 'codex'];
const discoveredPlugins = [];

aiFolders.forEach(aiName => {
  const aiDir = path.join(BASE_DIR, aiName);
  if (fs.existsSync(aiDir)) {
    const entries = fs.readdirSync(aiDir, { withFileTypes: true });
    entries.forEach(entry => {
      if (entry.isDirectory()) {
        const potentialPlugin = path.join(aiDir, entry.name);
        const manifest = path.join(potentialPlugin, 'plugin.json');
        if (fs.existsSync(manifest)) {
          discoveredPlugins.push({
            aiName,
            name: entry.name,
            src: potentialPlugin,
            dst: path.join(GLOBAL_PLUGINS_DIR, entry.name)
          });
        }
      }
    });
  }
});

console.log(`🔍 Keşfedilen macOS Yapay Zeka Eklentileri (${discoveredPlugins.length} adet):\n`);
discoveredPlugins.forEach(p => {
  console.log(`  • [${p.aiName.toUpperCase()}] ${p.name} -> ${p.src}`);
});

// 2. EKLENTİLERİ SİSTEME AKTAR
console.log('\n📦 2. Eklentiler macOS sistemine aktarılıyor...');
discoveredPlugins.forEach(p => {
  console.log(`   🚀 Aktarılıyor: ${p.name}...`);
  copyDirRecursive(p.src, p.dst);
  console.log(`   ✅ ${p.name} başarıyla aktarıldı.`);
});

// 3. CONFIG.JSON CERRAHİ BİRLEŞTİRME
console.log('\n⚙️ 3. config.json cerrahi birleştirme yapılıyor...');
const configJsonPath = path.join(CONFIG_DIR, 'config.json');
let configData = {};
if (fs.existsSync(configJsonPath)) {
  try {
    fs.copyFileSync(configJsonPath, configJsonPath + '.bak');
    configData = JSON.parse(fs.readFileSync(configJsonPath, 'utf8'));
  } catch (e) {}
}
if (!configData.userSettings) configData.userSettings = {};
if (!configData.plugins) configData.plugins = {};
discoveredPlugins.forEach(p => { configData.plugins[p.name] = { enabled: true }; });
fs.writeFileSync(configJsonPath, JSON.stringify(configData, null, 4) + '\n', 'utf8');
console.log('   ✅ config.json güncellendi (macOS eklentileri aktif edildi).');

// 4. HOOKS.JSON (POSIX YOLLARI)
console.log('\n🪝 4. hooks.json kancaları bağlanıyor...');
const hooksJsonPath = path.join(CONFIG_DIR, 'hooks.json');
let hooksData = {};
if (fs.existsSync(hooksJsonPath)) {
  try { hooksData = JSON.parse(fs.readFileSync(hooksJsonPath, 'utf8')); } catch (e) {}
}
hooksData['agent-surgical-guard'] = {
  "enabled": true,
  "PreToolUse": [
    {
      "matcher": "write_to_file|replace_file_content",
      "hooks": [
        {
          "type": "command",
          "command": "node plugins/gemini-master-suite/scripts/pre_tool_guard.js",
          "timeout": 5
        }
      ]
    }
  ],
  "PostToolUse": [
    {
      "matcher": "write_to_file|replace_file_content",
      "hooks": [
        {
          "type": "command",
          "command": "node plugins/gemini-master-suite/scripts/ensure_utf8_nobom.js",
          "timeout": 5
        }
      ]
    }
  ]
};
fs.writeFileSync(hooksJsonPath, JSON.stringify(hooksData, null, 2) + '\n', 'utf8');
console.log('   ✅ hooks.json kancaları doğrulandı.');

// 5. ANTIGRAVITY RESMİ CLI İLE TESCİL
const agyCmd = getAgyExecutable();
let canRun = false;
try {
  execSync(`"${agyCmd}" --version`, { stdio: ['pipe', 'pipe', 'pipe'] });
  canRun = true;
} catch (e) {
  canRun = fs.existsSync(agyCmd);
}

if (canRun) {
  console.log(`\n🚀 5. Antigravity resmi CLI ile tescil ediliyor (${path.basename(agyCmd)})...\n`);
  discoveredPlugins.forEach(p => {
    try {
      const out = execSync(`"${agyCmd}" plugin install "${p.dst}"`, { stdio: ['pipe', 'pipe', 'pipe'] }).toString('utf8');
      console.log(`   ✔ ${p.name} tescil edildi: ` + out.trim().split('\n')[0]);
    } catch (err) {
      console.log(`   ℹ️ ${p.name} tescil notu: ` + err.message.split('\n')[0]);
    }
  });
}

// 6. MANIFEST BÜTÜNLÜK GARANTİSİ
discoveredPlugins.forEach(p => {
  copyDirRecursive(p.src, p.dst);
  const srcM = path.join(p.src, 'plugin.json');
  const dstM = path.join(p.dst, 'plugin.json');
  if (fs.existsSync(srcM)) fs.copyFileSync(srcM, dstM);
});

console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log('🎉 MACOS KURULUMU BAŞARIYLA TAMAMLANDI!');
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
