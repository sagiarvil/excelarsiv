/**
 * self_updater.js — Antigravity & AI Master Suite Otonom Güncelleme ve Versiyonlama Motoru
 * 
 * Bu betik, proje sırasında eklentide yapılan iyileştirmeleri, yeni eklenen
 * araçları, kuralları ve becerileri tespit eder; eklentinin sürümünü (SemVer)
 * otomatik artırır, CHANGELOG.md ve README.md sürüm geçmişine işler,
 * Desktop ile sistem kurulu eklenti dizinini çift yönlü senkronize eder
 * ve Antigravity CLI (agy) ile sisteme anında yeniden tescil eder.
 * 
 * Kullanım:
 *   node scripts/self_updater.js --patch "Küçük hata düzeltmesi veya kural iyileştirmesi"
 *   node scripts/self_updater.js --minor "Yeni araç veya beceri entegrasyonu"
 *   node scripts/self_updater.js --major "Kapsamlı mimari güncelleme"
 *   node scripts/self_updater.js --sync   (Sürüm değiştirmeden dosyaları eşitle)
 *   node scripts/self_updater.js --status (Mevcut sürüm ve dosya durumu)
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execSync } = require('child_process');

const USER_HOME = process.env.USERPROFILE || process.env.HOME || os.homedir();
const GEMINI_ROOT = path.join(USER_HOME, '.gemini');
const CONFIG_PLUGINS_DIR = path.join(GEMINI_ROOT, 'config', 'plugins');
function getAgyExecutable() {
  const isWin = process.platform === 'win32';
  const localBin = path.join(GEMINI_ROOT, 'bin', isWin ? 'agy' : 'agy');
  if (fs.existsSync(localBin)) return localBin;
  try {
    const whichCmd = isWin ? 'where agy' : 'which agy';
    const out = execSync(whichCmd, { stdio: ['pipe', 'pipe', 'pipe'] }).toString().trim();
    if (out) return out.split('\n')[0].trim();
  } catch (e) {}
  return isWin ? 'agy' : 'agy';
}
const AGY_EXE = getAgyExecutable();
const DESKTOP_PLUGINS_DIR = path.join(USER_HOME, 'Desktop', 'plugins');

const SCRIPT_DIR = __dirname;
const PLUGIN_ROOT = path.resolve(SCRIPT_DIR, '..');
const MANIFEST_PATH = path.join(PLUGIN_ROOT, 'plugin.json');

if (!fs.existsSync(MANIFEST_PATH)) {
  console.error('❌ HATA: plugin.json bulunamadı! Yol: ' + MANIFEST_PATH);
  process.exit(1);
}

const pluginName = path.basename(PLUGIN_ROOT);
const aiFolder = pluginName.startsWith('claude') ? 'claude' : 'gemini';
const desktopPluginDir = path.join(DESKTOP_PLUGINS_DIR, aiFolder, pluginName);
const globalPluginDir = path.join(CONFIG_PLUGINS_DIR, pluginName);

// Eğer manifest 0 byte ise diğer dizinden kurtar
let manifestRaw = fs.readFileSync(MANIFEST_PATH, 'utf8');
if (!manifestRaw.trim()) {
  const fallbackPath = PLUGIN_ROOT.includes('Desktop') 
    ? path.join(globalPluginDir, 'plugin.json') 
    : path.join(desktopPluginDir, 'plugin.json');
  if (fs.existsSync(fallbackPath) && fs.statSync(fallbackPath).size > 0) {
    manifestRaw = fs.readFileSync(fallbackPath, 'utf8');
    fs.writeFileSync(MANIFEST_PATH, manifestRaw, 'utf8');
  }
}

let manifest = {};
try {
  manifest = JSON.parse(manifestRaw);
} catch (e) {
  console.error('❌ HATA: plugin.json JSON olarak ayrıştırılamadı!');
  process.exit(1);
}

function logHeader(title) {
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log(title);
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}

function countItems(dirPath) {
  if (!fs.existsSync(dirPath)) return 0;
  return fs.readdirSync(dirPath).filter(f => !f.startsWith('.') && f !== '__pycache__' && f !== 'node_modules').length;
}

function bumpVersion(currentVer, bumpType) {
  const parts = currentVer.split('.').map(n => parseInt(n, 10) || 0);
  while (parts.length < 3) parts.push(0);

  if (bumpType === 'major') {
    parts[0] += 1;
    parts[1] = 0;
    parts[2] = 0;
  } else if (bumpType === 'minor') {
    parts[1] += 1;
    parts[2] = 0;
  } else {
    parts[2] += 1;
  }
  return parts.join('.');
}

function copyDirRecursive(src, dst) {
  if (!fs.existsSync(dst)) fs.mkdirSync(dst, { recursive: true });
  const items = fs.readdirSync(src);
  for (const item of items) {
    if (item === 'node_modules' || item === '.git' || item === '__pycache__') continue;
    const s = path.join(src, item);
    const d = path.join(dst, item);
    const stat = fs.statSync(s);
    if (stat.isDirectory()) {
      copyDirRecursive(s, d);
    } else {
      fs.copyFileSync(s, d);
    }
  }
}

// Argümanları oku
const args = process.argv.slice(2);
let bumpType = null;
let changeMessage = '';

for (let i = 0; i < args.length; i++) {
  const arg = args[i];
  if (arg === '--patch' || arg === '-p') {
    bumpType = 'patch';
    changeMessage = args[i + 1] || 'Hata onarımı ve kural iyileştirmesi';
    i++;
  } else if (arg === '--minor' || arg === '-m') {
    bumpType = 'minor';
    changeMessage = args[i + 1] || 'Yeni araç, beceri veya özellik eklendi';
    i++;
  } else if (arg === '--major') {
    bumpType = 'major';
    changeMessage = args[i + 1] || 'Kapsamlı mimari güncelleme';
    i++;
  } else if (arg === '--sync' || arg === '-s') {
    bumpType = 'sync';
  } else if (arg === '--status') {
    bumpType = 'status';
  } else if (!changeMessage && !arg.startsWith('-')) {
    changeMessage = arg;
  }
}

if (!bumpType) {
  bumpType = 'patch';
  if (!changeMessage) changeMessage = 'Otonom periyodik geliştirme ve güncelleme';
}

const agentsCount = countItems(path.join(PLUGIN_ROOT, 'agents'));
const skillsCount = countItems(path.join(PLUGIN_ROOT, 'skills'));
const rulesCount = countItems(path.join(PLUGIN_ROOT, 'rules'));
const scriptsCount = countItems(path.join(PLUGIN_ROOT, 'scripts'));

logHeader('🔄 ANTIGRAVITY EKLENTİ OTONOM GÜNCELLEME MOTORU (SELF-EVOLUTION)');
console.log('Eklenti Adı    : ' + pluginName);
console.log('Yapay Zeka Tipi: ' + aiFolder.toUpperCase());
console.log('Mevcut Sürüm   : ' + (manifest.version || '1.0.0'));
console.log('Konum          : ' + PLUGIN_ROOT);
console.log('Yetenek Havuzu : ' + agentsCount + ' Ajan | ' + skillsCount + ' Beceri | ' + rulesCount + ' Kural | ' + scriptsCount + ' Betik');

if (bumpType === 'status') {
  console.log('\n✅ Durum kontrolü tamamlandı.');
  process.exit(0);
}

const oldVersion = manifest.version || '1.0.0';
let newVersion = oldVersion;

if (bumpType !== 'sync') {
  newVersion = bumpVersion(oldVersion, bumpType);
  console.log('Yeni Sürüm     : ' + newVersion + ' (' + bumpType.toUpperCase() + ')');
  console.log('Açıklama       : ' + changeMessage);

  // 1. MANIFEST GÜNCELLEME (plugin.json)
  manifest.version = newVersion;
  if (!manifest.capabilities) manifest.capabilities = {};
  manifest.capabilities.agents = agentsCount;
  manifest.capabilities.skills = skillsCount;
  manifest.capabilities.rules = rulesCount;
  manifest.capabilities.scripts = scriptsCount;
  manifest.capabilities.lastUpdated = new Date().toISOString();

  fs.writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
  console.log('\n📝 1. plugin.json güncellendi: ' + oldVersion + ' -> ' + newVersion);

  // 2. INSTALLED_VERSION.JSON GÜNCELLEME
  const installedVerPath = path.join(PLUGIN_ROOT, 'installed_version.json');
  fs.writeFileSync(installedVerPath, JSON.stringify({ version: newVersion, lastUpdated: new Date().toISOString() }, null, 2) + '\n', 'utf8');
  console.log('📝 2. installed_version.json güncellendi.');

  // 3. CHANGELOG.MD GÜNCELLEME
  const changelogPath = path.join(PLUGIN_ROOT, 'CHANGELOG.md');
  const dateStr = new Date().toISOString().slice(0, 10);
  const timeStr = new Date().toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' });
  const logEntry = `\n### [${newVersion}] — ${dateStr} ${timeStr} (${bumpType.toUpperCase()})\n- **Değişiklik:** ${changeMessage}\n- **Yetenek Durumu:** ${agentsCount} Ajan, ${skillsCount} Beceri, ${rulesCount} Kural, ${scriptsCount} Betik\n- **Yazar:** WebTasarimOfisim / Antigravity AI Engine\n`;

  let currentChangelog = '';
  if (fs.existsSync(changelogPath)) {
    currentChangelog = fs.readFileSync(changelogPath, 'utf8');
  } else {
    currentChangelog = `# 📜 ${pluginName.toUpperCase()} — DEĞİŞİKLİK GÜNLÜĞÜ (CHANGELOG)\n\nBu dosya, eklentide yapılan tüm otonom güncellemeleri ve sürüm değişikliklerini kayıt altına alır.\n`;
  }

  const headerEnd = currentChangelog.indexOf('\n\n');
  if (headerEnd !== -1) {
    currentChangelog = currentChangelog.slice(0, headerEnd + 2) + logEntry + currentChangelog.slice(headerEnd + 2);
  } else {
    currentChangelog += logEntry;
  }
  fs.writeFileSync(changelogPath, currentChangelog, 'utf8');
  console.log('📝 3. CHANGELOG.md güncellendi.');
} else {
  console.log('\nℹ️ --sync modu: Sürüm artırılmadan dosyalar senkronize ediliyor...');
}

// 4. ANTIGRAVITY CLI İLE TESCİL (agy plugin install)
if (fs.existsSync(AGY_EXE)) {
  console.log('\n🚀 4. Antigravity CLI ile eklenti tescil ediliyor...');
  try {
    const targetRegister = fs.existsSync(globalPluginDir) ? globalPluginDir : PLUGIN_ROOT;
    const regOut = execSync(`"${AGY_EXE}" plugin install "${targetRegister}"`, { stdio: ['pipe', 'pipe', 'pipe'] }).toString('utf8');
    console.log('   ✔ Tescil başarılı: ' + regOut.trim().split('\n')[0]);
  } catch (err) {
    console.log('   ℹ️ CLI tescil notu: ' + err.message.split('\n')[0]);
  }
}

// 5. MASAÜSTÜ VE SİSTEM EKLENTİ DİZİNİ ÇİFT YÖNLÜ SENKRONİZASYON (agy tescilinden hemen sonra çalıştır ve manifesti koru)
console.log('\n🔄 5. Çift Yönlü Dizin Senkronizasyonu yürütülüyor...');

if (PLUGIN_ROOT.includes('Desktop') && fs.existsSync(CONFIG_PLUGINS_DIR)) {
  copyDirRecursive(PLUGIN_ROOT, globalPluginDir);
  fs.writeFileSync(path.join(globalPluginDir, 'plugin.json'), JSON.stringify(manifest, null, 2) + '\n', 'utf8');
  fs.writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
  console.log('   ✅ Küresel eklenti dizini ve Masaüstü senkronize edildi (Manifest korundu).');
} else if (PLUGIN_ROOT.includes('.gemini') && fs.existsSync(DESKTOP_PLUGINS_DIR)) {
  copyDirRecursive(PLUGIN_ROOT, desktopPluginDir);
  fs.writeFileSync(path.join(desktopPluginDir, 'plugin.json'), JSON.stringify(manifest, null, 2) + '\n', 'utf8');
  fs.writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
  console.log('   ✅ Masaüstü eklenti dizini ve Küresel senkronize edildi (Manifest korundu).');
}

// 6. UTF-8 BOM DENETİMİ VE DÜZELTME
const bomScanTool = path.join(SCRIPT_DIR, 'bom_utf8_scan.js');
if (fs.existsSync(bomScanTool)) {
  console.log('\n🛡️ 6. UTF-8 BOM koruma taraması çalıştırılıyor...');
  try {
    execSync(`node "${bomScanTool}" --fix "${PLUGIN_ROOT}"`, { stdio: ['pipe', 'pipe', 'pipe'] });
    console.log('   ✅ Tüm eklenti dosyaları BOM-suz UTF-8 olarak doğrulandı.');
  } catch (err) {
    console.log('   ℹ️ BOM tarayıcı notu: ' + err.message.split('\n')[0]);
  }
}

logHeader('🎉 GÜNCELLEME TAMAMLANDI! Eklenti ' + newVersion + ' sürümüne başarıyla yükseltildi.');
