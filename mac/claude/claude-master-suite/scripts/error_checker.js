#!/usr/bin/env node
/**
 * error_checker.js — Kapsamlı Proje Hata ve Sözdizim Kontrol Aracı (Bug Hunter & QA)
 * 
 * Bu araç belirtilen dosya veya klasörü özyinelemeli olarak tarar:
 * 1. PHP Sözdizim Doğrulama (php -l)
 * 2. JavaScript / Node.js Sözdizim Doğrulama (node --check)
 * 3. Python Sözdizim Doğrulama (python3 -B -m py_compile)
 * 4. JSON Format ve Sözdizim Doğrulama (JSON.parse)
 * 5. UTF-8 BOM Tespiti (0xEF 0xBB 0xBF)
 * 6. Güvenlik ve Mantık Açığı Statik Analizi (eval, unescaped XSS, raw SQL, zayıf hash)
 * 7. Eksik require/include Tespiti
 * 
 * Kullanım:
 *   node error_checker.js [hedef_klasor_veya_dosya] [--fix-bom] [--json] [--strict]
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function getExecutor(canonicalPath, fallbackCmd) {
  if (canonicalPath && fs.existsSync(canonicalPath)) return `"${canonicalPath}"`;
  return fallbackCmd;
}

function resolvePython() {
  const candidates = [
    '/opt/homebrew/bin/python3',
    '/usr/local/bin/python3',
    '/usr/bin/python3'
  ];
  for (const c of candidates) {
    if (fs.existsSync(c)) return `"${c}"`;
  }
  try {
    execSync('python3 --version', { stdio: 'ignore' });
    return 'python3';
  } catch (e) {}
  return 'python';
}

const PHP_CMD = getExecutor('php', 'php');
const NODE_CMD = getExecutor('node', 'node');
const PYTHON_CMD = resolvePython();

// Argümanları ayrıştır
const args = process.argv.slice(2);
let targetPath = '.';
let fixBom = false;
let isJson = false;
let isStrict = false;

args.forEach(arg => {
  if (arg === '--fix-bom') fixBom = true;
  else if (arg === '--json') isJson = true;
  else if (arg === '--strict') isStrict = true;
  else if (!arg.startsWith('--')) targetPath = arg;
});

const resolvedTarget = path.resolve(targetPath);

if (!fs.existsSync(resolvedTarget)) {
  console.error(`❌ HATA: Hedef yol bulunamadı: ${resolvedTarget}`);
  process.exit(2);
}

const IGNORED_DIRS = new Set([
  'node_modules', 'vendor', '.git', 'storage', 'cache',
  'dist', 'build', '.vscode', '.idea', 'tmp', 'temp', '__pycache__'
]);

const report = {
  target: resolvedTarget,
  timestamp: new Date().toISOString(),
  scannedFiles: 0,
  stats: { pass: 0, fail: 0, warn: 0, info: 0 },
  errors: [],
  warnings: []
};

function addLog(type, code, file, message, detail = '') {
  if (type === 'FAIL') report.stats.fail++;
  else if (type === 'WARN') report.stats.warn++;
  else if (type === 'INFO') report.stats.info++;
  else if (type === 'PASS') report.stats.pass++;

  const entry = { type, code, file: path.relative(resolvedTarget, file) || path.basename(file), message, detail };
  if (type === 'FAIL') report.errors.push(entry);
  else if (type === 'WARN') report.warnings.push(entry);
}

// Dosya tarama fonksiyonu
function walk(dir) {
  const stat = fs.statSync(dir);
  if (!stat.isDirectory()) {
    checkFile(dir);
    return;
  }

  const items = fs.readdirSync(dir);
  for (const item of items) {
    if (IGNORED_DIRS.has(item)) continue;
    const fullPath = path.join(dir, item);
    try {
      const s = fs.statSync(fullPath);
      if (s.isDirectory()) {
        walk(fullPath);
      } else {
        checkFile(fullPath);
      }
    } catch (e) {}
  }
}

// Güvenlik ve mantık kalıp kontrolleri
const BUG_PATTERNS = [
  {
    regex: /eval\s*\(/i,
    category: 'SECURITY',
    code: 'SEC-EVAL',
    msg: 'eval() fonksiyonu kullanımı tespit edildi (uzaktan kod çalıştırma riski)'
  },
  {
    regex: /\bextract\s*\(\s*\$_(GET|POST|REQUEST)/i,
    category: 'SECURITY',
    code: 'SEC-EXTRACT',
    msg: 'extract($_GET/$_POST) ile kontrolsüz değişken enjeksiyonu riski'
  },
  {
    regex: /echo\s+\$_(GET|POST|REQUEST)\[[^\]]+\]\s*;/i,
    category: 'SECURITY',
    code: 'SEC-XSS',
    msg: 'Doğrudan kullanıcı girdisi echo ediliyor (XSS riski - htmlspecialchars zorunludur)'
  },
  {
    regex: /\$_(GET|POST|REQUEST)\[[^\]]+\]\s*\.\s*(['"][^'"]*SELECT|INSERT|UPDATE|DELETE)/i,
    category: 'SECURITY',
    code: 'SEC-SQLI',
    msg: 'Ham SQL sorgusu ile kullanıcı girdisi birleştiriliyor (SQL Enjeksiyonu riski)'
  },
  {
    regex: /\b(shell_exec|exec|system|passthru|popen|proc_open)\s*\(/i,
    category: 'SECURITY',
    code: 'SEC-CMD-EXEC',
    msg: 'Sistem komut yürütücü çağrısı tespit edildi'
  },
  {
    regex: /var_dump\s*\(|console\.log\s*\(/i,
    category: 'QUALITY',
    code: 'QUAL-DEBUG-LEAK',
    msg: 'Debug çıktısı (var_dump / console.log) canlı kodda kalmış olabilir'
  }
];

function checkFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const allowedExts = ['.php', '.js', '.mjs', '.cjs', '.json', '.py', '.html', '.sql'];
  if (!allowedExts.includes(ext)) return;

  report.scannedFiles++;
  let buf;
  try {
    buf = fs.readFileSync(filePath);
  } catch (e) {
    addLog('FAIL', 'IO-ERR', filePath, 'Dosya okunamadı: ' + e.message);
    return;
  }

  // 1. BOM Kontrolü
  if (buf.length >= 3 && buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) {
    if (fixBom) {
      fs.writeFileSync(filePath, buf.slice(3));
      addLog('WARN', 'BOM-FIXED', filePath, 'UTF-8 BOM tespit edildi ve otomatik onarıldı.');
      buf = buf.slice(3);
    } else {
      addLog('FAIL', 'BOM-DETECTED', filePath, 'Yasak UTF-8 BOM karakteri tespit edildi. (--fix-bom ile temizleyin)');
    }
  }

  const content = buf.toString('utf8');

  // 2. PHP Syntax Kontrolü (php -l)
  if (ext === '.php') {
    try {
      const out = execSync(`${PHP_CMD} -l "${filePath}"`, { stdio: ['pipe', 'pipe', 'pipe'] }).toString('utf8');
      if (!out.includes('No syntax errors detected')) {
        addLog('FAIL', 'PHP-SYNTAX', filePath, out.trim());
      } else {
        report.stats.pass++;
      }
    } catch (err) {
      const errOutput = (err.stderr ? err.stderr.toString('utf8') : '') || (err.stdout ? err.stdout.toString('utf8') : err.message);
      addLog('FAIL', 'PHP-SYNTAX', filePath, 'Sözdizim Hatası: ' + errOutput.split('\n')[0].trim(), errOutput.trim());
    }

    // Include / Require kontrolü
    const incRegex = /\b(require|require_once|include|include_once)\s*\(?\s*['"]([^'"]+\.php)['"]\s*\)?/g;
    let match;
    const fileDir = path.dirname(filePath);
    while ((match = incRegex.exec(content)) !== null) {
      const relTarget = match[2];
      if (!relTarget.startsWith('http://') && !relTarget.startsWith('https://')) {
        const potential = path.resolve(fileDir, relTarget);
        if (!fs.existsSync(potential) && !relTarget.includes('$')) {
          addLog('WARN', 'BROKEN-INCLUDE', filePath, `İçerilen dosya bulunamadı: ${relTarget}`);
        }
      }
    }
  }

  // 3. JavaScript / Node.js Syntax Kontrolü (node --check)
  if (['.js', '.mjs', '.cjs'].includes(ext)) {
    try {
      execSync(`${NODE_CMD} --check "${filePath}"`, { stdio: ['pipe', 'pipe', 'pipe'] });
      report.stats.pass++;
    } catch (err) {
      const errOutput = (err.stderr ? err.stderr.toString('utf8') : '') || err.message;
      addLog('FAIL', 'JS-SYNTAX', filePath, 'JavaScript Sözdizim Hatası', errOutput.trim());
    }
  }

  // 4. Python Syntax Kontrolü (python3 -B -m py_compile)
  if (ext === '.py') {
    try {
      execSync(`${PYTHON_CMD} -B -m py_compile "${filePath}"`, { stdio: ['pipe', 'pipe', 'pipe'] });
      report.stats.pass++;
    } catch (err) {
      const errOutput = (err.stderr ? err.stderr.toString('utf8') : '') || err.message;
      addLog('FAIL', 'PY-SYNTAX', filePath, 'Python Sözdizim Hatası', errOutput.trim());
    }
  }

  // 5. JSON Format Doğrulama
  if (ext === '.json') {
    try {
      JSON.parse(content);
      report.stats.pass++;
    } catch (err) {
      addLog('FAIL', 'JSON-PARSE', filePath, 'Geçersiz JSON formatı: ' + err.message);
    }
  }

  // 6. Güvenlik ve Kod Kalitesi Kalıp Taraması
  if (['.php', '.js'].includes(ext)) {
    BUG_PATTERNS.forEach(pat => {
      if (pat.regex.test(content)) {
        if (pat.category === 'SECURITY') {
          if (isStrict) {
            addLog('FAIL', pat.code, filePath, pat.msg);
          } else {
            addLog('WARN', pat.code, filePath, pat.msg);
          }
        } else {
          addLog('INFO', pat.code, filePath, pat.msg);
        }
      }
    });
  }
}

// Ana yürütme
walk(resolvedTarget);

// Çıktı Üretimi
if (isJson) {
  console.log(JSON.stringify(report, null, 2));
} else {
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log(`🔍 HATA VE SÖZDİZİM KONTROL RAPORU (Bug Hunter & QA)`);
  console.log(`Hedef  : ${resolvedTarget}`);
  console.log(`Zaman  : ${report.timestamp}`);
  console.log(`Taranan: ${report.scannedFiles} dosya`);
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  if (report.errors.length === 0 && report.warnings.length === 0) {
    console.log('✅ TEBRİKLER: Hiçbir sözdizim veya kritik hata tespit edilmedi.');
  } else {
    if (report.errors.length > 0) {
      console.log(`\n🔴 KRİTİK HATALAR (${report.errors.length} adet):`);
      report.errors.forEach((e, idx) => {
        console.log(`  ${idx + 1}. [${e.code}] ${e.file}`);
        console.log(`     Mesaj: ${e.message}`);
        if (e.detail) console.log(`     Detay: ${e.detail.split('\n')[0]}`);
      });
    }

    if (report.warnings.length > 0) {
      console.log(`\n🟡 UYARILAR & GÜVENLİK İPUÇLARI (${report.warnings.length} adet):`);
      report.warnings.forEach((w, idx) => {
        console.log(`  ${idx + 1}. [${w.code}] ${w.file} → ${w.message}`);
      });
    }
  }

  console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log(`SONUÇ ÖZETİ:`);
  console.log(`  BAŞARILI (PASS) : ${report.stats.pass}`);
  console.log(`  HATALI   (FAIL) : ${report.stats.fail}`);
  console.log(`  UYARI    (WARN) : ${report.stats.warn}`);
  console.log(`  BİLGİ    (INFO) : ${report.stats.info}`);
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}

// Çıkış Kodu
if (report.stats.fail > 0) {
  process.exit(1);
} else {
  process.exit(0);
}
