/**
 * tool_installer.js — Antigravity macOS Yerel Araç Yükleyici
 * Apple Silicon (M1/M2/M3/M4) ve Intel Mac Destekli
 * 
 * Homebrew ve yerel macOS kaynaklarından PHP, Composer, Git, Zsh/Bash,
 * Node.js ve Python araçlarının son sürümlerini otomatik kurar ve yapılandırır.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execSync } = require('child_process');
const https = require('https');

const USER_HOME = process.env.HOME || os.homedir();
const TOOLS_DIR = path.join(USER_HOME, '.gemini', 'tools');
const RUNTIME_RESOLVER_PATH = path.join(__dirname, 'runtime_resolver.js');

const { resolveAll, getAgyExecutable } = require(RUNTIME_RESOLVER_PATH);

function logHeader(title) {
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log(title);
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}

function runCmd(cmd, timeout = 120000) {
  try {
    return execSync(cmd, { stdio: ['pipe', 'pipe', 'pipe'], timeout }).toString().trim();
  } catch (e) {
    return null;
  }
}

function downloadFile(url, destPath) {
  return new Promise((resolve, reject) => {
    if (!fs.existsSync(path.dirname(destPath))) {
      fs.mkdirSync(path.dirname(destPath), { recursive: true });
    }
    const file = fs.createWriteStream(destPath);
    https.get(url, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        downloadFile(res.headers.location, destPath).then(resolve).catch(reject);
        return;
      }
      if (res.statusCode !== 200) return reject(new Error('HTTP ' + res.statusCode));
      res.pipe(file);
      file.on('finish', () => { file.close(); resolve(destPath); });
    }).on('error', (err) => { fs.unlink(destPath, () => {}); reject(err); });
  });
}

async function installComposerPortable(phpPath) {
  console.log('\n📦 Composer resmi kaynaktan indiriliyor (https://getcomposer.org/composer.phar)...');
  const compDir = path.join(TOOLS_DIR, 'composer');
  const pharPath = path.join(compDir, 'composer.phar');
  try {
    await downloadFile('https://getcomposer.org/composer.phar', pharPath);
    console.log('   ✅ composer.phar başarıyla indirildi: ' + pharPath);
    const phpCmd = phpPath.includes(' ') ? `"${phpPath}"` : phpPath;
    const shPath = path.join(compDir, 'composer');
    const shContent = `#!/usr/bin/env bash\nDIR="$(cd "$(dirname "\$0")" && pwd)"\nexec ${phpCmd} "\$DIR/composer.phar" "\$@"\n`;
    fs.writeFileSync(shPath, shContent, { encoding: 'utf8', mode: 0o755 });
    try { fs.chmodSync(shPath, 0o755); } catch (e) {}
    console.log('   ✅ composer çalıştırılabilir bash sarmalayıcısı oluşturuldu (macOS).');
    const v = runCmd(`"${shPath}" -V`);
    console.log('   🎉 Yüklendi: ' + (v ? v.split('\n')[0] : 'Composer Hazır'));
    return true;
  } catch (err) {
    console.error('   ❌ Composer indirme hatası:', err.message);
    return false;
  }
}

async function main() {
  const isCheck = process.argv.includes('--check');
  const isAll = process.argv.includes('--all');

  logHeader('🍎 macOS GELİŞTİRİCİ ARAÇLARI YÜKLEYİCİ (Homebrew & Native)');
  console.log('Kullanıcı : ' + USER_HOME);
  console.log('Platform  : 🍎 macOS (' + os.arch() + ')');
  console.log('CLI Aracı : ' + getAgyExecutable());
  console.log('Kabuk     : ' + (process.env.SHELL || '/bin/zsh'));
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

  console.log('🔍 1. Mevcut macOS Araçları Taranıyor...');
  const current = resolveAll(true);
  const t = current.tools;

  const hasBrew = runCmd('brew --version') !== null;
  console.log('   macOS Paket Yöneticisi (Homebrew): ' + (hasBrew ? '✅ Hazır' : '⚠️ Bulunamadı'));
  if (!hasBrew) {
    console.log('   💡 İPUCU: macOS üzerinde Homebrew kurulumu için Terminalden çalıştırın:');
    console.log('      /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"');
  }

  const tasks = [];
  if (!t.node) {
    console.log('❌ Node.js eksik.');
    if (hasBrew) tasks.push({ name: 'Node.js', cmd: 'brew install node' });
  } else {
    console.log('✅ Node.js: ' + t.node.version + ' (' + t.node.path + ')');
  }

  if (!t.python) {
    console.log('❌ Python eksik.');
    if (hasBrew) tasks.push({ name: 'Python', cmd: 'brew install python' });
  } else {
    console.log('✅ Python: ' + t.python.version + ' (' + t.python.path + ')');
  }

  if (!t.git) {
    console.log('❌ Git eksik.');
    if (hasBrew) tasks.push({ name: 'Git', cmd: 'brew install git' });
  } else {
    console.log('✅ Git: ' + t.git.version + ' (' + t.git.path + ')');
  }

  if (!t.php) {
    console.log('❌ PHP eksik.');
    if (hasBrew) tasks.push({ name: 'PHP', cmd: 'brew install php' });
    else console.log('   💡 Alternatif: Laravel Herd kurabilirsiniz: https://herd.laravel.com');
  } else {
    console.log('✅ PHP: ' + t.php.version + ' (' + t.php.path + ')');
  }

  if (!t.composer) {
    console.log('❌ Composer eksik.');
    if (hasBrew) tasks.push({ name: 'Composer', cmd: 'brew install composer' });
  } else {
    console.log('✅ Composer: ' + t.composer.version + ' (' + t.composer.path + ')');
  }

  if (isCheck) {
    console.log('\nℹ️  --check modu: Kurulum yapılmadı.');
    return;
  }

  if (tasks.length === 0 && t.composer) {
    console.log('\n🎉 Harika! Tüm temel macOS geliştirici araçları kurulu ve eksiksiz!');
    return;
  }

  if (tasks.length > 0) {
    console.log('\n⚡ Eksik Araçlar Kuruluyor (' + tasks.length + ' adet)...');
    for (const task of tasks) {
      console.log('\n▶ Kuruluyor: ' + task.name + ' (' + task.cmd + ')...');
      try {
        execSync(task.cmd, { stdio: 'inherit' });
        console.log('   ✅ ' + task.name + ' başarıyla kuruldu.');
      } catch (err) {
        console.error('   ❌ ' + task.name + ' kurulum hatası: ' + err.message);
      }
    }
  }

  // Eğer PHP var ama Composer yoksa ve brew başarısızsa portable composer.phar indir
  const afterPHP = resolveAll(false).tools.php;
  const afterComp = resolveAll(false).tools.composer;
  if (afterPHP && !afterComp) {
    await installComposerPortable(afterPHP.path);
  }

  console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('🎯 Nihai Çalışma Zamanı Taranıyor...');
  resolveAll(true);
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
}

if (require.main === module) {
  main().catch((err) => {
    console.error('Kritik Hata:', err);
    process.exit(1);
  });
}

module.exports = { main, installComposerPortable };
