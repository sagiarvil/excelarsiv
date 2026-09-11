/**
 * runtime_resolver.js — Antigravity Çapraz Platform Çalışma Zamanı ve Yürütücü Çözümleyici
 * 
 * Hem macOS (Apple Silicon M1-M4 & Intel) hem de Windows (ve Linux) sistemlerinde;
 * Homebrew, Laravel Herd, Valet, MAMP, Laravel Herd, XAMPP, Laragon, WAMP veya bağımsız
 * kurulumlar olsun veya olmasın; PHP, Composer, Git, Bash/Zsh, Node.js ve Python
 * araçlarının en son/en yüksek sürümlerini dinamik olarak otomatik keşfeder ve sisteme bağlar.
 * 
 * Standart: Asla sabit kullanıcı veya tek bir işletim sistemi varsaymaz.
 */

"use strict";

const fs = require("fs");
const path = require("path");
const os = require("os");
const { execSync } = require("child_process");

const IS_WIN = process.platform === "win32";
const IS_MAC = process.platform === "darwin";
const IS_LINUX = process.platform === "linux";

const USER_HOME = process.env.USERPROFILE || process.env.HOME || os.homedir();
const CONFIG_DIR = path.join(USER_HOME, ".gemini", "config");
const RUNTIME_CACHE_PATH = path.join(CONFIG_DIR, "runtime_paths.json");

function isExecutable(filePath) {
  try {
    if (!fs.existsSync(filePath)) return false;
    const stat = fs.statSync(filePath);
    return stat.isFile();
  } catch (e) {
    return false;
  }
}

function runCmd(cmd, timeout = 3000) {
  try {
    return execSync(cmd, { stdio: ["pipe", "pipe", "pipe"], timeout }).toString().trim();
  } catch (e) {
    return null;
  }
}

function parseSemVer(str) {
  const m = str.match(/(\d+)\.(\d+)(?:\.(\d+))?/);
  if (!m) return [0, 0, 0];
  return [parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3] || 0, 10)];
}

function compareSemVer(a, b) {
  for (let i = 0; i < 3; i++) {
    if (a[i] > b[i]) return 1;
    if (a[i] < b[i]) return -1;
  }
  return 0;
}

// 1. NODE.JS & NPM
function resolveNode() {
  const pathNode = runCmd("node -v");
  if (pathNode) {
    return { name: "Node.js", path: "node", version: pathNode, isGlobal: true };
  }
  const candidates = [];
  if (IS_WIN) {
    candidates.push(
      "C:\\Program Files\\nodejs\\node",
      "C:\\Program Files (x86)\\nodejs\\node",
      path.join(USER_HOME, "AppData", "Roaming", "nvm", "current", "node")
    );
  } else if (IS_MAC) {
    candidates.push(
      "/opt/homebrew/bin/node",
      "/usr/local/bin/node",
      "/usr/bin/node",
      path.join(USER_HOME, ".local", "share", "fnm", "current", "bin", "node")
    );
    const nvmDir = path.join(USER_HOME, ".nvm", "versions", "node");
    if (fs.existsSync(nvmDir)) {
      try {
        const subs = fs.readdirSync(nvmDir).sort().reverse();
        for (const sub of subs) {
          candidates.push(path.join(nvmDir, sub, "bin", "node"));
        }
      } catch (e) {}
    }
  } else {
    candidates.push("/usr/bin/node", "/usr/local/bin/node");
  }

  for (const c of candidates) {
    if (isExecutable(c)) {
      const v = runCmd(`"${c}" -v`);
      if (v) return { name: "Node.js", path: c, version: v, isGlobal: false };
    }
  }
  return null;
}

// 2. PYTHON
function resolvePython() {
  const pathPy = runCmd("python3 --version") || runCmd("python --version") || (IS_WIN ? runCmd("py -3 --version") : null);
  if (pathPy) {
    return { name: "Python", path: "python3", version: pathPy, isGlobal: true };
  }
  const candidates = [];
  if (IS_WIN) {
    const programsDir = path.join(USER_HOME, "AppData", "Local", "Programs", "Python");
    if (fs.existsSync(programsDir)) {
      try {
        const subs = fs.readdirSync(programsDir).sort().reverse();
        for (const sub of subs) {
          candidates.push(path.join(programsDir, sub, "python3"));
        }
      } catch (e) {}
    }
    candidates.push(
      "C:\\Python314\\python3",
      "C:\\Python312\\python3",
      "C:\\Python311\\python3",
      "C:\\Laravel Herd\\bin\\python\\python-3.14.7\\python3"
    );
  } else if (IS_MAC) {
    candidates.push(
      "/opt/homebrew/bin/python3",
      "/usr/local/bin/python3",
      "/Library/Frameworks/Python.framework/Versions/Current/bin/python3",
      path.join(USER_HOME, ".pyenv", "shims", "python3"),
      "/usr/bin/python3"
    );
  } else {
    candidates.push("/usr/bin/python3", "/usr/local/bin/python3");
  }

  for (const c of candidates) {
    if (isExecutable(c)) {
      const v = runCmd(`"${c}" --version`);
      if (v) return { name: "Python", path: c, version: v, isGlobal: false };
    }
  }
  return null;
}

// 3. GIT & BASH / ZSH
function resolveGit() {
  const pathGit = runCmd("git --version");
  const gitObj = { name: "Git", path: pathGit ? "git" : null, version: pathGit, bash: null };

  if (IS_WIN) {
    const commonGitRoots = [
      "C:\\Program Files\\Git",
      "C:\\Program Files (x86)\\Git",
      path.join(USER_HOME, "AppData", "Local", "Programs", "Git")
    ];
    let gitRoot = null;
    for (const root of commonGitRoots) {
      if (fs.existsSync(root)) {
        gitRoot = root;
        break;
      }
    }
    if (!gitObj.path && gitRoot) {
      const gitExe = path.join(gitRoot, "cmd", "git");
      if (isExecutable(gitExe)) {
        gitObj.path = gitExe;
        gitObj.version = runCmd(`"${gitExe}" --version`);
      }
    }
    const pathBash = runCmd("bash --version");
    if (pathBash) {
      gitObj.bash = { path: "bash", version: pathBash.split("\n")[0] };
    } else if (gitRoot) {
      const bashExe = path.join(gitRoot, "bin", "bash");
      if (isExecutable(bashExe)) {
        const bv = runCmd(`"${bashExe}" --version`);
        gitObj.bash = { path: bashExe, version: bv ? bv.split("\n")[0] : "Git Bash" };
      }
    }
  } else {
    const macGitCandidates = ["/usr/bin/git", "/opt/homebrew/bin/git", "/usr/local/bin/git"];
    if (!gitObj.path) {
      for (const gc of macGitCandidates) {
        if (isExecutable(gc)) {
          gitObj.path = gc;
          gitObj.version = runCmd(`"${gc}" --version`);
          break;
        }
      }
    }
    const shellCandidates = ["/bin/zsh", "/bin/bash", "/opt/homebrew/bin/bash", "/usr/bin/bash"];
    for (const sc of shellCandidates) {
      if (isExecutable(sc)) {
        const sv = runCmd(`"${sc}" --version`);
        gitObj.bash = { path: sc, version: sv ? sv.split("\n")[0] : path.basename(sc) };
        break;
      }
    }
  }

  return gitObj.path ? gitObj : null;
}

// 4. PHP (EVRENSEL ÇOKLU ORTAM VE SÜRÜM SEÇİCİ - WINDOWS & MACOS)
function resolvePhp() {
  const discovered = [];

  // A) PATH PHP
  const pathPhp = runCmd("php -v");
  if (pathPhp) {
    discovered.push({ path: "php", versionRaw: pathPhp.split("\n")[0], isGlobal: true });
  }

  // B) İşletim sistemine göre dizin taraması
  const searchDirs = [];
  const searchFiles = [];

  if (IS_WIN) {
    searchDirs.push(
      "C:\\php",
      "C:\\tools\\php",
      "C:\\Program Files\\php",
      "C:\\xampp\\php",
      "C:\\laragon\\bin\\php",
      "C:\\wamp64\\bin\\php",
      path.join(USER_HOME, ".config", "herd", "bin"),
      "C:\\Laravel Herd\\bin\\php"
    );
  } else if (IS_MAC) {
    searchFiles.push(
      "/opt/homebrew/bin/php",
      "/usr/local/bin/php",
      "/opt/homebrew/opt/php/bin/php",
      "/opt/homebrew/opt/php@8.4/bin/php",
      "/opt/homebrew/opt/php@8.3/bin/php",
      "/opt/homebrew/opt/php@8.2/bin/php",
      "/usr/local/opt/php@8.4/bin/php",
      "/usr/local/opt/php@8.3/bin/php",
      "/usr/bin/php",
      "/opt/local/bin/php",
      path.join(USER_HOME, "Library", "Application Support", "Herd", "bin", "php")
    );
    const herdPhpDir = path.join(USER_HOME, "Library", "Application Support", "Herd", "config", "php");
    if (fs.existsSync(herdPhpDir)) searchDirs.push(herdPhpDir);

    const mampPhpDir = "/Applications/MAMP/bin/php";
    if (fs.existsSync(mampPhpDir)) searchDirs.push(mampPhpDir);
  } else {
    searchFiles.push("/usr/bin/php", "/usr/local/bin/php", "/usr/bin/php8.4", "/usr/bin/php8.3", "/usr/bin/php8.2");
  }

  for (const f of searchFiles) {
    if (isExecutable(f)) {
      const v = runCmd(`"${f}" -v`);
      if (v) discovered.push({ path: f, versionRaw: v.split("\n")[0], isGlobal: false });
    }
  }

  for (const sDir of searchDirs) {
    if (!fs.existsSync(sDir)) continue;

    const directNames = IS_WIN ? ["php"] : ["php", "bin/php"];
    for (const dName of directNames) {
      const directPhp = path.join(sDir, dName);
      if (isExecutable(directPhp)) {
        const v = runCmd(`"${directPhp}" -v`);
        if (v) discovered.push({ path: directPhp, versionRaw: v.split("\n")[0], isGlobal: false });
      }
    }

    try {
      const subs = fs.readdirSync(sDir, { withFileTypes: true });
      for (const sub of subs) {
        if (sub.isDirectory()) {
          const subNames = IS_WIN ? ["php"] : ["php", "bin/php"];
          for (const sName of subNames) {
            const subPhp = path.join(sDir, sub.name, sName);
            if (isExecutable(subPhp)) {
              const v = runCmd(`"${subPhp}" -v`);
              if (v) discovered.push({ path: subPhp, versionRaw: v.split("\n")[0], isGlobal: false });
            }
          }
        }
      }
    } catch (e) {}
  }

  if (discovered.length === 0) return null;

  discovered.sort((a, b) => {
    const verA = parseSemVer(a.versionRaw);
    const verB = parseSemVer(b.versionRaw);
    return compareSemVer(verB, verA);
  });

  const best = discovered[0];
  return {
    name: "PHP",
    path: best.path,
    version: best.versionRaw,
    allDiscovered: discovered.map(d => ({ path: d.path, version: d.versionRaw }))
  };
}

// 5. COMPOSER (EVRENSEL ÇÖZÜMLEME - WINDOWS & MACOS)
function resolveComposer(phpPath = "php") {
  const pathComp = runCmd("composer -V");
  if (pathComp) {
    return { name: "Composer", path: "composer", version: pathComp.split("\n")[0], isGlobal: true };
  }

  const candidates = [];
  if (IS_WIN) {
    candidates.push(
      "C:\\ProgramData\\ComposerSetup\\bin\\composer",
      path.join(USER_HOME, "AppData", "Roaming", "Composer", "vendor", "bin", "composer"),
      "C:\\Laravel Herd\\bin\\composer\\composer",
      "C:\\Laravel Herd\\bin\\composer\\composer.phar",
      "C:\\tools\\composer\\composer.phar",
      "C:\\ProgramData\\ComposerSetup\\bin\\composer.phar",
      path.join(USER_HOME, ".gemini", "tools", "composer", "composer"),
      path.join(USER_HOME, ".gemini", "tools", "composer", "composer.phar")
    );
  } else if (IS_MAC) {
    candidates.push(
      "/opt/homebrew/bin/composer",
      "/usr/local/bin/composer",
      path.join(USER_HOME, ".composer", "vendor", "bin", "composer"),
      path.join(USER_HOME, ".config", "composer", "vendor", "bin", "composer"),
      path.join(USER_HOME, "Library", "Application Support", "Herd", "bin", "composer"),
      path.join(USER_HOME, ".gemini", "tools", "composer", "composer"),
      path.join(USER_HOME, ".gemini", "tools", "composer", "composer.phar")
    );
  } else {
    candidates.push(
      "/usr/local/bin/composer",
      "/usr/bin/composer",
      path.join(USER_HOME, ".config", "composer", "vendor", "bin", "composer"),
      path.join(USER_HOME, ".gemini", "tools", "composer", "composer")
    );
  }

  for (const c of candidates) {
    if (fs.existsSync(c)) {
      if (c.endsWith(".phar")) {
        const phpExec = phpPath.includes(" ") ? `"${phpPath}"` : phpPath;
        const v = runCmd(`${phpExec} "${c}" -V`);
        if (v) return { name: "Composer", path: `${phpExec} "${c}"`, rawPath: c, version: v.split("\n")[0], isPhar: true };
      } else {
        const v = runCmd(`"${c}" -V`);
        if (v) return { name: "Composer", path: c, version: v.split("\n")[0], isGlobal: false };
      }
    }
  }

  return null;
}

// 6. VERİTABANI VE SERVİS İSTEMCİLERİ (MySQL, PostgreSQL, Redis)
function resolveDatabaseClients() {
  const clients = {};

  // MySQL
  const mysqlOut = runCmd("mysql --version");
  if (mysqlOut) {
    clients.mysql = { path: "mysql", version: mysqlOut.split("\n")[0] };
  } else if (IS_WIN && isExecutable("/opt/homebrew/bin/mysql")) {
    clients.mysql = { path: "/opt/homebrew/bin/mysql", version: "MySQL (Homebrew / DBngin)" };
  } else if (IS_MAC && isExecutable("/opt/homebrew/bin/mysql")) {
    clients.mysql = { path: "/opt/homebrew/bin/mysql", version: "Homebrew MySQL" };
  }

  // PostgreSQL
  const psqlOut = runCmd("psql --version");
  if (psqlOut) {
    clients.psql = { path: "psql", version: psqlOut.split("\n")[0] };
  } else if (IS_WIN && isExecutable("/opt/homebrew/bin/psql")) {
    clients.psql = { path: "/opt/homebrew/bin/psql", version: "PostgreSQL (Homebrew / DBngin)" };
  } else if (IS_MAC && isExecutable("/opt/homebrew/bin/psql")) {
    clients.psql = { path: "/opt/homebrew/bin/psql", version: "Homebrew PostgreSQL" };
  }

  // Redis
  const redisOut = runCmd("redis-cli --version");
  if (redisOut) {
    clients.redis = { path: "redis-cli", version: redisOut.split("\n")[0] };
  } else if (IS_WIN && isExecutable("/opt/homebrew/bin/redis-cli")) {
    clients.redis = { path: "/opt/homebrew/bin/redis-cli", version: "Redis (Homebrew / DBngin)" };
  } else if (IS_MAC && isExecutable("/opt/homebrew/bin/redis-cli")) {
    clients.redis = { path: "/opt/homebrew/bin/redis-cli", version: "Homebrew Redis" };
  }

  return clients;
}

// 7. ANTIGRAVITY CLI YÜRÜTÜCÜ BULUCU (agy / agy)
function getAgyExecutable() {
  const localBin = path.join(USER_HOME, ".gemini", "bin", IS_WIN ? "agy" : "agy");
  if (fs.existsSync(localBin)) return localBin;
  try {
    const whichCmd = IS_WIN ? "where agy" : "which agy";
    const out = runCmd(whichCmd);
    if (out) return out.split("\n")[0].trim();
  } catch (e) {}
  return IS_WIN ? "agy" : "agy";
}

// TÜM SİSTEMİ TARA VE RAPORLA
function resolveAll(saveCache = true) {
  const node = resolveNode();
  const python = resolvePython();
  const git = resolveGit();
  const php = resolvePhp();
  const composer = resolveComposer(php ? php.path : "php");
  const databaseClients = resolveDatabaseClients();
  const agy = getAgyExecutable();

  const result = {
    platform: process.platform,
    arch: os.arch(),
    isWindows: IS_WIN,
    isMac: IS_MAC,
    isLinux: IS_LINUX,
    resolvedAt: new Date().toISOString(),
    agyExecutable: agy,
    tools: {
      node,
      python,
      git,
      php,
      composer,
      databaseClients
    }
  };

  if (saveCache) {
    if (!fs.existsSync(CONFIG_DIR)) fs.mkdirSync(CONFIG_DIR, { recursive: true });
    fs.writeFileSync(RUNTIME_CACHE_PATH, JSON.stringify(result, null, 2) + "\n", "utf8");
  }

  return result;
}

// CLI Olarak Çalıştırıldığında
if (require.main === module) {
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("🔍 macOS GELİŞTİRİCİ ÇALIŞMA ZAMANI RAPORU (Apple Silicon & Intel)");
  console.log("Kullanıcı : " + USER_HOME);
  console.log("Platform  : " + (IS_MAC ? "🍎 macOS" : IS_WIN ? "🪟 Windows" : "🐧 Linux") + " (" + process.platform + " " + os.arch() + ")");
  console.log("CLI Aracı : " + getAgyExecutable());
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

  const res = resolveAll(true);
  const t = res.tools;

  console.log("  1. Node.js  : " + (t.node ? `✅ ${t.node.version} (${t.node.path})` : "❌ BULUNAMADI"));
  console.log("  2. Python   : " + (t.python ? `✅ ${t.python.version} (${t.python.path})` : "❌ BULUNAMADI"));
  console.log("  3. Git      : " + (t.git ? `✅ ${t.git.version} (${t.git.path})` : "❌ BULUNAMADI"));
  console.log("  4. Kabuk    : " + (t.git && t.git.bash ? `✅ ${t.git.bash.version} (${t.git.bash.path})` : "❌ BULUNAMADI"));
  console.log("  5. PHP      : " + (t.php ? `✅ ${t.php.version} (${t.php.path})` : "❌ BULUNAMADI"));
  if (t.php && t.php.allDiscovered && t.php.allDiscovered.length > 1) {
    console.log(`     ℹ️ Toplam ${t.php.allDiscovered.length} PHP sürümü keşfedildi, en yükseği seçildi.`);
  }
  console.log("  6. Composer : " + (t.composer ? `✅ ${t.composer.version} (${t.composer.path})` : "❌ BULUNAMADI"));

  console.log("\n📁 Yapılandırma Kaydedildi: " + RUNTIME_CACHE_PATH);
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
}

module.exports = {
  resolveAll,
  resolveNode,
  resolvePython,
  resolveGit,
  resolvePhp,
  resolveComposer,
  resolveDatabaseClients,
  getAgyExecutable,
  IS_WIN,
  IS_MAC,
  IS_LINUX
};
