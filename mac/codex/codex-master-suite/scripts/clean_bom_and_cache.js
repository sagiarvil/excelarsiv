const fs = require('fs');
const path = require('path');

function walkDir(dir, callback) {
  fs.readdirSync(dir).forEach(f => {
    let dirPath = path.join(dir, f);
    let isDirectory = fs.statSync(dirPath).isDirectory();
    if (isDirectory) {
      if (f !== 'vendor' && f !== 'node_modules' && f !== '.git') {
        walkDir(dirPath, callback);
      }
    } else {
      callback(path.join(dir, f));
    }
  });
}

const rootDir = '$HOME/Sites/satis';
let count = 0;

walkDir(rootDir, filePath => {
  if (filePath.endsWith('.php') || filePath.endsWith('.json') || filePath.endsWith('.html') || filePath.endsWith('.js') || filePath.endsWith('.css')) {
    const buf = fs.readFileSync(filePath);
    if (buf.length >= 3 && buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) {
      const cleanBuf = buf.slice(3);
      fs.writeFileSync(filePath, cleanBuf);
      console.log(`[BOM TEMİZLENDİ]: ${filePath.replace(rootDir, '')}`);
      count++;
    }
  }
});

console.log(`Toplam ${count} dosyadaki UTF-8 BOM temizlendi.`);

// Cache klasörünü boşalt
const cacheDir = '$HOME/Sites/satis/storage/cache';
if (fs.existsSync(cacheDir)) {
  fs.readdirSync(cacheDir).forEach(f => {
    try {
      const fp = path.join(cacheDir, f);
      if (fs.statSync(fp).isFile()) fs.unlinkSync(fp);
    } catch (e) {}
  });
  console.log('PageCache (storage/cache) tamamen temizlendi.');
}
