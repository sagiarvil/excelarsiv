import fs from 'fs';
import path from 'path';
import sharp from 'sharp';

const sourceDir = 'urun-gorselleri';
const destDir = 'public/images/kapak';

if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir, { recursive: true });
}

// Extract slugs from kapak.ts
const kapakFile = fs.readFileSync('src/lib/kapak.ts', 'utf8');
const match = kapakFile.match(/const PREMIUM_KAPAK_SLUGS = new Set<string>\(\[\s*([\s\S]*?)\s*\]\);/);
if (!match) {
  console.error("Could not find PREMIUM_KAPAK_SLUGS in kapak.ts");
  process.exit(1);
}
const slugs = match[1].split(',').map(s => s.trim().replace(/['"]/g, '')).filter(s => s.length > 0);
console.log(`Found ${slugs.length} slugs.`);

// Get all files in sourceDir
const files = fs.readdirSync(sourceDir).filter(f => f.endsWith('.png') || f.endsWith('.jpg') || f.endsWith('.jpeg'));

// Sort files by size (descending) to get the best ones
const fileStats = files.map(f => {
  const p = path.join(sourceDir, f);
  return { file: f, size: fs.statSync(p).size };
});
fileStats.sort((a, b) => b.size - a.size);

// Take exactly slugs.length (124) best images
if (fileStats.length < slugs.length) {
    console.error(`Not enough unique images! Found ${fileStats.length}, need ${slugs.length}`);
    process.exit(1);
}
const selectedFiles = fileStats.slice(0, slugs.length).map(f => f.file);
console.log(`Selected ${selectedFiles.length} unique highest-quality images.`);

const targetWidth = 3840;
const targetHeight = 2400;

async function processAll() {
  let count = 0;
  for (let i = 0; i < slugs.length; i++) {
    const slug = slugs[i];
    const sourceImage = path.join(sourceDir, selectedFiles[i]);
    const destPath = path.join(destDir, `${slug}.png`);
    
    // Process unique image with no repetitive color shifts, just pure enhancement
    await sharp(sourceImage)
      .resize({
        width: targetWidth,
        height: targetHeight,
        fit: 'contain',
        background: '#f8fafc',
        kernel: sharp.kernel.lanczos3
      })
      .modulate({
        saturation: 1.05,
        lightness: 1.01
      })
      .sharpen({
        sigma: 1.2,
        m1: 1.0,
        m2: 0.5
      })
      .png({
        quality: 100,
        compressionLevel: 9
      })
      .toFile(destPath);
      
    count++;
    if (count % 20 === 0) {
      console.log(`Processed ${count}/${slugs.length}...`);
    }
  }
  console.log(`✅ All ${slugs.length} UNIQUE images processed to Premium Ultra Pro 4K!`);
}

processAll().catch(err => {
  console.error("Error processing images:", err);
  process.exit(1);
});
