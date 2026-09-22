import fs from 'fs';
import path from 'path';
import sharp from 'sharp';
import crypto from 'crypto';

const sourceDir = 'urun-gorselleri';
const destDir = 'public/images/kapak';

// The 5 highest quality images by file size
const premiumTemplates = [
  'Ekran Resmi 2026-09-21 21.06.31.png',
  'Ekran Resmi 2026-09-21 20.50.11.png',
  'Ekran Resmi 2026-09-21 21.06.19.png',
  'Ekran Resmi 2026-09-21 21.02.19.png',
  'Ekran Resmi 2026-09-21 20.43.13.png'
];

if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir, { recursive: true });
}

const targetWidth = 3840;
const targetHeight = 2400;

function hashStringToNumber(str) {
  const hash = crypto.createHash('md5').update(str).digest('hex');
  return parseInt(hash.substring(0, 8), 16);
}

// Extract slugs from kapak.ts
const kapakFile = fs.readFileSync('src/lib/kapak.ts', 'utf8');
const match = kapakFile.match(/const PREMIUM_KAPAK_SLUGS = new Set<string>\(\[\s*([\s\S]*?)\s*\]\);/);
if (!match) {
  console.error("Could not find PREMIUM_KAPAK_SLUGS in kapak.ts");
  process.exit(1);
}
const slugs = match[1].split(',').map(s => s.trim().replace(/['"]/g, '')).filter(s => s.length > 0);

async function processAll() {
  let count = 0;
  for (const slug of slugs) {
    const destPath = path.join(destDir, `${slug}.png`);
    
    const hashVal = hashStringToNumber(slug);
    
    const templateIdx = hashVal % premiumTemplates.length;
    const sourceImage = path.join(sourceDir, premiumTemplates[templateIdx]);
    
    // Determine a hue rotation (0 to 360)
    // To make them distinctly different, we'll pick steps
    const hueShift = (hashVal % 12) * 30; // 0, 30, 60... 330
    
    await sharp(sourceImage)
      .modulate({
        hue: hueShift,
        saturation: 1.15,
        lightness: 1.02
      })
      .resize({
        width: targetWidth,
        height: targetHeight,
        fit: 'contain',
        background: '#f8fafc',
        kernel: sharp.kernel.lanczos3
      })
      .sharpen({
        sigma: 1.5,
        m1: 1.2,
        m2: 0.8
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
  console.log(`✅ All ${slugs.length} images processed to Premium Ultra Pro 4K!`);
}

processAll().catch(err => {
  console.error("Error processing images:", err);
  process.exit(1);
});
