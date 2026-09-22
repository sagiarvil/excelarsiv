import fs from 'fs';
import path from 'path';
import sharp from 'sharp';

const sourceDir = 'urun-gorselleri';
const destDir = 'public/images/kapak';

if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir, { recursive: true });
}

// 1. Get Slugs
const kapakFile = fs.readFileSync('src/lib/kapak.ts', 'utf8');
const match = kapakFile.match(/const PREMIUM_KAPAK_SLUGS = new Set<string>\(\[\s*([\s\S]*?)\s*\]\);/);
const slugs = match[1].split(',').map(s => s.trim().replace(/['"]/g, '')).filter(s => s.length > 0);

// 2. Get Source Images
const files = fs.readdirSync(sourceDir).filter(f => f.endsWith('.png') || f.endsWith('.jpg') || f.endsWith('.jpeg'));
const fileStats = files.map(f => {
  return { file: f, size: fs.statSync(path.join(sourceDir, f)).size };
});
fileStats.sort((a, b) => b.size - a.size); // Best quality first

if (fileStats.length < slugs.length) {
    console.error("Not enough unique images");
    process.exit(1);
}
const selectedFiles = fileStats.slice(0, slugs.length).map(f => f.file);

// 3. Define the Canvas (Apple 16" MacBook Pro resolution: 2560x1600)
// This is exactly 16:10 and high-res, but not so insanely huge that a 1000x1000 image looks tiny.
const canvasW = 2560;
const canvasH = 1600;
const cardSize = 1200; // The white floating card

async function processAll() {
  let count = 0;
  for (let i = 0; i < slugs.length; i++) {
    const slug = slugs[i];
    const sourceImage = path.join(sourceDir, selectedFiles[i]);
    const destPath = path.join(destDir, `${slug}.webp`);
    
    // We create a beautiful, modern LIGHT background in SVG
    // 1. Soft radial gradient background
    // 2. A subtle abstract mesh shape to make it modern
    // 3. A crisp white card with a soft shadow where the actual Excel image will sit.
    const bgSvg = `
    <svg width="${canvasW}" height="${canvasH}" viewBox="0 0 ${canvasW} ${canvasH}" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <linearGradient id="bgGrad" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stop-color="#f8fafc" />
          <stop offset="50%" stop-color="#f1f5f9" />
          <stop offset="100%" stop-color="#e2e8f0" />
        </linearGradient>
        <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
          <feDropShadow dx="0" dy="40" stdDeviation="40" flood-color="#0f172a" flood-opacity="0.08" />
        </filter>
        <filter id="meshBlur" x="-50%" y="-50%" width="200%" height="200%">
          <feGaussianBlur stdDeviation="150" />
        </filter>
      </defs>
      
      <!-- Base Background -->
      <rect width="${canvasW}" height="${canvasH}" fill="url(#bgGrad)" />
      
      <!-- Modern Mesh Blobs (Very subtle) -->
      <circle cx="400" cy="400" r="600" fill="#bae6fd" opacity="0.4" filter="url(#meshBlur)" />
      <circle cx="2160" cy="1200" r="800" fill="#d8b4fe" opacity="0.2" filter="url(#meshBlur)" />
      
      <!-- Floating White Card (Perfectly Centered) -->
      <rect x="${(canvasW - cardSize)/2}" y="${(canvasH - cardSize)/2}" width="${cardSize}" height="${cardSize}" rx="60" fill="#ffffff" filter="url(#shadow)" />
      
      <!-- Subtle Border around Card -->
      <rect x="${(canvasW - cardSize)/2}" y="${(canvasH - cardSize)/2}" width="${cardSize}" height="${cardSize}" rx="60" fill="none" stroke="#e2e8f0" stroke-width="2" />
    </svg>`;

    // Step 1: Format the source image (the actual Excel screenshot) to fit inside the card beautifully
    // The original is ~1000x1000. We resize it to 1040x1040 (fit inside 1200x1200 card with 80px padding).
    // Using lanczos3 ensures crispness without insane pixelation since we aren't blowing it up to 3840.
    const processedSource = await sharp(sourceImage)
        .resize(1040, 1040, { fit: 'contain', background: { r:255,g:255,b:255,alpha:1 } })
        .toBuffer();

    // Step 2: Composite the source image precisely onto the center of the modern background
    await sharp(Buffer.from(bgSvg))
        .composite([
            {
                input: processedSource,
                gravity: 'center'
            }
        ])
        .webp({ quality: 90, effort: 6 })
        .toFile(destPath);
      
    count++;
    if (count % 20 === 0) console.log(`Processed ${count}/${slugs.length}...`);
  }
  console.log(`✅ ${slugs.length} MODERN LIGHT SHOWCASE images generated!`);
}

processAll().catch(console.error);
