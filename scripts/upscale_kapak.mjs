import fs from 'fs';
import path from 'path';
import sharp from 'sharp';

const kapakDir = path.join(process.cwd(), 'public/images/kapak');

async function processImages() {
  const files = fs.readdirSync(kapakDir).filter(f => f.endsWith('.png'));
  console.log(`Found ${files.length} images to process to 4K Premium.`);

  for (const file of files) {
    const filePath = path.join(kapakDir, file);
    const tempPath = path.join(kapakDir, `temp_${file}`);
    
    try {
      // 4K resolution at 16:10 aspect ratio is 3840x2400
      await sharp(filePath)
        .resize({
          width: 3840,
          height: 2400,
          fit: 'contain', // Ensure the whole image is visible (tam olarak gorunmeli)
          background: { r: 248, g: 250, b: 252, alpha: 1 }, // #f8fafc to match background
          kernel: sharp.kernel.lanczos3 // High-quality upscaling
        })
        .sharpen({
          sigma: 1.5,
          m1: 1.2,
          m2: 0.8
        }) // Subtle sharpening for crisp text and edges
        .png({ quality: 100, compressionLevel: 6 }) // Max quality
        .toFile(tempPath);

      // Replace original with the new 4K version
      fs.renameSync(tempPath, filePath);
      console.log(`✅ Upscaled and fitted: ${file}`);
    } catch (err) {
      console.error(`❌ Failed to process ${file}:`, err);
    }
  }
  
  console.log('🎉 All images processed to 4K Premium quality!');
}

processImages();
