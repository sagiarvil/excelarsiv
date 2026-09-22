import fs from 'fs';
import path from 'path';
import sharp from 'sharp';

const dir = 'urun-gorselleri';
const files = fs.readdirSync(dir).filter(f => f.match(/\.(png|jpe?g|webp)$/i));

async function analyze() {
  const stats = [];
  for (const file of files) {
    const filePath = path.join(dir, file);
    const metadata = await sharp(filePath).metadata();
    const stat = fs.statSync(filePath);
    stats.push({
      file,
      width: metadata.width,
      height: metadata.height,
      pixels: metadata.width * metadata.height,
      size: stat.size
    });
  }
  
  stats.sort((a, b) => a.pixels - b.pixels);
  console.log("Bottom 10 by resolution:");
  console.table(stats.slice(0, 10));

  stats.sort((a, b) => a.size - b.size);
  console.log("\nBottom 10 by file size:");
  console.table(stats.slice(0, 10));
}

analyze();
