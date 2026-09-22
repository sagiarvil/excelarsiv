import fs from 'fs';
import path from 'path';
import sharp from 'sharp';

const sourceDir = 'urun-gorselleri';
const files = fs.readdirSync(sourceDir).filter(f => f.endsWith('.png') || f.endsWith('.jpg'));
let sizes = [];

async function run() {
  for (const f of files) {
    const p = path.join(sourceDir, f);
    const stat = fs.statSync(p);
    const meta = await sharp(p).metadata();
    sizes.push({file: f, width: meta.width, height: meta.height, sizeKB: Math.round(stat.size/1024)});
  }
  sizes.sort((a, b) => a.sizeKB - b.sizeKB);
  console.log('--- 10 WORST IMAGES ---');
  console.log(sizes.slice(0, 10));
  console.log('--- 10 BEST IMAGES ---');
  console.log(sizes.slice(-10));
}
run();
