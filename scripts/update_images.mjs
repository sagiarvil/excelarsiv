import fs from 'fs';
import path from 'path';

const templatesDir = path.join(process.cwd(), 'src/content/templates');
const sourceImagesDir = path.join(process.cwd(), 'urun-gorselleri');
const targetImagesDir = path.join(process.cwd(), 'public/images/kapak');
const kapakFile = path.join(process.cwd(), 'src/lib/kapak.ts');

const templates = fs.readdirSync(templatesDir).filter(f => f.endsWith('.mdx')).map(f => f.replace('.mdx', ''));
const sourceImages = fs.readdirSync(sourceImagesDir).filter(f => f.endsWith('.png'));

if (sourceImages.length < templates.length) {
    console.error(`Not enough images! Have ${sourceImages.length}, need ${templates.length}`);
    process.exit(1);
}

// Copy images
templates.forEach((slug, index) => {
    const sourceImage = sourceImages[index];
    const sourcePath = path.join(sourceImagesDir, sourceImage);
    const targetPath = path.join(targetImagesDir, `${slug}.png`);
    fs.copyFileSync(sourcePath, targetPath);
    console.log(`Copied ${sourceImage} to ${slug}.png`);
});

// Update kapak.ts
let kapakContent = fs.readFileSync(kapakFile, 'utf-8');

const slugArrayString = templates.map(slug => `  "${slug}",`).join('\n');
const newPremiumSlugs = `const PREMIUM_KAPAK_SLUGS = new Set<string>([\n${slugArrayString}\n]);`;

// Replace the PREMIUM_KAPAK_SLUGS definition
kapakContent = kapakContent.replace(/const PREMIUM_KAPAK_SLUGS = new Set<string>\(\[[\s\S]*?\]\);/, newPremiumSlugs);

// Replace .webp with .png
kapakContent = kapakContent.replace(/.webp/g, '.png');

fs.writeFileSync(kapakFile, kapakContent, 'utf-8');
console.log('Updated src/lib/kapak.ts');
