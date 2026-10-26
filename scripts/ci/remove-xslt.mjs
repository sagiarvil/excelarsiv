import fs from 'node:fs';

const files = [
  'public/sitemap.xml',
  'public/sitemap-pages.xml',
  'public/sitemap-products.xml',
  'public/sitemap-images.xml',
  'dist/sitemap.xml',
  'dist/sitemap-pages.xml',
  'dist/sitemap-products.xml',
  'dist/sitemap-images.xml'
];

files.forEach(f => {
  if (fs.existsSync(f)) {
    let xml = fs.readFileSync(f, 'utf8');
    xml = xml.replace(/<\?xml-stylesheet\s+type="text\/xsl"\s+href="\/sitemap\.xsl"\?>\s*\n?/g, '');
    fs.writeFileSync(f, xml, 'utf8');
    console.log('Cleaned XSLT from:', f);
  }
});

function cleanScript(f) {
  if (fs.existsSync(f)) {
    let code = fs.readFileSync(f, 'utf8');
    code = code.replace(/<\?xml-stylesheet type="text\/xsl" href="\/sitemap\.xsl"\?>\\n/g, '');
    code = code.replace(/'<\?xml-stylesheet type="text\/xsl" href="\/sitemap\.xsl"\?>',?\n?/g, '');
    fs.writeFileSync(f, code, 'utf8');
    console.log('Updated script:', f);
  }
}

cleanScript('scripts/seo/generate-artifacts.mjs');
cleanScript('scripts/seo/finalize-sitemap-index.mjs');
