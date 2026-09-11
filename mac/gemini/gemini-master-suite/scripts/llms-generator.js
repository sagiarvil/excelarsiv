const fs = require('fs');
const path = require('path');

const targetDir = process.argv[2];
if (!targetDir) {
    console.log("Kullanım: node llms-generator.js <hedef_klasor>");
    process.exit(1);
}

const llmsContent = `# Domain
Projeniz Hakkında Genel Bilgi

> Summary
Bu dosya yapay zeka ajanları ve LLM'ler için tasarlanmıştır.

## Mimariler ve Alt Linkler
- [Core Architecture](/llms/core.md)
- [Sayfalar](/llms/pages/index.md)
`;

const llmsDir = path.join(targetDir, 'llms');
const llmsPagesDir = path.join(llmsDir, 'pages');

try {
    if (!fs.existsSync(llmsDir)) fs.mkdirSync(llmsDir, { recursive: true });
    if (!fs.existsSync(llmsPagesDir)) fs.mkdirSync(llmsPagesDir, { recursive: true });
    
    fs.writeFileSync(path.join(targetDir, 'llms.txt'), llmsContent);
    fs.writeFileSync(path.join(llmsDir, 'core.md'), '# Core System\nLLM Yönergeleri buraya...');
    fs.writeFileSync(path.join(llmsPagesDir, 'index.md'), '# Pages\nSayfa alt kırılımları...');
    
    console.log(`\n=== SÜPER SEO MOTORU (LLMS HUB OLUŞTURUCU V3.0) ===`);
    console.log(`[+] BAŞARILI: ${path.resolve(targetDir)} dizinine \`llms.txt\` hub ve alt dizinleri oluşturuldu.`);
} catch (err) {
    console.error(`[X] HATA: ${err.message}`);
}
