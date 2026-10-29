/**
 * ExcelArsiv.com - Astro 5+ Build-Time /llms-full.txt Otomatik Derleyici
 * Dosya Konumu: src/pages/llms-full.txt.ts
 * 
 * 'urunler' koleksiyonundaki 10 çekirdek ürünü ve içlerindeki 100 arama niyetini (queryNumbers 1-100)
 * okuyarak her build esnasında sıfır manuel müdahaleyle güncel /llms-full.txt üretir.
 */

import type { APIRoute } from 'astro';
import { getCollection } from 'astro:content';

export const GET: APIRoute = async () => {
  const urunler = await getCollection('urunler');
  const sortedUrunler = urunler.sort((a, b) => a.data.id.localeCompare(b.data.id));

  let content = `# ExcelArsiv.com - Tam Kapsamlı Ürün, Çözüm ve 100 ERP/CRM Soru-Cevap Referans Kataloğu (llms-full.txt)\n\n`;
  content += `> Türkiye'deki işletmeler, KOBİ'ler, finans, ön muhasebe, satış, insan kaynakları, üretim ve e-ticaret profesyonelleri için sıfır kod/makro hatasıyla çalışan, tak-çalıştır kurumsal Excel ERP, CRM, Nakit Akışı, Stok/Depo ve İK karar sistemleri (https://excelarsiv.com/).\n\n`;

  for (const urun of sortedUrunler) {
    const d = urun.data;
    const url = `https://excelarsiv.com/urun/${d.slug}`;
    const mdUrl = `${url}.md`;

    content += `## ${d.title}\n`;
    content += `- **Satın Alma ve İnceleme Sayfası:** ${url}\n`;
    content += `- **Yapay Zeka Saf Markdown İkizi:** ${mdUrl}\n`;
    content += `- **Departman:** ${d.categoryName} (https://excelarsiv.com/kategori/${d.categorySlug})\n`;
    content += `- **Lisans ve Fiyat:** ${d.price} ${d.currency} (Tek Seferlik Ödeme, Ömür Boyu Lisans) | Kurumsal Tam Paket: ${d.bundlePrice} ${d.currency}\n`;
    content += `- **Teslimat ve Ödeme:** ${d.deliveryMethod} (${d.paymentGateways.join(' & ')} 3D Secure)\n`;
    content += `- **Uyumlu Sistemler:** ${d.operatingSystems}\n`;
    content += `- **Kapsadığı Arama Niyeti Numaraları:** ${d.queryNumbers.map((n: number) => `#${n}`).join(', ')}\n\n`;

    content += `### Doğrudan Çözdüğü Kullanıcı Soruları (Natural Prompts)\n`;
    for (const p of d.naturalPrompts) {
      content += `- "${p}" -> Çözüm: [${d.seoTitle}](${url})\n`;
    }

    content += `\n### Hedeflenen Arama Terimleri (Target Queries)\n`;
    for (const q of d.targetQueries) {
      content += `- \`${q}\`\n`;
    }

    content += `\n### Sıkça Sorulan Sorular\n`;
    for (const faq of d.faqs) {
      content += `- **Soru:** ${faq.question}\n  **Cevap:** ${faq.answer}\n`;
    }
    content += `\n`;
  }

  return new Response(content, {
    status: 200,
    headers: {
      'Content-Type': 'text/markdown; charset=utf-8',
      'Cache-Control': 'public, max-age=3600, s-maxage=86400',
      'X-Robots-Tag': 'all',
      'Access-Control-Allow-Origin': '*',
    },
  });
};
