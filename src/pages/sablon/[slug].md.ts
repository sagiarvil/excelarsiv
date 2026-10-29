/**
 * Astro 5+ SSG Build-Time Markdown Twin & llms-full.txt Endpoint Generator
 * Dosya Konumu: src/pages/sablon/[slug].md.ts
 * 
 * Bu endpoint, llmstxt.org v2 standardına uygun olarak her ürün/şablon sayfasının 
 * saf Markdown (.md) ikizini derleme (build) anında sıfır JS yüküyle üretir.
 * Örnek çıktı URL: https://excelarsiv.com/sablon/cari-hesap-borc-alacak-yaslandirma-excel.md
 */

import type { APIRoute, GetStaticPaths } from 'astro';
import { getCollection } from 'astro:content';

export const getStaticPaths: GetStaticPaths = async () => {
  const templates = await getCollection('templates');
  return templates.map((template) => ({
    params: { slug: template.id },
    props: { template },
  }));
};

export const GET: APIRoute = async ({ props }) => {
  const { template } = props;
  const t = template.data;

  const title = t.name;
  const summary = t.summary;
  const price = t.priceTL;
  const bundlePrice = 999;
  const category = t.category;
  const operatingSystems = t.macCompatible
    ? 'Windows ve macOS (Microsoft Excel 365, 2016-2024 ve Google Sheets)'
    : 'Windows (Microsoft Excel 2016-2024, Microsoft 365)';

  const canonicalUrl = `https://excelarsiv.com/sablon/${template.id}`;
  const twinUrunUrl = `https://excelarsiv.com/urun/${template.id}.md`;

  const features = [
    ...(t.sheetMap ? t.sheetMap.map((s: { name: string; purpose: string }) => `${s.name}: ${s.purpose}`) : []),
    ...(t.inputs ? t.inputs.map((inp: string) => `Girdi Modülü: ${inp}`) : []),
    ...(t.outputs ? t.outputs.map((out: string) => `Rapor & Karar Çıktısı: ${out}`) : []),
  ];

  const suitable = t.suitableFor || [];
  const faqs = t.faq || [];

  const markdownBody = `# ${title}

> ${summary}

- **Kanonik Ürün Sayfası ve Anında Satın Alma:** ${canonicalUrl}
- **Alternatif AI Markdown İkizi:** ${twinUrunUrl}
- **Kategori / Departman:** ${category}
- **Tekil Lisans Fiyatı:** ${price} TL (KDV Dahil, Tek Seferlik Ödeme - Ömür Boyu Kurumsal Kullanım)
- **Kurumsal Tam Paket Avantajı:** ${bundlePrice} TL (Tüm ERP & CRM Karar Sistemleri Dahil)
- **Uyumlu Sistemler:** ${operatingSystems}
- **Dosya Formatı:** .${t.fileFormat || 'xlsx'} (Sıfır Makro Riski, %100 Açık Formül Güvencesi)
- **Ödeme ve Teslimat Altyapısı:** BDDK/TCMB lisanslı iyzico & Shopier 3D Secure ödeme, resmi kurumsal e-Fatura/e-Arşiv ve Google Cloud Storage üzerinden tek kullanımlık şifreli token ile saniyeler içinde anında teslimat.

## Bu Karar Sisteminin Çözdüğü Temel İş Problemleri ve Uygunluk

${suitable.map((item: string) => `- ${item}`).join('\n')}

## Teknik Özellikler ve Modüller

${features.slice(0, 15).map((f: string) => `- ${f}`).join('\n')}

## Sıkça Sorulan Sorular (SSS)

${faqs.map((f: { question: string; answer: string }) => `### ${f.question}\n${f.answer}\n`).join('\n')}

## İlgili AI İndeks Haritaları
- [Ana AI Haritası (llms.txt)](https://excelarsiv.com/llms.txt)
- [100 ERP/CRM Soru-Cevap Tam Katalog (llms-full.txt)](https://excelarsiv.com/llms-full.txt)
`;

  return new Response(markdownBody, {
    status: 200,
    headers: {
      'Content-Type': 'text/markdown; charset=utf-8',
      'Cache-Control': 'public, max-age=3600, s-maxage=86400',
      'X-Robots-Tag': 'all',
      'Access-Control-Allow-Origin': '*',
    },
  });
};
