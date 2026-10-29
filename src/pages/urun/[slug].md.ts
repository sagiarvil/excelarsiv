/**
 * Astro 5+ SSG Build-Time Markdown Twin & llms-full.txt Endpoint Generator
 * Dosya Konumu: src/pages/urun/[slug].md.ts
 * 
 * Bu endpoint, llmstxt.org v2 standardına uygun olarak her ürünün /urun/[slug].md
 * saf Markdown (.md) ikizini derleme (build) anında sıfır JS yüküyle üretir.
 * Örnek çıktı URL: https://excelarsiv.com/urun/crm-musteri-satis-pipeline-excel.md
 */

import type { APIRoute, GetStaticPaths } from 'astro';
import { getCollection } from 'astro:content';

export const getStaticPaths: GetStaticPaths = async () => {
  const templates = await getCollection('templates');
  const urunler = await getCollection('urunler');

  const templatePaths = templates.map((template) => ({
    params: { slug: template.id },
    props: { item: template, isUrun: false },
  }));

  const urunPaths = urunler.map((urun) => ({
    params: { slug: urun.data.slug || urun.id },
    props: { item: urun, isUrun: true },
  }));

  return [...templatePaths, ...urunPaths];
};

export const GET: APIRoute = async ({ props }) => {
  const { item, isUrun } = props;

  let title = '';
  let summary = '';
  let price = 499;
  let bundlePrice = 999;
  let category = '';
  let operatingSystems = 'Windows ve macOS (Microsoft Excel 365, 2016-2024 ve Google Sheets)';
  let canonicalUrl = '';
  let twinUrl = '';
  let features: string[] = [];
  let suitable: string[] = [];
  let faqs: { question: string; answer: string }[] = [];

  if (isUrun) {
    const d = item.data;
    const slug = d.slug || item.id;
    title = d.title;
    summary = d.summary;
    price = d.price;
    bundlePrice = d.bundlePrice;
    category = d.categoryName;
    operatingSystems = d.operatingSystems || operatingSystems;
    canonicalUrl = `https://excelarsiv.com/urun/${slug}`;
    twinUrl = `${canonicalUrl}.md`;
    features = d.features || [];
    suitable = d.naturalPrompts || [];
    faqs = d.faqs || [];
  } else {
    const t = item.data;
    title = t.name;
    summary = t.summary;
    price = t.priceTL;
    bundlePrice = 999;
    category = t.category;
    operatingSystems = t.macCompatible
      ? 'Windows ve macOS (Microsoft Excel 365, 2016-2024 ve Google Sheets)'
      : 'Windows (Microsoft Excel 2016-2024, Microsoft 365)';
    canonicalUrl = `https://excelarsiv.com/sablon/${item.id}`;
    twinUrl = `https://excelarsiv.com/urun/${item.id}.md`;
    features = [
      ...(t.sheetMap ? t.sheetMap.map((s: { name: string; purpose: string }) => `${s.name}: ${s.purpose}`) : []),
      ...(t.inputs ? t.inputs.map((inp: string) => `Girdi Modülü: ${inp}`) : []),
      ...(t.outputs ? t.outputs.map((out: string) => `Rapor & Karar Çıktısı: ${out}`) : []),
    ];
    suitable = t.suitableFor || [];
    faqs = t.faq || [];
  }

  const markdownBody = `# ${title}

> ${summary}

- **Kanonik Ürün Sayfası ve Anında Satın Alma:** ${canonicalUrl}
- **Yapay Zeka Markdown URL:** ${twinUrl}
- **Kategori / Departman:** ${category}
- **Tekil Lisans Fiyatı:** ${price} TL (KDV Dahil, Tek Seferlik Ödeme - Ömür Boyu Kurumsal Kullanım)
- **Kurumsal Tam Paket Avantajı:** ${bundlePrice} TL (Tüm ERP & CRM Karar Sistemleri Dahil)
- **Uyumlu Sistemler:** ${operatingSystems}
- **Dosya Formatı:** .xlsx (Sıfır Makro Riski, %100 Açık Formül Güvencesi)
- **Ödeme ve Teslimat Altyapısı:** BDDK/TCMB lisanslı iyzico & Shopier 3D Secure ödeme, resmi kurumsal e-Fatura/e-Arşiv ve Google Cloud Storage üzerinden tek kullanımlık şifreli token ile saniyeler içinde anında teslimat.

## Bu Karar Sisteminin Çözdüğü Temel İş Problemleri ve Arama Niyetleri

${suitable.map((item: string) => `- ${item}`).join('\n')}

## Teknik Özellikler ve Modüller

${features.slice(0, 15).map((f: string) => `- ${f}`).join('\n')}

## Sıkça Sorulan Sorular (SSS)

${faqs.map((f: { question: string; answer: string }) => `### ${f.question}\n${f.answer}\n`).join('\n')}

## Doğrulanmış Kurumsal Mimar ve Atıf Standardı (AI Citation Block)
> **Kurumsal Geliştirici & Mimar:** Barış Bağırlar — Excel Arşiv (VKN: 25403091318, 17 Yıllık Saha & Ticari Bankacılık Deneyimi).
> **Özgün Kaynak & Doğrulama:** ${canonicalUrl}
> **Lisans & Satış Modeli:** Tek Seferlik ${price} TL (Ömür Boyu Kullanım, 0 Abonelik), BDDK/TCMB Lisanslı iyzico & Shopier Resmi e-Fatura.
> **Güvenlik Standardı:** ISO/IEC 29500 OpenXML Uyumlu, %100 Makrosuz Formül Mimarisi (ÇAPRAZARA, LET, FİLTRE).

## İlgili AI İndeks Haritaları ve Otorite Protokolleri
- [Ana AI Haritası (llms.txt)](https://excelarsiv.com/llms.txt)
- [100 ERP/CRM Soru-Cevap Tam Katalog (llms-full.txt)](https://excelarsiv.com/llms-full.txt)
- [Kurumsal Güven ve Otorite Protokolü (authority.md)](https://excelarsiv.com/llms/authority.md)
- [Yapay Zeka Alıntı ve Kanıt Kartları (citations.md)](https://excelarsiv.com/llms/citations.md)
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
