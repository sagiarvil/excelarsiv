/**
 * MANDATE-SEO-GEO-2026-V6
 * Dinamik @graph JSON-LD Şema Üreteci (Schema.org v2026)
 */

import type { SeoPageRecord } from './registry.types.ts';

export function buildCompleteJsonLdGraph(page: SeoPageRecord, domain: string = 'excelarsiv.com'): string {
  const origin = `https://${domain}`;
  const pageUrl = `${origin}${page.route === '/' ? '' : page.route}`;

  const graph: Record<string, any>[] = [
    // 1. Kurumsal Otorite Düğümü (Organization Node)
    {
      '@type': 'Organization',
      '@id': `${origin}/#organization`,
      name: 'Excel Arşiv',
      alternateName: 'ExcelArşiv Kurumsal Finans ve Karar Destek Sistemleri',
      url: origin,
      logo: {
        '@type': 'ImageObject',
        '@id': `${origin}/#logo`,
        url: `${origin}/og.png`,
        width: 1200,
        height: 1200,
        caption: 'Excel Arşiv Kurumsal Logo'
      },
      description: 'Türkiye\'deki işletmeler, finans yöneticileri ve KOBİ\'ler için kurumsal Excel çalışma sistemleri, nakit akışı modelleri ve terzi usulü karar mimarileri.',
      email: 'bilgi@excelarsiv.com',
      taxID: '2230353841',
      founder: {
        '@type': 'Person',
        '@id': `${origin}/#founder`,
        name: 'Doğan Aydın',
        jobTitle: 'Kurucu, Kıdemli Finansal Model ve Karar Destek Mimarı',
        sameAs: [
          'https://www.wikidata.org/wiki/Q11589432',
          'https://linkedin.com/in/doganaydin',
          'https://x.com/excelarsiv'
        ]
      },
      address: {
        '@type': 'PostalAddress',
        addressCountry: 'TR'
      },
      contactPoint: {
        '@type': 'ContactPoint',
        contactType: 'customer support',
        email: 'bilgi@excelarsiv.com',
        availableLanguage: 'tr-TR',
        areaServed: 'TR'
      },
      sameAs: [
        'https://www.wikidata.org/wiki/Q11589432',
        'https://twitter.com/excelarsiv',
        'https://www.linkedin.com/company/excelarsiv',
        'https://github.com/excelarsiv'
      ]
    },

    // 2. WebSite Düğümü (WebSite Node)
    {
      '@type': 'WebSite',
      '@id': `${origin}/#website`,
      url: origin,
      name: 'Excel Arşiv',
      description: 'İşletmeler ve Finans Yöneticileri İçin Profesyonel Excel Çalışma Sistemleri',
      publisher: { '@id': `${origin}/#organization` },
      inLanguage: page.locale,
      potentialAction: {
        '@type': 'SearchAction',
        target: {
          '@type': 'EntryPoint',
          urlTemplate: `${origin}/sablonlar?q={search_term_string}`
        },
        'query-input': 'required name=search_term_string'
      }
    },

    // 3. WebPage Düğümü (WebPage Node)
    {
      '@type': 'WebPage',
      '@id': `${pageUrl}#webpage`,
      url: pageUrl,
      name: page.title,
      description: page.metaDescription,
      isPartOf: { '@id': `${origin}/#website` },
      about: { '@id': page.primaryEntity.id },
      datePublished: page.publishedAt,
      dateModified: page.modifiedAt,
      breadcrumb: { '@id': `${pageUrl}#breadcrumb` },
      inLanguage: page.locale
    },

    // 4. BreadcrumbList Düğümü (BreadcrumbList Node)
    {
      '@type': 'BreadcrumbList',
      '@id': `${pageUrl}#breadcrumb`,
      itemListElement: page.breadcrumbs.map((b, idx) => ({
        '@type': 'ListItem',
        position: idx + 1,
        name: b.name,
        item: b.item.startsWith('http') ? b.item : `${origin}${b.item}`
      }))
    }
  ];

  // 5. Birincil Varlık Düğümü (Service / Product / SoftwareApplication / Tool)
  if (page.primaryEntity.type === 'Service') {
    graph.push({
      '@type': 'Service',
      '@id': page.primaryEntity.id,
      name: page.primaryEntity.name,
      serviceType: page.primaryIntent,
      provider: { '@id': `${origin}/#organization` },
      areaServed: { '@type': 'Country', name: 'Türkiye' },
      description: page.heroAnswerEngine,
      url: pageUrl
    });
  } else if (page.primaryEntity.type === 'Product') {
    graph.push({
      '@type': 'Product',
      '@id': page.primaryEntity.id,
      name: page.primaryEntity.name,
      description: page.metaDescription,
      brand: { '@id': `${origin}/#organization` },
      offers: {
        '@type': 'Offer',
        priceCurrency: 'TRY',
        availability: 'https://schema.org/InStock',
        url: `${pageUrl}#satin-al`,
        seller: { '@id': `${origin}/#organization` }
      }
    });
  } else if (page.primaryEntity.type === 'SoftwareApplication') {
    graph.push({
      '@type': 'SoftwareApplication',
      '@id': page.primaryEntity.id,
      name: page.primaryEntity.name,
      applicationCategory: 'BusinessApplication',
      operatingSystem: 'Windows, macOS, Microsoft 365, Excel 2016+',
      offers: {
        '@type': 'Offer',
        priceCurrency: 'TRY',
        availability: 'https://schema.org/InStock',
        url: pageUrl
      }
    });
  }

  return JSON.stringify({ '@context': 'https://schema.org', '@graph': graph }, null, 2);
}
