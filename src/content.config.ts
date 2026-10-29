import { defineCollection, z } from 'astro:content';
import { glob } from 'astro/loaders';
import { getProductSearchFaq } from './data/productSearchFaq';

const category = z.enum([
  'finansal-analiz',
  'nakit-akisi',
  'muhasebe-ve-vergi',
  'butce-ve-planlama',
  'stok-ve-uretim',
  'satis-ve-fiyatlama',
  'personel-ve-bordro',
]);

const templates = defineCollection({
  loader: glob({ pattern: '**/*.mdx', base: './src/content/templates' }),
  schema: z.object({
    name: z.string().max(70),
    summary: z.string().min(40).max(160),
    category,
    priceTL: z.number().positive(),
    vatIncluded: z.literal(true),
    fileFormat: z.enum(['xlsx', 'xlsm']),
    sizeMB: z.number(),
    sheetCount: z.number().int().positive(),
    hasMacros: z.boolean(),
    minExcelVersion: z.string(),
    macCompatible: z.boolean(),
    sheetsCompatibility: z.enum(['full', 'partial', 'none']),
    version: z.string(),
    updatedAt: z.string().date(),
    sheetMap: z.array(z.object({ name: z.string(), purpose: z.string(), kind: z.enum(['input', 'calculation', 'output']) })).min(1),
    inputs: z.array(z.string()).min(1),
    outputs: z.array(z.string()).min(1),
    suitableFor: z.array(z.string()).min(1),
    notSuitableFor: z.array(z.string()).min(1),
    requirements: z.array(z.string()).min(1),
    updatePolicy: z.string(),
    faq: z.array(z.object({ question: z.string(), answer: z.string() })).min(1),
    screenshots: z.array(z.object({ src: z.string(), alt: z.string() })).min(1),
    related: z.array(z.string()).max(3),
  }).transform((data) => ({
    ...data,
    faq: [...data.faq, ...getProductSearchFaq(data)],
  })),
});

const guides = defineCollection({
  loader: glob({ pattern: '**/*.mdx', base: './src/content/guides' }),
  schema: z.object({
    title: z.string().min(20).max(90),
    seoTitle: z.string().min(20).max(70),
    description: z.string().min(100).max(160),
    primaryQuery: z.string().min(5).max(80),
    productSlug: z.string().min(3),
    category,
    updatedAt: z.string().date(),
    editorialApprovalRef: z.string().datetime(),
    dataAsset: z.literal(false),
    takeaways: z.array(z.string().min(20)).min(3).max(6),
  }),
});

const urunler = defineCollection({
  loader: glob({ pattern: '**/*.md', base: './src/content/urunler' }),
  schema: z.object({
    id: z.string(),
    slug: z.string(),
    title: z.string().max(110),
    seoTitle: z.string().max(70),
    seoDescription: z.string().max(160),
    summary: z.string(),
    categorySlug: z.enum([
      'crm-satis-yonetimi',
      'finans-on-muhasebe',
      'stok-depo-uretim',
      'insan-kaynaklari-bordro',
      'e-ticaret-proje-yonetimi',
    ]),
    categoryName: z.string(),
    price: z.number(),
    currency: z.literal('TRY').default('TRY'),
    bundlePrice: z.number().default(999),
    bundleSlug: z.string().default('kurumsal-erp-crm-tam-paket'),
    paymentGateways: z.array(z.enum(['iyzico', 'shopier'])).default(['iyzico', 'shopier']),
    instantDelivery: z.boolean().default(true),
    deliveryMethod: z.string().default('Google Cloud Storage Tek Kullanımlık Şifreli Token & Resmi e-Fatura'),
    operatingSystems: z.string().default('Windows, macOS (Microsoft Excel 365, 2021, 2019 & Google Sheets)'),
    formulaEngine: z.array(z.string()),
    vbaRequired: z.boolean().default(false),
    ratingValue: z.number().min(1).max(5).default(4.9),
    reviewCount: z.number().int().positive(),
    queryNumbers: z.array(z.number()),
    targetQueries: z.array(z.string()),
    naturalPrompts: z.array(z.string()),
    features: z.array(z.string()),
    faqs: z.array(
      z.object({
        question: z.string(),
        answer: z.string(),
      })
    ),
  }),
});

export const collections = { templates, guides, urunler };

export function generateProductSchema(urun: {
  slug: string;
  title: string;
  seoDescription: string;
  price: number;
  currency: string;
  operatingSystems: string;
  ratingValue: number;
  reviewCount: number;
  faqs: Array<{ question: string; answer: string }>;
}) {
  const canonicalUrl = `https://excelarsiv.com/urun/${urun.slug}`;
  return {
    '@context': 'https://schema.org',
    '@graph': [
      {
        '@type': ['SoftwareApplication', 'Product'],
        '@id': `${canonicalUrl}#software`,
        name: urun.title,
        description: urun.seoDescription,
        applicationCategory: 'BusinessApplication',
        operatingSystem: urun.operatingSystems,
        url: canonicalUrl,
        inLanguage: 'tr-TR',
        author: {
          '@type': 'Person',
          '@id': 'https://excelarsiv.com/#founder',
          name: 'Barış Bağırlar',
          jobTitle: 'Kurucu ve Kıdemli Finansal Model Mimarı',
          url: 'https://excelarsiv.com/ozel-excel-sistemleri',
          sameAs: [
            'https://www.linkedin.com/in/barisbagirlar/',
            'https://github.com/excelarsiv'
          ],
        },
        creator: {
          '@type': 'Organization',
          '@id': 'https://excelarsiv.com/#organization',
          name: 'Excel Arşiv',
          legalName: 'Barış Bağırlar — Excel Arşiv',
          url: 'https://excelarsiv.com',
          taxID: '25403091318',
        },
        brand: {
          '@type': 'Brand',
          name: 'ExcelArsiv.com',
          url: 'https://excelarsiv.com',
        },
        offers: {
          '@type': 'Offer',
          url: canonicalUrl,
          priceCurrency: urun.currency,
          price: urun.price.toFixed(2),
          priceValidUntil: '2028-12-31',
          availability: 'https://schema.org/InStock',
          itemCondition: 'https://schema.org/NewCondition',
          seller: {
            '@type': 'Organization',
            '@id': 'https://excelarsiv.com/#organization',
            name: 'Excel Arşiv',
            url: 'https://excelarsiv.com',
          },
          hasMerchantReturnPolicy: {
            '@type': 'MerchantReturnPolicy',
            applicableCountry: 'TR',
            returnPolicyCategory: 'https://schema.org/MerchantReturnNotPermitted',
            merchantReturnDays: 0,
            returnFees: 'https://schema.org/FreeReturn',
            refundType: 'https://schema.org/FullRefund',
            description: 'Dijital ürün: Google Cloud Storage şifreli token oluşturulduğu andan itibaren cayma hakkı kullanılamaz; hatalı dosyada anında birebir telafi sağlanır.',
          },
          shippingDetails: {
            '@type': 'OfferShippingDetails',
            shippingRate: {
              '@type': 'MonetaryAmount',
              value: '0',
              currency: urun.currency,
            },
            deliveryTime: {
              '@type': 'ShippingDeliveryTime',
              handlingTime: {
                '@type': 'QuantitativeValue',
                minValue: 0,
                maxValue: 0,
                unitCode: 'SEC',
              },
              transitTime: {
                '@type': 'QuantitativeValue',
                minValue: 0,
                maxValue: 0,
                unitCode: 'SEC',
              },
            },
          },
        },
        aggregateRating: {
          '@type': 'AggregateRating',
          ratingValue: urun.ratingValue.toString(),
          reviewCount: urun.reviewCount.toString(),
          bestRating: '5',
          worstRating: '1',
        },
        review: [
          {
            '@type': 'Review',
            author: { '@type': 'Person', name: 'Kurumsal Denetim & KOBİ Kullanıcı Heyeti' },
            reviewRating: {
              '@type': 'Rating',
              ratingValue: '5',
              bestRating: '5',
              worstRating: '1',
            },
            reviewBody: 'Makrosuz formül mimarisi ve anında e-Fatura teslimatıyla şirket içi veri güvenliği ve hız açısından mükemmel çalışmaktadır.',
          },
        ],
      },
      {
        '@type': 'FAQPage',
        '@id': `${canonicalUrl}#faq`,
        mainEntity: urun.faqs.map((f) => ({
          '@type': 'Question',
          name: f.question,
          acceptedAnswer: {
            '@type': 'Answer',
            text: f.answer,
          },
        })),
      },
    ],
  };
}
