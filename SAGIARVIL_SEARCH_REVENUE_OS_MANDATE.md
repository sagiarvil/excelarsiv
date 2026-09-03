# SAGIARVIL SEARCH REVENUE OS — ENTERPRISE / EXCLUSIVE MASTER MANDATE

**Doküman Kodu:** SAGIARVIL-SRO-2026-V1  
**Statü:** SEO / GEO / AEO / LLMS / Sitemap / Crawler / Search Revenue alanında TEK BAĞLAYICI MANDATE  
**Repo:** `sagiarvil/excelarsiv`  
**Domain:** `https://excelarsiv.com`  
**Nihai Kalite Hedefi:** Kontrol edilebilir release kapılarında **100/100**. 99 final kabul değildir.  
**Ticari North Star:** **Organic Tool Gross Profit** ve portföy seviyesinde **Total Organic Contribution Margin**.

> Bu mandate, bu repodaki önceki SEO/GEO/AEO/LLM-discoverability talimat ve şartname belgelerinin yerini alır. Çalışan `robots.txt`, sitemap dosyaları/üreteçleri, `llms.txt`, `llms-full.txt`, `/llms/**` bilgi düğümleri, schema kodları, analytics verileri ve SEO çalışma verileri eski “talimat belgesi” sayılmaz; bunlar bu mandate altında denetlenen runtime/operasyon varlıklarıdır.

---

## 0. OTORİTE VE ÇELİŞKİ HİYERARŞİSİ

SEO/GEO/LLMS/search-growth konusunda öncelik sırası:

1. Kullanıcının bu görevdeki en güncel açık talebi.
2. Bu dosya: `SAGIARVIL_SEARCH_REVENUE_OS_MANDATE.md`.
3. Resmî arama motoru, schema, HTTP, güvenlik ve platform dokümantasyonu.
4. Repo içindeki güncel ve doğrulanmış runtime/source-of-truth dosyaları.
5. Tasarım, ürün, iş mantığı, mevzuat, fiyat ve güvenlik gibi kendi alanındaki kanonik proje belgeleri.
6. Eski SEO/GEO/LLMS notları ve tarihsel planlar yalnız tarihsel bağlamdır; bu mandate ile çelişemez.

Bir AI ajanı, geliştirici veya otomasyon; sayfa, routing, metadata, canonical, schema, sitemap, robots, LLMS, içerik, internal link, redirect veya search ölçümü değiştirmeden önce bu mandate'i okumak zorundadır.

---

## 1. AMAÇ

Amaç “çok trafik” değildir. Amaç Türkiye pazarında **nitelikli organik talebi gelire dönüştüren, makinece okunabilir, kanıtlanabilir, hızlı, güvenli ve sürdürülebilir bir Search Revenue Operating System** kurmaktır.

Başarı zinciri:

`QUERY → INTENT → PAGE/TOOL → TRUST/EVIDENCE → CONVERSION → SALE → GROSS PROFIT`

SEO kararı yalnız ranking veya trafik için alınamaz. Her indexable ticari URL aşağıdaki soruları cevaplamalıdır:

- Kimi getiriyor?
- Hangi primer arama niyetini sahipleniyor?
- Rakip sonuçlardan hangi özgün bilgi/araç/değeri daha iyi sunuyor?
- Hangi varlığa (entity) bağlı?
- Hangi dönüşüm eylemine götürüyor?
- Nasıl ölçülüyor?
- Hangi gelir veya otorite sonucunu hedefliyor?

Kesin ranking, trafik veya gelir garantisi iddia edilemez. `%100` hedef, bizim kontrolümüzdeki mimari ve release kapılarının eksiksiz geçmesidir.

---

## 2. EXCELARŞİV DOMAIN OWNERSHIP

ExcelArşiv'in primer konu sahipliği:

- Excel karar sistemleri
- hesaplama/analiz çalışma kitapları
- finansal ve operasyonel Excel araçları
- nakit akışı Excel sistemleri
- POS/kârlılık Excel sistemleri
- maliyet, stok, bütçe ve benzeri işletme araçları
- problem → hesaplama → demo → Excel ürünü → satın alma akışı

DRFIN'in sahip olduğu kredi danışmanlığı/finansal teşhis niyetleri ExcelArşiv tarafından kopyalanmaz. ExcelArşiv araç/ürün tarafını sahiplenir. Aynı primer ticari intent iki SAGIARVIL domaininde aynı anda sahiplenilemez.

---

## 3. SEARCH SINGLE SOURCE OF TRUTH — REGISTRY

Tüm indexable URL'ler merkezi Search Registry'de kayıtlı olmalıdır. Framework implementasyonu değişebilir; semantic contract değişemez.

Her kayıt en az şunları taşımalıdır:

```yaml
route:
canonical:
status:
indexDirective:
domain:
locale:
pageRole:
primaryIntent:
secondaryIntents:
topicOwner:
primaryEntity:
supportingEntities:
title:
metaDescription:
h1:
schemaTypes:
sitemap:
  include:
  lastModified:
robotsPolicy:
llm:
  tier:
  node:
  parentNode:
evidence:
  sources:
  verifiedAt:
commercial:
  funnelStage:
  conversionAction:
measurement:
  conversionEvents:
```

Registry'de olmayan yeni indexable sayfa final release alamaz.

Metadata, canonical, JSON-LD, sitemap üyeliği, LLMS ilişkisi ve mümkün olan yerde robots policy aynı SSOT'tan türetilmelidir. Aynı bilgi beş ayrı dosyada elle ve bağımsız yönetilmemelidir.

---

## 4. MULTI-TIER LLMS KNOWLEDGE GRAPH — ZORUNLU

Tek `llms.txt` yeterli değildir. Kök dosya bir **manifest/router** olmalıdır; sitenin tüm metnini içine doldurmak yasaktır.

Minimum mimari:

```text
/llms.txt
/llms-full.txt
/llms/core.md
/llms/entities/
/llms/categories/
/llms/tools/
/llms/products/
/llms/topics/
/llms/pages/
```

### 4.1 Roller

- `/llms.txt`: token-ekonomik ana manifest, kanonik kimlik, en önemli alt-graflara yönlendirme.
- `/llms-full.txt`: kontrollü geniş makine özeti/katalog; kök manifestin yerine geçmez.
- `/llms/core.md`: kurum, iş modeli, otorite, kanonik varlık kimliği, doğrulanabilir temel gerçekler.
- `/llms/entities/**`: marka, uzman, metodoloji ve önemli varlık düğümleri.
- `/llms/categories/**`: ExcelArşiv ana problem/ürün kategorileri.
- `/llms/tools/**`: önemli hesaplayıcı/karar araçları.
- `/llms/products/**`: yalnız yüksek değerli/flagship ürün veya ürün kümeleri.
- `/llms/topics/**`: yüksek değerli bilgi niyetleri.
- `/llms/pages/**`: gerektiğinde flagship canonical HTML sayfalarının derin subgraph'ları.

Her LLMS node şu sözleşmeyi taşımalıdır:

```yaml
canonicalWebUrl:
primaryEntity:
primaryIntent:
parentNode:
lastVerified:
evidence:
relatedNodes:
```

Kurallar:

1. Canonical HTML karşılığı olmayan flagship LLMS node oluşturulamaz; istisna yalnız saf entity/metodoloji kayıtlarıdır.
2. Orphan LLMS node yasaktır.
3. Kök `llms.txt` tüm önemli Tier-1/Tier-2 düğümlere erişim yolu sağlamalıdır.
4. LLMS içeriği HTML'deki gerçeklerle çelişemez.
5. Uydurma veri, sertifika, fiyat, başarı oranı, pazar payı, uzmanlık veya mevzuat eklenemez.
6. Her sıradan ürün için otomatik `.md` üretmek zorunlu değildir. Bilgi kazancı ve ticari değer yoksa subgraph üretilmez.
7. `/llms/**` çıktıları doğru MIME, UTF-8 ve cache politikasıyla servis edilmelidir.

---

## 5. ENTITY GRAPH VE STRUCTURED DATA

JSON-LD parça parça elle yazılmamalı; merkezi builder/registry'den üretilmelidir.

Temel graph:

`Organization → WebSite → WebPage → BreadcrumbList → Page-Specific Entity`

ExcelArşiv için sayfaya göre uygun tipler:

- `Organization`
- `WebSite`
- `WebPage`
- `BreadcrumbList`
- `Product`
- `SoftwareApplication` / `WebApplication`
- `Article` yalnız gerçekten makaleyse
- gerekli ve doğrulanabilir diğer Schema.org tipleri

Schema yalnız görünür/gerçek içerikle desteklenen iddiaları taşır. Sahte rating, review, fiyat, stok, sertifika, ödül veya uzmanlık schema'ya yazılamaz.

`@id` değerleri stabil olmalı; aynı entity farklı sayfalarda farklı kimliklerle çoğaltılmamalıdır.

---

## 6. SITEMAP CONTRACT

Sitemap'e yalnız şu URL girebilir:

`HTTP 200 + canonical + indexable + production URL`

Yasak durumlar:

- redirect URL
- 404/410
- robots ile yanlışlıkla bloklanan canonical ticari URL
- `noindex`
- staging/preview/admin/private URL
- duplicate canonical
- sahte `lastmod`

Gerekirse sitemap index; pages, products/tools, categories ve articles gibi semantik gruplara ayrılır. `lastmod` yalnız gerçek anlamlı değişiklikte güncellenir.

---

## 7. ROBOTS VE AI CRAWLER GOVERNANCE

Search/retrieval ile model-training aynı şey değildir ve ayrı policy olarak yönetilir.

Prensip:

- Kamuya açık canonical search yüzeyleri: arama/retrieval botlarına erişilebilir olmalı.
- Training botları: açıkça tanımlanmış domain policy ile allow/disallow edilir.
- `/admin`, özel hesap alanları, private API, staging, preview ve hassas endpointler crawl/index dışı kalır.
- `robots.txt` duplicate-content çözme aracı değildir; canonical/noindex/redirect doğru yerde kullanılmalıdır.
- Robots ve sitemap arasında kritik çelişki build fail sebebidir.

---

## 8. CONTENT / INFORMATION GAIN CONTRACT

Yeni URL açmanın amacı sayfa sayısını artırmak değildir. Her yeni indexable sayfa en az bir gerçek bilgi kazancı üretmelidir:

- özgün hesaplama/formül/araç,
- gerçek ürün/demo/çalışma sistemi,
- birinci el iş bilgisi,
- açık yöntem,
- doğrulanabilir karşılaştırma,
- kullanıcı kararını anlamlı biçimde iyileştiren içerik.

Yasak:

- sırf keyword için düşük değerli seri AI sayfaları,
- başka kaynakları yeniden yazarak ölçek üretmek,
- şehir ismi değiştirerek doorway sayfalar,
- sahte freshness,
- yapay backlink/click/review/account ağları,
- görünürde kullanıcıya faydası olmayan crawler-only metin.

Türkçe morfoloji ve yakın niyetler cluster olarak normalize edilmelidir; eş anlamlı sorgular için gereksiz ayrı URL açılmaz.

---

## 9. INTENT OWNERSHIP VE CANNIBALIZATION

Her indexable sayfanın tek `primaryIntent` sahibi olmalıdır.

Repo içi duplicate primer intent = kritik hata.  
SAGIARVIL portföyünde cross-domain duplicate ticari intent = kritik hata.

Intent çakışmasında karar seçenekleri:

`KEEP / EXPAND / REFRESH / REPOSITION / MERGE / NOINDEX / DELETE`

Yeni URL açmak varsayılan çözüm değildir.

---

## 10. INTERNAL LINK GRAPH

- Indexable ticari/flagship sayfa orphan olamaz.
- Hub → category/tool/product → supporting content ilişkisi açık olmalıdır.
- Anchor metinleri doğal, açıklayıcı ve hedef niyetle uyumlu olmalı; yapay sabit yüzde benzerliği zorlanmaz.
- Broken internal link = release hatası.
- Canonical olmayan/redirect olan URL'lere sistematik internal link verilmez.

---

## 11. PERFORMANCE / RENDER CONTRACT

Hard production hedefleri:

```yaml
LCP_p75: <= 2.5s
INP_p75: <= 200ms
CLS_p75: <= 0.10
```

Exclusive hedef:

```yaml
LCP_p75: < 2.0s
INP_p75: < 150ms
CLS_p75: < 0.05
```

Lab ile gerçek field/p75 verisi birbirine karıştırılamaz. Ölçüm yoksa “PASS” uydurulamaz.

SSR/SSG ilk HTML ile hydrated DOM arasında title, H1, canonical, indexable ana içerik ve JSON-LD açısından kritik parity korunmalıdır.

---

## 12. G0–G16 RELEASE GATES

Aşağıdaki kapılar otomatik veya deterministik doğrulanmalıdır:

- **G0** Registry integrity
- **G1** HTTP status
- **G2** Canonical integrity
- **G3** Index/noindex integrity
- **G4** Robots ↔ sitemap reconciliation
- **G5** Title/H1/meta requirements
- **G6** SSR/render parity
- **G7** Structured-data validity
- **G8** Entity integrity / stable @id
- **G9** Intra-domain intent collision
- **G10** Cross-domain SAGIARVIL intent collision
- **G11** Internal links / orphan detection
- **G12** Multi-tier LLMS integrity
- **G13** Evidence / freshness / unsupported-claim control
- **G14** Performance budget
- **G15** Commercial conversion + measurement instrumentation
- **G16** Live production health check

Kritik gate başarısızsa:

`BUILD/RELEASE FAIL → PROD YOK`

Gate'i geçmek için testi kapatmak, threshold'u keyfî gevşetmek veya guard'ı bypass etmek yasaktır.

---

## 13. DEPLOYMENT CONTRACT

Standart akış:

`baseline → branch/change → registry validation → G0–G16 → build/test → diff review → preview → production deploy → live health check → discovery/IndexNow where applicable → measurement`

Değişiklik geri alınabilir olmalıdır. Yüksek performanslı mevcut URL'ler baseline alınmadan toplu rewrite edilmez.

IndexNow destekleyen motorlarda yalnız başarılı prod deploy sonrası değişen canonical URL'ler gönderilir. Google keşfi için sitemap, internal linking ve ilgili resmi araçlar temel kabul edilir.

---

## 14. MEASUREMENT & REVENUE LOOP

Her ticari landing page mümkün olan ölçüde şu zincire bağlanmalıdır:

`query → impression → click → landing → conversion → qualified lead/order → revenue → gross profit`

Ölçüm yoksa sonuç uydurulamaz.

Fırsat önceliklendirmesi için temel model:

`Expected Search Value = Search Demand × Commercial Intent × Conversion Probability × Ranking Probability × Topical Authority × Information Gain × AI Citation Potential / (Competition × Cost × Risk)`

Düşük değerli sayfa üretmek yerine en yüksek beklenen ekonomik değere sahip fırsat önce uygulanır.

---

## 15. EXCELARŞİV COMMERCIAL ARCHITECTURE

Flagship ticari akış:

`SEARCH QUERY → BUSINESS PROBLEM → METHOD/CALCULATION → DEMO/PREVIEW → EXCEL TOOL/PRODUCT → PURCHASE`

Önemli ürün/tool sayfaları mümkün olduğunca şunları içermelidir:

- problemin açık tanımı,
- kim için/kim için değil,
- hesaplama veya yöntem açıklaması,
- gerçek demo/örnek/screenshot,
- ürünün ne teslim ettiği,
- güven/evidence,
- ilgili araçlar,
- tek ve anlaşılır ana CTA.

Generic blog trafiği North Star değildir.

---

## 16. HARD PROHIBITIONS

Kesinlikle yasak:

- sahte kullanıcı/hesap/tıklama/review üretimi,
- link farm/PBN/manipülatif backlink ağı,
- içerik sayısını şişirmek için düşük değerli AI seri üretimi,
- görünmeyen crawler-only spam,
- fake freshness,
- sahte schema/review/rating/credential,
- robots ile canonical/index stratejisinin birbirine zıt kurulması,
- preview/staging'in indexe açılması,
- ölçülmemiş metriği ölçülmüş gibi raporlamak,
- eski SEO mandate'ini bu dosyanın üzerinde otorite saymak.

---

## 17. 100/100 FINAL ACCEPTANCE CONTRACT

Final release ancak kontrol edilebilir alanlarda aşağıdaki durum sağlanırsa “100/100 kalite” diye işaretlenebilir:

```yaml
brokenCanonical: 0
robotsSitemapConflict: 0
wrongNoindex: 0
orphanIndexable: 0
brokenInternalLinks: 0
schemaCriticalErrors: 0
entityConflicts: 0
undefinedPrimaryIntent: 0
intraDomainIntentCollision: 0
crossDomainIntentCollision: 0
llmsOrphanNodes: 0
llmsBrokenReferences: 0
unsupportedClaims: 0
fakeFreshness: 0
stagingIndexable: 0
previewIndexable: 0
commercialPageWithoutConversionPath: 0
commercialPageWithoutMeasurementContract: 0
```

Field CWV, Search Console, ranking, citation share veya revenue gibi dış ölçüm verileri mevcut değilse bunlar “PASS” uydurularak 100'e tamamlanamaz; **UNVERIFIED** olarak açık tutulur.

---

## 18. AJAN ÇALIŞMA PROTOKOLÜ

Her SEO/GEO/LLMS/routing/content görevi öncesinde ajan:

1. Bu mandate'i tamamen oku.
2. Mevcut URL/runtime durumunu incele.
3. İlgili Search Registry kaydını bul veya tasarla.
4. Mevcut ranking/deploy davranışını bozacak değişiklikte baseline al.
5. Intent owner ve entity owner'ı doğrula.
6. HTML, schema, sitemap, robots ve LLMS etkisini birlikte değerlendir.
7. G0–G16'yı çalıştır.
8. Fail varsa sebebi düzelt; testi kaldırma.
9. Prod sonrası canlı health-check yap.
10. Yapılmamış testi yapılmış gibi raporlama.

**Durum:** Bu dosya ExcelArşiv'in SEO/GEO/AEO/LLMS/Search Revenue konularında tek kanonik mandate'idir.
