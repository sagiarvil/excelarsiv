# SUPER-SEO-MOTOR-V3: CANLI SİTE DENETİM, KÖK-ONARIM VE OTOMASYON ŞARTNAMESİ

**Standart:** Mandate V6.0 / Engine V3.0 Enterprise Architecture ($5M AI Search Standardı)  
**Hedef Kapsam:** Tüm Araçlarla Canlı Tarama, Hata Raporlama, PHP-Developer ile BOM/DOM Onarımı, Güncelleme Paketi & DOM Tam Doğrulaması  
**Dosya Kodlaması:** BOM'suz Saf UTF-8 (UTF-8 Without BOM)  

---

## 🧭 6 ADIMLI MASTER OTONOM YÜRÜTME DÖNGÜSÜ

Herhangi bir SEO analizi tetiklendiğinde AI ajan istisnasız aşağıdaki **6 Adımlı Süreç Döngüsünü** işletir:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│ 1. ADIM: TAŞINABİLİR YOL ÇÖZÜMLEMESİ & ESKİ KALINTILARIN SIFIRLANMASI        │
│    - [SEOTEST], [PROJE_KÖKÜ], [GUNCELLEME], [RAPOR] yolları çözülür.        │
│    - Varsa eski docs, rapor ve guncelleme kalıntıları TAMAMEN TEMİZLENİR.   │
│    - Canlı hedef domain ([HEDEF_URL]) kullanıcıdan alınır.                  │
└──────────────────────────────────────┬──────────────────────────────────────┘
                                       ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 2. ADIM: TÜM RESMİ ARAÇLARLA CANLI TARAMA (FULL TOOLCHAIN EXECUTION)         │
│    - heading_semantics_check.js  (H1-H6 başlık hiyerarşisi & landmark)      │
│    - opengraph_check.js          (OG:title, image, url & Twitter cards)     │
│    - schema_validator.js         (JSON-LD @graph & Wikidata QID)            │
│    - sitemap_robots_check.js     (robots.txt, sitemap.xml, /llms.txt)       │
│    - internal_link_check.js      (İç bağlantı ağı & 404 kırık link avı)     │
│    - web_vitals_hints.js         (14KB AST bütçesi & Core Web Vitals)       │
│    - audit_engine_v3.js          (18 Motorlu Engine V3 genel puanlama)      │
└──────────────────────────────────────┬──────────────────────────────────────┘
                                       ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 3. ADIM: HATALARIN DETERMINİSTİK RAPORLANMASI                                │
│    - [SEOTEST]/rapor/test-sonucu.md (8 temel teknik kontrol maddesi)        │
│    - [SEOTEST]/rapor/RAPOR.md       (18 Motor skoru & P0-P3 hata reçetesi)  │
└──────────────────────────────────────┬──────────────────────────────────────┘
                                       ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 4. ADIM: PHP-DEVELOPER İLE KÖK-ONARIM (BOM & DOM KURALINA KATI UYUM)         │
│    - ag-php-developer devreye alınır, MVC kod tabanı incelenir.             │
│    - KATI BOM KURALI: 0xEF, 0xBB, 0xBF kesinlikle yasak! Saf UTF-8.         │
│    - KATI DOM KURALI: Her sayfada tek H1, H1->H2->H3 hiyerarşisi,           │
│      ilk 100px .hero-answer, data-chunk-id ve @graph şeması.                │
│    - STATİK OTOMASYON: SeoGenerator ile robots.txt ve llms.txt'nin          │
│      statik yazılması ve controller üzerinden 24 saatte bir güncellenmesi.  │
└──────────────────────────────────────┬──────────────────────────────────────┘
                                       ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 5. ADIM: GÜNCELLEME PAKETİ VE ZIP DERLEMESİ                                  │
│    - Onarılan dosyalar [GUNCELLEME]/ altına orijinal göreli yollarıyla konur.│
│    - Sunucuya tek tıkla yükleme için [SEOTEST]/guncelleme.zip derlenir.     │
└──────────────────────────────────────┬──────────────────────────────────────┘
                                       ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 6. ADIM: GÜNCELLEME SONRASI DOM & BOM TAM TESTİ                              │
│    - tools/bom_utf8_scan.js --fix [GUNCELLEME] çalıştırılır.                │
│    - tools/dom_utf8_full_test.js çalıştırılır (10/10 PASS Şartı!).          │
│    - php -l ile tüm dosyaların sözdizimi doğrulanır.                        │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 📦 `[GUNCELLEME]` DİZİNİ VE ZIP ARŞİVİ

Düzeltilen tüm dosyalar hem klasör ağacı hem de zip olarak hazırlanır:

```text
[SEOTEST]/guncelleme/ (ve [SEOTEST]/guncelleme.zip)
├── app/
│   ├── Controllers/
│   ├── Routes/
│   └── Views/
└── public/
    ├── .well-known/
    ├── llms.txt
    └── robots.txt
```

---

## 🧪 KALİTE KAPISI (QUALITY GATE)

> **TESLİMAT DEMİR KANUNU:**
> `dom_utf8_full_test.js` çalıştığında **10/10 PASS (%100 Başarılı)** çıktısı vermeden, BOM taramasında tek bir uyarı dahi kalmadan ve tüm PHP dosyaları `No syntax errors detected` olmadan süreç tamamlanamaz.