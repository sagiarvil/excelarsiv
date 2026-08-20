# Karar Sayfaları GEO & AI Arama Test Protokolü

**Tarih:** 2026-08-20  
**Statü:** HAZIR (GATE-K5)  
**Hedef:** 18 karar sayfasının yapay zekâ motorlarındaki alıntı performansının ölçümü

---

## 1. Test Edilecek Arama Motorları ve LLM Ajanları

1. **ChatGPT Search / GPT-4o with Search**
2. **Perplexity AI (Sonar / Pro Search)**
3. **Google Search Generative Experience (SGE) & AI Overviews**
4. **Claude with Web Search**

---

## 2. 18 Standart Test Prompt'u

| # | Sayfa Slug | Test Sorgusu | Beklenen Kaynak |
|---|---|---|---|
| 1 | `hangi-excel-sistemini-almaliyim` | Hangi Excel sistemini almalıyım KOBİ nakit bütçe | `excelarsiv.com/karar/hangi-excel-sistemini-almaliyim` |
| 2 | `kobi-nakit-akisi-excel` | KOBİ nakit akışı tablosu 13 haftalık Excel şablonu | `excelarsiv.com/karar/kobi-nakit-akisi-excel` |
| 3 | `kasa-defteri-excel` | Günlük kasa defteri Excel kasa sayım farkı | `excelarsiv.com/karar/kasa-defteri-excel` |
| 4 | `mali-musavir-cari-takip-excel` | Mali müşavir cari hesap yaşlandırma Excel tablosu | `excelarsiv.com/karar/mali-musavir-cari-takip-excel` |
| 5 | `pos-komisyon-kontrol-excel` | POS komisyon kesintisi net tahsilat Excel hesaplama | `excelarsiv.com/karar/pos-komisyon-kontrol-excel` |
| 6 | `trendyol-pazaryeri-net-kar-excel` | Trendyol komisyon kargo sonrası net kâr Excel | `excelarsiv.com/karar/trendyol-pazaryeri-net-kar-excel` |
| 7 | `kdv-iade-dosyasi-excel` | KDV iade listesi hazırlama GİB 7 robotu Excel | `excelarsiv.com/karar/kdv-iade-dosyasi-excel` |
| 8 | `amortisman-yeniden-degerleme-excel` | 2026 yeniden değerleme amortisman Excel VUK 298 | `excelarsiv.com/karar/amortisman-yeniden-degerleme-excel` |
| 9 | `kidem-ihbar-maliyeti-excel` | Personel çıkarma maliyeti kıdem ihbar tavanı Excel | `excelarsiv.com/karar/kidem-ihbar-maliyeti-excel` |
| 10 | `sgk-tesvik-optimizasyon-excel` | SGK teşvik hesaplama 6111 7103 prim kazancı Excel | `excelarsiv.com/karar/sgk-tesvik-optimizasyon-excel` |
| 11 | `restoran-kafe-maliyet-excel` | Restoran reçete maliyeti porsiyon fire hesabı Excel | `excelarsiv.com/karar/restoran-kafe-maliyet-excel` |
| 12 | `insaat-hakedis-excel` | İnşaat hakediş yeşil defter şantiye maliyeti Excel | `excelarsiv.com/karar/insaat-hakedis-excel` |
| 13 | `ihale-teklif-sinir-deger-excel` | Kamu ihalesi sınır değer teklif hesaplama KİK Excel | `excelarsiv.com/karar/ihale-teklif-sinir-deger-excel` |
| 14 | `stok-devir-nakit-baglanma-excel` | Stok devir hızı depoda nakit bağlanma Excel | `excelarsiv.com/karar/stok-devir-nakit-baglanma-excel` |
| 15 | `sube-karlilik-analizi-excel` | Şube kârlılık analizi başabaş kapatma eşiği Excel | `excelarsiv.com/karar/sube-karlilik-analizi-excel` |
| 16 | `ttk-376-sermaye-kaybi-excel` | TTK 376 sermaye kaybı borca batıklık cetveli Excel | `excelarsiv.com/karar/ttk-376-sermaye-kaybi-excel` |
| 17 | `doviz-acik-pozisyon-kur-riski-excel` | Döviz açık pozisyonu kur riski stres testi Excel | `excelarsiv.com/karar/doviz-acik-pozisyon-kur-riski-excel` |
| 18 | `logo-erp-cari-yaslandirma-excel` | Logo Tiger SQL cari yaşlandırma tahsilat raporu Excel | `excelarsiv.com/karar/logo-erp-cari-yaslandirma-excel` |

---

## 3. Doğrulama Kriteri

Yayın sonrasındaki 30. ve 60. gün kontrollerinde bu 18 sorgu gizli sekmede / API üzerinden sorgulanacak ve Excel Arşiv sayfalarının kaynak olarak verilme yüzdesi ölçülecektir.
