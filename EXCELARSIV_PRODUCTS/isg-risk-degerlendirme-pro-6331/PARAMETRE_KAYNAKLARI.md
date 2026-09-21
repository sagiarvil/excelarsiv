# PARAMETRE KAYNAKLARI — İSG Risk Değerlendirme Pro (6331)

| Anahtar | Değer | Kaynak | Yürürlük |
|---|---|---|---|
| raporTarihi | 10.08.2026 | Kullanıcı / dönem kapanışı | 01.01.2026 |
| zincirTolerans | 15 skor | İş kuralı — İlkSkor↔HedefSkor sapma | 01.01.2026 |
| isg_matrisTipi | L | Kullanıcı seçimi (L veya FK) | 01.01.2026 |
| isg_esikGecikmeGun | 30 gün | İç politika | 01.01.2026 |
| isg_esikGeciken | 5 adet | İç politika — karar eşiği | 01.01.2026 |
| isg_esikKritik | 3 adet | İç politika — karar eşiği | 01.01.2026 |
| isg_esikKritikSeviye | 20 | L-matris kritik eşik (örnek) | 01.01.2026 |
| isg_esikYuksek | 12 | L-matris yüksek eşik (örnek) | 01.01.2026 |
| isg_esikOrta | 6 | L-matris orta eşik (örnek) | 01.01.2026 |
| isg_gecikmeOranGun | %0,05 | Finans varsayımı (örnek) | 01.01.2026 |
| isg_senaryoTemkinli | %115 | Senaryo motoru | 01.01.2026 |
| isg_senaryoBaz | %100 | Senaryo motoru | 01.01.2026 |
| isg_senaryoIyimser | %85 | Senaryo motoru | 01.01.2026 |
| isg_yenilemeTarihi | 10.08.2027 | Yıllık yenileme takvimi | 01.01.2026 |
| isg_olcekHedef | 5.000 | Manda A4 Ö1 | 01.01.2026 |

Durum haritası SPEC `akis.durum_haritasi` ile AYARLAR `tblDurumHaritasi` tablosundan gelir. Skor eşikleri örnek L-matris kalibrasyonudur; Fine-Kinney seçildiğinde kullanıcı eşikleri yükseltmelidir. Dosya oranların güncelliğini taahhüt etmez. Mevzuat belirsizliği: `matris_secimi_L_vs_FK` (KILAVUZ).
