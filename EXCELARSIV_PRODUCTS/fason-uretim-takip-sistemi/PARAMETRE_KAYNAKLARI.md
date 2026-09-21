# PARAMETRE KAYNAKLARI — Fason Üretim Takip Sistemi

| Anahtar | Kaynak | Yürürlük | Açıklama |
|---|---|---|---|
| raporTarihi | Kullanıcı / dönem kapanışı | 01.01.2026 | TODAY yok; sabit rapor tarihi |
| zincirTolerans | İş kuralı — verilen/(dönen+fire) | 01.01.2026 | ZİNCİR KIRIK eşiği (adet) |
| fut_esikGecikmeGun | İç politika | 01.01.2026 | Plan teslime göre gecikme |
| fut_esikGecikenUyari / Kritik | İç politika | 01.01.2026 | UYAR / DURDUR adet eşikleri |
| fut_kovaEsik1..3 | Yaşlandırma | 01.01.2026 | 0-7 / 8-14 / 15-30 / 30+ |
| fut_hedefFire | Fire modeli | 01.01.2026 | Hedef fire oranı |
| fut_esikFireUyari / Kritik | FUT-005 | 01.01.2026 | Fire karar eşikleri |
| fut_senaryoTemkinli/Baz/Iyimser | Senaryo motoru | 01.01.2026 | Maliyet senaryo çarpanları |
| fut_tornadoOran | Duyarlılık | 01.01.2026 | Tornado etki oranı |
| fut_yuzdelikOran | İstatistik | 01.01.2026 | PERCENTILE (P90) |
| fut_kuralFUT001..005 | SPEC mevzuat.kurallar | 01.01.2026 | Durum / geçiş / zincir / kapalı / fire |

Belirsizlik notu: fire ölçüm yorumu işletmeden işletmeye değişir; eşikler AYARLAR’dan güncellenir.
