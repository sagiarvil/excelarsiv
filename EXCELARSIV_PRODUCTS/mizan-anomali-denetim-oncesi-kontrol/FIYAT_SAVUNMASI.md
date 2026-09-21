# FİYAT SAVUNMASI — Mizan Anomali & Denetim Öncesi Kontrol

**Satış fiyatı:** 9.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 29.700 TL

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 20.000 | Denetim öncesi mizan tarama / anomali danışmanlığı |
| Y2 — Kurulum ikamesi | 10.000 | Mizan↔profil eşleştirme, MOTOR, MAD kök neden, 8 grafik |
| **Toplam ikame** | **30.000** | 30.000 ≥ 29.700 ✓ |

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle mizan Excel | Normalize + duplike hesap | `mad_normAnahtar` |
| ERP mizan dökümü | Dört kova bütünlük | `mad_kovaButunluk` |
| Dönem sonu kontrol listesi | TANIMSIZ kök neden | `mad_tanimsizKova` |
| Sezgisel bakış | Eşleşme oranı veriden | `mad_eslesmeOrani` |
| Sınırsız satır vaadi | 50.000 ölçek | `mad_olcekHedef` |

## En zayıf nokta

Örnek 35 satırdır; gerçek mizan binlerce hesap satırı olabilir. Tablo kapasitesi 500; ölçek sözleşmesi SPEC'te 50.000'dir.
