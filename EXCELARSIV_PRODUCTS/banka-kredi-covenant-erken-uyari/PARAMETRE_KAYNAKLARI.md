# Parametre Kaynakları — Banka Kredi + Covenant Erken Uyarı

| Parametre | Kaynak | Not |
|---|---|---|
| BKE-001 DSCR min | Tipik kredi sözleşmesi covenant maddesi | Ürün varsayılanı 1,20; kullanıcı değiştirir |
| BKE-002 Borç/FAVÖK max | Tipik kaldıraç covenantı | Ürün varsayılanı 3,50 |
| BKE-003 Cari min | Likidite covenantı | Ürün varsayılanı 1,00 |
| BKE-004 Erken uyarı bandı | İç politika | %10 sapma bandı |
| BKE-005 Faiz yükü max | İç politika | Faiz×bakiye/FAVÖK ≤ 0,40 |
| Yorum A/B | Covenant eşik yorumu | B: DSCR min +0,05 |

Dış API / canlı kur bağlantısı yoktur (Ç15).
