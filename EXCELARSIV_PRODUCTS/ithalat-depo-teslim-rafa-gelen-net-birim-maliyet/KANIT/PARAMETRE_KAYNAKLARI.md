# PARAMETRE KAYNAKLARI — İTHALAT DEPO TESLİM (RAFA GELEN) NET BİRİM MALİYET

Tüm katsayı ve eşikler AYARLAR sayfasındadır. Kaynağı `Varsayım` işaretli değerler kullanıcı
tarafından doğrulanmalıdır.

| Anahtar | Değer | Birim | Kaynak |
|---|---|---|---|
| firmaUnvani | Örnek Şirket A.Ş. | metin | Kullanıcı |
| raporTarihi | =TODAY() | tarih | Kullanıcı |
| VarsayilanKur | 38,50 | ₺/USD | TCMB gösterge kuru; kullanıcı doğrulamalı |
| OncekiKur | 35,00 | ₺/USD | TCMB gösterge kuru; kullanıcı doğrulamalı |
| GumrukOranVarsayilan | %5 | oran | İthalat Rejimi Kararı; GTİP bazlı |
| OtvOranVarsayilan | %10 | oran | 4760 s. ÖTV Kanunu Liste; GTİP bazlı |
| KdvOranVarsayilan | %20 | oran | 3065 s. KDV Kanunu md.21 |
| NavlunSigortaOranVarsayilan | %5 | oran | İthalat Rejimi; götürü sigorta uygulaması |
| BolumPaydaTaban | 0,0000001 | katsayı | Varsayım |
| VergiYukEsik | %25 | oran | Varsayım |
| MaliyetFaktorEsik | 1,80 | katsayı | Varsayım |
| AnomaliZEsik | 2,00 | oran | Varsayım |
| MutabakatEsikOran | %5 | oran | Varsayım |
| VeriKaliteIyiEsik | 90 | puan | Varsayım |
| VeriKaliteOrtaEsik | 70 | puan | Varsayım |
| VeriKaliteTaban | 100 | puan | Varsayım |
| VeriKaliteFailCeza | 20 | puan | Varsayım |
| VeriKaliteUyariCeza | 10 | puan | Varsayım |
| KayitKapasite | 3008 | adet | Tasarım |
| SenaryoKotumserKat | %10 | oran | Varsayım |
| SenaryoIyimserKat | %10 | oran | Varsayım |
| DuyarliDegisim | %5 | oran | Varsayım |
| ParetoEsikOran | %30 | oran | Varsayım |
| SenaryoEtkiEsik | %15 | oran | Varsayım |
| P90Kuantil | 0,90 | oran | Varsayım |
| IskontoOrani | %10 | oran | Varsayım |
| PesinOdemeTutari | 0 | ₺ | Varsayım |
| HhiYuksekEsik | 0,50 | oran | Varsayım |
| NpvEsikOran | 0,50 | oran | Varsayım |

## Mevzuat dayanakları

- **4458 s. Gümrük Kanunu:** gümrük kıymeti (FOB/CIF), beyan ve gümrük vergisi tahakkuku.
- **3065 s. KDV Kanunu md.21:** ithalatta KDV matrahı = CIF + gümrük vergisi + ÖTV + İGV + KKDF.
- **4760 s. ÖTV Kanunu:** ithalatta ÖTV matrahı = CIF + gümrük vergisi (özel tüketim vergisi listesi).
- **İthalat Rejimi Kararı:** gümrük vergisi oranları; sigorta götürü bedeli uygulaması.
