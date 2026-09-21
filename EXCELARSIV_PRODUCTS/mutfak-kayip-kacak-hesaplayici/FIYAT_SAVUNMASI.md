# FİYAT SAVUNMASI — MUTFAK KAYIP/KAÇAK HESAPLAYICI

## D01 — İkame maliyeti analizi (satış fiyatı: 2.490 ₺)

Satış fiyatı çapası, manda v4 D01 kuralına göre danışman ikamesi (Y1) + kurulum ikamesi (Y2) toplamının satış fiyatının en az 3 katı olmasını gerektirir.

| Kalem | Değer | Gerekçe |
|---|---|---|
| Satış fiyatı | 2.490 ₺ | Ürünün perakende fiyatı |
| Y1 — Danışman ikamesi | 9.000 ₺ | Kayıp/kaçağın stok mutabakatı ve reçete teorik tüketimiyle tespiti, hammadde bazında kayıp tutarı ve oranı çıkarımı, karar kapısı kurgusu için danışmanlık (3 gün × 3.000 ₺) |
| Y2 — Kurulum ikamesi | 3.000 ₺ | Reçete ve stok verisinin toplanması, dönem başı/alım/dönem sonu stok düzeninin kurulması, eşik ve senaryo parametrelerinin belirlenmesi (1 gün) |
| Toplam ikame | 12.000 ₺ | Y1 + Y2 |
| Eşik (3 × satış) | 7.470 ₺ | 2.490 ₺ × 3 |
| Sonuç | 12.000 ₺ ≥ 7.470 ₺ → GEÇTİ | D01 sağlanır |

Y3 (önlenen hata) hesaba dahil edilmemiştir. Gerçek hesapta ayrıca şunlar eklenir: tespit edilemeyen kaçak nedeniyle oluşan hammadde gideri, mutfak kaybının yönetici raporu yerine sezgiyle yönetilmesi, elle toplanan verilerdeki mutabakat hataları.

## D02 — Serbest alternatif analizi

Alıcının bu ürünü satın almadan aynı işi bedavaya nasıl yapabileceği ve ürünün hangi noktalarda fark yarattığı:

| # | Serbest alternatif | Ürünün sunduğu ayrım | Dosya karşılığı |
|---|---|---|---|
| 1 | Satış adedi ile reçete miktarını çarpıp hammadde bazında teorik tüketimi elle çıkarmak | Satış adedi ile reçete miktarını çarpıp hammadde bazında teorik tüketimi tek tabloda canlı formülle üretir; alternatifte her ürün ve hammadde elle çarpılır | `MutfakToplamKayip` |
| 2 | Dönem başı + alım − dönem sonu stok ile fiili tüketimi, kayıp miktarını, oranını ve tutarını manuel hesaplamak | Dönem başı + alım − dönem sonu stok ile fiili tüketimi, kayıp miktarını, oranını ve tutarını otomatik hesaplar; alternatifte mutabakat elle yürütülür | `MutfakToplamKayipOrani` |
| 3 | Toplam kayıp oranını eşiklerle karşılaştıran karar belgesini el yazısı notlarla yönetmek | Toplam kayıp oranını eşiklerle karşılaştırarak KONTROL ALTINDA/İNCELE/KRİTİK KAYIP kararını gerekçesi ve önerilen aksiyonlarla üretir; alternatifte karar belgesi elle yazılır | `MutfakKarar` |
| 4 | Kayıp tutarının belirsizliğini ve en etkili değişkeni hiç ölçmeden yönetmek | Kayıp senaryosu bant aralığı ve hammadde duyarlılık (tornado) analizi ile kayıp tutarının belirsizlik aralığını ve en etkili değişkeni gösterir; alternatifte duyarlılık analizi yoktur | `modulO3Senaryo` |
| 5 | Kayıp yoğunlaşmasını ve veri kalitesini ölçmeden dönemlik raporu elle toplamak | Kayıp yoğunlaşmasını (HHI), kayıp sapmasını ve veri kalite skorunu ölçerek karar kapısına güven katmanı ekler; alternatifte bu analizler yoktur | `modulO8KaliteSkor` |

## Serbest alternatifle en çok örtüşen özellik

Alternatif 2'dir: Alıcı "stok farkını zaten aylık sayımla çıkarıyorum" der. Bu üründe fark, sayım verisinin üzerine **reçete bazlı teorik tüketimi, hammadde bazında kayıp oranı/tutarı analizini ve karar kapısını** kurmasıdır — yani "stok farkını görmek" değil, "kayıp ve kaçağın maliyeti ne, nerede yoğunlaşıyor ve ne yapmalıyım" sorularına gerekçeli cevap üretmektir.
