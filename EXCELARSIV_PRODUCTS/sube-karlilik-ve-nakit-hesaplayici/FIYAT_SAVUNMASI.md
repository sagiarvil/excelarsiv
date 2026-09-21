# FİYAT SAVUNMASI — ŞUBE KÂRLILIK VE NAKİT HESAPLAYICI

## D01 — İkame maliyeti analizi (satış fiyatı: 2.490 ₺)

Satış fiyatı çapası, manda v4 D01 kuralına göre danışman ikamesi (Y1) + kurulum ikamesi (Y2) toplamının satış fiyatının en az 3 katı olmasını gerektirir.

| Kalem | Değer | Gerekçe |
|---|---|---|
| Satış fiyatı | 2.490 ₺ | Ürünün perakende fiyatı |
| Y1 — Danışman ikamesi | 9.000 ₺ | Şube bazında gelir/gider kayıtlarından kârlılık, kâr marjı ve nakit pozisyonu çıkarımı, marj eşikleri ve karar kapısı kurgusu, senaryo ve duyarlılık analizi kurulumu için danışmanlık (3 gün × 3.000 ₺) |
| Y2 — Kurulum ikamesi | 3.000 ₺ | Şube gelir/gider ve tahsilat/ödeme verisinin toplanması, kalem ve eşik parametrelerinin belirlenmesi, rapor düzeninin kurulması (1 gün) |
| Toplam ikame | 12.000 ₺ | Y1 + Y2 |
| Eşik (3 × satış) | 7.470 ₺ | 2.490 ₺ × 3 |
| Sonuç | 12.000 ₺ ≥ 7.470 ₺ → GEÇTİ | D01 sağlanır |

Y3 (önlenen hata) hesaba dahil edilmemiştir. Gerçek hesapta ayrıca şunlar eklenir: tespit edilemeyen zararlı şube kaynaklı kâr kaybı, elle toplanan gelir/gider verilerindeki mutabakat hataları, tahmin edilmeyen nakit açıkları nedeniyle oluşan finansman gideri.

## D02 — Serbest alternatif analizi

Alıcının bu ürünü satın almadan aynı işi bedavaya nasıl yapabileceği ve ürünün hangi noktalarda fark yarattığı:

| # | Serbest alternatif | Ürünün sunduğu ayrım | Dosya karşılığı |
|---|---|---|---|
| 1 | Şube gelir ve giderlerini aylık hesap makinesiyle kalem kalem toplayıp kârlılığı elle çıkarmak | Şube bazlı gelir ve gider kayıtlarından işaretli tutar, net kâr ve kâr marjını tek tabloda canlı formülle üretir; alternatifte her kalem elle çarpılır | `SubeNetKar` |
| 2 | Tahsilat ve ödemeleri ayrı Excel sayfalarında tutup nakit pozisyonunu manuel birleştirmek | Tahsilat ve ödeme kayıtlarından nakit akışını ve toplam nakit pozisyonunu otomatik hesaplar; alternatifte mutabakat elle yürütülür | `SubeNakitPozisyonu` |
| 3 | Kâr marjını eşiklerle karşılaştıran karar belgesini el yazısı notlarla yönetmek | Kâr marjını eşiklerle karşılaştırarak KARLI/İNCELE/ZARARLI kararını gerekçesi ve önerilen aksiyonlarla üretir; alternatifte karar belgesi elle yazılır | `SubeKarar` |
| 4 | Gider kalemleri ve ciro belirsizliğini hiç ölçmeden şube yönetimini sezgiyle yürütmek | Ciro senaryo bandı ve gider duyarlılık (tornado) analizi ile kâr belirsizliğini ve en etkili değişkeni gösterir; alternatifte duyarlılık analizi yoktur | `modulO3Senaryo` |
| 5 | Şube gelir yoğunlaşmasını ve veri kalitesini ölçmeden dönemlik raporu elle toplamak | Şube gelir yoğunlaşmasını (HHI), gider sapmasını ve veri kalite skorunu ölçerek karar kapısına güven katmanı ekler; alternatifte bu analizler yoktur | `modulO8KaliteSkor` |

## Serbest alternatifle en çok örtüşen özellik

Alternatif 2'dir: Alıcı "tahsilat ve ödemeleri zaten Excel'de tutuyorum" der. Bu üründe fark, kayıt verisinin üzerine **şube bazında canlı kârlılık, marj eşikleriyle karar kapısı, ciro senaryosu ve nakit pozisyonu analizini** kurmasıdır — yani "kayıt tutmak" değil, "hangi şube kârlı, nakit pozisyonum ne ve ne yapmalıyım" sorularına gerekçeli cevap üretmektir.
