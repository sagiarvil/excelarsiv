# FİYAT SAVUNMASI — NAKLİYE MALİYETİ HESAPLAYICI

## D01 — İkame maliyeti analizi (satış fiyatı: 990 ₺)

Satış fiyatı çapası, manda v4 D01 kuralına göre danışman ikamesi (Y1) + kurulum ikamesi (Y2) toplamının satış fiyatının en az 3 katı olmasını gerektirir.

| Kalem | Değer | Gerekçe |
|---|---|---|
| Satış fiyatı | 990 ₺ | Ürünün perakende fiyatı |
| Y1 — Danışman ikamesi | 6.000 ₺ | Sefer bazında yakıt/bakım/sürücü/amortisman maliyet hesaplama, km ve ton-km başına birim maliyet ve kârlılık analizi, hasılat yoğunlaşması ve senaryo bandı kurulumu için danışmanlık (2 gün × 3.000 ₺) |
| Y2 — Kurulum ikamesi | 1.500 ₺ | Araç parametrelerinin (tüketim, bakım oranı, sürücü ücreti, amortisman) toplanması, sefer veri düzeninin kurulması, eşik ve senaryo parametrelerinin belirlenmesi (0,5 gün) |
| Toplam ikame | 7.500 ₺ | Y1 + Y2 |
| Eşik (3 × satış) | 2.970 ₺ | 990 ₺ × 3 |
| Sonuç | 7.500 ₺ ≥ 2.970 ₺ → GEÇTİ | D01 sağlanır |

Y3 (önlenen hata) hesaba dahil edilmemiştir. Gerçek hayatta ayrıca şunlar eklenir: kârsız seferlerin tespit edilememesi, elle toplanan sefer verilerindeki mutabakat hataları, filo maliyetinin yönetici raporu yerine sezgiyle yönetilmesi.

## D02 — Serbest alternatif analizi

Alıcının bu ürünü satın almadan aynı işi bedavaya nasıl yapabileceği ve ürünün hangi noktalarda fark yarattığı:

| # | Serbest alternatif | Ürünün sunduğu ayrım | Dosya karşılığı |
|---|---|---|---|
| 1 | Sefer maliyetini yakıt fişleri, bakım faturaları ve sürücü ücretlerini toplayarak hesap makinesiyle kalem kalem çıkarmak | Sefer mesafesi ve araç özelliklerinden yakıt, bakım, sürücü ve amortisman maliyetini tek tabloda canlı formülle üretir; alternatifte her kalem elle çarpılır | `NakliyeToplamMaliyet` |
| 2 | Araç özelliklerini (tüketim, bakım, amortisman) ayrı Excel sayfalarında tutup sefer bazında manuel birleştirmek | Toplam maliyeti mesafe ve ton-km bazında birim maliyete dönüştürüp hasılatla karşılaştırarak net kâr ve kâr marjını otomatik hesaplar; alternatifte mutabakat elle yürütülür | `NakliyeBirimMaliyet` |
| 3 | Her sefer için hasılat-maliyet farkını aylarca elle toplayıp yönetici raporu hazırlamak | Kâr marjını eşiklerle karşılaştırarak KARLI/İNCELE/ZARARLI kararını gerekçesi ve önerilen aksiyonlarla üretir; alternatifte karar belgesi elle yazılır | `NakliyeKarar` |
| 4 | Birim maliyeti (TL/km, ton-km) hiç hesaplamadan taşıma fiyatını sezgisel belirlemek | Yakıt fiyatı senaryo bant aralığı ve maliyet bileşeni duyarlılık (tornado) analizi ile net kârın belirsizlik aralığını ve en etkili değişkeni gösterir; alternatifte duyarlılık analizi yoktur | `modulO3Senaryo` |
| 5 | Kârsız seferleri belirleyip iyileştirme aksiyonunu el yazısı notlarla yönetmek | Hasılat yoğunlaşmasını (HHI), maliyet sapmasını ve veri kalite skorunu ölçerek karar kapısına güven katmanı ekler; alternatifte bu analizler yoktur | `modulO8KaliteSkor` |

## Serbest alternatifle en çok örtüşen özellik

Alternatif 2'dir: Alıcı "araç maliyetlerini zaten ayrı sayfalarda tutuyorum" der. Bu üründe fark, araç parametrelerinin üzerine **sefer bazında canlı maliyet üretimi, birim maliyet (km ve ton-km), net kâr marjı analizi ve karar kapısını** kurmasıdır — yani "araç maliyetini görmek" değil, "bu sefer bana kâr ettiriyor mu, maliyet nerede yoğunlaşıyor ve ne yapmalıyım" sorularına gerekçeli cevap üretmektir.
