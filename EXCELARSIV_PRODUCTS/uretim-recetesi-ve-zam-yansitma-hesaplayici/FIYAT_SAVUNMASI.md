# FİYAT SAVUNMASI — ÜRETİM REÇETESİ & ZAM YANSITMA HESAPLAYICI

## D01 — İkame maliyeti analizi (satış fiyatı: 1.490 ₺)

Satış fiyatı çapası, manda v4 D01 kuralına göre danışman ikamesi (Y1) + kurulum ikamesi (Y2) toplamının satış fiyatının en az 3 katı olmasını gerektirir.

| Kalem | Değer | Gerekçe |
|---|---|---|
| Satış fiyatı | 1.490 ₺ | Ürünün perakende fiyatı |
| Y1 — Danışman ikamesi | 6.000 ₺ | Hammadde zamlarının reçete maliyetine etkisini modellemek, ürün bazında eski/yeni maliyet ve hedef marjı koruyan fiyat üretmek için danışmanlık (2 gün × 3.000 ₺) |
| Y2 — Kurulum ikamesi | 1.500 ₺ | Reçete verisinin toplanması, fire/işçilik/GÜG oranlarının belirlenmesi, eşik ve senaryo parametrelerinin kurulması (yarım gün) |
| Toplam ikame | 7.500 ₺ | Y1 + Y2 |
| Eşik (3 × satış) | 4.470 ₺ | 1.490 ₺ × 3 |
| Sonuç | 7.500 ₺ ≥ 4.470 ₺ → GEÇTİ | D01 sağlanır |

Y3 (önlenen hata) hesaba dahil edilmemiştir. Gerçek hesapta ayrıca şunlar eklenir: hatalı fiyatlama nedeniyle kaçan marj, eksik yansıtılan zam nedeniyle oluşan satış geliri kaybı, elle toplanan verilerdeki hesap hataları.

## D02 — Serbest alternatif analizi

Alıcının bu ürünü satın almadan aynı işi bedavaya nasıl yapabileceği ve ürünün hangi noktalarda fark yarattığı:

| # | Serbest alternatif | Ürünün sunduğu ayrım | Dosya karşılığı |
|---|---|---|---|
| 1 | Reçete maliyetini muhasebe/üretim sisteminden veya hesap makinesiyle kalem kalem çıkarmak | Kalem bazında miktar, fire ve zam oranlarından eski/yeni reçete maliyetini tek tabloda canlı formülle üretir; alternatifte her kalem elle hesap makinesiyle çıkarılır | `UretimToplamZamSonrasiMaliyet` |
| 2 | Hammadde zamlarını teklif ve faturalardan elle toplayıp yeni reçete maliyetini ayrı ayrı güncellemek | Zam yansıtma sonrası genel maliyet artışını ve hedef kâr marjını koruyan önerilen satış fiyatını ürün bazında üretir; alternatifte fiyat güncellemesi sezgiseldir | `UretimEnYuksekArtis` |
| 3 | Fire, işçilik ve genel üretim giderlerini ayrı Excel sayfalarında tutup birim maliyeti manuel birleştirmek | Genel maliyet artışını eşiklerle karşılaştırarak ZAM YANSIT/KISMİ YANSIT/FİYATI KORU kararını gerekçesi ve önerilen aksiyonlarla üretir; alternatifte karar belgesi elle yazılır | `UretimKarar` |
| 4 | Zam sonrası fiyatı hedef kâr marjına göre yeniden fiyatlamadan sezgisel olarak güncellemek | Zam senaryosu bant aralığı ve hammadde duyarlılık (tornado) analizi ile maliyet artışının belirsizlik aralığını ve en etkili değişkeni gösterir; alternatifte duyarlılık analizi yoktur | `modulO3Senaryo` |
| 5 | Her ürün için eski/yeni maliyet farkını aylarca elle toplayıp yönetici raporu ve zam gerekçesi hazırlamak | Maliyet yoğunlaşmasını (HHI), kalem fark sapmasını ve veri kalite skorunu ölçerek karar kapısına güven katmanı ekler; alternatifte bu analizler yoktur | `modulO8KaliteSkor` |

## Serbest alternatifle en çok örtüşen özellik

Alternatif 1'dir: Alıcı "reçete maliyetini zaten sistemimde görüyorum" der. Bu üründe fark, sistemin ham verisinin üzerine **fire ve zam duyarlılığını, ürün bazında fark analizini ve karar kapısını** kurmasıdır — yani "maliyeti görmek" değil, "zammı ne kadar yansıtmalıyım" sorusuna gerekçeli cevap üretmektir.
