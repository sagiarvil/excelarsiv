from pathlib import Path
import re

path = Path('src/pages/ozel-excel-sistemleri.astro')
source = path.read_text(encoding='utf-8')
text = source


def replace_once(old: str, new: str) -> None:
    global text
    if old not in text:
        raise SystemExit(f'missing expected text: {old[:80]}')
    text = text.replace(old, new, 1)


replace_once(
    'Ofisinizde her ay sessizce para ve ömür yutan 8 kritik kriz',
    'Ofisinizde her ay sessizce para ve ömür çürüten 8 kritik kriz',
)
replace_once('İNTERAKTİF DENETİM', 'İşinize özel Excel Tabloları')
replace_once(
    '"Esnaf bir ailenin çocuğu olarak büyüdüm, 17 yıl bankalarda masanın iki tarafını da yönettim."',
    'Sizi en iyi ben anlarım. Esnaf bir ailenin çocuğu olarak büyüdüm; 17 yıl bankalarda masanın iki tarafını da yönettim.',
)
replace_once(
    '17.500 TL - 29.500 TL (Tek Seferlik Ödeme)',
    '5.000 TL - 29.500 TL Arasında (Tek Seferlik Ödeme)',
)
replace_once(
    'Paket türünüze göre 30 ila 45 gün boyunca birebir canlı revizyon, personel eğitimi ve ince ayar desteği sağlıyoruz.',
    'Çalışmanın kapsamına göre teslim sonrası birebir revizyon, personel eğitimi ve ince ayar desteği sağlıyoruz.',
)

text = text.replace('905000000000', '905393333303')
text = text.replace(
    'class="w-5 h-5 fill-current" viewBox="0 0 24 24"',
    'class="w-2.5 h-2.5 fill-current opacity-70" viewBox="0 0 24 24"',
)
text = text.replace(
    '        <a href="#paketler" class="hover:text-brand-terracotta transition">Paketler & Fiyatlar</a>\n',
    '',
)

text, badge_count = re.subn(
    r'\n\s*<span class="text-xs font-mono text-emerald-700 bg-emerald-50 border border-emerald-200 px-3\.5 py-1\.5 rounded-full font-bold">\s*✓ Sıfır Tercüme Kaybı Garantisi\s*</span>',
    '',
    text,
    count=1,
)
if badge_count != 1:
    raise SystemExit(f'expected guarantee badge once, found {badge_count}')

package_start = '  <!-- PRODUCTIZED PACKAGES / MONETIZATION TIERS -->'
roi_marker = '  <!-- ROI COMPARISON TABLE -->'
if package_start not in text or roi_marker not in text:
    raise SystemExit('package section markers missing')
start = text.index(package_start)
end = text.index(roi_marker)
text = text[:start] + text[end:]

positioning = '''  <!-- FIELD EXPERIENCE / POSITIONING -->
  <section class="py-14 md:py-20 bg-brand-slate text-white border-b border-stone-800">
    <div class="max-w-4xl mx-auto px-4">
      <div class="grid md:grid-cols-[0.9fr_1.1fr] gap-8 items-start">
        <div>
          <span class="text-xs font-mono uppercase tracking-widest text-brand-terracotta font-bold">SİZİ EN İYİ BEN ANLARIM</span>
          <h2 class="text-2xl md:text-4xl font-extrabold mt-3 leading-tight">Sorunu bana uzun uzun tercüme etmek zorunda kalmazsınız.</h2>
        </div>
        <div class="space-y-4 text-sm md:text-base text-stone-300 leading-relaxed">
          <p>Mali müşavirlerin ay sonu yetiştirme baskısını, muhasebe personelinin ekstre ve mizan yükünü, finans yöneticisinin nakit açığını önceden görme ihtiyacını ve KOBİ sahibinin “kâr var ama para nerede?” sorusunu sahada yıllarca gördüm.</p>
          <p><strong class="text-white">Bu nedenle yalnız Excel formülü kurmuyorum.</strong> Önce işin finansal ve operasyonel mantığını okuyup nerede zaman kaybettiğinizi, hangi kontrolün eksik kaldığını ve hangi çıktının karar vermenizi hızlandıracağını belirliyorum.</p>
          <p>Muhasebe, mali müşavirlik, finans ve işletme tarafındaki gerçek sorunları bildiğim için ihtiyacı teknik dile çevirme yükünü size bırakmadan, doğrudan çalışabilecek bir Excel sistemi tasarlayabiliyorum.</p>
        </div>
      </div>
    </div>
  </section>

'''
replace_once(roi_marker, positioning + roi_marker)

replace_once(
    '''        <span class="text-xs font-mono uppercase tracking-widest text-brand-terracotta font-bold">AKLINIZDAKİ SORULAR</span>
        <h2 class="text-2xl md:text-3xl font-extrabold text-brand-slate mt-2">Sıkça Sorulan Sorular</h2>''',
    '''        <span class="text-xs font-mono uppercase tracking-widest text-brand-terracotta font-bold">AKLINIZDAKİ SORULAR</span>
        <h2 class="text-2xl md:text-4xl font-extrabold text-brand-slate mt-2">Sıkça Sorulan Sorular</h2>
        <p class="text-stone-600 text-sm md:text-base mt-4 leading-relaxed">Bu sorular masa başında üretilmiş değil. Mali müşavirlik ofislerinde, şirket finans ekiplerinde ve KOBİ'lerde yıllardır tekrar eden gerçek darboğazlardan geliyor. Saha tecrübem sayesinde sorunun yalnız görünen kısmını değil, arkasındaki muhasebe, nakit, kontrol ve karar ihtiyacını da okuyabiliyorum.</p>''',
)

faq_more = '''
        <div class="p-5 md:p-6 rounded-2xl bg-stone-50 border border-stone-200">
          <h3 class="font-bold text-sm md:text-base text-brand-slate mb-2">5. “Mali müşavirlik ofisinde en çok hangi işleri hızlandırabilirsiniz?”</h3>
          <p class="text-xs md:text-sm text-stone-600 leading-relaxed">Banka ekstrelerinin temizlenmesi ve sınıflandırılması, cari eşleştirme, mizan kontrolleri, KDV süreçleri, mükellef bazlı takip, dönemsel raporlama ve tekrar eden veri hazırlama işleri en sık karşılaştığım alanlardır. Hazır bir kalıp dayatmak yerine önce sizin ofisinizde zamanın gerçekten nerede kaybolduğunu bulurum.</p>
        </div>

        <div class="p-5 md:p-6 rounded-2xl bg-stone-50 border border-stone-200">
          <h3 class="font-bold text-sm md:text-base text-brand-slate mb-2">6. “Finans müdürü veya CFO tarafında hangi probleme çözüm üretirsiniz?”</h3>
          <p class="text-xs md:text-sm text-stone-600 leading-relaxed">Nakit akışı, çek-senet ve kredi vadeleri, tahsilat riski, banka limitleri, DSCR/NİS gibi kredi bakış açıları, kur senaryoları ve yönetici raporları. Bankacılıkta şirketlerin bu veriler üzerinden nasıl değerlendirildiğini yıllarca gördüğüm için yalnız rapor üretmek değil, yönetimin gerçekten hangi sinyali görmesi gerektiğini de tasarıma taşırım.</p>
        </div>

        <div class="p-5 md:p-6 rounded-2xl bg-stone-50 border border-stone-200">
          <h3 class="font-bold text-sm md:text-base text-brand-slate mb-2">7. “Mevcut Excel dosyamız çok dağınık. Baştan mı kurmak gerekir?”</h3>
          <p class="text-xs md:text-sm text-stone-600 leading-relaxed">Her zaman değil. Önce mevcut dosyanın veri girişlerini, formüllerini, tekrar eden adımlarını ve hata noktalarını incelerim. Sağlam kısımlar korunabilir; riskli ve kişiye bağlı alanlar yeniden kurgulanabilir. Amaç yeni bir dosya üretmek değil, mevcut işleyişinizi daha kontrollü ve sürdürülebilir hale getirmektir.</p>
        </div>

        <div class="p-5 md:p-6 rounded-2xl bg-stone-50 border border-stone-200">
          <h3 class="font-bold text-sm md:text-base text-brand-slate mb-2">8. “Hazır Excel şablonundan farkı nedir?”</h3>
          <p class="text-xs md:text-sm text-stone-600 leading-relaxed">Hazır şablon belirli bir probleme genel çözüm verir. Özel çalışma ise sizin veri kaynağınızı, personel alışkanlığınızı, kontrol noktalarınızı ve istediğiniz yönetim çıktısını esas alır. Bu nedenle işinizi bir yazılımcıya tercüme etmek yerine, finans ve muhasebe dilini bilen biriyle doğrudan ihtiyacın kendisi üzerinden ilerlersiniz.</p>
        </div>

        <div class="p-5 md:p-6 rounded-2xl bg-stone-50 border border-stone-200">
          <h3 class="font-bold text-sm md:text-base text-brand-slate mb-2">9. “Şirkette herkes farklı Excel kullanıyor; bunu tek düzene çevirebilir misiniz?”</h3>
          <p class="text-xs md:text-sm text-stone-600 leading-relaxed">Evet. En sık gördüğüm problemlerden biri aynı verinin farklı kişilerde farklı dosyalarda yaşamasıdır. Tek veri girişi, kontrollü hesaplama, ortak raporlama ve yetki mantığı kurularak kişiye bağlı dosyalar yerine ortak çalışma düzeni oluşturulabilir.</p>
        </div>

        <div class="p-5 md:p-6 rounded-2xl bg-stone-50 border border-stone-200">
          <h3 class="font-bold text-sm md:text-base text-brand-slate mb-2">10. “Bizi gerçekten anlayacağınızı nasıl biliyoruz?”</h3>
          <p class="text-xs md:text-sm text-stone-600 leading-relaxed">Çünkü aynı sorunların finansal sonuçlarını yıllarca bankacı, kredi değerlendiricisi ve danışman tarafında gördüm. Muhasebecinin veri yetiştirme baskısını, mali müşavirin mevzuat ve kontrol yükünü, finansçının nakit ve banka baskısını, işletme sahibinin ise karar için net bilgi ihtiyacını biliyorum. Çözümü bu gerçeklerin üzerine kuruyorum.</p>
        </div>
'''

faq_end = '''
      </div>

    </div>
  </section>

  <!-- CTA / DIRECT WHATSAPP BRIEF FUNNEL -->'''
replace_once(faq_end, faq_more + faq_end)

required = [
    'para ve ömür çürüten 8 kritik kriz',
    'İşinize özel Excel Tabloları',
    'Sizi en iyi ben anlarım.',
    '5.000 TL - 29.500 TL Arasında (Tek Seferlik Ödeme)',
    '905393333303',
    'SİZİ EN İYİ BEN ANLARIM',
    '10. “Bizi gerçekten anlayacağınızı nasıl biliyoruz?”',
]
for item in required:
    if item not in text:
        raise SystemExit(f'missing required result: {item}')

for forbidden in [
    'Özel Sistem Paketleri',
    'Sıfır Tercüme Kaybı Garantisi',
    '905000000000',
    '17.500 TL - 29.500 TL (Tek Seferlik Ödeme)',
]:
    if forbidden in text:
        raise SystemExit(f'forbidden text remains: {forbidden}')

path.write_text(text, encoding='utf-8')
print('service page refinement applied successfully')
