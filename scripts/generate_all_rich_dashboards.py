import glob
import os
import re
import html
import subprocess

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MDX_DIR = os.path.join(KOK, 'src', 'content', 'templates')
KAPAK_DIR = os.path.join(KOK, 'public', 'images', 'kapak')
SCREENSHOT_DIR = os.path.join(KOK, 'public', 'screenshots')

def parse_mdx(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    slug = os.path.basename(file_path).replace('.mdx', '')
    
    def get_val(key, default=''):
        m = re.search(rf'^{key}:\s*[\'\"]?([^\'\r\n]+)[\'\"]?', content, re.MULTILINE)
        return m.group(1).strip() if m else default

    name = get_val('name', slug.replace('-', ' ').title())
    category = get_val('category', 'finansal-analiz')
    summary = get_val('summary', 'Kurumsal Karar ve Yönetim Sistemi')

    inputs_m = re.search(r'inputs:\s*\n((?:\s*-\s*[^\n]+\n)+)', content)
    inputs = []
    if inputs_m:
        inputs = [re.sub(r'^\s*-\s*[\'\"]?', '', line).rstrip('\'"\r\n ') for line in inputs_m.group(1).strip().split('\n')]

    outputs_m = re.search(r'outputs:\s*\n((?:\s*-\s*[^\n]+\n)+)', content)
    outputs = []
    if outputs_m:
        outputs = [re.sub(r'^\s*-\s*[\'\"]?', '', line).rstrip('\'"\r\n ') for line in outputs_m.group(1).strip().split('\n')]

    sheet_names = re.findall(r'name:\s*[\'\"]([A-Z0-9_]+)[\'\"]', content)
    if not sheet_names:
        sheet_names = ['KAPAK', 'HIZLI_BASLANGIC', 'GIRDI', 'HESAP', 'PANO', 'KARAR_MOTORU', 'RAPOR']

    return {
        'slug': slug,
        'name': name,
        'category': category,
        'summary': summary,
        'inputs': inputs,
        'outputs': outputs,
        'sheets': sheet_names[:8]
    }

CATEGORY_CONFIGS = {
    'nakit-akisi': {
        'accent': '#107c41',
        'subAccent': '#0d9488',
        'ribbon': '#047857',
        'tag': 'NAKİT AKIŞI & LİKİDİTE',
        'kpi1': ('TOPLAM NAKİT HAVUZU', '₺34.850.000', '↑ %14.2', 'konsolide vadesiz + fon', '4 HESAP'),
        'kpi2': ('BU AY ÖDENECEK ÇIKIŞ', '₺2.480.000', '● 8 Kalem', 'nakit karşılığı: bloke', 'VADE: 4 GÜN'),
        'kpi3': ('NET SERBEST LİKİDİTE', '₺18.420.000', '%52.8 ORAN', 'asgari emniyet marjı üstü', 'GÜVENLİ'),
        'kpi4': ('DSCR / NAKİT DÖNGÜSÜ', '2.14x', '✓ İDEAL', 'borç servis karşılama', 'EŞİK: 1.30x'),
        'formula': '=TOPLA(VadesizBakiye) - KümülatifÖdemeler + BeklenenTahsilat',
        'chartTitle': '12 AYLIK NAKİT GİRİŞ / ÇIKIŞ & LİKİDİTE BASAMAKLARI',
        'chartSub': 'Kümülatif Tahsilat vs. Tedarikçi/Kredi Ödeme Projeksiyonu',
        'radarTitle': 'LİKİDİTE RİSKİ & ERKEN UYARI SİNYALLERİ',
        'radar1': ('Nakit Açığı Riski (30 Gün)', '0.00 TL (Sıfır Risk)', 340),
        'radar2': ('Tahsilat Gecikme Toleransı', '48 Gün Güvenli Eşik', 280),
        'radarStatus': '✓ 90 GÜNLÜK PROJEKSİYONDA NAKİT SIKIŞIKLIĞI TESPİT EDİLMEDİ',
        'advisorTitle': 'LİKİDİTE OPTİMİZASYON FIRSATI',
        'advisorBody': 'Atıl kalan ₺4.2M nakit rezervinin gecelik repoda değerlendirilmesiyle aylık ₺168.000 ek getiri sağlanabilir.',
        'advisorBtn': 'FONLAMA SENARYOSUNU İNCELE ›'
    },
    'finansal-analiz': {
        'accent': '#2563eb',
        'subAccent': '#0284c7',
        'ribbon': '#1d4ed8',
        'tag': 'FİNANSAL ANALİZ & YÖNETİM',
        'kpi1': ('İŞLETME DEĞERİ (EV)', '₺84.250.000', '↑ %18.4', 'indirgenmiş nakit akımı (DCF)', '10 YILLIK'),
        'kpi2': ('FAVÖK / EBITDA MARJI', '%28.6', '₺24.1M Yıllık', 'sektör ortalaması: %19.2', 'YÜKSEK'),
        'kpi3': ('NET BORÇ / FAVÖK', '1.24x', '✓ ÇOK GÜVENLİ', 'kaldıraç çarpanı sınırı: 3.0x', 'DÜŞÜK RİSK'),
        'kpi4': ('SERMAYE GETİRİSİ (ROIC)', '%34.2', '↑ %5.8', 'WACC üzeri katma değer', 'DEĞER ÜRETEN'),
        'formula': '=DCF_VALUATION(ProjeksiyonNakitAkislari; TerminalDeger; WACC)',
        'chartTitle': '5 YILLIK KÂRLILIK, FAVÖK & NAKİT YARATMA TRENDİ',
        'chartSub': 'Net Gelir Çarpanları ve Faaliyet Verimlilik Dinamikleri',
        'radarTitle': 'FİNANSAL SAĞLIK & İFLAS RİSKİ (ALTMAN Z-SCORE)',
        'radar1': ('Altman Z-Score Güvenlik Puanı', '3.85 (Güvenli Bölge > 2.9)', 330),
        'radar2': ('Faiz Karşılama Oranı (ICR)', '6.40x (Eşik: 2.50x)', 310),
        'radarStatus': '★ FİNANSAL GÖSTERGELER AAA KURUMSAL DERECELENDİRME UYUMLUDUR',
        'advisorTitle': 'STRATEJİK VALUASYON TAVSİYESİ',
        'advisorBody': 'Çalışma sermayesi optimizasyonu ile serbest nakit akımı %12 artırılarak şirket çarpanı 1.4 puan yükseltilebilir.',
        'advisorBtn': 'DEĞERLEME RAPORUNU ÇALIŞTIR ›'
    },
    'muhasebe-ve-vergi': {
        'accent': '#0f766e',
        'subAccent': '#059669',
        'ribbon': '#0f766e',
        'tag': 'VERGİ & MEVZUAT UYUMU',
        'kpi1': ('KORUNAN VERGİ MATRAHI', '₺42.600.000', 'Tam Beyan', 'GİB & VUK mevzuat denetimli', '12 DÖNEM'),
        'kpi2': ('KDV İADE / MAHSUP HAKKI', '₺5.840.000', 'Hazır Dosya', 'YMM & Vergi Dairesi onayına hazır', 'İNCELEMEDE'),
        'kpi3': ('YASAL VERGİ TASARRUFU', '₺1.920.000', '↑ %22.4', 'Ar-Ge, indirim ve istisnalar', 'AVANTAJ'),
        'kpi4': ('DENETİM RİSK PUANI', '0 / 100', '✓ SIFIR CEZA', 'çapraz mutabakat doğrulandı', 'TAM GÜVENLİ'),
        'formula': '=EĞER(MatrahFarkı=0; "MUTABIK"; "DÜZELTME_GEREKLİ")',
        'chartTitle': 'AYLIK KDV, KURUMLAR & STOPAJ YÜKÜ DAĞILIMI',
        'chartSub': 'Ödenecek vs. Devreden KDV ve Tevkifat Dengesi',
        'radarTitle': 'GİB İNCELEME & ANOMALİ TESPİT RADARI',
        'radar1': ('Kasa / Banka Adat Riski', '0.00 TL (Sıfır Anomali)', 350),
        'radar2': ('KKEG Sınır Uyum Derecesi', '%98.5 Tam Uyum', 320),
        'radarStatus': '✓ TÜM BEYANNAME KALEMLERİ VE MİZAN BAKİYELERİ KUSURSUZ DOĞRULANDI',
        'advisorTitle': 'VERGİ KALKANI & İSTİSNA TAVSİYESİ',
        'advisorBody': 'Yeniden değerleme ve amortisman planlaması ile gelecek çeyrekte ₺480.000 ek nakit vergi avantajı sağlanabilir.',
        'advisorBtn': 'VERGİ SAVUNMA DOSYASINI AÇ ›'
    },
    'butce-ve-planlama': {
        'accent': '#4f46e5',
        'subAccent': '#0284c7',
        'ribbon': '#4338ca',
        'tag': 'BÜTÇE & STRATEJİK PLANLAMA',
        'kpi1': ('YILLIK BÜTÇE HACMİ', '₺68.500.000', 'Onaylı', 'tüm departmanlar konsolide', '12 AY'),
        'kpi2': ('GERÇEKLEŞME / FARK', '-₺1.240.000', '↓ %1.8', 'bütçe tavanı aşılmadı', 'KONTROL ALTINDA'),
        'kpi3': ('FAALİYET SERBEST KÂRI', '₺16.850.000', '↑ %8.4', 'planlanan hedefin üzerinde', 'POZİTİF'),
        'kpi4': ('BÜTÇE DİSİPLİN PUANI', '%96.8', '✓ YÜKSEK', 'departman bütçe sadakati', 'BAŞARILI'),
        'formula': '=BÜTÇE_SAPMA_ORANI(GerçekleşenHarcama; HedefBütçe; LimitToleransı)',
        'chartTitle': 'DEPARTMAN BÜTÇE PERFORMANSI & AYLIK SAPMALAR',
        'chartSub': 'Hedeflenen Bütçe Tavanı vs. Gerçekleşen Fiili Harcamalar',
        'radarTitle': 'BÜTÇE AŞIM & MALİYET TAŞMA ERKEN UYARISI',
        'radar1': ('Operasyonel Gider (OPEX) Limiti', '%88 Kullanım (Güvenli)', 300),
        'radar2': ('Yatırım Bütçesi (CAPEX) Gerçekleşme', '%64 Plan Dahilinde', 240),
        'radarStatus': '✓ TÜM DEPARTMAN GİDERLERİ YILLIK ONAYLI TAVAN İÇERİSİNDEDİR',
        'advisorTitle': 'BÜTÇE REVİZYON TAVSİYESİ',
        'advisorBody': 'Pazarlama departmanı bütçesinden kalan kaynağın dijital altyapıya aktarılmasıyla ROI %14 artırılabilir.',
        'advisorBtn': 'REVİZYON SİMÜLASYONUNU ÇALIŞTIR ›'
    },
    'stok-ve-uretim': {
        'accent': '#d97706',
        'subAccent': '#ca8a04',
        'ribbon': '#b45309',
        'tag': 'ÜRETİM, STOK & MALİYET',
        'kpi1': ('TOPLAM STOK DEĞERİ', '₺18.420.000', 'Aktif', '3.450 SKU / 4 Depo Konsolide', 'FIFO / ORTALAMA'),
        'kpi2': ('OEE GENEL EKİPMAN', '%88.6', '↑ %4.2', 'dünya klasmanı hedef: %85', 'YÜKSEK VERİM'),
        'kpi3': ('STOK DEVİR HIZI', '8.4x / Yıl', '43 Gün', 'sektör ortalaması: 58 gün', 'HIZLI DÖNGÜ'),
        'kpi4': ('FİRE & ISKARTA ORANI', '%1.24', '↓ %0.6', 'üretim reçete toleransında', 'MİNİMAL KAYIP'),
        'formula': '=BİRİM_ÜRÜN_MALİYETİ(Hammadde; Direktİşçilik; GenelÜretimGideri)',
        'chartTitle': 'ABC STOK DAĞILIMI & KAPASİTE KULLANIM EĞRİSİ',
        'chartSub': 'Ürün Grubu Bazlı Satış Hacmi vs. Depo Maliyeti',
        'radarTitle': 'ÖLÜ STOK & MİNİMUM SEVİYE KRİTİK ALARMI',
        'radar1': ('Kritik Sipariş Eşiğindeki Kalemler', '4 Kalem (Otomatik Sipariş)', 260),
        'radar2': ('Depo Doluluk & Raf Kapasitesi', '%78.4 Optimum Denge', 290),
        'radarStatus': '✓ STOK KRİZİ VEYA HAT DURUŞ RİSKİ BULUNMAMAKTADIR',
        'advisorTitle': 'PARTİ BÜYÜKLÜĞÜ (EOQ) TAVSİYESİ',
        'advisorBody': 'A grubu 12 kritik hammadde siparişinin EOQ modeline çekilmesiyle yıllık ₺340.000 taşıma tasarrufu sağlanır.',
        'advisorBtn': 'EOQ ANALİZİNİ GÖRÜNTÜLE ›'
    },
    'satis-ve-fiyatlama': {
        'accent': '#7c3aed',
        'subAccent': '#2563eb',
        'ribbon': '#6d28d9',
        'tag': 'SATIŞ & KÂRLILIK ANALİZİ',
        'kpi1': ('TOPLAM SATIŞ CİROSU', '₺54.800.000', '↑ %24.6', 'çoklu kanal konsolide satış', 'YILLIK'),
        'kpi2': ('NET KÂR MARJI', '%22.8', '₺12.49M Net', 'komisyon ve kargo sonrası net', 'YÜKSEK'),
        'kpi3': ('ORTALAMA SEPET (AOV)', '₺1.850', '↑ %16.2', 'müşteri başına işlem hacmi', 'BÜYÜYEN'),
        'kpi4': ('KOMİSYON KAYBI TESPİTİ', '₺420.000', 'Geri Alındı', 'pazaryeri kesinti mutabakatı', 'KORUNDU'),
        'formula': '=NET_SATIŞ_KÂRI(SatışFiyatı; BirimMaliyet; PazaryeriKomisyonu; Kargo)',
        'chartTitle': 'KANAL BAZLI SATIŞ VE NET KÂRLILIK KARŞILAŞTIRMASI',
        'chartSub': 'Pazaryeri vs. Doğrudan Satış Efektif Kâr Marjları',
        'radarTitle': 'FİYATLAMA ELASTİKİYETİ & ZARARINA SATIŞ RADARI',
        'radar1': ('Eksi Marjlı Ürün Sayısı', '0 Ürün (Zararına Satış Yok)', 350),
        'radar2': ('Pazaryeri Hakediş Mutabakatı', '%100 Eksiksiz Tahsilat', 340),
        'radarStatus': '★ TÜM SATIŞ KANALLARINDA KOMİSYON VE FİYATLAMA HATASIZ ÇALIŞMAKTADIR',
        'advisorTitle': 'DİNAMİK FİYATLAMA FIRSATI',
        'advisorBody': 'Yüksek talep gören 18 üründe %4.5 fiyat ayarlaması yapılarak ciro kaybı olmadan net kâr ₺580.000 artırılabilir.',
        'advisorBtn': 'FİYAT SİMÜLATÖRÜNE GEÇ ›'
    },
    'personel-ve-bordro': {
        'accent': '#0284c7',
        'subAccent': '#0d9488',
        'ribbon': '#0369a1',
        'tag': 'İK, BORDRO & TAZMİNAT',
        'kpi1': ('TOPLAM PERSONEL KADROSU', '164 Kişi', 'Tam Zamanlı', '4 Departman / 2 Şube Konsolide', 'AKTİF'),
        'kpi2': ('AYLIK BORDRO MALİYETİ', '₺6.850.000', 'Brüt + SGK', 'işveren maliyeti dahil', 'DÜZENLİ'),
        'kpi3': ('KIDEM TAZMİNATI YÜKÜ', '₺9.420.000', 'Karşılık Ayrıldı', 'güncel tavan ücret üzerinden', 'FONLANDI'),
        'kpi4': ('SGK TEŞVİK TASARRUFU', '₺485.000 / Ay', '↑ %18.2', '7103, 6111 ve 5510 teşvikleri', 'OPTİMİZE'),
        'formula': '=İŞVEREN_NET_MALİYETİ(BrütÜcret; SGK_İşverenPayı; İşsizlikPayı; Teşvikler)',
        'chartTitle': 'DEPARTMAN BAZINDA İŞÇİLİK MALİYETLERİ VE SGK YÜKÜ',
        'chartSub': 'Brüt Ücret vs. Net Maaş vs. İşveren Yasal Yükümlülükleri',
        'radarTitle': 'İŞÇİLİK RİSKLERİ & DAVA ERKEN UYARI SİSTEMİ',
        'radar1': ('Fazla Mesai Limit Aşımı (Yıllık 270s)', '0 Personel (Yasal Sınırda)', 350),
        'radar2': ('Kullanılmayan Yıllık İzin Yükü', '142 Gün Kontrol Altında', 280),
        'radarStatus': '✓ TÜM PERSONEL BORDRO VE TAZMİNAT HESAPLAMALARI YASAL GÜVENCEDEDİR',
        'advisorTitle': 'TEŞVİK OPTİMİZASYON TAVSİYESİ',
        'advisorBody': 'Yeni işe alınan 6 personelin 6111 teşvikine yönlendirilmesiyle sonraki 18 ay boyunca ₺216.000 ek prim tasarrufu sağlanır.',
        'advisorBtn': 'TEŞVİK RAPORUNU İNDİR ›'
    }
}

def generate_svg(meta):
    cfg = CATEGORY_CONFIGS.get(meta['category'], CATEGORY_CONFIGS['finansal-analiz'])
    name = meta['name']
    slug = meta['slug']
    sheets = meta['sheets']
    inputs = meta['inputs']
    outputs = meta['outputs']

    # Gerçekçi Tablo Satırları (Outputs veya Inputs'tan beslenir)
    table_rows = []
    items = outputs if outputs else inputs
    if not items:
        items = ['Konsolide Analiz Kalemi', 'Dönemsel Operasyon', 'Mali Yükümlülük', 'Nihai Karar Göstergesi']
    
    # 4 satır doldur
    sample_amounts = ['₺8.450.000', '₺12.200.000', '₺4.850.000', '₺2.140.000']
    sample_subamounts = ['₺425.000', '₺580.400', '₺235.200', '₺118.000']
    sample_periods = ['12 Ay', '24 Ay', '8 Ay', '16 Ay']
    sample_badges = [
        ('✓ DÜZENLİ', '#dcfce7', '#15803d'),
        ('★ AVANTAJ', '#dbeafe', '#1d4ed8'),
        ('● YAKLAŞAN', '#fef3c7', '#b45309'),
        ('✓ TAM UYUM', '#ccfbf1', '#0f766e')
    ]

    for i in range(4):
        item_text = items[i % len(items)]
        if len(item_text) > 28:
            item_text = item_text[:26] + '...'
        
        badge_text, badge_bg, badge_color = sample_badges[i]
        table_rows.append({
            'col1': item_text,
            'col2': sample_amounts[i],
            'col3': sample_subamounts[i],
            'col4': sample_periods[i],
            'badge': badge_text,
            'badge_bg': badge_bg,
            'badge_color': badge_color
        })

    # Sheet listesi
    clean_sheets = []
    has_pano = False
    for s in sheets[:7]:
        clean_sheets.append(s)
        if 'PANO' in s or 'DASHBOARD' in s:
            has_pano = True
    if not has_pano:
        clean_sheets.insert(2, 'PANO')
    clean_sheets = clean_sheets[:6]

    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 750" width="1200" height="750" style="background:#0f172a; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;">
  <defs>
    <filter id="dropShadow" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="3" stdDeviation="5" flood-color="#000000" flood-opacity="0.12" />
    </filter>
    <filter id="glowGreen" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="0" stdDeviation="4" flood-color="#10b981" flood-opacity="0.4" />
    </filter>
    <linearGradient id="cardGrad1" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#ffffff"/>
      <stop offset="100%" stop-color="#f8fafc"/>
    </linearGradient>
    <linearGradient id="heroRibbon" x1="0" y1="0" x2="1" y2="0">
      <stop offset="0%" stop-color="{cfg['accent']}"/>
      <stop offset="100%" stop-color="#0f172a"/>
    </linearGradient>
  </defs>

  <!-- 1. DIŞ ÇERÇEVE & ARKA PLAN -->
  <rect width="1200" height="750" fill="#f1f5f9"/>

  <!-- 2. EXCEL ÜST PENCERE BAŞLIĞI -->
  <rect width="1200" height="38" fill="#1e293b"/>
  <circle cx="20" cy="19" r="6" fill="#ef4444"/>
  <circle cx="38" cy="19" r="6" fill="#f59e0b"/>
  <circle cx="56" cy="19" r="6" fill="#10b981"/>

  <g transform="translate(85, 11)">
    <rect width="22" height="18" rx="3" fill="#107c41"/>
    <text x="11" y="13" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">X</text>
    <text x="32" y="13" font-size="12" font-weight="600" fill="#f8fafc">{html.escape(slug)}.xlsx — Excel Arşiv Decision OS v15.1</text>
  </g>
  <rect x="960" y="8" width="220" height="22" rx="11" fill="#334155"/>
  <circle cx="975" cy="19" r="4" fill="#10b981" filter="url(#glowGreen)"/>
  <text x="988" y="23" font-size="10" font-weight="700" fill="#e2e8f0">{cfg['tag']}</text>

  <!-- 3. EXCEL RIBBON & FORMULA BAR -->
  <rect y="38" width="1200" height="32" fill="#ffffff" stroke="#cbd5e1" stroke-width="1"/>
  <text x="25" y="58" font-size="11" font-weight="600" fill="#64748b">Dosya</text>
  <rect x="70" y="44" width="56" height="26" rx="4" fill="#dcfce7"/>
  <text x="98" y="61" font-size="11" font-weight="800" fill="#15803d" text-anchor="middle">Pano</text>
  <text x="145" y="58" font-size="11" font-weight="600" fill="#64748b">Girdiler</text>
  <text x="205" y="58" font-size="11" font-weight="600" fill="#64748b">Hesap Motoru</text>
  <text x="300" y="58" font-size="11" font-weight="600" fill="#64748b">Duyarlılık</text>
  <text x="380" y="58" font-size="11" font-weight="600" fill="#64748b">Stres Testi</text>
  <text x="460" y="58" font-size="11" font-weight="600" fill="#64748b">Yönetim Raporu</text>

  <rect y="70" width="1200" height="26" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1"/>
  <rect x="15" y="73" width="55" height="20" rx="3" fill="#ffffff" stroke="#cbd5e1" stroke-width="1"/>
  <text x="42" y="87" font-size="11" font-weight="700" fill="#334155" text-anchor="middle">A1</text>
  <text x="85" y="87" font-size="12" font-style="italic" font-weight="700" fill="#0f766e">fx</text>
  <line x1="105" y1="74" x2="105" y2="92" stroke="#cbd5e1" stroke-width="1"/>
  <text x="115" y="87" font-size="11" font-family="'JetBrains Mono',monospace" fill="#0369a1">{html.escape(cfg['formula'])}</text>

  <!-- 4. DASHBOARD HEADER KONSOLU -->
  <g transform="translate(20, 105)">
    <text x="0" y="24" font-size="18" font-weight="900" fill="#0f172a" letter-spacing="-0.3">{html.escape(name[:55].upper())}</text>
    <text x="0" y="42" font-size="11" font-weight="600" fill="#64748b">Yönetici Karar Kokpiti, Otomatik Doğrulama ve Dinamik Hesaplama Sistemi · 2026</text>

    <rect x="850" y="6" width="140" height="34" rx="6" fill="#ffffff" stroke="#cbd5e1" stroke-width="1"/>
    <text x="862" y="27" font-size="11" font-weight="700" fill="#334155">Dönem: 2026 / Tam</text>
    <text x="975" y="27" font-size="10" fill="#64748b">▼</text>

    <rect x="1000" y="6" width="160" height="34" rx="6" fill="{cfg['accent']}" filter="url(#dropShadow)"/>
    <text x="1080" y="27" font-size="11" font-weight="800" fill="#ffffff" text-anchor="middle">★ DOĞRULANMIŞ</text>
  </g>

  <!-- 5. DÖRT ADET KRİTİK YÖNETİCİ KPI KARTI -->
  <g transform="translate(20, 158)">
    <!-- KART 1 -->
    <g transform="translate(0, 0)">
      <rect width="275" height="90" rx="8" fill="url(#cardGrad1)" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <rect x="0" y="0" width="5" height="90" rx="2" fill="{cfg['accent']}"/>
      <text x="16" y="23" font-size="10.5" font-weight="700" fill="#64748b">{cfg['kpi1'][0]}</text>
      <rect x="195" y="10" width="68" height="18" rx="9" fill="#dbeafe"/>
      <text x="229" y="23" font-size="8.5" font-weight="800" fill="#1d4ed8" text-anchor="middle">{cfg['kpi1'][4]}</text>
      <text x="16" y="54" font-size="22" font-weight="900" fill="#0f172a">{cfg['kpi1'][1]}</text>
      <text x="16" y="75" font-size="10" font-weight="700" fill="#16a34a">{cfg['kpi1'][2]}</text>
      <text x="68" y="75" font-size="10" font-weight="500" fill="#64748b">{cfg['kpi1'][3]}</text>
    </g>

    <!-- KART 2 -->
    <g transform="translate(295, 0)">
      <rect width="275" height="90" rx="8" fill="url(#cardGrad1)" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <rect x="0" y="0" width="5" height="90" rx="2" fill="#dc2626"/>
      <text x="16" y="23" font-size="10.5" font-weight="700" fill="#64748b">{cfg['kpi2'][0]}</text>
      <rect x="180" y="10" width="83" height="18" rx="9" fill="#fee2e2"/>
      <text x="221" y="23" font-size="8.5" font-weight="800" fill="#b91c1c" text-anchor="middle">{cfg['kpi2'][4]}</text>
      <text x="16" y="54" font-size="22" font-weight="900" fill="#b91c1c">{cfg['kpi2'][1]}</text>
      <text x="16" y="75" font-size="10" font-weight="700" fill="#dc2626">{cfg['kpi2'][2]}</text>
      <text x="80" y="75" font-size="10" font-weight="500" fill="#64748b">{cfg['kpi2'][3]}</text>
    </g>

    <!-- KART 3 -->
    <g transform="translate(590, 0)">
      <rect width="275" height="90" rx="8" fill="url(#cardGrad1)" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <rect x="0" y="0" width="5" height="90" rx="2" fill="#16a34a"/>
      <text x="16" y="23" font-size="10.5" font-weight="700" fill="#64748b">{cfg['kpi3'][0]}</text>
      <rect x="195" y="10" width="68" height="18" rx="9" fill="#dcfce7"/>
      <text x="229" y="23" font-size="8.5" font-weight="800" fill="#15803d" text-anchor="middle">{cfg['kpi3'][4]}</text>
      <text x="16" y="54" font-size="22" font-weight="900" fill="#0f172a">{cfg['kpi3'][1]}</text>
      <rect x="16" y="66" width="160" height="7" rx="3.5" fill="#e2e8f0"/>
      <rect x="16" y="66" width="95" height="7" rx="3.5" fill="#16a34a"/>
      <text x="186" y="73" font-size="9.5" font-weight="700" fill="#64748b">{cfg['kpi3'][2]}</text>
    </g>

    <!-- KART 4 -->
    <g transform="translate(885, 0)">
      <rect width="275" height="90" rx="8" fill="url(#cardGrad1)" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <rect x="0" y="0" width="5" height="90" rx="2" fill="{cfg['subAccent']}"/>
      <text x="16" y="23" font-size="10.5" font-weight="700" fill="#64748b">{cfg['kpi4'][0]}</text>
      <rect x="195" y="10" width="68" height="18" rx="9" fill="#ccfbf1"/>
      <text x="229" y="23" font-size="8.5" font-weight="800" fill="#0f766e" text-anchor="middle">{cfg['kpi4'][4]}</text>
      <text x="16" y="54" font-size="22" font-weight="900" fill="#0f766e">{cfg['kpi4'][1]}</text>
      <text x="16" y="75" font-size="10" font-weight="700" fill="#0f766e">{cfg['kpi4'][2]}</text>
      <text x="82" y="75" font-size="10" font-weight="500" fill="#64748b">{cfg['kpi4'][3]}</text>
    </g>
  </g>

  <!-- 6. ORTA BÖLÜM: GRAFİK & DETAYLI ÇİZELGE -->
  <g transform="translate(20, 262)">
    <!-- SOL GRAFİK -->
    <g transform="translate(0, 0)">
      <rect width="460" height="270" rx="10" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <text x="16" y="25" font-size="12" font-weight="800" fill="#0f172a">{cfg['chartTitle'][:45]}</text>
      <text x="16" y="40" font-size="10" font-weight="500" fill="#64748b">{cfg['chartSub'][:55]}</text>

      <rect x="285" y="15" width="10" height="10" rx="2" fill="{cfg['accent']}"/>
      <text x="300" y="24" font-size="9" font-weight="600" fill="#475569">Aktif Hacim</text>
      <rect x="375" y="15" width="10" height="10" rx="2" fill="#dc2626"/>
      <text x="390" y="24" font-size="9" font-weight="600" fill="#475569">Maliyet / Yük</text>

      <g transform="translate(35, 58)">
        <line x1="0" y1="0" x2="400" y2="0" stroke="#f1f5f9" stroke-width="1"/>
        <line x1="0" y1="40" x2="400" y2="40" stroke="#f1f5f9" stroke-width="1"/>
        <line x1="0" y1="80" x2="400" y2="80" stroke="#f1f5f9" stroke-width="1"/>
        <line x1="0" y1="120" x2="400" y2="120" stroke="#f1f5f9" stroke-width="1"/>
        <line x1="0" y1="160" x2="400" y2="160" stroke="#cbd5e1" stroke-width="1.5"/>

        <text x="-8" y="5" font-size="9" fill="#94a3b8" text-anchor="end">100%</text>
        <text x="-8" y="45" font-size="9" fill="#94a3b8" text-anchor="end">75%</text>
        <text x="-8" y="85" font-size="9" fill="#94a3b8" text-anchor="end">50%</text>
        <text x="-8" y="125" font-size="9" fill="#94a3b8" text-anchor="end">25%</text>
        <text x="-8" y="164" font-size="9" fill="#94a3b8" text-anchor="end">0</text>

        <!-- Sütunlar -->
        <rect x="15" y="45" width="16" height="115" rx="2" fill="{cfg['accent']}"/><rect x="33" y="85" width="16" height="75" rx="2" fill="#dc2626"/>
        <text x="32" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">D1</text>

        <rect x="75" y="35" width="16" height="125" rx="2" fill="{cfg['accent']}"/><rect x="93" y="80" width="16" height="80" rx="2" fill="#dc2626"/>
        <text x="92" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">D2</text>

        <rect x="135" y="25" width="16" height="135" rx="2" fill="{cfg['accent']}"/><rect x="153" y="70" width="16" height="90" rx="2" fill="#dc2626"/>
        <text x="152" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">D3</text>

        <rect x="195" y="40" width="16" height="120" rx="2" fill="{cfg['accent']}"/><rect x="213" y="90" width="16" height="70" rx="2" fill="#dc2626"/>
        <text x="212" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">D4</text>

        <rect x="255" y="20" width="16" height="140" rx="2" fill="{cfg['accent']}"/><rect x="273" y="95" width="16" height="65" rx="2" fill="#dc2626"/>
        <text x="272" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">D5</text>

        <rect x="315" y="15" width="16" height="145" rx="2" fill="{cfg['accent']}"/><rect x="333" y="110" width="16" height="50" rx="2" fill="#dc2626"/>
        <text x="332" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">D6</text>

        <path d="M 32 55 L 92 45 L 152 35 L 212 50 L 272 30 L 332 20" fill="none" stroke="#10b981" stroke-width="3"/>
        <circle cx="332" cy="20" r="4" fill="#10b981"/>
        <rect x="340" y="12" width="62" height="18" rx="4" fill="#10b981"/>
        <text x="371" y="25" font-size="8.5" font-weight="800" fill="#ffffff" text-anchor="middle">+%32 Verim</text>
      </g>
    </g>

    <!-- SAĞ VERİ ÇİZELGESİ -->
    <g transform="translate(475, 0)">
      <rect width="685" height="270" rx="10" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <text x="16" y="25" font-size="12" font-weight="800" fill="#0f172a">KONSOLİDE OPERASYON VE KARAR MATRİSİ</text>
      <text x="16" y="40" font-size="10" font-weight="500" fill="#64748b">Anlık Hesaplama Sonuçları, Dinamik Formül Çıktıları ve Durum Puanlaması</text>

      <!-- Tablo Başlık -->
      <g transform="translate(16, 52)">
        <rect width="653" height="26" rx="4" fill="#1e293b"/>
        <text x="12" y="17" font-size="9.5" font-weight="800" fill="#ffffff">PARAMETRE / KALEM</text>
        <text x="300" y="17" font-size="9.5" font-weight="800" fill="#ffffff" text-anchor="end">ANA TUTAR</text>
        <text x="400" y="17" font-size="9.5" font-weight="800" fill="#ffffff" text-anchor="end">FARK / MARJ</text>
        <text x="470" y="17" font-size="9.5" font-weight="800" fill="#ffffff" text-anchor="end">VADE</text>
        <text x="590" y="17" font-size="9.5" font-weight="800" fill="#ffffff" text-anchor="middle">KONTROL DURUMU</text>
      </g>

      <!-- Satırlar -->'''

    y_pos = 82
    for r in table_rows:
        bg = '#f8fafc' if (y_pos // 36) % 2 == 0 else '#ffffff'
        svg += f'''
      <g transform="translate(16, {y_pos})">
        <rect width="653" height="32" fill="{bg}"/>
        <text x="12" y="20" font-size="10.5" font-weight="700" fill="#0f172a">{html.escape(r['col1'])}</text>
        <text x="300" y="20" font-size="11" font-weight="800" fill="#0f172a" text-anchor="end">{r['col2']}</text>
        <text x="400" y="20" font-size="11" font-weight="800" fill="#dc2626" text-anchor="end">{r['col3']}</text>
        <text x="470" y="20" font-size="10" font-weight="600" fill="#64748b" text-anchor="end">{r['col4']}</text>
        <rect x="535" y="7" width="110" height="18" rx="9" fill="{r['badge_bg']}"/>
        <text x="590" y="20" font-size="9" font-weight="800" fill="{r['badge_color']}" text-anchor="middle">{r['badge']}</text>
      </g>'''
        y_pos += 36

    svg += f'''
      <!-- Toplam Satırı -->
      <g transform="translate(16, 226)">
        <rect width="653" height="30" rx="4" fill="#0f172a"/>
        <text x="12" y="19" font-size="10.5" font-weight="800" fill="#ffffff">KONSOLİDE SİSTEM TOPLAMI</text>
        <text x="300" y="19" font-size="11.5" font-weight="900" fill="#38bdf8" text-anchor="end">₺42.850.000</text>
        <text x="400" y="19" font-size="11.5" font-weight="900" fill="#f87171" text-anchor="end">₺3.420.000</text>
        <text x="590" y="19" font-size="9.5" font-weight="800" fill="#4ade80" text-anchor="middle">TAM DOĞRULANDI (0 HATA)</text>
      </g>
    </g>
  </g>

  <!-- 7. ALT KISIM: 3 ADET KOKPİT KARTI -->
  <g transform="translate(20, 546)">
    <!-- SOL: DONUT -->
    <g transform="translate(0, 0)">
      <rect width="360" height="145" rx="10" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <text x="16" y="22" font-size="11.5" font-weight="800" fill="#0f172a">KONSOLİDE PAY VE AĞIRLIK</text>
      
      <g transform="translate(70, 82)">
        <circle cx="0" cy="0" r="44" fill="none" stroke="#e2e8f0" stroke-width="16"/>
        <circle cx="0" cy="0" r="44" fill="none" stroke="{cfg['accent']}" stroke-width="16" stroke-dasharray="116 276" stroke-dashoffset="0"/>
        <circle cx="0" cy="0" r="44" fill="none" stroke="{cfg['subAccent']}" stroke-width="16" stroke-dasharray="83 276" stroke-dashoffset="-116"/>
        <circle cx="0" cy="0" r="44" fill="none" stroke="#f59e0b" stroke-width="16" stroke-dasharray="41 276" stroke-dashoffset="-199"/>
        <circle cx="0" cy="0" r="44" fill="none" stroke="#8b5cf6" stroke-width="16" stroke-dasharray="36 276" stroke-dashoffset="-240"/>
        <text x="0" y="4" font-size="10.5" font-weight="900" fill="#0f172a" text-anchor="middle">100%</text>
      </g>

      <g transform="translate(150, 42)">
        <rect x="0" y="0" width="9" height="9" rx="2" fill="{cfg['accent']}"/><text x="15" y="8" font-size="9.5" font-weight="600" fill="#334155">Ana Segment (%42)</text>
        <rect x="0" y="22" width="9" height="9" rx="2" fill="{cfg['subAccent']}"/><text x="15" y="30" font-size="9.5" font-weight="600" fill="#334155">İkincil Grup (%30)</text>
        <rect x="0" y="44" width="9" height="9" rx="2" fill="#f59e0b"/><text x="15" y="52" font-size="9.5" font-weight="600" fill="#334155">Destek Birimi (%15)</text>
        <rect x="0" y="66" width="9" height="9" rx="2" fill="#8b5cf6"/><text x="15" y="74" font-size="9.5" font-weight="600" fill="#334155">Rezervler (%13)</text>
      </g>
    </g>

    <!-- ORTA: ERKEN UYARI -->
    <g transform="translate(375, 0)">
      <rect width="405" height="145" rx="10" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <text x="16" y="22" font-size="11.5" font-weight="800" fill="#0f172a">{cfg['radarTitle']}</text>

      <g transform="translate(16, 36)">
        <text x="0" y="12" font-size="10.5" font-weight="700" fill="#334155">{cfg['radar1'][0]}</text>
        <text x="370" y="12" font-size="10.5" font-weight="800" fill="#15803d" text-anchor="end">{cfg['radar1'][1]}</text>
        <rect x="0" y="18" width="370" height="7" rx="3.5" fill="#f1f5f9"/>
        <rect x="0" y="18" width="{cfg['radar1'][2]}" height="7" rx="3.5" fill="#16a34a"/>
      </g>

      <g transform="translate(16, 70)">
        <text x="0" y="12" font-size="10.5" font-weight="700" fill="#334155">{cfg['radar2'][0]}</text>
        <text x="370" y="12" font-size="10.5" font-weight="800" fill="#15803d" text-anchor="end">{cfg['radar2'][1]}</text>
        <rect x="0" y="18" width="370" height="7" rx="3.5" fill="#f1f5f9"/>
        <rect x="0" y="18" width="{cfg['radar2'][2]}" height="7" rx="3.5" fill="#16a34a"/>
      </g>

      <g transform="translate(16, 106)">
        <rect width="370" height="26" rx="6" fill="#ecfdf5" stroke="#a7f3d0" stroke-width="1"/>
        <text x="185" y="17" font-size="9.5" font-weight="800" fill="#047857" text-anchor="middle">{cfg['radarStatus']}</text>
      </g>
    </g>

    <!-- SAĞ: KARAR MOTORU -->
    <g transform="translate(795, 0)">
      <rect width="365" height="145" rx="10" fill="url(#heroRibbon)" filter="url(#dropShadow)"/>
      <rect x="14" y="12" width="100" height="20" rx="10" fill="#ffffff" fill-opacity="0.18"/>
      <text x="64" y="25" font-size="9" font-weight="900" fill="#ffffff" text-anchor="middle">KARAR MOTORU</text>

      <text x="16" y="52" font-size="13" font-weight="900" fill="#ffffff">{cfg['advisorTitle']}</text>
      <text x="16" y="72" font-size="10.5" font-weight="500" fill="#dcfce7">{html.escape(cfg['advisorBody'][:95])}</text>
      
      <rect x="16" y="96" width="333" height="32" rx="6" fill="#ffffff"/>
      <text x="182" y="116" font-size="10.5" font-weight="900" fill="{cfg['accent']}" text-anchor="middle">{cfg['advisorBtn']}</text>
    </g>
  </g>

  <!-- 8. EXCEL ÇALIŞMA SAYFASI SEKMELERİ -->
  <rect y="705" width="1200" height="45" fill="#e2e8f0" stroke="#cbd5e1" stroke-width="1"/>
  <g transform="translate(15, 715)">
    <text x="5" y="18" font-size="13" font-weight="700" fill="#64748b">◀  ▶</text>'''

    sheet_x = 45
    for s_name in clean_sheets:
        is_active = ('PANO' in s_name or 'DASHBOARD' in s_name)
        s_width = max(80, len(s_name) * 8 + 24)
        if is_active:
            svg += f'''
    <rect x="{sheet_x}" y="-3" width="{s_width}" height="38" fill="#ffffff" rx="4" stroke="#107c41" stroke-width="2"/>
    <rect x="{sheet_x}" y="32" width="{s_width}" height="3" fill="#ffffff"/>
    <text x="{sheet_x + s_width//2}" y="19" font-size="10" font-weight="900" fill="#107c41" text-anchor="middle">{html.escape(s_name)}</text>'''
        else:
            svg += f'''
    <rect x="{sheet_x}" y="0" width="{s_width}" height="35" fill="#f8fafc" rx="4"/>
    <text x="{sheet_x + s_width//2}" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">{html.escape(s_name)}</text>'''
        sheet_x += s_width + 8

    svg += f'''
    <circle cx="{sheet_x + 15}" cy="17" r="11" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="1"/>
    <text x="{sheet_x + 15}" y="21" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">+</text>

    <text x="1150" y="18" font-size="10" font-weight="600" fill="#64748b" text-anchor="end">HAZIR · 100% · HESAPLAMA: OTOMATİK</text>
  </g>
</svg>
'''
    # XML entity güvenliği: Kaçış yapılmamış tüm & işaretlerini &amp; yap
    svg = re.sub(r'&(?!(amp|lt|gt|quot|apos);)', '&amp;', svg)
    return svg

def run_batch():
    files = glob.glob(os.path.join(MDX_DIR, '*.mdx'))
    print(f"Generating rich dashboards for {len(files)} templates...")
    count = 0
    for f in files:
        meta = parse_mdx(f)
        svg_str = generate_svg(meta)
        slug = meta['slug']
        target_svg = os.path.join(KAPAK_DIR, f"{slug}.svg")
        with open(target_svg, 'w', encoding='utf-8') as out:
            out.write(svg_str)
        
        # PNG ekran görüntüsü 1 olarak da derle
        subprocess.run(['qlmanage', '-t', '-s', '1200', '-o', '/tmp', target_svg], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        tmp_png = f'/tmp/{slug}.svg.png'
        dest_png = os.path.join(SCREENSHOT_DIR, f"{slug}-1.png")
        if os.path.exists(tmp_png):
            os.replace(tmp_png, dest_png)
        count += 1

    print(f"Successfully generated {count} SVG & PNG dashboard covers in {KAPAK_DIR} and {SCREENSHOT_DIR}")

if __name__ == '__main__':
    run_batch()
