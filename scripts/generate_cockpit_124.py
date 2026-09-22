import json
import os
import math

with open('scripts/products_meta.json', 'r', encoding='utf-8') as f:
    products = json.load(f)

print(f"124 ürün için Enterprise Cockpit Kapak Üretimi başlatılıyor (Toplam: {len(products)})...")

def escape_xml(unsafe):
    if not unsafe:
        return ''
    return str(unsafe).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&apos;')

def polar_to_cartesian(cx, cy, r, angle_deg):
    rad = ((angle_deg - 90) * math.pi) / 180.0
    return cx + r * math.cos(rad), cy + r * math.sin(rad)

def donut_segment(cx, cy, r_in, r_out, start_ang, end_ang, fill):
    p1 = polar_to_cartesian(cx, cy, r_out, end_ang)
    p2 = polar_to_cartesian(cx, cy, r_out, start_ang)
    p3 = polar_to_cartesian(cx, cy, r_in, end_ang)
    p4 = polar_to_cartesian(cx, cy, r_in, start_ang)
    large = 0 if end_ang - start_ang <= 180 else 1
    d = f"M {p1[0]:.1f} {p1[1]:.1f} A {r_out} {r_out} 0 {large} 0 {p2[0]:.1f} {p2[1]:.1f} L {p4[0]:.1f} {p4[1]:.1f} A {r_in} {r_in} 0 {large} 1 {p3[0]:.1f} {p3[1]:.1f} Z"
    return f'<path d="{d}" fill="{fill}" />'

# 16 Profesyonel Analitik Grafik Motoru
def render_chart(chart_type, metric_title, metric_val, sub_text, color_primary="#107c41", color_accent="#2563eb"):
    if chart_type == 'waterfall':
        return f"""
        <rect width="360" height="430" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="25" y="32" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a">{escape_xml(metric_title)}</text>
        <rect x="210" y="16" width="130" height="28" rx="6" fill="#dcfce7" />
        <text x="275" y="35" font-family="system-ui" font-size="12" font-weight="900" fill="#0d5c3a" text-anchor="middle">{escape_xml(metric_val)}</text>
        <line x1="25" y1="365" x2="335" y2="365" stroke="#cbd5e1" stroke-width="2" />
        
        <rect x="35" y="240" width="35" height="125" rx="4" fill="#1e293b" /><text x="52" y="385" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H1</text>
        <rect x="85" y="180" width="35" height="60" rx="4" fill="{color_primary}" /><text x="102" y="385" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H2</text>
        <rect x="135" y="240" width="35" height="70" rx="4" fill="#dc2626" /><text x="152" y="385" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H3</text>
        <rect x="185" y="150" width="35" height="90" rx="4" fill="{color_primary}" /><text x="202" y="385" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H4</text>
        <rect x="235" y="220" width="35" height="70" rx="4" fill="#dc2626" /><text x="252" y="385" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H5</text>
        <rect x="285" y="90" width="40" height="275" rx="6" fill="{color_primary}" /><text x="305" y="385" font-family="system-ui" font-size="10" font-weight="800" fill="{color_primary}" text-anchor="middle">TOPLAM</text>
        
        <path d="M 52 240 L 102 180 L 152 240 L 202 150 L 252 220 L 305 90" fill="none" stroke="{color_accent}" stroke-width="3" stroke-linecap="round" />
        <circle cx="305" cy="90" r="5" fill="{color_accent}" />
        <text x="25" y="415" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b">{escape_xml(sub_text)}</text>
        """
    elif chart_type == 'gauge':
        return f"""
        <rect width="360" height="430" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="25" y="32" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a">{escape_xml(metric_title)}</text>
        
        <circle cx="180" cy="190" r="105" fill="none" stroke="#e2e8f0" stroke-width="22" stroke-dasharray="330" stroke-dashoffset="82" transform="rotate(135 180 190)" />
        <circle cx="180" cy="190" r="105" fill="none" stroke="{color_primary}" stroke-width="22" stroke-dasharray="330" stroke-dashoffset="115" transform="rotate(135 180 190)" stroke-linecap="round" />
        
        <text x="180" y="180" font-family="system-ui" font-size="34" font-weight="900" fill="#0f172a" text-anchor="middle">{escape_xml(metric_val)}</text>
        <text x="180" y="208" font-family="system-ui" font-size="11" font-weight="800" fill="{color_primary}" text-anchor="middle">TAM UYGUNLUK</text>
        
        <rect x="25" y="295" width="310" height="110" rx="12" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5" />
        <text x="45" y="325" font-family="system-ui" font-size="11" font-weight="700" fill="#475569">Denetim Standartı:</text>
        <text x="45" y="350" font-family="system-ui" font-size="9.5" font-weight="800" fill="{color_primary}">✓ {escape_xml(sub_text)}</text>
        <text x="45" y="380" font-family="system-ui" font-size="10" font-weight="600" fill="#64748b">Tüm formüller şeffaf ve denetlenebilir.</text>
        """
    elif chart_type == 'scurve':
        return f"""
        <rect width="360" height="430" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="25" y="32" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a">{escape_xml(metric_title)}</text>
        <rect x="220" y="16" width="120" height="28" rx="6" fill="#dbeafe" />
        <text x="280" y="35" font-family="system-ui" font-size="12" font-weight="900" fill="#1e40af" text-anchor="middle">{escape_xml(metric_val)}</text>
        
        <line x1="25" y1="365" x2="335" y2="365" stroke="#cbd5e1" stroke-width="2" />
        <path d="M 35 340 Q 140 330 200 180 T 320 80" fill="none" stroke="{color_primary}" stroke-width="5" stroke-linecap="round" />
        <circle cx="320" cy="80" r="7" fill="{color_primary}" />
        <circle cx="200" cy="180" r="5" fill="{color_accent}" />
        
        <g transform="translate(35, 120)">
          <rect width="140" height="50" rx="8" fill="#ffffff" stroke="#e2e8f0" stroke-width="1" />
          <text x="15" y="22" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b">Başabaş Eşiği</text>
          <text x="15" y="40" font-family="system-ui" font-size="13" font-weight="900" fill="{color_accent}">4. Çeyrek</text>
        </g>
        
        <text x="25" y="415" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b">{escape_xml(sub_text)}</text>
        """
    elif chart_type == 'stacked':
        return f"""
        <rect width="360" height="430" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="25" y="32" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a">{escape_xml(metric_title)}</text>
        <line x1="25" y1="365" x2="335" y2="365" stroke="#cbd5e1" stroke-width="2" />
        
        <!-- Stack 1 -->
        <rect x="50" y="260" width="55" height="105" rx="4" fill="#cbd5e1" />
        <rect x="50" y="190" width="55" height="70" rx="4" fill="{color_accent}" />
        <rect x="50" y="120" width="55" height="70" rx="4" fill="{color_primary}" />
        <text x="77" y="385" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">2024</text>
        
        <!-- Stack 2 -->
        <rect x="150" y="240" width="55" height="125" rx="4" fill="#cbd5e1" />
        <rect x="150" y="160" width="55" height="80" rx="4" fill="{color_accent}" />
        <rect x="150" y="80" width="55" height="80" rx="4" fill="{color_primary}" />
        <text x="177" y="385" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">2025</text>
        
        <!-- Stack 3 -->
        <rect x="250" y="210" width="55" height="155" rx="4" fill="#cbd5e1" />
        <rect x="250" y="120" width="55" height="90" rx="4" fill="{color_accent}" />
        <rect x="250" y="40" width="55" height="80" rx="4" fill="{color_primary}" />
        <text x="277" y="385" font-family="system-ui" font-size="10" font-weight="800" fill="{color_primary}" text-anchor="middle">2026 PRO</text>
        
        <text x="25" y="415" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b">{escape_xml(sub_text)}</text>
        """
    elif chart_type == 'donut':
        s1 = donut_segment(180, 200, 75, 120, 0, 190, color_primary)
        s2 = donut_segment(180, 200, 75, 120, 190, 290, color_accent)
        s3 = donut_segment(180, 200, 75, 120, 290, 360, "#f59e0b")
        return f"""
        <rect width="360" height="430" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="25" y="32" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a">{escape_xml(metric_title)}</text>
        {s1} {s2} {s3}
        <text x="180" y="195" font-family="system-ui" font-size="24" font-weight="900" fill="#0f172a" text-anchor="middle">{escape_xml(metric_val)}</text>
        <text x="180" y="218" font-family="system-ui" font-size="10" font-weight="800" fill="{color_primary}" text-anchor="middle">OPTİMUM</text>
        
        <g transform="translate(30, 345)">
          <circle cx="10" cy="10" r="5" fill="{color_primary}" /><text x="25" y="14" font-family="system-ui" font-size="11" font-weight="600" fill="#334155">Ana Pay (%52)</text>
          <circle cx="125" cy="10" r="5" fill="{color_accent}" /><text x="140" y="14" font-family="system-ui" font-size="11" font-weight="600" fill="#334155">İkincil (%28)</text>
          <circle cx="230" cy="10" r="5" fill="#f59e0b" /><text x="245" y="14" font-family="system-ui" font-size="11" font-weight="600" fill="#334155">Diğer (%20)</text>
        </g>
        <text x="25" y="415" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b">{escape_xml(sub_text)}</text>
        """
    else: # radar / matris / bar
        return f"""
        <rect width="360" height="430" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="25" y="32" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a">{escape_xml(metric_title)}</text>
        <g transform="translate(180, 200)">
          <polygon points="0,-110 100,-40 65,85 -65,85 -100,-40" fill="none" stroke="#cbd5e1" stroke-width="1.5" />
          <polygon points="0,-75 70,-25 45,60 -45,60 -70,-25" fill="none" stroke="#e2e8f0" stroke-width="1.5" />
          <polygon points="0,-100 85,-35 55,75 -50,65 -80,-30" fill="{color_primary}" fill-opacity="0.25" stroke="{color_primary}" stroke-width="3" />
          <circle cx="0" cy="-100" r="5" fill="{color_primary}" /><circle cx="85" cy="-35" r="5" fill="{color_primary}" /><circle cx="55" cy="75" r="5" fill="{color_primary}" /><circle cx="-50" cy="65" r="5" fill="{color_primary}" /><circle cx="-80" cy="-30" r="5" fill="{color_primary}" />
          <text x="0" y="-120" font-family="system-ui" font-size="10" font-weight="800" fill="#475569" text-anchor="middle">VERİM</text>
          <text x="110" y="-35" font-family="system-ui" font-size="10" font-weight="800" fill="#475569">KÂR</text>
          <text x="75" y="100" font-family="system-ui" font-size="10" font-weight="800" fill="#475569">HIZ</text>
          <text x="-75" y="100" font-family="system-ui" font-size="10" font-weight="800" fill="#475569" text-anchor="end">GÜVEN</text>
          <text x="-110" y="-35" font-family="system-ui" font-size="10" font-weight="800" fill="#475569" text-anchor="end">LİKİT</text>
        </g>
        <rect x="25" y="325" width="310" height="70" rx="10" fill="#ffffff" stroke="#e2e8f0" stroke-width="1" />
        <text x="45" y="352" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">Konsolide Performans Skoru</text>
        <text x="45" y="378" font-family="system-ui" font-size="18" font-weight="900" fill="{color_primary}">{escape_xml(metric_val)} · %96.4 Uyum</text>
        <text x="25" y="415" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b">{escape_xml(sub_text)}</text>
        """

def generate_cockpit_svg(prod, index):
    slug = prod['slug']
    title = prod['title']
    category = prod.get('category', 'FİNANSAL YÖNETİM')

    safe_title = escape_xml(title)
    safe_cat = escape_xml((category or 'FİNANSAL YÖNETİM').upper())

    # Konuya göre dinamik içerik üretimi (124 ürün için özel kurgulanmış değerler)
    # Temel anahtar kelimelere göre kategori eşleşmesi
    is_nakit = any(k in slug for k in ['nakit', 'kasa', 'tahsilat', 'pos', 'vade', 'cek', 'para'])
    is_vergi = any(k in slug for k in ['vergi', 'degerleme', 'amortisman', 'sgk', 'bordro', 'tesvik', 'kdv', 'ymm', 'matrah', 'ttk', 'konkordato'])
    is_maliyet = any(k in slug for k in ['maliyet', 'recete', 'fiyat', 'zam', 'ithalat', 'navlun', 'kargo', 'ihracat', 'desi', 'fire', 'menu', 'yem', 'mutfak'])
    is_yatirim = any(k in slug for k in ['yatirim', 'fizibilite', 'gayrimenkul', 'arsa', 'otel', 'ges', 'restoran', 'magaza', 'klinik', 'franchise', 'kira', 'proje', 'filo', 'sarj', 'sera', 'sirket'])
    is_ihale = any(k in slug for k in ['ihale', 'teklif', 'hakedis', 'kik', 'santiye', 'taseron', 'asiri-dusuk'])

    if is_nakit:
        kpis = [
            {'label': 'Mevcut Nakit Hacmi', 'value': '₺4.850.000'},
            {'label': 'Ortalama Tahsilat (DSO)', 'value': '34 Gün'},
            {'label': 'Öngörülen Net Likidite', 'value': '+₺2.420.000'}
        ]
        table_rows = [
            {'col1': 'Banka / Cari Tahsilat', 'col2': '₺3.850.000', 'col3': '%42.0', 'status': '✓ TAHSİL EDİLDİ', 'statusType': 'good'},
            {'col1': 'Sabit Gider & Rezerv', 'col2': '-₺820.000', 'col3': '%9.0', 'status': 'AYRILDI', 'statusType': 'normal'},
            {'col1': 'Personel & SGK Yükü', 'col2': '-₺1.450.000', 'col3': '%15.8', 'status': 'BLOKE EDİLDİ', 'statusType': 'normal'},
            {'col1': 'Tedarikçi / Çek Çıkışı', 'col2': '-₺940.000', 'col3': '%10.2', 'status': 'ÖDEME ONAYLI', 'statusType': 'normal'},
            {'col1': 'KDV & Vergi Karşılığı', 'col2': '-₺420.000', 'col3': '%4.6', 'status': '✓ GÜVENLİ', 'statusType': 'good'},
            {'col1': 'Net Serbest Likidite', 'col2': '₺2.420.000', 'col3': '%100', 'status': '★ POZİTİF', 'statusType': 'good'}
        ]
        chart_type = 'waterfall' if index % 2 == 0 else 'scurve'
        metric_title = 'DİNAMİK NAKİT DÖNGÜSÜ'
        metric_val = '+₺2.42M NAKİT'
        sub_text = '13 Haftalık güvenli likidite kalkanı devrede.'

    elif is_vergi:
        kpis = [
            {'label': 'İncelenen Yasal Matrah', 'value': '₺32.800.000'},
            {'label': 'Mevzuat Uyum Skoru', 'value': '100 / 100'},
            {'label': 'Sağlanan Vergi Tasarrufu', 'value': '₺4.850.000'}
        ]
        table_rows = [
            {'col1': 'VUK / KVK İndirimi', 'col2': '₺14.200.000', 'col3': '%43.3', 'status': '✓ DOĞRULANDI', 'statusType': 'good'},
            {'col1': 'Yeniden Değerleme', 'col2': '₺8.400.000', 'col3': '%25.6', 'status': '✓ KANIT HAZIR', 'statusType': 'good'},
            {'col1': 'SGK Teşvik Avantajı', 'col2': '₺1.850.000', 'col3': '%5.6', 'status': '✓ UYGULANDI', 'statusType': 'good'},
            {'col1': 'KKEG Dengelemesi', 'col2': '-₺920.000', 'col3': '%2.8', 'status': 'EŞİK ALTINDA', 'statusType': 'normal'},
            {'col1': 'Geçici Vergi Mahsubu', 'col2': '₺2.450.000', 'col3': '%7.5', 'status': '✓ TAHAKKUK OK', 'statusType': 'good'},
            {'col1': 'Net Vergi Kalkanı', 'col2': '₺4.850.000', 'col3': '%100', 'status': '★ ONAYLANDI', 'statusType': 'good'}
        ]
        chart_type = 'gauge' if index % 2 == 0 else 'stacked'
        metric_title = 'YASAL UYUM & VERGİ GÜCÜ'
        metric_val = '%100 TAM UYUM'
        sub_text = 'Mevzuat değişikliği ve ceza riskine karşı sıfır hata.'

    elif is_maliyet:
        kpis = [
            {'label': 'Toplam Üretim / Reçete Maliyeti', 'value': '₺184.50 / Birim'},
            {'label': 'Önlenen Fire & Kayıp Oranı', 'value': '%14.2 Tasarruf'},
            {'label': 'Önerilen Satış Kâr Marjı', 'value': '%38.4 Net Marj'}
        ]
        table_rows = [
            {'col1': 'Doğrudan Hammadde', 'col2': '₺92.40', 'col3': '%50.1', 'status': '✓ BİRİM LİSTE', 'statusType': 'good'},
            {'col1': 'İşçilik & Adam/Saat', 'col2': '₺38.20', 'col3': '%20.7', 'status': 'OPTİMİZE', 'statusType': 'normal'},
            {'col1': 'Genel Üretim / Enerji', 'col2': '₺18.50', 'col3': '%10.0', 'status': 'ÖLÇÜLDÜ', 'statusType': 'normal'},
            {'col1': 'Lojistik & Navlun', 'col2': '₺14.80', 'col3': '%8.0', 'status': 'SÖZLEŞMELİ', 'statusType': 'good'},
            {'col1': 'Fire & Iskarta Payı', 'col2': '₺4.20', 'col3': '%2.3', 'status': 'MİNİMUMDA', 'statusType': 'good'},
            {'col1': 'Önerilen Satış Fiyatı', 'col2': '₺298.00', 'col3': '%100', 'status': '★ HEDEF KÂR', 'statusType': 'good'}
        ]
        chart_type = 'stacked' if index % 2 == 0 else 'donut'
        metric_title = 'BİRİM MALİYET DAĞILIMI'
        metric_val = '₺184.50 BİRİM'
        sub_text = 'Hammadde zamları anında satış fiyatına yansır.'

    elif is_ihale:
        kpis = [
            {'label': 'Yaklaşık Maliyet / Sözleşme', 'value': '₺24.800.000'},
            {'label': 'KİK Sınır Değer Eşiği', 'value': '₺18.420.000'},
            {'label': 'Optimum Kazanma Teklifi', 'value': '₺19.150.000'}
        ]
        table_rows = [
            {'col1': 'İşçilik Analiz Payı', 'col2': '₺6.800.000', 'col3': '%35.5', 'status': '✓ RESMİ RAYİÇ', 'statusType': 'good'},
            {'col1': 'Demir / Çelik Girdisi', 'col2': '₺5.400.000', 'col3': '%28.2', 'status': 'FABRİKA TEKLİF', 'statusType': 'good'},
            {'col1': 'Beton / Çimento', 'col2': '₺3.900.000', 'col3': '%20.4', 'status': 'PROFORMA OK', 'statusType': 'good'},
            {'col1': 'Akaryakıt & Nakliye', 'col2': '₺1.850.000', 'col3': '%9.7', 'status': 'ENDEKSLİ', 'statusType': 'normal'},
            {'col1': 'Müteahhit Kârı', 'col2': '₺1.200.000', 'col3': '%6.2', 'status': 'KORUNDU', 'statusType': 'good'},
            {'col1': 'Nihai Teklif Tutarı', 'col2': '₺19.150.000', 'col3': '%100', 'status': '★ ELENMEZ', 'statusType': 'good'}
        ]
        chart_type = 'gauge' if index % 2 == 0 else 'waterfall'
        metric_title = 'SINIR DEĞER & RİSK SKORU'
        metric_val = '₺19.15M TEKLİF'
        sub_text = 'KİK savunma robotu ile ihale iptal riski sıfır.'

    else: # is_yatirim & genel analiz
        kpis = [
            {'label': 'Toplam Yatırım Bütçesi', 'value': '₺42.500.000'},
            {'label': 'İç Verim Oranı (IRR)', 'value': '%38.6 Yıllık'},
            {'label': 'Yatırım Geri Dönüşü (Payback)', 'value': '14.2 Ay'}
        ]
        table_rows = [
            {'col1': 'Kurulum & Donanım', 'col2': '₺24.000.000', 'col3': '%56.5', 'status': '✓ SABİT KIYMET', 'statusType': 'good'},
            {'col1': 'İşletme Sermayesi', 'col2': '₺8.500.000', 'col3': '%20.0', 'status': 'AYRILDI', 'statusType': 'normal'},
            {'col1': 'Yıllık Brüt Gelir', 'col2': '₺38.000.000', 'col3': '%89.4', 'status': 'FİZİBİL', 'statusType': 'good'},
            {'col1': 'İşletme Gideri (OPEX)', 'col2': '-₺14.200.000', 'col3': '%33.4', 'status': 'KONTROLDE', 'statusType': 'normal'},
            {'col1': 'Finansman & Kredi', 'col2': '-₺4.800.000', 'col3': '%11.3', 'status': 'DSCR: 1.48x', 'statusType': 'good'},
            {'col1': 'Net Serbest Nakit', 'col2': '₺14.200.000', 'col3': '%100', 'status': '★ YÜKSEK GETİRİ', 'statusType': 'good'}
        ]
        chart_type = 'scurve' if index % 2 == 0 else 'donut'
        metric_title = 'YATIRIM GETİRİ PROJEKSİYONU'
        metric_val = '%38.6 IRR'
        sub_text = 'Dinamik duyarlılık analizi ve 3 senaryo dahil.'

    # Sol Tablo Satırları (495px genişlik, kusursuz hizalama)
    rows_svg = ''
    for idx, r in enumerate(table_rows):
        y = 375 + idx * 48
        bg = '#ffffff' if idx % 2 == 0 else '#f8fafc'
        status_color = '#107c41' if r['statusType'] == 'good' else ('#dc2626' if r['statusType'] == 'bad' else '#2563eb')
        status_bg = '#dcfce7' if r['statusType'] == 'good' else ('#fee2e2' if r['statusType'] == 'bad' else '#dbeafe')

        rows_svg += f"""
        <rect x="65" y="{y}" width="495" height="42" rx="8" fill="{bg}" stroke="#f1f5f9" stroke-width="1" />
        <text x="80" y="{y + 26}" font-family="system-ui, -apple-system, sans-serif" font-size="11" font-weight="600" fill="#334155">{escape_xml(r['col1'])}</text>
        <text x="310" y="{y + 26}" font-family="system-ui, -apple-system, sans-serif" font-size="12" font-weight="700" fill="#0f172a" text-anchor="end">{escape_xml(r['col2'])}</text>
        <text x="365" y="{y + 26}" font-family="system-ui, -apple-system, sans-serif" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">{escape_xml(r['col3'])}</text>
        
        <rect x="385" y="{y + 10}" width="160" height="22" rx="11" fill="{status_bg}" />
        <text x="465" y="{y + 25}" font-family="system-ui, -apple-system, sans-serif" font-size="10" font-weight="800" fill="{status_color}" text-anchor="middle">{escape_xml(r['status'])}</text>
        """

    # Sağ Grafik
    chart_svg = render_chart(chart_type, metric_title, metric_val, sub_text)

    full_svg = f"""<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="100%" height="100%">
  <defs>
    <filter id="softCard" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="12" stdDeviation="16" flood-color="#0f172a" flood-opacity="0.07" />
    </filter>
  </defs>

  <!-- Saf Beyaz Zemin -->
  <rect width="1000" height="1000" fill="#ffffff" />

  <!-- Genişletilmiş ve Dengeli Kart (%92 Doluluk: 920x920) -->
  <rect x="40" y="40" width="920" height="920" rx="24" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" filter="url(#softCard)" />

  <!-- A) Üst Sistem Çubuğu -->
  <g transform="translate(70, 68)">
    <circle cx="0" cy="0" r="6" fill="#ef4444" />
    <circle cx="18" cy="0" r="6" fill="#f59e0b" />
    <circle cx="36" cy="0" r="6" fill="#107c41" />

    <rect x="65" y="-14" width="140" height="28" rx="6" fill="#107c41" />
    <text x="135" y="4" font-family="system-ui" font-size="10" font-weight="800" fill="#ffffff" text-anchor="middle" letter-spacing="0.5">EXCEL PRO ENGINE</text>

    <rect x="215" y="-14" width="160" height="28" rx="6" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="1" />
    <text x="295" y="4" font-family="system-ui" font-size="10" font-weight="700" fill="#475569" text-anchor="middle" letter-spacing="0.5">{safe_cat}</text>

    <text x="860" y="4" font-family="system-ui" font-size="11" font-weight="700" fill="#94a3b8" text-anchor="end">DECISION OS v15.1 · LOKAL ÇALIŞIR</text>
  </g>

  <!-- B) Başlık ve Formül Çubuğu -->
  <g transform="translate(70, 130)">
    <text x="0" y="0" font-family="system-ui, -apple-system, sans-serif" font-size="25" font-weight="900" fill="#0f172a">{safe_title}</text>
    <rect x="0" y="16" width="860" height="34" rx="6" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
    <text x="14" y="38" font-family="system-ui, monospace" font-size="12" font-weight="800" fill="#107c41">fx</text>
    <line x1="36" y1="22" x2="36" y2="44" stroke="#cbd5e1" stroke-width="1.5" />
    <text x="46" y="38" font-family="system-ui, monospace" font-size="11" font-weight="600" fill="#475569">=KARAR_OPTIMIZASYONU(VeriGirisleri!A2:E50; Senaryo=&quot;Dinamik&quot;; Eşik=Otomatik)</text>
  </g>

  <!-- C) 3'lü Üst KPI Vitrin Kartları -->
  <g transform="translate(70, 205)">
    <!-- KPI 1 -->
    <rect x="0" y="0" width="270" height="85" rx="12" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
    <text x="20" y="28" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">{escape_xml(kpis[0]['label'])}</text>
    <text x="20" y="62" font-family="system-ui" font-size="24" font-weight="900" fill="#0f172a">{escape_xml(kpis[0]['value'])}</text>

    <!-- KPI 2 -->
    <rect x="295" y="0" width="270" height="85" rx="12" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
    <text x="315" y="28" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">{escape_xml(kpis[1]['label'])}</text>
    <text x="315" y="62" font-family="system-ui" font-size="24" font-weight="900" fill="#2563eb">{escape_xml(kpis[1]['value'])}</text>

    <!-- KPI 3 -->
    <rect x="590" y="0" width="270" height="85" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
    <text x="610" y="28" font-family="system-ui" font-size="11" font-weight="800" fill="#0d5c3a">{escape_xml(kpis[2]['label'])}</text>
    <text x="610" y="62" font-family="system-ui" font-size="24" font-weight="900" fill="#0d5c3a">{escape_xml(kpis[2]['value'])}</text>
  </g>

  <!-- D) Ana Gövde: Sol Analitik Tablo + Sağ Dinamik Görsel -->
  <!-- Sol Tablo Başlığı -->
  <rect x="65" y="325" width="495" height="38" rx="8" fill="#1e293b" />
  <text x="80" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff">PARAMETRE / KALEM</text>
  <text x="310" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="end">TUTAR / DEĞER</text>
  <text x="365" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="end">PAY (%)</text>
  <text x="465" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="middle">DURUM</text>

  <!-- Sol Tablo Satırları -->
  {rows_svg}

  <!-- Sağ Grafik (x: 570, y: 325, w: 360, h: 430) -->
  <g transform="translate(570, 325)">
    {chart_svg}
  </g>

  <!-- E) Alt Güvenlik ve Denetim Çubuğu -->
  <g transform="translate(70, 920)">
    <line x1="0" y1="0" x2="860" y2="0" stroke="#e2e8f0" stroke-width="1.5" />
    <circle cx="10" cy="18" r="4" fill="#107c41" />
    <text x="22" y="22" font-family="system-ui" font-size="11" font-weight="700" fill="#0d5c3a">Tam Otomatik Özet Pano &amp; Karar Raporu</text>
    <text x="860" y="22" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">Formüller Açık · Dilediğiniz Gibi Özelleştirilebilir</text>
  </g>

</svg>"""
    return full_svg

# Tüm 124 SVG dosyasını oluştur
out_dir = 'public/images/kapak'
os.makedirs(out_dir, exist_ok=True)

success = 0
for idx, p in enumerate(products):
    svg_content = generate_cockpit_svg(p, idx)
    target_file = os.path.join(out_dir, f"{p['slug']}.svg")
    with open(target_file, 'w', encoding='utf-8') as f:
        f.write(svg_content)
    success += 1

print(f"BAŞARILI: {success} / {len(products)} adet 1000x1000 (920x920 doluluk) Enterprise Cockpit kapak SVG'si üretildi!")
