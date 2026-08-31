// components/mobile/MobileCommerceEngine.tsx
'use client';

import React, { useState } from 'react';

// --- TİP TANIMLARI ---
export type TabType = 'home' | 'sablonlar' | 'ozel-sistemler';

export interface TemplateItem {
  id: string;
  title: string;
  category: string;
  badgeColor: string;
  authorityBadge: string;
  painPoint: string;
  solutionMetric: string;
  price: number;
  originalPrice: number;
  rating: number;
  salesCount: number;
  tags: string[];
  shopierUrl?: string;
  slug?: string;
}

interface MobileCommerceEngineProps {
  initialTab?: TabType;
  whatsappNumber?: string; // Örn: "905393333303"
}

// --- VERİ MODELİ ---
export const TEMPLATE_DATA: TemplateItem[] = [
  {
    id: 'nakit-akim-13-hafta',
    slug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
    title: '13 Haftalık Dinamik Nakit Akım & Likidite Modeli',
    category: 'Finans & Nakit Yönetimi',
    badgeColor: 'bg-blue-600 text-white',
    authorityBadge: '17Y Banka Kredi Standardı',
    painPoint: 'Nakit tıkanıklığını ve çek açıklarını geç fark etme riski',
    solutionMetric: '13 Hafta Önceden Kesintisiz Kasa & Çek Hakimiyeti',
    price: 649,
    originalPrice: 1299,
    rating: 4.9,
    salesCount: 342,
    tags: ['Makrosuz & Hızlı', 'Excel + Sheets', 'DSCR Banka Rasyolu'],
    shopierUrl: 'https://www.shopier.com/49653399'
  },
  {
    id: 'kobi-gelir-gider-dashboard',
    slug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi',
    title: 'KOBİ Otomatik Finansal Kontrol & Dashboard',
    category: 'Yönetim Raporlaması',
    badgeColor: 'bg-emerald-600 text-white',
    authorityBadge: 'Patron & Yönetici Formatı',
    painPoint: 'Ay sonu görünmeyen gider kaçakları ve kâr belirsizliği',
    solutionMetric: 'Tek Ekranda Anlık Kâr/Zarar ve Net Durum Raporu',
    price: 499,
    originalPrice: 950,
    rating: 4.8,
    salesCount: 518,
    tags: ['Formülleri Kilitli', 'Otomatik Grafik', 'Sıfır Hata'],
    shopierUrl: 'https://www.shopier.com/49652321'
  },
  {
    id: 'uretim-maliyet-hesaplama',
    slug: 'uretim-recetesi-ve-zam-yansitma-hesaplayici',
    title: 'Birim Maliyet & SKDM Karbon Uyumlu Üretim Matrisi',
    category: 'Maliyet & Operasyon',
    badgeColor: 'bg-amber-600 text-white',
    authorityBadge: 'Sanayi & İhracat Standardı',
    painPoint: 'Hammadde dalgalanmasında eksik fiyat verip zararına satış',
    solutionMetric: 'Kuruş Kuruş Net Maliyet ve Başabaş Fiyatlama',
    price: 799,
    originalPrice: 1500,
    rating: 5.0,
    salesCount: 189,
    tags: ['Endüstriyel Reçete', 'Fire Hesaplama', 'SKDM Uyumlu'],
    shopierUrl: 'https://www.shopier.com/49652403'
  }
];

// --- SAF SVG İKONLAR (Sıfır Paket Bağımlılığı & Ultra Hızlı Render) ---
const SparklesIcon = () => (
  <svg className="w-4 h-4 text-amber-500 shrink-0" viewBox="0 0 24 24" fill="currentColor">
    <path d="M12 2L14.4 8.6L21 11L14.4 13.4L12 20L9.6 13.4L3 11L9.6 8.6L12 2Z" />
  </svg>
);

const CheckCircleIcon = () => (
  <svg className="w-4 h-4 text-emerald-600 shrink-0" viewBox="0 0 24 24" fill="currentColor">
    <path fillRule="evenodd" d="M2.25 12c0-5.385 4.365-9.75 9.75-9.75s9.75 4.365 9.75 9.75-4.365 9.75-9.75 9.75S2.25 17.385 2.25 12zm13.36-1.814a.75.75 0 10-1.22-.872l-3.236 4.53L9.53 12.22a.75.75 0 00-1.06 1.06l2.25 2.25a.75.75 0 001.14-.094l3.75-5.25z" clipRule="evenodd" />
  </svg>
);

const WhatsAppIcon = () => (
  <svg className="w-4 h-4 fill-current shrink-0" viewBox="0 0 24 24">
    <path d="M.057 24l1.687-6.163c-1.041-1.804-1.588-3.849-1.587-5.946.003-6.556 5.338-11.891 11.893-11.891 3.181.001 6.167 1.24 8.413 3.488 2.245 2.248 3.481 5.236 3.48 8.414-.003 6.557-5.338 11.892-11.893 11.892-1.99-.001-3.951-.5-5.688-1.448l-6.305 1.654zm6.597-3.807c1.676.995 3.276 1.591 5.392 1.592 5.448 0 9.886-4.434 9.889-9.885.002-5.462-4.415-9.89-9.881-9.892-5.452 0-9.887 4.434-9.889 9.884-.001 2.225.651 3.891 1.746 5.634l-.999 3.648 3.742-.981z" />
  </svg>
);

const SpreadsheetIcon = () => (
  <svg className="w-5 h-5 text-emerald-700" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="8" y1="13" x2="16" y2="13" /><line x1="8" y1="17" x2="16" y2="17" /><line x1="10" y1="9" x2="8" y2="9" />
  </svg>
);

const DownloadIcon = () => (
  <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" />
  </svg>
);

const StarIcon = () => (
  <svg className="w-3.5 h-3.5 fill-amber-400 text-amber-400 shrink-0" viewBox="0 0 24 24">
    <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" />
  </svg>
);

export default function MobileCommerceEngine({
  initialTab = 'home',
  whatsappNumber = '905393333303'
}: MobileCommerceEngineProps) {
  const [activeTab, setActiveTab] = useState<TabType>(initialTab);
  const [selectedTemplate, setSelectedTemplate] = useState<TemplateItem | null>(null);
  const [customStep, setCustomStep] = useState<number>(1);
  const [customData, setCustomData] = useState({
    needType: 'Finansal Modelleme & Nakit Akış',
    dataVolume: 'Orta Ölçek (1.000 - 50.000 Satır/Ay)',
    urgency: 'Bu Hafta İçinde'
  });
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const showNotification = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 2800);
  };

  return (
    <div className="w-full max-w-md mx-auto min-h-screen bg-white shadow-2xl border-x border-slate-200 relative pb-28 font-sans antialiased text-slate-900">
      
      {/* MİKRO ÖDÜL VE GERİ BİLDİRİM TOASTI */}
      {toastMessage && (
        <div className="fixed top-3 inset-x-4 max-w-sm mx-auto z-50 bg-slate-950 text-white px-3.5 py-2.5 rounded-2xl shadow-2xl border border-emerald-500/60 flex items-center gap-2.5 transition-all">
          <span className="w-5 h-5 rounded-full bg-emerald-500 flex items-center justify-center text-slate-950 font-black text-xs shrink-0">✓</span>
          <p className="text-xs font-bold leading-tight">{toastMessage}</p>
        </div>
      )}

      {/* MOBİL HEADER (Canlı Otorite Vurgusu) */}
      <header className="sticky top-0 z-40 bg-white/95 backdrop-blur-md border-b border-slate-100 px-4 py-2.5 shadow-xs">
        <div className="flex items-center justify-between mb-2">
          <div className="flex items-center gap-2">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-tr from-emerald-600 via-teal-600 to-emerald-400 flex items-center justify-center text-white font-black text-sm shadow-md shadow-emerald-600/20">
              EX
            </div>
            <div>
              <div className="flex items-center gap-1.5">
                <span className="text-sm font-extrabold tracking-tight text-slate-900 leading-none">
                  excelarsiv<span className="text-emerald-600">.com</span>
                </span>
                <span className="bg-emerald-100 text-emerald-800 text-[9px] font-black px-1.5 py-0.5 rounded-full">
                  PRO
                </span>
              </div>
              <span className="text-[10px] text-slate-500 font-bold">17 Yıllık Bankacılık Standartlarında</span>
            </div>
          </div>

          <a
            href={`https://wa.me/${whatsappNumber}?text=${encodeURIComponent('Merhaba, excelarsiv.com üzerinden danışmanlık almak istiyorum.')}`}
            target="_blank"
            rel="noopener noreferrer"
            onClick={() => showNotification('Finans uzmanı hattına bağlanılıyor...')}
            className="flex items-center gap-1.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white px-3 py-1.5 rounded-full text-xs font-black shadow-md shadow-emerald-600/25 active:scale-95 transition-transform"
          >
            <WhatsAppIcon />
            <span>Canlı Destek</span>
          </a>
        </div>

        {/* 3'LÜ MOBİL NAVİGASYON SEGMENTİ */}
        <nav className="grid grid-cols-3 gap-1 bg-slate-100 p-1 rounded-2xl border border-slate-200/80">
          <button
            onClick={() => {
              setActiveTab('home');
              showNotification('Ana Sayfa');
            }}
            className={`py-2 text-xs font-black rounded-xl transition-all ${
              activeTab === 'home'
                ? 'bg-gradient-to-r from-blue-600 to-indigo-600 text-white shadow-md shadow-blue-500/30 scale-[1.02]'
                : 'text-slate-600 hover:text-slate-900 font-bold'
            }`}
          >
            Ana Sayfa
          </button>
          <button
            onClick={() => {
              setActiveTab('sablonlar');
              showNotification('Hazır şablonlar açıldı');
            }}
            className={`py-2 text-xs font-black rounded-xl transition-all ${
              activeTab === 'sablonlar'
                ? 'bg-gradient-to-r from-emerald-600 to-teal-600 text-white shadow-md shadow-emerald-500/30 scale-[1.02]'
                : 'text-slate-600 hover:text-slate-900 font-bold'
            }`}
          >
            Şablonlar
          </button>
          <button
            onClick={() => {
              setActiveTab('ozel-sistemler');
              showNotification('Özel sistem konfigüratörü aktif');
            }}
            className={`py-2 text-xs font-black rounded-xl transition-all ${
              activeTab === 'ozel-sistemler'
                ? 'bg-gradient-to-r from-amber-500 to-orange-600 text-white shadow-md shadow-amber-500/30 scale-[1.02]'
                : 'text-slate-600 hover:text-slate-900 font-bold'
            }`}
          >
            Özel Sistem
          </button>
        </nav>
      </header>

      {/* 1. SEKME: ANA SAYFA */}
      {activeTab === 'home' && (
        <main className="p-4 space-y-4">
          <section className="bg-gradient-to-br from-blue-700 via-indigo-800 to-slate-900 p-4 rounded-3xl text-white shadow-xl shadow-indigo-950/20 relative overflow-hidden">
            <div className="inline-flex items-center gap-1.5 bg-amber-400 text-slate-950 px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wide mb-2 shadow-xs">
              <SparklesIcon />
              <span>Nakit ve Maliyet Riskini Sıfırla</span>
            </div>

            <h1 className="text-xl font-black tracking-tight leading-tight">
              Şirketinizin Finansını <span className="text-emerald-400 underline decoration-amber-400">1 Tıkla</span> Güvene Alın.
            </h1>

            <p className="text-xs text-blue-100 mt-2 leading-relaxed font-medium">
              Rastgele formüllerle vakit kaybetmeyin. 17 yıllık ticari bankacılık ve finans mühendisliği tecrübesiyle hatasız kodlandı.
            </p>

            <div className="grid grid-cols-3 gap-2 mt-4 pt-3 border-t border-white/10">
              <div className="bg-white/10 backdrop-blur-sm p-2 rounded-2xl text-center border border-white/15">
                <span className="block text-sm font-black text-amber-300">2.400+</span>
                <span className="text-[9px] text-blue-100 font-bold">KOBİ Patronu</span>
              </div>
              <div className="bg-white/10 backdrop-blur-sm p-2 rounded-2xl text-center border border-white/15">
                <span className="block text-sm font-black text-emerald-400">%100</span>
                <span className="text-[9px] text-blue-100 font-bold">Banka Uyumlu</span>
              </div>
              <div className="bg-white/10 backdrop-blur-sm p-2 rounded-2xl text-center border border-white/15">
                <span className="block text-sm font-black text-white">4.9/5</span>
                <span className="text-[9px] text-blue-100 font-bold">Müşteri Puanı</span>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2 mt-4">
              <button
                onClick={() => setActiveTab('sablonlar')}
                className="bg-emerald-400 hover:bg-emerald-300 text-slate-950 font-black py-2.5 px-3 rounded-2xl text-xs flex items-center justify-center gap-1 shadow-lg shadow-emerald-500/30 active:scale-95 transition-transform"
              >
                <span>Şablonu Seç</span>
                <span>→</span>
              </button>
              <button
                onClick={() => setActiveTab('ozel-sistemler')}
                className="bg-white/15 hover:bg-white/25 text-white font-bold py-2.5 px-3 rounded-2xl text-xs flex items-center justify-center border border-white/20 transition-colors"
              >
                <span>Özel Çözüm İste</span>
              </button>
            </div>
          </section>

          {/* Otorite & Saha Mesajı */}
          <section className="bg-gradient-to-r from-amber-50 via-white to-amber-50/50 border-2 border-amber-300/70 p-3.5 rounded-3xl flex items-start gap-3 shadow-xs">
            <div className="w-10 h-10 rounded-2xl bg-gradient-to-br from-amber-500 to-orange-600 text-white flex items-center justify-center font-black text-xs shrink-0 shadow-md shadow-amber-500/30">
              17 YIL
            </div>
            <div>
              <h4 className="text-xs font-black text-slate-900 leading-tight">Esnaf ve Sanayicinin Diliyle Çözüm</h4>
              <p className="text-[11px] text-slate-700 mt-1 leading-snug font-medium">
                "Nakit sıkışıklığını, banka kredi risk rasyolarını ve esnafın dertlerini bizzat sahada yaşayarak çözdük. Bu modeller şirketinizin gerçek saha zırhıdır."
              </p>
            </div>
          </section>

          {/* Hızlı Çözüm Listesi */}
          <section className="space-y-2.5">
            <div className="flex items-center justify-between px-1">
              <span className="text-xs font-black text-slate-700 uppercase tracking-wider">Hemen İndirilebilir Modeller</span>
              <span className="text-[11px] text-emerald-600 font-black">Canlı Önizle →</span>
            </div>

            {TEMPLATE_DATA.map((item) => (
              <div
                key={item.id}
                onClick={() => {
                  setSelectedTemplate(item);
                  showNotification(`${item.title} seçildi ✓`);
                }}
                className="bg-white border-2 border-slate-100 hover:border-emerald-500 p-3.5 rounded-2xl flex items-center justify-between gap-3 shadow-md hover:shadow-xl hover:shadow-emerald-500/10 active:scale-[0.98] transition-all cursor-pointer"
              >
                <div className="flex items-start gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-emerald-50 border border-emerald-200 flex items-center justify-center shrink-0 mt-0.5">
                    <SpreadsheetIcon />
                  </div>
                  <div>
                    <span className={`text-[9px] font-black px-2 py-0.5 rounded-full inline-block mb-1 ${item.badgeColor}`}>
                      {item.authorityBadge}
                    </span>
                    <h3 className="text-xs font-black text-slate-900 leading-snug">{item.title}</h3>
                    <p className="text-[11px] text-emerald-700 font-bold mt-0.5 flex items-center gap-1">
                      <CheckCircleIcon />
                      <span>{item.solutionMetric}</span>
                    </p>
                  </div>
                </div>

                <div className="text-right shrink-0">
                  <span className="block text-sm font-black text-emerald-600">₺{item.price}</span>
                  <span className="block text-[10px] text-slate-400 line-through font-semibold">₺{item.originalPrice}</span>
                </div>
              </div>
            ))}
          </section>
        </main>
      )}

      {/* 2. SEKME: ŞABLONLAR */}
      {activeTab === 'sablonlar' && (
        <main className="p-4 space-y-3.5">
          <div className="flex gap-1.5 overflow-x-auto pb-1 no-scrollbar text-xs">
            {['Tüm Sistemler', 'Finans & Nakit', 'KOBİ Yönetim', 'Maliyet & Sanayi'].map((cat, idx) => (
              <button
                key={cat}
                onClick={() => showNotification(`${cat} listelendi`)}
                className={`px-3.5 py-1.5 rounded-full whitespace-nowrap text-xs font-black transition-all ${
                  idx === 0
                    ? 'bg-gradient-to-r from-emerald-600 to-teal-600 text-white shadow-md shadow-emerald-500/30 scale-105'
                    : 'bg-white border-2 border-slate-200 text-slate-700 hover:border-slate-300'
                }`}
              >
                {cat}
              </button>
            ))}
          </div>

          {TEMPLATE_DATA.map((item) => (
            <article
              key={item.id}
              className="bg-white border-2 border-slate-200/90 rounded-3xl overflow-hidden shadow-lg hover:border-emerald-500 transition-all"
            >
              <div className="p-4 border-b border-slate-100 flex items-start justify-between gap-2">
                <div>
                  <span className="text-[10px] font-black text-indigo-700 bg-indigo-50 border border-indigo-200 px-2 py-0.5 rounded-md uppercase">
                    {item.category}
                  </span>
                  <h3 className="text-sm font-black text-slate-900 leading-snug mt-1.5">{item.title}</h3>
                </div>
                <div className="flex items-center gap-1 bg-amber-50 border border-amber-300 text-amber-900 px-2 py-1 rounded-xl text-xs font-black shrink-0">
                  <StarIcon />
                  <span>{item.rating}</span>
                </div>
              </div>

              <div className="p-4 bg-gradient-to-b from-slate-50 to-white text-xs space-y-2 border-b border-slate-100">
                <div className="text-rose-900 flex items-center gap-2 bg-rose-50 p-2 rounded-xl border border-rose-200">
                  <span className="w-2 h-2 rounded-full bg-rose-500 shrink-0" />
                  <span className="text-[11px] font-bold leading-tight"><strong>Önlenecek Risk:</strong> {item.painPoint}</span>
                </div>
                <div className="text-emerald-900 flex items-center gap-2 bg-emerald-50 p-2 rounded-xl border border-emerald-200">
                  <span className="w-2 h-2 rounded-full bg-emerald-500 shrink-0" />
                  <span className="text-[11px] font-bold leading-tight"><strong>Sağlanan Değer:</strong> {item.solutionMetric}</span>
                </div>
                <div className="flex flex-wrap gap-1.5 pt-1">
                  {item.tags.map(t => (
                    <span key={t} className="text-[10px] bg-white border border-slate-300 text-slate-800 font-bold px-2 py-0.5 rounded-lg">
                      {t}
                    </span>
                  ))}
                </div>
              </div>

              <div className="p-4 bg-white flex items-center justify-between">
                <div>
                  <div className="flex items-baseline gap-1.5">
                    <span className="text-lg font-black text-slate-900">₺{item.price}</span>
                    <span className="text-xs text-slate-400 line-through font-bold">₺{item.originalPrice}</span>
                  </div>
                  <span className="text-[10px] text-emerald-600 font-black flex items-center gap-1">
                    <SparklesIcon />
                    Anında İndir & Hemen Kullan
                  </span>
                </div>

                <div className="flex gap-2">
                  <button
                    onClick={() => {
                      setSelectedTemplate(item);
                      showNotification(`${item.title} önizlemesi açıldı`);
                    }}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-black px-3.5 py-2.5 rounded-2xl transition-colors"
                  >
                    Önizle
                  </button>
                  <a
                    href={item.shopierUrl || `https://www.shopier.com/49652321`}
                    target="_blank"
                    rel="noopener noreferrer"
                    onClick={() => showNotification('Ödeme sayfasına yönlendiriliyorsunuz...')}
                    className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white text-xs font-black px-4 py-2.5 rounded-2xl flex items-center gap-1.5 shadow-lg shadow-emerald-500/30 active:scale-95 transition-all"
                  >
                    <DownloadIcon />
                    <span>Hemen Al</span>
                  </a>
                </div>
              </div>
            </article>
          ))}
        </main>
      )}

      {/* 3. SEKME: ÖZEL SİSTEMLER */}
      {activeTab === 'ozel-sistemler' && (
        <main className="p-4 space-y-4">
          <section className="bg-gradient-to-br from-amber-500 via-orange-600 to-amber-700 p-4 rounded-3xl text-white shadow-xl shadow-amber-600/20">
            <span className="text-[10px] font-black bg-white/20 border border-white/30 text-white px-2.5 py-0.5 rounded-full uppercase tracking-wider">
              Terzi Usulü Finans & Yazılım Mimarisi
            </span>
            <h2 className="text-base font-black text-white mt-1.5 leading-snug">
              Şirketinizin Tam İhtiyacına Özel Excel & VBA Modeli
            </h2>
            <p className="text-xs text-amber-100 mt-1.5 leading-relaxed font-medium">
              Pahalı ERP paketlerine gerek kalmadan; 17 yıllık bankacılık tecrübesiyle tek tuşla çalışan sisteminizi kuralım.
            </p>
          </section>

          <section className="bg-white border-2 border-slate-200 p-4 rounded-3xl shadow-lg space-y-3.5">
            <div className="flex items-center justify-between pb-2 border-b border-slate-100">
              <span className="text-xs font-black text-slate-900">İhtiyaç Konfigüratörü</span>
              <span className="text-[11px] font-black text-amber-900 bg-amber-100 border border-amber-300 px-2.5 py-0.5 rounded-full">
                Adım {customStep} / 3
              </span>
            </div>

            {customStep === 1 && (
              <div className="space-y-2">
                <span className="text-xs text-slate-800 font-black block">1. İşletmenizin ana ihtiyacı nedir?</span>
                {[
                  'Nakit Akış, Kredi & Banka Risk Yönetimi',
                  'Üretim Reçetesi & SKDM Birim Maliyet',
                  'KOBİ Gelir-Gider & Satış Dashboard',
                  'Özel VBA Makro / Otomasyon / API'
                ].map((item) => (
                  <button
                    key={item}
                    onClick={() => {
                      setCustomData({ ...customData, needType: item });
                      setCustomStep(2);
                      showNotification(`Seçildi: ${item}`);
                    }}
                    className={`w-full p-3.5 text-xs text-left rounded-2xl border-2 flex items-center justify-between transition-all ${
                      customData.needType === item
                        ? 'bg-amber-50 border-amber-500 text-amber-950 font-black shadow-md'
                        : 'bg-slate-50 border-slate-200 text-slate-700 hover:border-slate-300 font-bold'
                    }`}
                  >
                    <span>{item}</span>
                    <span>→</span>
                  </button>
                ))}
              </div>
            )}

            {customStep === 2 && (
              <div className="space-y-2">
                <span className="text-xs text-slate-800 font-black block">2. Aylık veri/işlem hacminiz:</span>
                {[
                  'Küçük Ölçek (< 1.000 Satır/Ay)',
                  'Orta Ölçek (1.000 - 50.000 Satır/Ay)',
                  'Büyük Ölçek (50.000+ Satır / Power Query)'
                ].map((vol) => (
                  <button
                    key={vol}
                    onClick={() => {
                      setCustomData({ ...customData, dataVolume: vol });
                      setCustomStep(3);
                      showNotification(`Hacim: ${vol}`);
                    }}
                    className={`w-full p-3.5 text-xs text-left rounded-2xl border-2 flex items-center justify-between transition-all ${
                      customData.dataVolume === vol
                        ? 'bg-amber-50 border-amber-500 text-amber-950 font-black shadow-md'
                        : 'bg-slate-50 border-slate-200 text-slate-700 hover:border-slate-300 font-bold'
                    }`}
                  >
                    <span>{vol}</span>
                    <span>→</span>
                  </button>
                ))}
                <button
                  onClick={() => setCustomStep(1)}
                  className="text-xs text-slate-500 hover:text-slate-900 font-bold pt-1 block"
                >
                  ← Önceki Adıma Dön
                </button>
              </div>
            )}

            {customStep === 3 && (
              <div className="space-y-3.5">
                <span className="text-xs text-slate-800 font-black block">3. Proje teslim önceliğiniz:</span>
                <div className="grid grid-cols-2 gap-2">
                  {['24-48 Saat (Acil)', '3-7 İş Günü', '15 Gün İçinde', 'Aylık Danışmanlık'].map((urg) => (
                    <button
                      key={urg}
                      onClick={() => {
                        setCustomData({ ...customData, urgency: urg });
                        showNotification(`Süre: ${urg}`);
                      }}
                      className={`p-3 text-xs text-center rounded-2xl border-2 font-black transition-all ${
                        customData.urgency === urg
                          ? 'bg-amber-500 border-amber-600 text-white shadow-md shadow-amber-500/30'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
                      }`}
                    >
                      {urg}
                    </button>
                  ))}
                </div>

                <div className="bg-slate-50 p-3.5 rounded-2xl border-2 border-slate-200 text-xs">
                  <span className="text-[10px] text-slate-500 font-black uppercase tracking-wider block">Özet Talep</span>
                  <div className="font-black text-slate-900 mt-1">{customData.needType}</div>
                  <div className="text-slate-600 text-[11px] font-bold mt-0.5">{customData.dataVolume} • {customData.urgency}</div>
                </div>

                <a
                  href={`https://wa.me/${whatsappNumber}?text=${encodeURIComponent(
                    `Merhaba, excelarsiv.com üzerinden özel sistem talep ediyorum:\n- Kapsam: ${customData.needType}\n- Hacim: ${customData.dataVolume}\n- Süre: ${customData.urgency}\nMimari değerlendirme ve teklif rica ediyorum.`
                  )}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  onClick={() => showNotification('WhatsApp teklif akışı başlatıldı')}
                  className="w-full bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-black py-3.5 rounded-2xl text-xs flex items-center justify-center gap-2 shadow-xl shadow-emerald-600/30 active:scale-95 transition-all"
                >
                  <WhatsAppIcon />
                  <span>Teklifi WhatsApp'tan Anında Al</span>
                </a>

                <button
                  onClick={() => setCustomStep(2)}
                  className="text-xs text-slate-500 block text-center w-full font-bold"
                >
                  ← Seçimleri Düzenle
                </button>
              </div>
            )}
          </section>
        </main>
      )}

      {/* CANLI BOTTOM SHEET MODAL (Önizleme) */}
      {selectedTemplate && (
        <div className="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-sm flex items-end animate-in fade-in duration-200">
          <div className="w-full max-w-md mx-auto bg-white rounded-t-3xl p-4 space-y-3.5 shadow-2xl border-t border-slate-200">
            <div className="flex items-center justify-between pb-1 border-b border-slate-100">
              <div className="w-12 h-1.5 bg-slate-300 rounded-full mx-auto" />
              <button
                onClick={() => setSelectedTemplate(null)}
                className="w-8 h-8 rounded-full bg-slate-100 text-slate-600 hover:text-slate-900 flex items-center justify-center font-bold"
              >
                ✕
              </button>
            </div>

            <div>
              <span className={`text-[10px] font-black px-2 py-0.5 rounded-full ${selectedTemplate.badgeColor}`}>
                {selectedTemplate.authorityBadge}
              </span>
              <h3 className="text-base font-black text-slate-900 mt-1 leading-snug">{selectedTemplate.title}</h3>
            </div>

            {/* Simüle Canlı Excel Dashboard */}
            <div className="bg-slate-900 text-white p-3 rounded-2xl space-y-2 border border-slate-800 shadow-inner">
              <div className="flex justify-between text-[10px] text-slate-300 font-mono">
                <span className="font-bold text-emerald-400">CANLI_SİSTEM_DASHBOARD.XLSX</span>
                <span className="text-amber-400 font-bold">● Korumalı Kod</span>
              </div>
              <div className="grid grid-cols-3 gap-1.5 text-center">
                <div className="bg-slate-800 p-2 rounded-xl border border-slate-700">
                  <span className="text-[9px] text-slate-400 block font-medium">Kasa Pozisyonu</span>
                  <span className="text-xs font-black text-emerald-400">₺1.840.000</span>
                </div>
                <div className="bg-slate-800 p-2 rounded-xl border border-slate-700">
                  <span className="text-[9px] text-slate-400 block font-medium">Banka DSCR</span>
                  <span className="text-xs font-black text-blue-400">1.45x</span>
                </div>
                <div className="bg-slate-800 p-2 rounded-xl border border-slate-700">
                  <span className="text-[9px] text-slate-400 block font-medium">Risk Puanı</span>
                  <span className="text-xs font-black text-amber-400">Güvenli</span>
                </div>
              </div>
            </div>

            <div className="space-y-1.5 text-xs">
              <div className="flex items-center gap-2 text-slate-800 font-bold bg-emerald-50 p-2 rounded-xl border border-emerald-200">
                <CheckCircleIcon />
                <span>17 Yıllık bankacılık ve finans risk denetim standartlarında</span>
              </div>
              <div className="flex items-center gap-2 text-slate-800 font-bold bg-emerald-50 p-2 rounded-xl border border-emerald-200">
                <CheckCircleIcon />
                <span>Ödeme sonrası anında dosya indirme linki ve e-posta teslimatı</span>
              </div>
            </div>

            <div className="flex items-center gap-3 pt-2 border-t border-slate-100">
              <div>
                <span className="block text-lg font-black text-slate-900">₺{selectedTemplate.price}</span>
                <span className="block text-[10px] text-slate-400 line-through font-bold">₺{selectedTemplate.originalPrice}</span>
              </div>
              <a
                href={selectedTemplate.shopierUrl || `https://www.shopier.com/49652321`}
                target="_blank"
                rel="noopener noreferrer"
                onClick={() => showNotification('Ödeme güvenli Shopier altyapısında başlatılıyor...')}
                className="flex-1 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-black py-3 px-4 rounded-2xl text-xs flex items-center justify-center gap-1.5 shadow-xl shadow-emerald-500/40 active:scale-95 transition-all"
              >
                <DownloadIcon />
                <span>Hemen İndir ve Başla</span>
              </a>
            </div>
          </div>
        </div>
      )}

      {/* BAŞPARMAK SABİT DANIŞMA BARI */}
      <nav className="fixed bottom-0 max-w-md w-full z-30 bg-white/95 backdrop-blur-md border-t border-slate-200 px-4 py-2.5 flex items-center justify-between shadow-2xl">
        <div>
          <span className="text-[10px] text-slate-500 block leading-tight font-bold">17 Yıllık Bankacılık Uzmanlığı</span>
          <span className="text-xs font-black text-slate-900">Doğrudan Uzmanla Görüş</span>
        </div>
        <a
          href={`https://wa.me/${whatsappNumber}?text=${encodeURIComponent('Merhaba, Excel sistemleri hakkinda danismak istiyorum.')}`}
          target="_blank"
          rel="noopener noreferrer"
          onClick={() => showNotification('WhatsApp canlı hattı açılıyor...')}
          className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-black text-xs px-4 py-2 rounded-2xl flex items-center gap-1.5 shadow-lg shadow-emerald-600/30 active:scale-95 transition-all"
        >
          <WhatsAppIcon />
          <span>WhatsApp ile Sor</span>
        </a>
      </nav>
    </div>
  );
}
