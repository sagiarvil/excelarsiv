#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'node:fs';

function once(html, from, to, label) {
  const count = html.split(from).length - 1;
  if (count !== 1) throw new Error(`SPECIAL FUNNEL: expected 1 ${label}, found ${count}`);
  return html.replace(from, to);
}

export function applySpecialFunnelOptimization({ specialPath = 'dist/ozel-excel-sistemleri/index.html' } = {}) {
  let html = readFileSync(specialPath, 'utf8');

  const styles = `
<style data-special-funnel-v1>
  .special-v3 .funnel-band{padding:34px 0 8px;background:#fff}
  .special-v3 .funnel-authority{display:grid;grid-template-columns:minmax(0,1fr) minmax(300px,.72fr);gap:18px;align-items:stretch}
  .special-v3 .authority-card,.special-v3 .compat-card,.special-v3 .diagnostic-card,.special-v3 .funnel-card{border:1px solid var(--line);border-radius:18px;background:#fff;box-shadow:0 1px 2px rgba(16,24,40,.03)}
  .special-v3 .authority-card{padding:24px;display:flex;gap:16px;align-items:center;background:linear-gradient(135deg,#0b162a,#111e34);color:#fff}
  .special-v3 .authority-badge{width:48px;height:48px;flex:0 0 48px;display:grid;place-items:center;border-radius:14px;background:rgba(255,255,255,.09);border:1px solid rgba(255,255,255,.13);font-weight:900}
  .special-v3 .authority-card small{display:block;color:#aebcd0;font-size:12px!important;font-weight:700;letter-spacing:.04em;text-transform:uppercase}
  .special-v3 .authority-card strong{display:block;margin-top:4px;font-size:19px;line-height:1.3}
  .special-v3 .authority-card p{margin:7px 0 0;color:#d5deea!important;font-size:14px!important;line-height:1.55!important}
  .special-v3 .compat-card{padding:22px;background:#fff7f6;border-color:#f6d2ce}
  .special-v3 .compat-card strong{display:block;color:#9f2d24;font-size:12px;letter-spacing:.06em;text-transform:uppercase}
  .special-v3 .compat-card p{margin:8px 0 0;color:#6b3b37!important;font-size:14px!important;line-height:1.6!important}

  .special-v3 .diagnostic-wrap{padding:28px 0 64px;background:#fff}
  .special-v3 .diagnostic-card{padding:28px;background:linear-gradient(180deg,#fff,#fbfcfd)}
  .special-v3 .diagnostic-head{display:flex;justify-content:space-between;gap:20px;align-items:end;margin-bottom:22px}
  .special-v3 .diagnostic-head h2{margin:0;font-size:clamp(26px,3vw,38px);letter-spacing:-.035em;line-height:1.08}
  .special-v3 .diagnostic-head p{max-width:540px;margin:8px 0 0;color:var(--muted)!important;font-size:15px!important;line-height:1.65!important}
  .special-v3 .profile-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
  .special-v3 .profile-option{appearance:none;text-align:left;border:1px solid var(--line-2);border-radius:14px;padding:18px;background:#fff;color:var(--ink);cursor:pointer;transition:.18s ease}
  .special-v3 .profile-option:hover,.special-v3 .profile-option:focus-visible{border-color:#12b76a;box-shadow:0 0 0 4px rgba(18,183,106,.09);outline:none;transform:translateY(-1px)}
  .special-v3 .profile-option b{display:block;font-size:15px}.special-v3 .profile-option span{display:block;margin-top:6px;color:var(--muted);font-size:13px;line-height:1.5}
  .special-v3 .diagnostic-result{display:none;margin-top:16px;padding:18px;border-radius:14px;border:1px solid var(--green-line);background:var(--green-soft)}
  .special-v3 .diagnostic-result.is-visible{display:block}.special-v3 .diagnostic-result b{color:#05603a}.special-v3 .diagnostic-result p{margin:6px 0 14px;color:#3d6150!important;font-size:14px!important}
  .special-v3 .diagnostic-actions{display:flex;gap:10px;flex-wrap:wrap}

  .special-v3 .funnel-extensions{padding:68px 0;background:#f7f9f8;border-top:1px solid var(--line);border-bottom:1px solid var(--line)}
  .special-v3 .funnel-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px;margin-top:26px}
  .special-v3 .funnel-card{padding:24px;display:flex;flex-direction:column;min-height:290px}
  .special-v3 .funnel-card .funnel-kicker{font-size:11px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:var(--green)}
  .special-v3 .funnel-card h3{margin:10px 0 8px;font-size:24px;line-height:1.2;letter-spacing:-.03em}
  .special-v3 .funnel-card p{margin:0;color:var(--muted)!important;font-size:15px!important;line-height:1.65!important}
  .special-v3 .funnel-card ul{list-style:none;padding:0;margin:18px 0 0;display:grid;gap:9px}.special-v3 .funnel-card li{position:relative;padding-left:22px;color:#475467;font-size:14px!important;line-height:1.5!important}.special-v3 .funnel-card li::before{content:'✓';position:absolute;left:0;color:#067647;font-weight:900}
  .special-v3 .funnel-card .btn{margin-top:auto;align-self:flex-start}
  .special-v3 .onboard-note{margin-top:18px;padding:14px;border-left:4px solid #98a2b3;background:#f8fafc;border-radius:8px;color:#475467;font-size:13px;line-height:1.55}
  .special-v3 .radar-row{margin-top:18px;padding:16px;border:1px solid var(--line);border-radius:13px;background:#fff;display:flex;align-items:center;justify-content:space-between;gap:16px}
  .special-v3 .radar-status{display:flex;align-items:center;gap:9px;font-weight:800}.special-v3 .radar-dot{width:9px;height:9px;border-radius:50%;background:#f04438;box-shadow:0 0 0 5px #fef3f2}

  .special-v3 .lead-modal[hidden]{display:none!important}.special-v3 .lead-modal{position:fixed;inset:0;z-index:120;display:grid;place-items:center;padding:20px;background:rgba(11,22,42,.62);backdrop-filter:blur(6px)}
  .special-v3 .lead-dialog{width:min(460px,100%);padding:28px;border-radius:20px;background:#fff;border:1px solid var(--line);box-shadow:0 28px 70px rgba(11,22,42,.28)}
  .special-v3 .lead-dialog h3{margin:0;font-size:25px;letter-spacing:-.03em}.special-v3 .lead-dialog p{margin:8px 0 18px;color:var(--muted)!important;font-size:14px!important;line-height:1.6!important}.special-v3 .lead-dialog input{width:100%;min-height:48px;padding:0 14px;border:1px solid var(--line-2);border-radius:11px;background:#fbfcfd}.special-v3 .lead-dialog input:focus{outline:3px solid rgba(18,128,68,.12);border-color:var(--green)}
  .special-v3 .lead-dialog-actions{display:grid;gap:9px;margin-top:12px}.special-v3 .lead-close{border:0;background:transparent;color:var(--muted);font-weight:700;cursor:pointer;padding:8px}

  @media(max-width:840px){.special-v3 .funnel-authority,.special-v3 .funnel-grid{grid-template-columns:1fr}.special-v3 .profile-grid{grid-template-columns:1fr}.special-v3 .diagnostic-head{align-items:flex-start;flex-direction:column}.special-v3 .funnel-card{min-height:0}}
  @media(max-width:560px){.special-v3 .authority-card{align-items:flex-start}.special-v3 .diagnostic-card,.special-v3 .funnel-card{padding:20px}.special-v3 .radar-row{align-items:flex-start;flex-direction:column}}
</style>`;

  const topFunnel = `
<section class="funnel-band" aria-label="Deneyim ve uyumluluk"><div class="wrap funnel-authority">
  <article class="authority-card"><div class="authority-badge">EA</div><div><small>Tasarım ve finansal denetim yaklaşımı</small><strong>20 yıllık finans ve bilanço tecrübesi ile iş akışına göre Excel sistemi</strong><p>Önce muhasebe ve finans kuralı kurulur; ekran ve otomasyon bu kurala göre tasarlanır.</p></div></article>
  <aside class="compat-card"><strong>Teknik uyumluluk notu</strong><p>Makro ve ileri otomasyon içeren teslimatlar Windows masaüstü Excel sürümleri için planlanır. Mac, Numbers ve Google Sheets uyumu proje kapsamına göre ayrıca değerlendirilir.</p></aside>
</div></section>
<section class="diagnostic-wrap" aria-labelledby="diagnostic-title"><div class="wrap"><div class="diagnostic-card">
  <div class="diagnostic-head"><div><p class="eyebrow">60 saniyelik ihtiyaç teşhisi</p><h2 id="diagnostic-title">Size hangi Excel sistemi gerekir?</h2><p>İşletme tipinizi seçin; ilk görüşmede hangi problem alanından başlamamız gerektiğini görün.</p></div><button type="button" class="btn btn-secondary" data-open-demo>Demo Talebi</button></div>
  <div class="profile-grid" role="group" aria-label="İşletme profili">
    <button type="button" class="profile-option" data-profile="esnaf"><b>Küçük İşletme / KOBİ</b><span>Günlük kasa, nakit akışı, tahsilat ve ödeme kontrolü.</span></button>
    <button type="button" class="profile-option" data-profile="yonetim"><b>Şirket / Finans Yönetimi</b><span>Kârlılık, banka-kredi, bütçe, nakit ve yönetim raporlaması.</span></button>
    <button type="button" class="profile-option" data-profile="smmm"><b>SMMM / YMM / Muhasebe Ofisi</b><span>Mükellef takibi, cari kontrol, ofis operasyonu ve standart raporlama.</span></button>
  </div>
  <div class="diagnostic-result" data-result="esnaf"><b>Önerilen başlangıç:</b><p>Kasa + 13 haftalık nakit akışı + tahsilat/ödeme görünürlüğü. Önce günlük finansal kontrolü tek ekranda toplamak en yüksek faydayı verir.</p><div class="diagnostic-actions"><a class="btn btn-primary" href="#iletisim">Bu Yapıyı Konuşalım →</a><a class="btn btn-secondary" href="/sablonlar">Hazır Sistemleri İncele</a></div></div>
  <div class="diagnostic-result" data-result="yonetim"><b>Önerilen başlangıç:</b><p>Banka limit-risk, yönetim dashboard'u, gerçek kârlılık ve nakit projeksiyonu. Yönetim kararlarının tek veri omurgasından üretilmesi hedeflenir.</p><div class="diagnostic-actions"><a class="btn btn-primary" href="#iletisim">Finansal Röntgen Başlat →</a><a class="btn btn-secondary" href="#moduller">Modülleri Gör</a></div></div>
  <div class="diagnostic-result" data-result="smmm"><b>Önerilen başlangıç:</b><p>Mizan/muavin kontrolü, cari-tahsilat görünürlüğü ve mükellef bazlı standart operasyon paneli. Tekrarlanan manuel kontrolleri azaltmaya odaklanır.</p><div class="diagnostic-actions"><a class="btn btn-primary" href="#iletisim">Ofis Yapısını Analiz Edelim →</a><button type="button" class="btn btn-secondary" data-open-demo>Demo Talebi</button></div></div>
</div></div></section>`;

  const extensions = `
<section class="funnel-extensions" aria-labelledby="funnel-extension-title"><div class="wrap"><p class="eyebrow">Kurumsal kullanım ve teslim sonrası</p><h2 class="section-title" id="funnel-extension-title">Tek dosya değil, sürdürülebilir kullanım düzeni</h2><p class="section-copy">Özel sistemin değeri teslim günü değil; her ay aynı kontrolün daha hızlı, daha görünür ve daha az manuel emekle çalışmasında ortaya çıkar.</p>
<div class="funnel-grid">
  <article class="funnel-card"><span class="funnel-kicker">B2B / Ofis Kullanımı</span><h3>Mali müşavir ve finans ekipleri için toplu kullanım</h3><p>Aynı kontrol standardını birden fazla şirket veya mükellefte uygulayacak ofisler için kapsam, yetki, marka ve kullanım modeli proje bazında tasarlanır.</p><ul><li>Ofise veya kuruma özel görünüm</li><li>Standart veri giriş ve kontrol disiplini</li><li>Yetki, koruma ve teslim dokümantasyonu</li></ul><a class="btn btn-primary" href="#iletisim">Kurumsal Lisansı Görüşelim →</a></article>
  <article class="funnel-card"><span class="funnel-kicker">Partnerlik</span><h3>Referans ve çözüm ortaklığı modeli</h3><p>Müşterilerine düzenli finansal kontrol sistemi kurmak isteyen danışman, SMMM/YMM ve çözüm ortakları için işbirliği modeli ayrıca değerlendirilebilir.</p><ul><li>Kuruma uygun kullanım senaryosu</li><li>Referans akışı için şeffaf süreç</li><li>Gerekirse white-label kapsam değerlendirmesi</li></ul><a class="btn btn-secondary" href="#iletisim">Partnerlik Talebi Oluştur</a></article>
  <article class="funnel-card"><span class="funnel-kicker">Onboarding</span><h3>İlk 5 dakikada neyin nereye girileceği belli olur</h3><p>Teslim edilen sistemde veri giriş alanları, hesaplanan alanlar ve kontrol noktaları birbirinden ayrılır. Amaç boş bir Excel dosyası değil, kullanım sırası belli bir çalışma düzenidir.</p><div class="onboard-note"><strong>Performans disiplini:</strong> Çok yüksek satır hacmi, ağır makro veya yıllar arası arşiv ihtiyacı varsa veri hacmi proje başında test edilir; tek bir sabit satır limiti tüm sistemlere zorla uygulanmaz.</div><a class="btn btn-secondary" href="#surec">Teslimat Sürecini Gör</a></article>
  <article class="funnel-card"><span class="funnel-kicker">Mevzuat Radarı</span><h3>Mevzuata bağlı dosyalarda sürüm takibi</h3><p>Vergi, SGK, yeniden değerleme veya benzeri mevzuata bağlı hesaplama sistemlerinde değişiklik gerektiren durumlar ürün ve sürüm notları üzerinden takip edilir.</p><div class="radar-row"><div><div class="radar-status"><span class="radar-dot"></span> Güncelleme gerektiren dosyalar işaretlenir</div><p style="margin-top:7px">Eski oranı sessizce kullanmak yerine sürüm ve güncelleme ihtiyacı görünür tutulur.</p></div><a class="btn btn-secondary" href="/sablonlar">Ürünleri Gör</a></div></article>
</div></div></section>`;

  const modal = `
<div class="lead-modal" data-lead-modal hidden role="dialog" aria-modal="true" aria-labelledby="lead-modal-title"><div class="lead-dialog">
  <h3 id="lead-modal-title">Demo talebi oluşturun</h3><p>E-posta adresinizi girin. Adresi mevcut iletişim formuna aktaracağız; talebinizi oradan tamamlayabilirsiniz.</p>
  <label for="lead-email" style="font-size:13px;font-weight:800;display:block;margin-bottom:7px">E-posta adresi</label><input id="lead-email" type="email" autocomplete="email" placeholder="ornek@firma.com" />
  <div class="lead-dialog-actions"><button type="button" class="btn btn-primary" data-transfer-lead>İletişim Formuna Aktar →</button><button type="button" class="lead-close" data-close-demo>İptal</button></div>
</div></div>`;

  const script = `
<script data-special-funnel-js>
(() => {
  const results = [...document.querySelectorAll('[data-result]')];
  document.querySelectorAll('[data-profile]').forEach((button) => {
    button.addEventListener('click', () => {
      results.forEach((el) => el.classList.toggle('is-visible', el.dataset.result === button.dataset.profile));
      const active = document.querySelector('[data-result="' + button.dataset.profile + '"]');
      active?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    });
  });
  const modal = document.querySelector('[data-lead-modal]');
  const email = document.getElementById('lead-email');
  const open = () => { if (!modal) return; modal.hidden = false; setTimeout(() => email?.focus(), 20); };
  const close = () => { if (modal) modal.hidden = true; };
  document.querySelectorAll('[data-open-demo]').forEach((el) => el.addEventListener('click', open));
  document.querySelectorAll('[data-close-demo]').forEach((el) => el.addEventListener('click', close));
  modal?.addEventListener('click', (event) => { if (event.target === modal) close(); });
  document.addEventListener('keydown', (event) => { if (event.key === 'Escape') close(); });
  document.querySelector('[data-transfer-lead]')?.addEventListener('click', () => {
    const value = email?.value?.trim() || '';
    const target = document.querySelector('#iletisim input[type="email"]');
    if (target && value) { target.value = value; target.dispatchEvent(new Event('input', { bubbles: true })); }
    close();
    document.getElementById('iletisim')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    setTimeout(() => target?.focus(), 450);
  });
})();
</script>`;

  const diagnosticSchema = `
<script type="application/ld+json" data-special-funnel-schema>{"@context":"https://schema.org","@type":"WebApplication","name":"Excel Arşiv Özel Excel Sistemi İhtiyaç Teşhisi","url":"https://excelarsiv.com/ozel-excel-sistemleri","applicationCategory":"BusinessApplication","operatingSystem":"Web","description":"İşletme profiline göre özel Excel finans, muhasebe ve yönetim sistemi ihtiyacını sınıflandıran etkileşimli teşhis aracı.","provider":{"@type":"Organization","name":"Excel Arşiv","url":"https://excelarsiv.com"}}</script>`;

  html = once(html, '</head>', `${diagnosticSchema}${styles}</head>`, 'head');
  html = once(html, '<section id="saha"', `${topFunnel}<section id="saha"`, 'saha section anchor');
  html = once(html, '<section id="surec"', `${extensions}<section id="surec"`, 'process section anchor');
  html = once(html, '</body>', `${modal}${script}</body>`, 'body close');

  for (const token of [
    'data-special-funnel-v1',
    '60 saniyelik ihtiyaç teşhisi',
    'Mali müşavir ve finans ekipleri için toplu kullanım',
    'Mevzuat Radarı',
    'data-special-funnel-js',
    'data-special-funnel-schema',
  ]) {
    if (!html.includes(token)) throw new Error(`SPECIAL FUNNEL: missing token ${token}`);
  }
  if (html.includes('Ahmet Y. (Mali Müşavir)') || html.includes('Haftalık 12 Saat Tasarruf')) {
    throw new Error('SPECIAL FUNNEL: unverified testimonial content must not be published');
  }

  writeFileSync(specialPath, html, 'utf8');
  console.log('SPECIAL FUNNEL PASS — diagnostic, B2B, onboarding, radar, modal and truthful schema injected.');
}
