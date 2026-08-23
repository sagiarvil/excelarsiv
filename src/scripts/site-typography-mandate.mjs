#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const MANDATE = 'data-ea-typography-mandate="chat-readable-v2"';
const FONT_LINK = '<link data-ea-typography-font href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">';

const typographyStyles = `
<style ${MANDATE}>
  :root{
    --ea-font-sans:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
    --ea-type-body:16px;
    --ea-type-nav:16px;
    --ea-type-small:14px;
    --ea-type-h1:clamp(42px,5vw,64px);
    --ea-type-h2:clamp(32px,3.8vw,48px);
    --ea-type-h3:clamp(22px,2.2vw,29px);
    --ea-leading-body:1.65;
  }
  html{font-size:16px!important}
  body,button,input,select,textarea{font-family:var(--ea-font-sans)!important}
  body{font-size:16px!important;line-height:var(--ea-leading-body)!important;text-rendering:optimizeLegibility;-webkit-font-smoothing:antialiased}
  header nav a,header [role="navigation"] a,nav[aria-label] a{font-size:16px!important;line-height:1.35!important}
  main p,main li,main td,main th,main details,main summary{font-family:var(--ea-font-sans)!important;font-size:16px!important;line-height:1.65!important}
  main button,main a[class*="btn"],main a[class*="button"]{font-size:16px!important}
  main input,main select,main textarea{font-size:16px!important;line-height:1.45!important}
  main small{font-size:14px!important;line-height:1.5!important}
  footer,footer a{font-family:var(--ea-font-sans)!important;font-size:14px!important;line-height:1.55!important}
  @media(max-width:760px){
    :root{--ea-type-body:16px;--ea-type-nav:16px;--ea-type-small:14px}
    body{font-size:16px!important}
    main p,main li,main td,main th,main details,main summary{font-size:16px!important}
  }
</style>`;

const specialSalesStyles = `
<style data-special-sales-v5>
  .special-v3 .nav-inner{height:76px!important;gap:28px!important}
  .special-v3 .brand{min-width:270px!important;gap:12px!important}
  .special-v3 .brand-mark{overflow:visible!important}
  .special-v3 .brand-logo{width:52px!important;height:52px!important;max-height:none!important;object-fit:contain!important}
  .special-v3 .brand-wordmark{display:inline-flex!important;flex-direction:column!important;justify-content:center!important;margin-left:10px!important;white-space:nowrap!important}
  .special-v3 .brand-wordmark strong{font-size:22px!important;line-height:1.05!important;letter-spacing:-.035em!important;color:#111827!important;font-weight:800!important}
  .special-v3 .brand-wordmark small{margin-top:4px!important;font-size:11px!important;line-height:1.2!important;letter-spacing:.12em!important;text-transform:uppercase!important;color:#667085!important;font-weight:700!important}
  .special-v3 .nav-links{gap:30px!important;font-size:16px!important;font-weight:750!important}
  .special-v3 .nav-links a{font-size:16px!important;line-height:1.25!important;letter-spacing:-.012em!important}
  .special-v3 .nav .btn{font-size:16px!important;min-height:50px!important;padding-inline:22px!important}

  .special-v3 .hero-copy,.special-v3 .section-copy{font-size:17px!important;line-height:1.7!important}
  .special-v3 .proof b{font-size:18px!important}.special-v3 .proof span{font-size:14px!important;line-height:1.5!important}
  .special-v3 .dashboard-head span,.special-v3 .dash-kpi small,.special-v3 .dash-row{font-size:13px!important}
  .special-v3 .dash-kpi b{font-size:15px!important}

  .special-v3 .pain-grid{gap:16px!important}
  .special-v3 .pain{padding:28px!important}
  .special-v3 .pain h3{font-size:23px!important;line-height:1.28!important}
  .special-v3 .pain p{font-size:16px!important;line-height:1.72!important;color:#5d6b82!important}
  .special-v3 .pain.primary{min-height:100%!important;justify-content:flex-start!important}
  .special-v3 .pain.primary .pain-foot{margin-top:24px!important}
  .special-v3 .risk-chip{font-size:13px!important;padding:7px 10px!important}

  .special-v3 .comparison-head div{font-size:15px!important;line-height:1.45!important;padding:18px 20px!important}
  .special-v3 .comparison-row>div{font-size:16px!important;line-height:1.6!important;padding:18px 20px!important}
  .special-v3 .comparison-row .new{padding-left:48px!important}

  .special-v3 .module{padding:26px!important}
  .special-v3 .module h3{font-size:23px!important;line-height:1.3!important}
  .special-v3 .module p,.special-v3 .module li{font-size:16px!important;line-height:1.62!important}
  .special-v3 .module-meta span{font-size:13px!important;padding:7px 9px!important}

  .special-v3 .step h3{font-size:22px!important}.special-v3 .step p{font-size:16px!important;line-height:1.65!important}

  .special-v3 .close-grid{grid-template-columns:1fr 1fr!important;gap:24px!important;align-items:stretch!important}
  .special-v3 .faq-panel{display:flex!important;flex-direction:column!important;padding:32px!important;height:100%!important}
  .special-v3 .faq-panel .section-title{max-width:540px!important;font-size:clamp(32px,3vw,42px)!important}
  .special-v3 .faq-panel .section-copy{font-size:16px!important;line-height:1.72!important}
  .special-v3 .faq-list{flex:1!important;margin-top:24px!important;gap:11px!important}
  .special-v3 .faq-list details{background:#fff!important}
  .special-v3 .faq-list summary{padding:17px 18px!important;font-size:16px!important;line-height:1.5!important}
  .special-v3 .faq-list details p{padding:0 18px 18px!important;font-size:16px!important;line-height:1.72!important}
  .special-v3 .contact{height:100%!important;padding:32px!important}
  .special-v3 .contact .section-copy{font-size:16px!important;line-height:1.72!important}
  .special-v3 .contact-points div{font-size:15px!important;line-height:1.55!important;padding:13px 14px!important}
  .special-v3 .field label{font-size:14px!important}
  .special-v3 .field input,.special-v3 .field select,.special-v3 .field textarea{font-size:16px!important}
  .special-v3 .privacy{font-size:13px!important}

  .special-v3 footer .footer-inner{min-height:96px!important;gap:30px!important;font-size:14px!important}
  .special-v3 footer .footer-inner>div:first-child{display:flex!important;align-items:center!important;gap:18px!important;min-width:430px!important;color:#b7c3d4!important}
  .special-v3 footer .footer-logo{width:48px!important;height:48px!important;max-height:none!important;filter:none!important;object-fit:contain!important}
  .special-v3 footer .footer-wordmark{display:inline-flex!important;flex-direction:column!important;justify-content:center!important;white-space:nowrap!important}
  .special-v3 footer .footer-wordmark strong{font-size:19px!important;line-height:1.05!important;color:#fff!important;letter-spacing:-.025em!important}
  .special-v3 footer .footer-wordmark small{margin-top:4px!important;font-size:12px!important;line-height:1.25!important;color:#9fb0c4!important;letter-spacing:.08em!important;text-transform:uppercase!important}

  @media(max-width:1100px){
    .special-v3 .brand{min-width:220px!important}.special-v3 .nav-links{gap:20px!important}.special-v3 .nav-links a{font-size:15px!important}
  }
  @media(max-width:980px){
    .special-v3 .close-grid{grid-template-columns:1fr!important}.special-v3 .faq-panel,.special-v3 .contact{height:auto!important}
    .special-v3 footer .footer-inner>div:first-child{min-width:0!important}
  }
  @media(max-width:720px){
    .special-v3 .nav-inner{height:68px!important}.special-v3 .brand{min-width:0!important}.special-v3 .brand-logo{width:44px!important;height:44px!important}
    .special-v3 .brand-wordmark strong{font-size:19px!important}.special-v3 .brand-wordmark small{font-size:9px!important}
    .special-v3 .pain,.special-v3 .module,.special-v3 .faq-panel,.special-v3 .contact{padding:22px!important}
    .special-v3 footer .footer-inner>div:first-child{align-items:flex-start!important}.special-v3 footer .footer-wordmark strong{font-size:18px!important}
  }
</style>`;

const oldPain = 'Borç-alacak mantığını, ters bakiye veren 100/102 hesapları, kapanmayan 320 avanslarını ve tevkifatı bilmezler. Size sundukları sistem ilk ay sonu kapanışında patlar; mizan tutturmak yine sizin manuel ameleliğinize kalır.';
const newPain = 'Borç-alacak mantığını, hesapların normal bakiye yönünü, 100 Kasa ve 102 Bankalar hesaplarında oluşan ters bakiyelerin ne anlattığını, 120–320 mahsuplarını, kapanmayan 320 avanslarını, tevkifat/KDV etkisini ve dönem sonu kayıt zincirini yalnız birer kolon olarak gören sistemler finansal gerçeği kaçırır. Sorun veri tabanı kurmak değil; mizanı kontrol edilebilir, mutabık ve yönetim kararına dönüşebilir hâle getirmektir. Yanlış eşleştirilmiş tek hesap; cari bakiyeyi, nakit görünümünü, kârlılık analizini ve yönetim raporunu aynı anda bozar. Sonuçta ilk ay sonu kapanışında ekip yine Excel’e döner, ters bakiye avına çıkar, manuel mahsup yapar ve yazılımın üretmesi gereken kontrolü insan emeğiyle tamamlar. Bunun devamında alacak yaşlandırması yanlış okunur, net işletme sermayesi ihtiyacı hatalı görünür, banka limiti ve kredi kullanım kararı eksik veriye dayanır; patrona sunulan rapor ile muhasebenin gerçek bakiyesi birbirinden kopar. Biz sistemi ekran sayısıyla değil; mizanın tutması, istisnaların görünmesi, hesaplar arası ilişkinin doğrulanması ve yönetim raporunun aynı veri üzerinden güvenilir biçimde üretilmesiyle ölçüyoruz. Amaç daha çok tablo üretmek değil; ay sonu kapanışından nakit planına kadar aynı finansal mantığın tek sistem içinde denetlenebilir çalışmasını sağlamaktır.';

const newFaq = '<p class="eyebrow">Size Özel Excel Sistemi</p><h2 class="section-title">Karar Vermeden Önce Bilmeniz Gerekenler</h2><p class="section-copy">Buradaki amaç hazır bir tablo satmak değil; işletmenizde bugün elle kontrol edilen finans, muhasebe ve yönetim işini daha görünür, daha hızlı ve daha az hata üreten bir Excel sistemine dönüştürmektir.</p><div class="faq-list"><details open><summary>Bir yazılımcıdan temel farkınız ne?</summary><p>İşe ekran ve kod tarafından değil, mizan, nakit akışı, tahsilat riski, banka limiti, faiz yükü ve yönetim kararının nasıl üretileceği tarafından başlıyoruz. Önce finansal kuralı doğru kuruyor, sonra Excel mimarisini bu kurala hizmet edecek şekilde inşa ediyoruz.</p></details><details><summary>Hazır şablon yerine size özel sistem ne zaman gerekir?</summary><p>Mevcut Excel dosyalarınız birbirine bağlanmıyorsa, aynı veri birden fazla kez giriliyorsa, patron raporu için her ay manuel çalışma yapılıyorsa veya hazır ERP raporu karar vermeye yetmiyorsa özel sistem anlamlıdır. Sistem sizin gerçek iş akışınıza, kolonlarınıza, kontrol noktalarınıza ve raporlama ihtiyacınıza göre kurulur.</p></details><details><summary>Logo, Luca, Zirve veya Mikro verilerini yeniden mi gireceğiz?</summary><p>Hayır. Mevcut muhasebe/ERP sisteminizden aldığınız mizan, muavin, cari veya Excel/TXT dökümleri mümkün olan yerde doğrudan giriş kaynağı olarak kullanılır. Amaç ikinci bir veri giriş yükü yaratmak değil, elinizdeki veriyi kontrol ve karar ekranına çevirmektir.</p></details><details><summary>Teslimden sonra size bağımlı kalır mıyız?</summary><p>Hayır. Formüller ve hesap mantığı açık teslim edilir; sistem kara kutu hâline getirilmez. Ekibiniz hangi alanın girdi, hangi alanın hesap, hangi kontrolün neyi denetlediğini görebilir. Böylece küçük revizelerde sürekli geliştirici bekleme zorunluluğu azalır.</p></details><details><summary>İlk analizde ne göndermeliyim, çalışan örneği ne zaman görürüm?</summary><p>Tıkandığınız mevcut Excel dosyası, örnek mizan/muavin çıktısı veya görmek istediğiniz yönetim raporu yeterlidir. İlk görüşmede darboğazı ve karar ihtiyacını netleştirir, kapsam uygunsa gerçek yapınıza benzeyen çalışan ilk prototipi 48 saat içinde görünür hâle getiririz.</p></details></div>';

const headerLogo = '<img class="brand-logo" src="/images/brand/excelarsiv-header-logo.png" alt="Excel Arşiv" width="420" height="120" decoding="async">';
const headerLockup = '<img class="brand-logo" src="/images/brand/excelarsiv-header-logo.png" alt="Excel Arşiv" width="120" height="120" decoding="async"><span class="brand-wordmark"><strong>Excel Arşiv</strong><small>Finans & Yönetim Mimarisi</small></span>';
const footerLogo = '<img class="footer-logo" src="/images/brand/excelarsiv-header-logo.png" alt="Excel Arşiv" width="420" height="120" loading="lazy" decoding="async">';
const footerLockup = '<img class="footer-logo" src="/images/brand/excelarsiv-header-logo.png" alt="Excel Arşiv" width="120" height="120" loading="lazy" decoding="async"><span class="footer-wordmark"><strong>Excel Arşiv</strong><small>Finans & Yönetim Mimarisi</small></span>';

function walkHtml(dir, out = []) {
  if (!existsSync(dir)) return out;
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    const stat = statSync(path);
    if (stat.isDirectory()) walkHtml(path, out);
    else if (stat.isFile() && name.endsWith('.html')) out.push(path);
  }
  return out;
}

function replaceExactly(html, from, to, label) {
  const count = html.split(from).length - 1;
  if (count !== 1) throw new Error(`TYPOGRAPHY MANDATE: expected 1 ${label}, found ${count}`);
  return html.replace(from, to);
}

function replaceFaqByVisibleAnchors(html) {
  const title = 'Sıkça Sorulan Sorular';
  const finalAnswer = 'Özel Excel modelleri ERP dökümlerini tek tıkla karar mekanizmasına çevirir.';
  const titleAt = html.indexOf(title);
  if (titleAt < 0 || html.indexOf(title, titleAt + title.length) >= 0) throw new Error('TYPOGRAPHY MANDATE: FAQ title anchor is missing or ambiguous');
  const start = html.lastIndexOf('<p', titleAt);
  const answerAt = html.indexOf(finalAnswer, titleAt);
  const end = answerAt < 0 ? -1 : html.indexOf('</div>', answerAt);
  if (start < 0 || answerAt < 0 || end < 0) throw new Error('TYPOGRAPHY MANDATE: FAQ visible anchors could not resolve panel range');
  return `${html.slice(0, start)}${newFaq}${html.slice(end + '</div>'.length)}`;
}

export function applySiteTypographyMandate({ distDir = 'dist', specialPath = join('dist', 'ozel-excel-sistemleri', 'index.html') } = {}) {
  const htmlFiles = walkHtml(distDir);
  if (htmlFiles.length < 10) throw new Error(`TYPOGRAPHY MANDATE: suspicious HTML count ${htmlFiles.length}`);

  for (const path of htmlFiles) {
    let html = readFileSync(path, 'utf8');
    if (!html.includes('data-ea-typography-font')) html = html.replace('</head>', `${FONT_LINK}${typographyStyles}</head>`);
    else if (!html.includes('data-ea-typography-mandate="chat-readable-v2"')) html = html.replace('</head>', `${typographyStyles}</head>`);
    writeFileSync(path, html, 'utf8');
  }

  let special = readFileSync(specialPath, 'utf8');
  special = replaceExactly(special, oldPain, newPain, 'expanded mizan pain copy');
  special = replaceFaqByVisibleAnchors(special);
  special = replaceExactly(special, headerLogo, headerLockup, 'premium header wordmark');
  special = replaceExactly(special, footerLogo, footerLockup, 'premium footer wordmark');
  special = replaceExactly(special, '</head>', `${specialSalesStyles}</head>`, 'special sales styles');
  writeFileSync(specialPath, special, 'utf8');

  const finalFiles = walkHtml(distDir);
  const missingMandate = finalFiles.filter((path) => !readFileSync(path, 'utf8').includes('data-ea-typography-mandate="chat-readable-v2"'));
  if (missingMandate.length) throw new Error(`TYPOGRAPHY MANDATE: missing on ${missingMandate.length} HTML files: ${missingMandate.slice(0, 5).join(', ')}`);

  const specialFinal = readFileSync(specialPath, 'utf8');
  for (const token of [
    'data-special-sales-v5',
    'brand-wordmark',
    'footer-wordmark',
    '>Excel Arşiv</strong>',
    'Bir yazılımcıdan temel farkınız ne?',
    'Hazır şablon yerine size özel sistem ne zaman gerekir?',
    'Logo, Luca, Zirve veya Mikro verilerini yeniden mi gireceğiz?',
    'Teslimden sonra size bağımlı kalır mıyız?',
    'İlk analizde ne göndermeliyim, çalışan örneği ne zaman görürüm?',
    'Amaç daha çok tablo üretmek değil;',
  ]) {
    if (!specialFinal.includes(token)) throw new Error(`TYPOGRAPHY MANDATE: special systems token missing: ${token}`);
  }
  if (specialFinal.includes('Finans ve Güvenlik Hakkında Merak Edilenler')) throw new Error('TYPOGRAPHY MANDATE: obsolete FAQ heading survived');
  console.log(`TYPOGRAPHY MANDATE PASS — ${finalFiles.length}/${finalFiles.length} HTML pages use 16px readable baseline; special systems brand + sales symmetry gate passed.`);
}
