#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const MANDATE = 'data-ea-typography-mandate="cimpactpro-v1"';
const FONT_LINK = '<link data-ea-typography-font href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">';

const typographyStyles = `
<style ${MANDATE}>
  :root{
    --ea-font-sans:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
    --ea-type-body:16px;
    --ea-type-nav:15px;
    --ea-type-small:14px;
    --ea-type-h1:clamp(42px,5vw,64px);
    --ea-type-h2:clamp(32px,3.8vw,48px);
    --ea-type-h3:clamp(21px,2.2vw,28px);
    --ea-leading-body:1.65;
  }
  html{font-size:16px}
  body,button,input,select,textarea{font-family:var(--ea-font-sans)!important}
  body{font-size:var(--ea-type-body);line-height:var(--ea-leading-body);text-rendering:optimizeLegibility;-webkit-font-smoothing:antialiased}
  header nav a,header [role="navigation"] a,nav[aria-label] a{font-size:var(--ea-type-nav)!important;line-height:1.35!important}
  main p,main li,main details,main summary{font-family:var(--ea-font-sans)!important}
  footer,footer a{font-family:var(--ea-font-sans)!important;font-size:var(--ea-type-small)}
  @media(max-width:760px){:root{--ea-type-body:15px;--ea-type-nav:15px;--ea-type-small:13px}}
</style>`;

const specialSalesStyles = `
<style data-special-sales-v4>
  .special-v3 .nav-links{gap:26px!important;font-size:15px!important;font-weight:750!important}
  .special-v3 .nav-links a{font-size:15px!important;line-height:1.25!important;letter-spacing:-.01em!important}
  .special-v3 .nav .btn{font-size:15px!important;min-height:48px!important;padding-inline:20px!important}
  .special-v3 .pain.primary{min-height:100%!important}
  .special-v3 .pain.primary p{font-size:14px!important;line-height:1.78!important;color:#5d6b82!important}
  .special-v3 .close-grid{grid-template-columns:1fr 1fr!important;gap:24px!important;align-items:stretch!important}
  .special-v3 .faq-panel{display:flex!important;flex-direction:column!important;padding:30px!important;height:100%!important}
  .special-v3 .faq-panel .section-title{max-width:520px!important;font-size:clamp(30px,3vw,40px)!important}
  .special-v3 .faq-panel .section-copy{font-size:15px!important;line-height:1.7!important}
  .special-v3 .faq-list{flex:1!important;margin-top:24px!important;gap:10px!important}
  .special-v3 .faq-list details{background:#fff!important}
  .special-v3 .faq-list summary{padding:16px 17px!important;font-size:14px!important;line-height:1.45!important}
  .special-v3 .faq-list details p{padding:0 17px 17px!important;font-size:14px!important;line-height:1.7!important}
  .special-v3 .contact{height:100%!important;padding:30px!important}
  .special-v3 .contact .section-copy{font-size:16px!important;line-height:1.7!important}
  .special-v3 .contact-points div{font-size:13px!important;line-height:1.55!important}
  .special-v3 .field label{font-size:12px!important}
  .special-v3 .field input,.special-v3 .field select,.special-v3 .field textarea{font-size:14px!important}
  .special-v3 .privacy{font-size:11px!important}
  @media(max-width:980px){.special-v3 .close-grid{grid-template-columns:1fr!important}.special-v3 .faq-panel,.special-v3 .contact{height:auto!important}}
</style>`;

const oldPain = 'Borç-alacak mantığını, ters bakiye veren 100/102 hesapları, kapanmayan 320 avanslarını ve tevkifatı bilmezler. Size sundukları sistem ilk ay sonu kapanışında patlar; mizan tutturmak yine sizin manuel ameleliğinize kalır.';
const newPain = 'Borç-alacak mantığını, hesapların normal bakiye yönünü, 100 Kasa ve 102 Bankalar hesaplarında oluşan ters bakiyelerin ne anlattığını, 120–320 mahsuplarını, kapanmayan 320 avanslarını, tevkifat/KDV etkisini ve dönem sonu kayıt zincirini yalnız birer kolon olarak gören sistemler finansal gerçeği kaçırır. Sorun veri tabanı kurmak değil; mizanı kontrol edilebilir, mutabık ve yönetim kararına dönüşebilir hâle getirmektir. Yanlış eşleştirilmiş tek hesap; cari bakiyeyi, nakit görünümünü, kârlılık analizini ve yönetim raporunu aynı anda bozar. Sonuçta ilk ay sonu kapanışında ekip yine Excel’e döner, ters bakiye avına çıkar, manuel mahsup yapar ve yazılımın üretmesi gereken kontrolü insan emeğiyle tamamlar. Biz sistemi ekran sayısıyla değil; mizanın tutması, istisnaların görünmesi ve yönetim raporunun aynı veri üzerinden güvenilir biçimde üretilmesiyle ölçüyoruz.';

const newFaq = '<p class="eyebrow">Size Özel Excel Sistemi</p><h2 class="section-title">Karar Vermeden Önce Bilmeniz Gerekenler</h2><p class="section-copy">Buradaki amaç hazır bir tablo satmak değil; işletmenizde bugün elle kontrol edilen finans, muhasebe ve yönetim işini daha görünür, daha hızlı ve daha az hata üreten bir Excel sistemine dönüştürmektir.</p><div class="faq-list"><details open><summary>Bir yazılımcıdan temel farkınız ne?</summary><p>İşe ekran ve kod tarafından değil, mizan, nakit akışı, tahsilat riski, banka limiti, faiz yükü ve yönetim kararının nasıl üretileceği tarafından başlıyoruz. Önce finansal kuralı doğru kuruyor, sonra Excel mimarisini bu kurala hizmet edecek şekilde inşa ediyoruz.</p></details><details><summary>Hazır şablon yerine size özel sistem ne zaman gerekir?</summary><p>Mevcut Excel dosyalarınız birbirine bağlanmıyorsa, aynı veri birden fazla kez giriliyorsa, patron raporu için her ay manuel çalışma yapılıyorsa veya hazır ERP raporu karar vermeye yetmiyorsa özel sistem anlamlıdır. Sistem sizin gerçek iş akışınıza, kolonlarınıza, kontrol noktalarınıza ve raporlama ihtiyacınıza göre kurulur.</p></details><details><summary>Logo, Luca, Zirve veya Mikro verilerini yeniden mi gireceğiz?</summary><p>Hayır. Mevcut muhasebe/ERP sisteminizden aldığınız mizan, muavin, cari veya Excel/TXT dökümleri mümkün olan yerde doğrudan giriş kaynağı olarak kullanılır. Amaç ikinci bir veri giriş yükü yaratmak değil, elinizdeki veriyi kontrol ve karar ekranına çevirmektir.</p></details><details><summary>Teslimden sonra size bağımlı kalır mıyız?</summary><p>Hayır. Formüller ve hesap mantığı açık teslim edilir; sistem kara kutu hâline getirilmez. Ekibiniz hangi alanın girdi, hangi alanın hesap, hangi kontrolün neyi denetlediğini görebilir. Böylece küçük revizelerde sürekli geliştirici bekleme zorunluluğu azalır.</p></details><details><summary>İlk analizde ne göndermeliyim, çalışan örneği ne zaman görürüm?</summary><p>Tıkandığınız mevcut Excel dosyası, örnek mizan/muavin çıktısı veya görmek istediğiniz yönetim raporu yeterlidir. İlk görüşmede darboğazı ve karar ihtiyacını netleştirir, kapsam uygunsa gerçek yapınıza benzeyen çalışan ilk prototipi 48 saat içinde görünür hâle getiririz.</p></details></div>';

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

function replaceRegexExactly(html, pattern, replacement, label) {
  const matches = [...html.matchAll(pattern)];
  if (matches.length !== 1) throw new Error(`TYPOGRAPHY MANDATE: expected 1 ${label}, found ${matches.length}`);
  return html.replace(pattern, replacement);
}

export function applySiteTypographyMandate({ distDir = 'dist', specialPath = join('dist', 'ozel-excel-sistemleri', 'index.html') } = {}) {
  const htmlFiles = walkHtml(distDir);
  if (htmlFiles.length < 10) throw new Error(`TYPOGRAPHY MANDATE: suspicious HTML count ${htmlFiles.length}`);

  for (const path of htmlFiles) {
    let html = readFileSync(path, 'utf8');
    if (!html.includes('data-ea-typography-font')) html = html.replace('</head>', `${FONT_LINK}${typographyStyles}</head>`);
    else if (!html.includes(MANDATE)) html = html.replace('</head>', `${typographyStyles}</head>`);
    writeFileSync(path, html, 'utf8');
  }

  let special = readFileSync(specialPath, 'utf8');
  special = replaceExactly(special, oldPain, newPain, 'expanded mizan pain copy');
  special = replaceRegexExactly(
    special,
    /<p class="eyebrow">Sıkça Sorulan Sorular<\/p>\s*<h2 class="section-title">[\s\S]*?<\/h2>\s*<div class="faq-list">[\s\S]*?<\/div>/g,
    newFaq,
    'special sales FAQ',
  );
  special = replaceExactly(special, '</head>', `${specialSalesStyles}</head>`, 'special sales styles');
  writeFileSync(specialPath, special, 'utf8');

  const finalFiles = walkHtml(distDir);
  const missingMandate = finalFiles.filter((path) => !readFileSync(path, 'utf8').includes(MANDATE));
  if (missingMandate.length) throw new Error(`TYPOGRAPHY MANDATE: missing on ${missingMandate.length} HTML files: ${missingMandate.slice(0, 5).join(', ')}`);

  const specialFinal = readFileSync(specialPath, 'utf8');
  for (const token of [
    'data-special-sales-v4',
    'Bir yazılımcıdan temel farkınız ne?',
    'Hazır şablon yerine size özel sistem ne zaman gerekir?',
    'Logo, Luca, Zirve veya Mikro verilerini yeniden mi gireceğiz?',
    'Teslimden sonra size bağımlı kalır mıyız?',
    'İlk analizde ne göndermeliyim, çalışan örneği ne zaman görürüm?',
    'Biz sistemi ekran sayısıyla değil; mizanın tutması',
  ]) {
    if (!specialFinal.includes(token)) throw new Error(`TYPOGRAPHY MANDATE: special systems token missing: ${token}`);
  }
  if (specialFinal.includes('Finans ve Güvenlik Hakkında Merak Edilenler')) throw new Error('TYPOGRAPHY MANDATE: obsolete FAQ heading survived');
  console.log(`TYPOGRAPHY MANDATE PASS — ${finalFiles.length}/${finalFiles.length} HTML pages standardized; special systems sales copy + symmetry gate passed.`);
}
