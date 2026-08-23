#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'node:fs';

function once(html, from, to, label) {
  const count = html.split(from).length - 1;
  if (count !== 1) throw new Error(`SPECIAL COPY POLISH: expected 1 ${label}, found ${count}`);
  return html.replace(from, to);
}

function regexOnce(html, pattern, replacement, label) {
  const matches = html.match(pattern);
  if (!matches || matches.length !== 1) throw new Error(`SPECIAL COPY POLISH: expected 1 ${label}, found ${matches?.length ?? 0}`);
  return html.replace(pattern, replacement);
}

export function applySpecialCopyPolish({ specialPath = 'dist/ozel-excel-sistemleri/index.html' } = {}) {
  let html = readFileSync(specialPath, 'utf8');

  html = regexOnce(html, /<span[^>]*>Ters Pazarlama<\/span>/g, '', 'topbar tag');
  html = regexOnce(
    html,
    /Yazılımcı jargonu değil,\s*<b[^>]*>20 yıllık finans ve bilanço aklı<\/b>\s*konuşuyoruz\./g,
    '<b>20 yıllık finans ve bilanço tecrübesi ile</b> konuşuyoruz.',
    'topbar sales sentence',
  );

  html = once(
    html,
    '20 Yıllık Bankacılık & Bilanço Deneyimi · Sıfır Kod Bağımlılığı',
    'İhtiyaca yönelik Excel tabloları hazırlıyoruz',
    'hero positioning label',
  );

  const flatPain = 'Borç-alacak mantığını, hesapların normal bakiye yönünü, 100 Kasa ve 102 Bankalar hesaplarında oluşan ters bakiyelerin ne anlattığını, 120–320 mahsuplarını, kapanmayan 320 avanslarını, tevkifat/KDV etkisini ve dönem sonu kayıt zincirini yalnız birer kolon olarak gören sistemler finansal gerçeği kaçırır. Sorun veri tabanı kurmak değil; mizanı kontrol edilebilir, mutabık ve yönetim kararına dönüşebilir hâle getirmektir. Yanlış eşleştirilmiş tek hesap; cari bakiyeyi, nakit görünümünü, kârlılık analizini ve yönetim raporunu aynı anda bozar. Sonuçta ilk ay sonu kapanışında ekip yine Excel’e döner, ters bakiye avına çıkar, manuel mahsup yapar ve yazılımın üretmesi gereken kontrolü insan emeğiyle tamamlar. Bunun devamında alacak yaşlandırması yanlış okunur, net işletme sermayesi ihtiyacı hatalı görünür, banka limiti ve kredi kullanım kararı eksik veriye dayanır; patrona sunulan rapor ile muhasebenin gerçek bakiyesi birbirinden kopar. Biz sistemi ekran sayısıyla değil; mizanın tutması, istisnaların görünmesi, hesaplar arası ilişkinin doğrulanması ve yönetim raporunun aynı veri üzerinden güvenilir biçimde üretilmesiyle ölçüyoruz. Amaç daha çok tablo üretmek değil; ay sonu kapanışından nakit planına kadar aynı finansal mantığın tek sistem içinde denetlenebilir çalışmasını sağlamaktır.';

  const structuredPain = 'Borç-alacak mantığını, hesapların normal bakiye yönünü, 100 Kasa ve 102 Bankalar hesaplarında oluşan ters bakiyelerin ne anlattığını, 120–320 mahsuplarını, kapanmayan 320 avanslarını, tevkifat/KDV etkisini ve dönem sonu kayıt zincirini yalnızca birer veri kolonu olarak gören sistemler finansal gerçeği kaçırır.</p><p>Asıl mesele veri tabanı kurmak değildir. Mizanı kontrol edilebilir, mutabık ve yönetim kararına dönüşebilir hâle getirmektir. Yanlış eşleştirilmiş tek bir hesap; cari bakiyeyi, nakit görünümünü, kârlılık analizini ve yönetim raporunu aynı anda bozabilir. İlk ay sonu kapanışında ekip yeniden Excel’e döner, ters bakiye avına çıkar, manuel mahsup yapar ve yazılımın üretmesi gereken kontrolü insan emeğiyle tamamlar.</p><p>Bu bozulma yalnız muhasebe tarafında kalmaz. Alacak yaşlandırması yanlış okunur, net işletme sermayesi ihtiyacı hatalı görünür, banka limiti ve kredi kullanım kararı eksik veriye dayanır; patrona sunulan rapor ile muhasebenin gerçek bakiyesi birbirinden kopar.</p><p>Biz sistemi ekran sayısıyla değil; mizanın tutması, istisnaların görünmesi, hesaplar arası ilişkinin doğrulanması ve yönetim raporunun aynı veri üzerinden güvenilir biçimde üretilmesiyle ölçüyoruz. Amaç daha çok tablo üretmek değil; ay sonu kapanışından nakit planına kadar aynı finansal mantığın tek sistem içinde denetlenebilir çalışmasını sağlamaktır.';

  html = once(html, flatPain, structuredPain, 'structured mizan pain copy');

  const copyStyles = `\n<style data-special-copy-polish>\n  .special-v3 .pain.primary p + p{margin-top:16px!important}\n  .special-v3 .topbar .wrap{gap:14px!important}\n  .special-v3 .topbar .wrap>span:first-child{font-size:14px!important;line-height:1.45!important}\n</style>`;
  html = once(html, '</head>', `${copyStyles}</head>`, 'copy polish styles');

  for (const token of [
    '20 yıllık finans ve bilanço tecrübesi ile',
    'İhtiyaca yönelik Excel tabloları hazırlıyoruz',
    'Asıl mesele veri tabanı kurmak değildir.',
    'data-special-copy-polish',
  ]) {
    if (!html.includes(token)) throw new Error(`SPECIAL COPY POLISH: token missing: ${token}`);
  }
  if (html.includes('Ters Pazarlama') || html.includes('Yazılımcı jargonu değil')) {
    throw new Error('SPECIAL COPY POLISH: obsolete topbar copy survived');
  }

  writeFileSync(specialPath, html, 'utf8');
  console.log('SPECIAL COPY POLISH PASS — topbar, positioning label and structured mizan copy updated.');
}
