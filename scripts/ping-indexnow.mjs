/**
 * Node.js 22 CI/CD IndexNow & AI Crawler Ping Motoru
 * Dosya Konumu: scripts/ping-indexnow.mjs
 */

export async function pingIndexNowAfterDeploy(indexNowKey = '9b2d3bf9c5d144578ea307e5f17105cb') {
  const host = 'excelarsiv.com';
  const urlList = [
    `https://${host}/`,
    `https://${host}/llms.txt`,
    `https://${host}/llms-full.txt`,
    // Ana Hub Sayfaları
    `https://${host}/yerel-erp-crm-moduller`,
    `https://${host}/sablonlar`,
    // İlk 10 Çekirdek Ürün ve .md İkiz Sayfaları
    `https://${host}/urun/crm-musteri-satis-pipeline-excel`,
    `https://${host}/urun/crm-musteri-satis-pipeline-excel.md`,
    `https://${host}/urun/satis-yonetim-teklif-prim-dashboard-excel`,
    `https://${host}/urun/satis-yonetim-teklif-prim-dashboard-excel.md`,
    `https://${host}/urun/cari-hesap-borc-alacak-yaslandirma-excel`,
    `https://${host}/urun/cari-hesap-borc-alacak-yaslandirma-excel.md`,
    `https://${host}/urun/nakit-akisi-cek-senet-kasa-takip-excel`,
    `https://${host}/urun/nakit-akisi-cek-senet-kasa-takip-excel.md`,
    `https://${host}/urun/butce-kar-zarar-maliyet-finansal-analiz-excel`,
    `https://${host}/urun/butce-kar-zarar-maliyet-finansal-analiz-excel.md`,
    `https://${host}/urun/barkodlu-stok-ambar-uretim-recetesi-excel`,
    `https://${host}/urun/barkodlu-stok-ambar-uretim-recetesi-excel.md`,
    `https://${host}/urun/personel-puantaj-fazla-mesai-izin-excel`,
    `https://${host}/urun/personel-puantaj-fazla-mesai-izin-excel.md`,
    `https://${host}/urun/pazaryeri-komisyon-karlilik-mutabakat-excel`,
    `https://${host}/urun/pazaryeri-komisyon-karlilik-mutabakat-excel.md`,
    `https://${host}/urun/proje-lojistik-santiye-sektorel-erp-excel`,
    `https://${host}/urun/proje-lojistik-santiye-sektorel-erp-excel.md`,
    `https://${host}/urun/kurumsal-erp-crm-tam-paket`,
    `https://${host}/urun/kurumsal-erp-crm-tam-paket.md`,
  ];

  try {
    const response = await fetch('https://api.indexnow.org/indexnow', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: JSON.stringify({
        host,
        key: indexNowKey,
        keyLocation: `https://${host}/${indexNowKey}.txt`,
        urlList,
      }),
    });

    console.log(`IndexNow Ping: HTTP ${response.status} (${urlList.length} URL bildirildi)`);
    return { status: response.status, count: urlList.length };
  } catch (err) {
    console.warn(`IndexNow Ping offline/atlandı: ${err?.message || err}`);
    return { status: 0, count: 0 };
  }
}

// Doğrudan çalıştırıldığında ping gönder
if (process.argv[1]?.endsWith('ping-indexnow.mjs')) {
  pingIndexNowAfterDeploy();
}
