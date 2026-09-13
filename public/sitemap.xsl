<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
                xmlns:html="http://www.w3.org/TR/REC-html40"
                xmlns:sitemap="http://www.sitemaps.org/schemas/sitemap/0.9"
                xmlns:image="http://www.google.com/schemas/sitemap-image/1.1"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output method="html" version="1.0" encoding="UTF-8" indent="yes"/>
  <xsl:template match="/">
    <html lang="tr">
      <head>
        <title>XML Sitemap | ExcelArşiv</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
        <style>
          :root {
            --bg: #090d0b;
            --card-bg: #111815;
            --border: #1e2d27;
            --text: #f0fdf4;
            --text-muted: #86efac;
            --text-sub: #94a3b8;
            --accent: #10b981;
            --accent-hover: #34d399;
            --badge-bg: rgba(16, 185, 129, 0.1);
            --badge-border: rgba(16, 185, 129, 0.25);
          }
          @media (prefers-color-scheme: light) {
            :root {
              --bg: #f8fafc;
              --card-bg: #ffffff;
              --border: #e2e8f0;
              --text: #0f172a;
              --text-muted: #059669;
              --text-sub: #64748b;
              --accent: #059669;
              --accent-hover: #047857;
              --badge-bg: #ecfdf5;
              --badge-border: #a7f3d0;
            }
          }
          * { box-sizing: border-box; margin: 0; padding: 0; }
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: var(--bg);
            color: var(--text);
            padding: 2rem 1rem;
            font-size: 13px;
            line-height: 1.5;
          }
          .container {
            max-width: 1200px;
            margin: 0 auto;
          }
          .header {
            margin-bottom: 2rem;
            padding-bottom: 1.5rem;
            border-bottom: 1px solid var(--border);
          }
          .brand {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            font-size: 0.85rem;
            font-weight: 700;
            color: var(--accent);
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 0.5rem;
          }
          .header h1 {
            font-size: 1.6rem;
            font-weight: 700;
            letter-spacing: -0.02em;
            margin-bottom: 0.5rem;
          }
          .header p {
            color: var(--text-sub);
            font-size: 0.875rem;
          }
          .stats-bar {
            display: flex;
            gap: 0.75rem;
            margin-top: 1rem;
            flex-wrap: wrap;
          }
          .stat-pill {
            display: inline-flex;
            align-items: center;
            padding: 0.25rem 0.75rem;
            background: var(--badge-bg);
            border: 1px solid var(--badge-border);
            border-radius: 9999px;
            color: var(--accent);
            font-weight: 600;
            font-size: 0.75rem;
          }
          .card {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
          }
          table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
          }
          th {
            background: var(--card-bg);
            padding: 0.75rem 1rem;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-sub);
            border-bottom: 1px solid var(--border);
          }
          td {
            padding: 0.75rem 1rem;
            border-bottom: 1px solid var(--border);
            vertical-align: middle;
            word-break: break-all;
          }
          tr:last-child td { border-bottom: none; }
          tr:hover td {
            background: rgba(255, 255, 255, 0.02);
          }
          a {
            color: var(--accent);
            text-decoration: none;
            transition: color 0.15s;
          }
          a:hover {
            color: var(--accent-hover);
            text-decoration: underline;
          }
          .footer {
            margin-top: 2rem;
            text-align: center;
            color: var(--text-sub);
            font-size: 0.75rem;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <div class="brand">◆ EXCELARŞİV FINANSAL SİSTEMLER</div>
            <xsl:choose>
              <xsl:when test="sitemap:sitemapindex">
                <h1>XML Site Haritası İndeksi</h1>
                <p>Bu ana indeks dosyası, arama motorları (Googlebot, Bing, Perplexity) için ExcelArşiv alt site haritalarını gruplar.</p>
                <div class="stats-bar">
                  <span class="stat-pill">Alt Harita Sayısı: <xsl:value-of select="count(sitemap:sitemapindex/sitemap:sitemap)"/></span>
                  <span class="stat-pill">Tip: Sitemaps.org İndeks</span>
                </div>
              </xsl:when>
              <xsl:otherwise>
                <h1>XML Site Haritası</h1>
                <p>Arama motorları ve AI tarayıcılar için yayınlanan kanonik URL ve veri listesi.</p>
                <div class="stats-bar">
                  <span class="stat-pill">Toplam URL: <xsl:value-of select="count(sitemap:urlset/sitemap:url)"/></span>
                  <span class="stat-pill">Tip: Sitemaps.org URL Seti</span>
                </div>
              </xsl:otherwise>
            </xsl:choose>
          </div>

          <div class="card">
            <table>
              <xsl:choose>
                <xsl:when test="sitemap:sitemapindex">
                  <thead>
                    <tr>
                      <th style="width: 70%;">Alt Site Haritası Adresi</th>
                      <th style="width: 30%;">Son Güncelleme Tarihi</th>
                    </tr>
                  </thead>
                  <tbody>
                    <xsl:for-each select="sitemap:sitemapindex/sitemap:sitemap">
                      <tr>
                        <td>
                          <a href="{sitemap:loc}">
                            <xsl:value-of select="sitemap:loc"/>
                          </a>
                        </td>
                        <td style="color: var(--text-sub);">
                          <xsl:value-of select="sitemap:lastmod"/>
                        </td>
                      </tr>
                    </xsl:for-each>
                  </tbody>
                </xsl:when>
                <xsl:otherwise>
                  <thead>
                    <tr>
                      <th style="width: 70%;">Sayfa / Şablon URL</th>
                      <th style="width: 30%;">Son Güncelleme Tarihi</th>
                    </tr>
                  </thead>
                  <tbody>
                    <xsl:for-each select="sitemap:urlset/sitemap:url">
                      <tr>
                        <td>
                          <a href="{sitemap:loc}">
                            <xsl:value-of select="sitemap:loc"/>
                          </a>
                        </td>
                        <td style="color: var(--text-sub);">
                          <xsl:value-of select="sitemap:lastmod"/>
                        </td>
                      </tr>
                    </xsl:for-each>
                  </tbody>
                </xsl:otherwise>
              </xsl:choose>
            </table>
          </div>

          <div class="footer">
            <p>© 2026 ExcelArşiv. Tüm Hakları Saklıdır. Kurumsal Finansal Excel &amp; Raporlama Kütüphanesi.</p>
          </div>
        </div>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>
