<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
                xmlns:html="http://www.w3.org/TR/REC-html40"
                xmlns:sitemap="http://www.sitemaps.org/schemas/sitemap/0.9"
                xmlns:xhtml="http://www.w3.org/1999/xhtml"
                xmlns:news="http://www.google.com/schemas/sitemap-news/0.9"
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
            --bg: #090b10;
            --card-bg: #11141d;
            --border: #1e2433;
            --text: #f1f5f9;
            --text-sub: #94a3b8;
            --accent: #10b981;
            --accent-hover: #059669;
            --badge-bg: rgba(59, 130, 246, 0.1);
            --badge-border: rgba(59, 130, 246, 0.25);
          }
          @media (prefers-color-scheme: light) {
            :root {
              --bg: #f8fafc;
              --card-bg: #ffffff;
              --border: #e2e8f0;
              --text: #0f172a;
              --text-sub: #64748b;
              --accent: #10b981;
              --accent-hover: #059669;
              --badge-bg: #eff6ff;
              --badge-border: #bfdbfe;
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
          .container { max-width: 1200px; margin: 0 auto; }
          .header { margin-bottom: 2rem; padding-bottom: 1.5rem; border-bottom: 1px solid var(--border); }
          .brand { display: inline-flex; align-items: center; gap: 0.5rem; font-size: 0.85rem; font-weight: 700; color: var(--accent); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.5rem; }
          .header h1 { font-size: 1.6rem; font-weight: 700; letter-spacing: -0.02em; margin-bottom: 0.5rem; }
          .header p { color: var(--text-sub); font-size: 0.875rem; }
          .stats-bar { display: flex; gap: 0.75rem; margin-top: 1rem; flex-wrap: wrap; }
          .stat-pill { display: inline-flex; align-items: center; padding: 0.25rem 0.75rem; background: var(--badge-bg); border: 1px solid var(--badge-border); border-radius: 9999px; color: var(--accent); font-weight: 600; font-size: 0.75rem; }
          .card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
          table { width: 100%; border-collapse: collapse; text-align: left; }
          th { background: var(--card-bg); padding: 0.75rem 1rem; font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; color: var(--text-sub); border-bottom: 1px solid var(--border); }
          td { padding: 0.75rem 1rem; border-bottom: 1px solid var(--border); vertical-align: middle; word-break: break-all; }
          tr:last-child td { border-bottom: none; }
          tr:hover td { background: rgba(255, 255, 255, 0.02); }
          a { color: var(--accent); text-decoration: none; transition: color 0.15s; }
          a:hover { color: var(--accent-hover); text-decoration: underline; }
          .lang-badge { display: inline-block; padding: 0.15rem 0.4rem; font-size: 0.7rem; font-weight: 600; border-radius: 4px; margin-right: 0.25rem; margin-bottom: 0.25rem; background: var(--badge-bg); border: 1px solid var(--badge-border); color: var(--accent); }
          .footer { margin-top: 2rem; text-align: center; color: var(--text-sub); font-size: 0.75rem; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <div class="brand">◆ EXCELARŞİV FİNANSAL SİSTEMLER</div>
            <xsl:choose>
              <xsl:when test="sitemap:sitemapindex">
                <h1>XML Site Haritası İndeksi</h1>
                <p>Arama motorları ve AI tarayıcılar (Googlebot, Bing, Perplexity) için alt site haritalarını gruplar.</p>
                <div class="stats-bar">
                  <span class="stat-pill">Alt Harita Sayısı: <xsl:value-of select="count(sitemap:sitemapindex/sitemap:sitemap)"/></span>
                  <span class="stat-pill">Tip: Sitemaps.org İndeks</span>
                </div>
              </xsl:when>
              <xsl:when test="//news:news">
                <h1>Google Haberler (News) XML Site Haritası</h1>
                <p>Google Haberler ve AI modelleri için yayınlanan güncel analitik bültenler.</p>
                <div class="stats-bar">
                  <span class="stat-pill">Haber Sayısı: <xsl:value-of select="count(sitemap:urlset/sitemap:url)"/></span>
                  <span class="stat-pill">Protokol: Google News 0.9</span>
                </div>
              </xsl:when>
              <xsl:when test="//image:image">
                <h1>Görsel (Image) XML Site Haritası</h1>
                <p>Google Görseller ve multimodal arama motorları için indekslenen şablon ve ekran görüntüleri.</p>
                <div class="stats-bar">
                  <span class="stat-pill">Sayfa Sayısı: <xsl:value-of select="count(sitemap:urlset/sitemap:url)"/></span>
                  <span class="stat-pill">Protokol: Google Image 1.1</span>
                </div>
              </xsl:when>
              <xsl:otherwise>
                <h1>XML Site Haritası (URL Seti)</h1>
                <p>Arama motorları ve AI tarayıcılar için yayınlanan kanonik sayfalar.</p>
                <div class="stats-bar">
                  <span class="stat-pill">Toplam URL: <xsl:value-of select="count(sitemap:urlset/sitemap:url)"/></span>
                  <span class="stat-pill">Protokol: Sitemaps.org 0.9</span>
                </div>
              </xsl:otherwise>
            </xsl:choose>
          </div>

          <div class="card">
            <table>
              <xsl:choose>
                <!-- 1. Sitemap Index Tablosu -->
                <xsl:when test="sitemap:sitemapindex">
                  <thead>
                    <tr>
                      <th style="width: 65%;">Alt Site Haritası URL</th>
                      <th style="width: 35%;">Son Güncelleme Tarihi</th>
                    </tr>
                  </thead>
                  <tbody>
                    <xsl:for-each select="sitemap:sitemapindex/sitemap:sitemap">
                      <tr>
                        <td>
                          <a href="{sitemap:loc}"><xsl:value-of select="sitemap:loc"/></a>
                        </td>
                        <td style="color: var(--text-sub);"><xsl:value-of select="sitemap:lastmod"/></td>
                      </tr>
                    </xsl:for-each>
                  </tbody>
                </xsl:when>

                <!-- 2. Google News Tablosu -->
                <xsl:when test="//news:news">
                  <thead>
                    <tr>
                      <th style="width: 45%;">Haber URL</th>
                      <th style="width: 35%;">Haber Başlığı</th>
                      <th style="width: 10%;">Dil</th>
                      <th style="width: 10%;">Yayın Tarihi</th>
                    </tr>
                  </thead>
                  <tbody>
                    <xsl:for-each select="sitemap:urlset/sitemap:url">
                      <tr>
                        <td>
                          <a href="{sitemap:loc}"><xsl:value-of select="sitemap:loc"/></a>
                        </td>
                        <td style="font-weight: 500;">
                          <xsl:value-of select="news:news/news:title"/>
                        </td>
                        <td>
                          <span class="lang-badge"><xsl:value-of select="news:news/news:publication/news:language"/></span>
                        </td>
                        <td style="color: var(--text-sub); font-size: 0.8rem;">
                          <xsl:value-of select="news:news/news:publication_date"/>
                        </td>
                      </tr>
                    </xsl:for-each>
                  </tbody>
                </xsl:when>

                <!-- 3. Google Image Tablosu -->
                <xsl:when test="//image:image">
                  <thead>
                    <tr>
                      <th style="width: 55%;">Sayfa URL</th>
                      <th style="width: 25%;">Son Güncelleme</th>
                      <th style="width: 20%;">Görsel Sayısı</th>
                    </tr>
                  </thead>
                  <tbody>
                    <xsl:for-each select="sitemap:urlset/sitemap:url">
                      <tr>
                        <td>
                          <a href="{sitemap:loc}"><xsl:value-of select="sitemap:loc"/></a>
                        </td>
                        <td style="color: var(--text-sub);"><xsl:value-of select="sitemap:lastmod"/></td>
                        <td>
                          <span class="stat-pill"><xsl:value-of select="count(image:image)"/> Görsel</span>
                        </td>
                      </tr>
                    </xsl:for-each>
                  </tbody>
                </xsl:when>

                <!-- 4. Standart URL Tablosu -->
                <xsl:otherwise>
                  <thead>
                    <tr>
                      <th style="width: 55%;">Sayfa URL</th>
                      <th style="width: 15%;">Son Güncelleme</th>
                      <th style="width: 10%;">Öncelik</th>
                      <th style="width: 20%;">Dil Varyantları</th>
                    </tr>
                  </thead>
                  <tbody>
                    <xsl:for-each select="sitemap:urlset/sitemap:url">
                      <tr>
                        <td>
                          <a href="{sitemap:loc}"><xsl:value-of select="sitemap:loc"/></a>
                        </td>
                        <td style="color: var(--text-sub);"><xsl:value-of select="sitemap:lastmod"/></td>
                        <td>
                          <xsl:if test="sitemap:priority">
                            <span class="stat-pill"><xsl:value-of select="sitemap:priority"/></span>
                          </xsl:if>
                        </td>
                        <td>
                          <xsl:for-each select="xhtml:link[@rel='alternate']">
                            <span class="lang-badge"><xsl:value-of select="@hreflang"/></span>
                          </xsl:for-each>
                        </td>
                      </tr>
                    </xsl:for-each>
                  </tbody>
                </xsl:otherwise>
              </xsl:choose>
            </table>
          </div>

          <div class="footer">
            <p>© 2026 ExcelArşiv. Tüm Hakları Saklıdır.</p>
          </div>
        </div>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>
