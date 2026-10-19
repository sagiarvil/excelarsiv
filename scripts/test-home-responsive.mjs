import puppeteer from 'puppeteer-core';
import path from 'node:path';
import fs from 'node:fs';
import http from 'node:http';

const CHROME_PATH = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const DIST_DIR = path.resolve('dist');

// Simple static server for dist
function startServer(port = 4173) {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      let reqPath = req.url.split('?')[0];
      if (reqPath === '/') reqPath = '/index.html';
      const filePath = path.join(DIST_DIR, reqPath);
      if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
        const ext = path.extname(filePath);
        const mimeTypes = {
          '.html': 'text/html; charset=utf-8',
          '.css': 'text/css; charset=utf-8',
          '.js': 'application/javascript; charset=utf-8',
          '.png': 'image/png',
          '.webp': 'image/webp',
          '.jpeg': 'image/jpeg',
          '.jpg': 'image/jpeg',
          '.svg': 'image/svg+xml',
          '.woff2': 'font/woff2'
        };
        res.writeHead(200, { 'Content-Type': mimeTypes[ext] || 'application/octet-stream' });
        fs.createReadStream(filePath).pipe(res);
      } else {
        res.writeHead(404);
        res.end('Not found');
      }
    });
    server.listen(port, () => resolve(server));
  });
}

async function run() {
  const port = 4188;
  const server = await startServer(port);
  const url = `http://localhost:${port}/`;
  console.log(`Test server running at ${url}`);

  const browser = await puppeteer.launch({
    executablePath: CHROME_PATH,
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu']
  });

  const viewports = [320, 360, 390, 768, 1024, 1280, 1366, 1440];
  const results = [];
  const screenshotDir = path.resolve('.screenshots');
  if (!fs.existsSync(screenshotDir)) fs.mkdirSync(screenshotDir, { recursive: true });

  const page = await browser.newPage();
  const consoleErrors = [];
  page.on('console', msg => {
    if (msg.type() === 'error') consoleErrors.push(msg.text());
  });
  page.on('pageerror', err => consoleErrors.push(err.message));

  for (const w of viewports) {
    await page.setViewport({ width: w, height: 900 });
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    await new Promise(r => setTimeout(r, 600));

    // 1. Horizontal overflow test
    const overflow = await page.evaluate(() => {
      const docEl = document.documentElement;
      const body = document.body;
      const scrollWidth = Math.max(docEl.scrollWidth, body.scrollWidth);
      const clientWidth = docEl.clientWidth;
      const badElements = [];
      const all = document.querySelectorAll('*');
      for (const el of all) {
        const r = el.getBoundingClientRect();
        if (r.right > clientWidth + 2) {
          badElements.push({
            tag: el.tagName,
            id: el.id,
            cls: typeof el.className === 'string' ? el.className.slice(0, 60) : '',
            right: Math.round(r.right),
            width: Math.round(r.width)
          });
        }
      }
      return {
        hasOverflow: scrollWidth > clientWidth,
        scrollWidth,
        clientWidth,
        diff: scrollWidth - clientWidth,
        badElements: badElements.slice(0, 5)
      };
    });

    // 2. Footer columns test
    const footerCheck = await page.evaluate(() => {
      const footer = document.querySelector('footer');
      if (!footer) return { ok: false, error: 'Footer not found' };
      const cols = Array.from(footer.querySelectorAll('.grid > div'));
      return {
        colCount: cols.length,
        colWidths: cols.map(c => c.getBoundingClientRect().width)
      };
    });

    // 3. Showcase cards count test
    const showcaseCount = await page.evaluate(() => {
      const cards = document.querySelectorAll('#templateGrid .template-card');
      return cards.length;
    });

    results.push({
      width: w,
      overflow,
      footerCheck,
      showcaseCount
    });

    // Capture screenshot at representative break points (360, 768, 1366)
    if ([360, 768, 1366].includes(w)) {
      await page.screenshot({ path: path.join(screenshotDir, `viewport-${w}.png`), fullPage: false });
    }
  }

  // Functional Interactive Tests at 1366px
  console.log('\n--- INTERACTIVE FUNCTIONAL TESTS (1366px) ---');
  await page.setViewport({ width: 1366, height: 900 });
  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await new Promise(r => setTimeout(r, 600));

  // 1. Test Chart and Simulator table recalculation
  const simulatorTest = await page.evaluate(() => {
    const kpi3Before = document.getElementById('kpi3Value')?.textContent;
    const firstInput = document.querySelector('.sim-cell-input');
    if (firstInput) {
      firstInput.value = '900000';
      firstInput.dispatchEvent(new Event('input', { bubbles: true }));
    }
    const kpi3After = document.getElementById('kpi3Value')?.textContent;
    const chartCanvas = document.getElementById('simulatorChart');
    const chartFallback = document.getElementById('chartFallback');
    const isFallbackHidden = chartFallback ? chartFallback.classList.contains('hidden') : true;

    return {
      kpi3Before,
      kpi3After,
      updated: kpi3Before !== kpi3After,
      chartCanvasExists: !!chartCanvas,
      isFallbackHidden
    };
  });
  console.log('Simulator test:', simulatorTest);

  // 2. Test Search & Filter on Showcase
  const filterTest = await page.evaluate(() => {
    // Click 'Finans' filter
    const finansBtn = document.querySelector('[data-cat="Finans"]');
    if (finansBtn) finansBtn.click();
    const visibleAfterCat = Array.from(document.querySelectorAll('#templateGrid .template-card')).filter(c => c.style.display !== 'none').length;

    // Search for non-existent keyword
    const searchInput = document.getElementById('heroSearchInput');
    searchInput.value = 'xyznonexistent999';
    searchInput.dispatchEvent(new Event('input', { bubbles: true }));
    const visibleAfterSearch = Array.from(document.querySelectorAll('#templateGrid .template-card')).filter(c => c.style.display !== 'none').length;
    const emptyStateVisible = !document.getElementById('templateEmptyState')?.classList.contains('hidden');

    // Click clear filters
    const clearBtn = document.getElementById('clearFiltersBtn');
    if (clearBtn) clearBtn.click();
    const visibleAfterClear = Array.from(document.querySelectorAll('#templateGrid .template-card')).filter(c => c.style.display !== 'none').length;

    return {
      visibleAfterCat,
      visibleAfterSearch,
      emptyStateVisible,
      visibleAfterClear
    };
  });
  console.log('Filter & Search test:', filterTest);

  // 3. Test Mobile Menu / Hamburger at 390px
  await page.setViewport({ width: 390, height: 844 });
  const menuTest = await page.evaluate(() => {
    const menuBtn = document.getElementById('mobileMenuBtn');
    const menuPanel = document.getElementById('mobileMenuPanel');
    if (!menuBtn || !menuPanel) return { ok: false };

    // Initially closed
    const initialExpanded = menuBtn.getAttribute('aria-expanded');
    const initialHidden = menuPanel.classList.contains('hidden');

    // Click to open
    menuBtn.click();
    const openExpanded = menuBtn.getAttribute('aria-expanded');
    const openHidden = menuPanel.classList.contains('hidden');

    // Press Escape to close
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    const escapeExpanded = menuBtn.getAttribute('aria-expanded');
    const escapeHidden = menuPanel.classList.contains('hidden');

    return {
      initialExpanded,
      initialHidden,
      openExpanded,
      openHidden,
      escapeExpanded,
      escapeHidden
    };
  });
  console.log('Mobile menu test:', menuTest);

  // 4. Test SSS Accordion
  const faqTest = await page.evaluate(() => {
    const firstFaq = document.querySelector('.faq-toggle');
    if (!firstFaq) return { ok: false };
    const content = firstFaq.nextElementSibling;
    const initialHidden = content.classList.contains('hidden');
    firstFaq.click();
    const afterClickHidden = content.classList.contains('hidden');
    return {
      initialHidden,
      afterClickHidden
    };
  });
  console.log('FAQ accordion test:', faqTest);

  // 5. Test Cart & Modal
  const cartTest = await page.evaluate(() => {
    const openCartBtn = document.getElementById('openCartBtn');
    const modal = document.getElementById('checkoutModal');
    if (!openCartBtn || !modal) return { ok: false };
    openCartBtn.click();
    const modalOpen = !modal.classList.contains('hidden');
    const itemCount = document.querySelectorAll('#cartItemsContainer > div').length;
    const closeBtn = document.getElementById('closeCheckoutModalBtn');
    if (closeBtn) closeBtn.click();
    const modalClosed = modal.classList.contains('hidden');
    return {
      modalOpen,
      itemCount,
      modalClosed
    };
  });
  console.log('Cart modal test:', cartTest);

  await browser.close();
  server.close();

  console.log('\n--- RESPONSIVE WIDTH RESULTS ---');
  results.forEach(r => {
    console.log(`[${r.width}px] Overflow: ${r.overflow.hasOverflow ? 'FAIL (' + r.overflow.diff + 'px)' : '0 (PASS)'} | Showcase: ${r.showcaseCount} cards | Footer cols: ${r.footerCheck.colCount}`);
    if (r.overflow.hasOverflow && r.overflow.badElements?.length) {
      console.log('   Bad elements:', JSON.stringify(r.overflow.badElements));
    }
  });

  console.log('\nConsole Errors count:', consoleErrors.length);
  if (consoleErrors.length > 0) {
    console.log('Errors:', consoleErrors);
  }
}

run().catch(err => {
  console.error('Test run failed:', err);
  process.exit(1);
});
