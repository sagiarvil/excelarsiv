'use strict';
/**
 * bom_utf8_scan.js — BOM'suz UTF-8 taraması (+ isteğe bağlı onarım, + cache temizliği).
 * Kullanım:
 *   node bom_utf8_scan.js [kök-dizin] [--fix] [--ext=php,json,html,js,css,xml,txt,md] [--cache=storage/cache] [--json]
 *   kök-dizin verilmezse otomatik proje kökü bulunur.
 */
const L = require('./_lib');
const { fs, path } = L;

const SKIP = new Set(['vendor', 'node_modules', '.git', 'dist', 'build', '.venv', 'storage']);

function walk(dir, cb) {
  for (const name of fs.readdirSync(dir)) {
    const p = path.join(dir, name);
    let st; try { st = fs.statSync(p); } catch { continue; }
    if (st.isDirectory()) { if (!SKIP.has(name)) walk(p, cb); }
    else cb(p);
  }
}

function main() {
  const args = L.parseArgs();
  const root = args._[0] ? path.resolve(args._[0]) : L.findProjectRoot();
  const exts = (args.ext || 'php,json,html,js,css,xml,txt,md').split(',').map((e) => '.' + e.trim().replace(/^\./, ''));
  const rep = L.makeReport('BOM / UTF-8 TARAMASI — ' + root);
  const hits = [];

  walk(root, (fp) => {
    if (!exts.includes(path.extname(fp).toLowerCase())) return;
    let buf; try { buf = fs.readFileSync(fp); } catch { return; }
    const rel = path.relative(root, fp);
    if (buf.length >= 3 && buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) {
      hits.push(rel);
      if (args.fix) { fs.writeFileSync(fp, buf.slice(3)); rep.pass('BOM-FIX', rel + ' → BOM silindi'); }
      else rep.fail('BOM', rel + ' BOM içeriyor');
    }
    // Tek satıra minify şüphesi (kaynak dosyada)
    if (['.php', '.js', '.css'].includes(path.extname(fp)) && buf.length > 3000 && !buf.includes(0x0A)) {
      rep.warn('MINIFY', rel + ' newline içermiyor (' + buf.length + ' bayt) — bozuk/minify olabilir');
    }
  });

  if (!hits.length) rep.pass('BOM', 'BOM içeren dosya yok');
  else if (!args.fix) rep.info('BOM', hits.length + ' dosyada BOM var — --fix ile onarılır');

  // Cache temizliği
  if (args.cache !== undefined) {
    const cd = path.join(root, typeof args.cache === 'string' ? args.cache : 'storage/cache');
    if (fs.existsSync(cd)) {
      let n = 0;
      for (const f of fs.readdirSync(cd)) {
        try { const fp = path.join(cd, f); if (fs.statSync(fp).isFile()) { fs.unlinkSync(fp); n++; } } catch {}
      }
      rep.info('CACHE', cd + ' → ' + n + ' dosya silindi');
    } else rep.warn('CACHE', 'cache dizini yok: ' + cd);
  }

  if (args.json) { console.log(JSON.stringify({ summary: rep.summary(), text: rep.toText(), files: hits }, null, 2)); process.exit(rep.summary().fail ? 1 : 0); }
  process.exit(rep.print());
}
main();
