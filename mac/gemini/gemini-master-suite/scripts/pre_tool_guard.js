// pre_tool_guard.js
'use strict';
const fs = require('fs');

try {
  const buf = Buffer.alloc(8192);
  let bytesRead = 0;
  try {
    bytesRead = fs.readSync(0, buf, 0, 8192, null);
  } catch (err) {}
  if (bytesRead > 0) {
    const raw = buf.toString('utf8', 0, bytesRead);
    const payload = JSON.parse(raw);
    const toolCall = payload.toolCall || {};
    const args = toolCall.args || {};
    const targetFile = (args.TargetFile || args.FilePath || args.AbsolutePath || '').toLowerCase();

    // 🛡️ Türkçe Mojibake ve Karakter Bütünlüğü Güvenlik Kalkanı
    if (args.CodeContent || args.ReplacementContent) {
      const content = args.CodeContent || args.ReplacementContent || '';
      if (/[ÝÐÞýðþ]/.test(content)) {
        const MOJIBAKE_FIX = { 'Ý': 'İ', 'Ð': 'Ğ', 'Þ': 'Ş', 'ý': 'ı', 'ð': 'ğ', 'þ': 'ş' };
        const fixed = content.replace(/[ÝÐÞýðþ]/g, m => MOJIBAKE_FIX[m] || m);
        if (args.CodeContent) args.CodeContent = fixed;
        if (args.ReplacementContent) args.ReplacementContent = fixed;
      }
    }

    const SENSITIVE_CORE = ['schema.sql', '.env', 'composer.lock', 'package-lock.json'];
    for (const s of SENSITIVE_CORE) {
      if (targetFile.endsWith(s)) {
        console.error('GÜVENLİK ENGELİ: ' + s + ' kısıtlıdır.');
        process.exit(1);
      }
    }
  }
} catch (e) {}
process.exit(0);
