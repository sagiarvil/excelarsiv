import fs from 'node:fs';
import path from 'node:path';

function updateAppJs(filePath) {
  let code = fs.readFileSync(filePath, 'utf8');

  // Replace submit listener
  const targetSubmit = `document.addEventListener('submit',e=>{const f=e.target;if(f.matches('[data-search-form]')){e.preventDefault();clearTimeout(timer);const q=new FormData(f).get('q')?.toString().trim().slice(0,120)||'';const path=f.querySelector('#catalogSearch')?current:'/sablonlar';navigate(path+'?q='+encodeURIComponent(q)+(kind==='model'?'&kind=model':''));return;}`;

  const replacementSubmit = `document.addEventListener('submit',e=>{const f=e.target;if(f.matches('[data-search-form]')){e.preventDefault();clearTimeout(timer);const q=new FormData(f).get('q')?.toString().trim().slice(0,120)||'';if($('catalogGrid')){query=q;limit=24;const inp=$('catalogSearch');if(inp)inp.value=q;setQueryUrl();renderCatalog();return;}const targetUrl=new URL('sablonlar/index.html?q='+encodeURIComponent(q)+(kind==='model'?'&kind=model':''),base).href;location.href=targetUrl;return;}`;

  if (code.includes(targetSubmit)) {
    code = code.replace(targetSubmit, replacementSubmit);
    fs.writeFileSync(filePath, code, 'utf8');
    console.log('Successfully updated search submit handler in:', filePath);
  } else {
    console.log('Could not find exact targetSubmit in:', filePath);
  }
}

updateAppJs('excelarsiv-premium-tum-sayfalar/site/assets/app.js');
if (fs.existsSync('excelarsiv-premium-tum-sayfalar/kaynak/app.js')) {
  updateAppJs('excelarsiv-premium-tum-sayfalar/kaynak/app.js');
}
