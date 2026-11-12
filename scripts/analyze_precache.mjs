import fs from 'fs';
import path from 'path';

function analyzePrecache() {
  const swPath = path.resolve('dist/sw.js');
  if (!fs.existsSync(swPath)) {
    console.log('dist/sw.js does not exist.');
    return;
  }
  const content = fs.readFileSync(swPath, 'utf8');
  const regex = /\{url:"([^"]+)",revision:([^}]+)\}/g;
  const entries = [];
  let m;
  while ((m = regex.exec(content)) !== null) {
    entries.push({ url: m[1], revision: m[2] });
  }
  console.log(`Total precache entries in dist/sw.js: ${entries.length}`);

  let totalSize = 0;
  const missing = [];
  const details = [];

  for (const entry of entries) {
    const filePath = path.resolve('dist', entry.url);
    if (fs.existsSync(filePath)) {
      const stat = fs.statSync(filePath);
      totalSize += stat.size;
      details.push({ url: entry.url, size: stat.size });
    } else {
      missing.push(entry.url);
    }
  }

  console.log(`Total precache uncompressed size: ${(totalSize / (1024 * 1024)).toFixed(2)} MB (${totalSize.toLocaleString()} B)`);
  if (missing.length > 0) {
    console.log(`Missing files in dist (${missing.length}):`, missing);
  }

  const wasmFiles = details.filter(d => d.url.endsWith('.wasm'));
  console.log('WASM files in precache:');
  for (const w of wasmFiles) {
    console.log(` - ${w.url}: ${(w.size / (1024 * 1024)).toFixed(2)} MB (${w.size.toLocaleString()} B)`);
  }
}

analyzePrecache();
