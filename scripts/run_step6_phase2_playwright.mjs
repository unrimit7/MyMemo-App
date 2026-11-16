// STEP6 2차 점검: 실제 앱에서 편집→저장→카드 반영, 원본 되돌리기(엑셀·PDF), 숨긴 행·열(K2), 큰 파일, PDF 내보내기 3방식을 확인하는 스크립트
// 실행: node scripts/run_step6_phase2_playwright.mjs  (PHASE=edit|k2|big|export 로 일부만, 5분 제한)
import { chromium } from 'playwright-core';
import { spawn, execSync } from 'child_process';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const outDir = path.join(root, 'docs', 'office-upgrade', 'out', 'step6');
const S = (...p) => path.join(root, 'docs', 'office-upgrade', 'samples', ...p);
const PORT = 5198;
const PHASES = (process.env.PHASE || 'edit,k2,export,big').split(',');
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const timer = setTimeout(() => { console.error('[TIMEOUT] 5분 초과, 강제 종료'); process.exit(2); }, 300000);
timer.unref();
const log = (...a) => console.log(`[${new Date().toISOString().slice(11, 19)}]`, ...a);
const results = {};
const rec = (kind, item, verdict, detail = '') => {
  results[kind] = results[kind] || {};
  results[kind][item] = { verdict, detail };
  log(`  ${verdict === 'OK' ? '✅' : verdict === 'FAIL' ? '❌' : '⚠️'} [${kind}] ${item}: ${detail}`);
};

// PREVIEW=1 이면 빌드본(dist)을 vite preview로 띄운다 (서비스 워커·오프라인 점검용)
function startVite() {
  const args = process.env.PREVIEW ? ['vite', 'preview', '--port', String(PORT), '--strictPort'] : ['vite', '--port', String(PORT), '--strictPort'];
  const proc = spawn('npx', args, { cwd: root, shell: true, stdio: ['ignore', 'pipe', 'pipe'] });
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error('Vite 30초 시간 초과')), 30000);
    // eslint-disable-next-line no-control-regex
    const onData = (d) => { const s = d.toString().replace(/\u001b\[[0-9;]*m/g, ''); if (s.includes(`localhost:${PORT}`) || s.includes('ready in')) { clearTimeout(t); resolve(proc); } };
    proc.stdout.on('data', onData); proc.stderr.on('data', onData);
    proc.on('exit', (c) => { clearTimeout(t); reject(new Error(`Vite 종료 code=${c}`)); });
  });
}
const killTree = (p) => { if (p) try { execSync(`taskkill /pid ${p.pid} /T /F`, { stdio: 'ignore' }); } catch { /* 이미 종료 */ } };
const dataUrl = (file, mime) => `data:${mime};base64,${fs.readFileSync(file).toString('base64')}`;
const b64 = (u) => Buffer.from(u.slice(u.indexOf(',') + 1), 'base64');

async function seed(page, atts) {
  let content = '# STEP6 2차 점검\n\n';
  for (const a of Object.values(atts)) content += `[${a.name}](attachment:${a.id})\n\n`;
  await page.goto(`http://localhost:${PORT}/`, { waitUntil: 'domcontentloaded' });
  await page.evaluate(async ({ atts, content }) => {
    await new Promise((resolve, reject) => {
      const req = indexedDB.open('MyMemo_DB', 1);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains('projects')) db.createObjectStore('projects', { keyPath: 'id' });
        if (!db.objectStoreNames.contains('notes')) { const s = db.createObjectStore('notes', { keyPath: 'id' }); s.createIndex('by-project', 'projectId'); s.createIndex('by-updated', 'updatedAt'); }
        if (!db.objectStoreNames.contains('settings')) db.createObjectStore('settings');
      };
      req.onsuccess = (e) => {
        const db = e.target.result;
        const tx = db.transaction(['projects', 'notes'], 'readwrite');
        tx.objectStore('projects').put({ id: 'proj-step6', name: 'STEP6', createdAt: '2026-11-16T00:00:00.000Z', isLocalOnly: true });
        tx.objectStore('notes').clear();
        tx.objectStore('notes').put({ id: 'note-step6', projectId: 'proj-step6', title: 'STEP6 2차', content, attachments: atts, isLocalOnly: true, syncStatus: 'synced', createdAt: '2026-11-16T00:00:00.000Z', updatedAt: '2026-11-16T00:00:00.000Z' });
        tx.oncomplete = resolve; tx.onerror = reject;
      };
      req.onerror = reject;
    });
    localStorage.setItem('mymemo_active_note_id', 'note-step6');
    localStorage.setItem('mymemo_active_project_id', 'ALL');
  }, { atts, content });
  await page.reload({ waitUntil: 'domcontentloaded' });
  await page.waitForSelector('.markdown-preview', { timeout: 20000 });
  await page.waitForTimeout(2500);
}

const readAtt = (page, id) => page.evaluate((id) => new Promise((res) => {
  const r = indexedDB.open('MyMemo_DB');
  r.onsuccess = () => { const g = r.result.transaction('notes').objectStore('notes').get('note-step6'); g.onsuccess = () => res(g.result?.attachments?.[id] || null); };
}), id);

const closeButtons = (page) => page.evaluate(() => Array.from(document.querySelectorAll('button[title*="닫기"]'))
  .filter((b) => { const r = b.getBoundingClientRect(); return r.width > 0 && !b.closest('.markdown-preview'); })
  .map((b) => { const r = b.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; }));
async function waitFor(fn, ms = 15000, step = 300) {
  const t = Date.now();
  while (Date.now() - t < ms) { if (await fn()) return true; await new Promise((r) => setTimeout(r, step)); }
  return false;
}
const hasText = (page, re) => page.evaluate((src) => Array.from(document.querySelectorAll('button, h3, p, span')).some((e) => e.getBoundingClientRect().width > 0 && new RegExp(src).test(e.textContent || '')), re);
async function clickText(page, text) {
  const loc = page.locator('button', { hasText: text }).last();
  if (!(await loc.count())) return false;
  await loc.click({ timeout: 5000 });
  return true;
}

// 편집기마다 버튼 이름이 다르다: '저장하지 않고 닫기'(엑셀·한글·PDF) / '저장 안 함'(워드)
async function discardAndClose(page) {
  if (await clickText(page, '저장하지 않고 닫기')) return true;
  return clickText(page, '저장 안 함');
}

async function excelValues(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const map = {};
  wb.eachSheet((ws) => ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (c) => {
    // 키 순서와 무관하게 비교한다
    const canon = (o) => JSON.stringify(o, (k, val) => (val && typeof val === 'object' && !Array.isArray(val) ? Object.fromEntries(Object.keys(val).sort().map((x) => [x, val[x]])) : val));
    const v = c.value && typeof c.value === 'object' ? canon(c.value.formula ? { f: c.value.formula } : c.value) : String(c.value);
    const slave = c.isMerged && c.master && c.master.address !== c.address;
    map[`${slave ? '병합내부:' : ''}${ws.name}!${c.address}`] = `${v}|${c.numFmt || ''}|${canon(c.border || {})}|${c.isMerged ? 'M' : ''}`;
  })));
  return map;
}

async function phaseEdit(page) {
  log('\n[엑셀] 편집 → 저장 → 카드 반영 → 원본 되돌리기');
  const file = S('real', '실파일B_견적서_MS엑셀.xlsx');
  const orig = dataUrl(file, XLSX);
  const pdfFile = S('step6', 'step6_가상_PDF_3쪽.pdf');
  const pdfOrig = dataUrl(pdfFile, 'application/pdf');
  await seed(page, {
    'att-x': { id: 'att-x', name: path.basename(file), type: XLSX, data: orig },
    'att-pdf': { id: 'att-pdf', name: path.basename(pdfFile), type: 'application/pdf', data: pdfOrig },
  });
  await page.locator('.excel-inline-viewer .excel-btn-edit').first().click();
  const ready = await waitFor(() => page.evaluate(() => typeof window.__excelIsReady === 'function' && window.__excelIsReady()), 30000);
  rec('excel', '편집기 준비', ready ? 'OK' : 'FAIL');
  await page.waitForTimeout(1500);
  // 사용자처럼 B2를 선택하고 글자를 친 뒤 Enter로 확정
  await typeCell(page, 'B2', 'STEP6수정');
  await page.waitForTimeout(800);
  await page.keyboard.press('Control+s');
  await page.waitForTimeout(1500);
  const warn = await page.evaluate(() => { const box = Array.from(document.querySelectorAll('div')).filter((d) => d.getBoundingClientRect().width > 0 && (d.innerText || '').startsWith('고급 서식 보존')).pop(); return box ? ['차트 (그래프)', '피벗 테이블', '삽입된 이미지 / 도형', 'VBA 매크로 코드', '스파크라인 (미니 차트)'].filter((t) => box.innerText.includes(t)) : []; });
  rec('excel', '저장 전 고급 요소 경고 내용', warn.length && !warn.some((t) => /매크로|이미지/.test(t)) ? 'OK' : 'WARN', JSON.stringify(warn));
  await clickText(page, '그래도 저장');
  await page.waitForTimeout(1500);
  const saved = await waitFor(async () => { const a = await readAtt(page, 'att-x'); return a && a.data !== orig; }, 20000);
  rec('excel', 'Ctrl+S 저장', saved ? 'OK' : 'FAIL');
  const att = await readAtt(page, 'att-x');
  rec('excel', '처음 저장 시 원본 보관', att?.originalData === orig ? 'OK' : 'FAIL', `originalData=${att?.originalData ? '있음' : '없음'}`);
  if (saved) {
    const buf = b64(att.data);
    fs.writeFileSync(path.join(outDir, 'step6_엑셀_B2수정_저장결과.xlsx'), buf);
    const [a, b] = [await excelValues(b64(orig)), await excelValues(buf)];
    const all = Object.keys({ ...a, ...b }).filter((k) => a[k] !== b[k]);
    const diff = all.filter((k) => !k.startsWith('병합내부:'));
    const inner = all.length - diff.length;
    rec('excel', '저장본 요소 비교 (값·수식·표시형식·테두리·병합)', diff.length === 1 && (b[diff[0]] || '').includes('STEP6수정') ? 'OK' : 'FAIL', `보이는 칸 차이 ${diff.length}개(${diff.join(', ')}), 병합 범위 내부(보이지 않는) 칸 서식 차이 ${inner}개`);
    const JSZip = (await import('jszip')).default;
    const names = Object.keys((await JSZip.loadAsync(buf)).files);
    rec('excel', '저장본 차트 보존', names.some((n) => n.startsWith('xl/charts/')) ? 'OK' : 'WARN', names.filter((n) => /chart|drawing/.test(n)).join(', ') || '차트 파일 없음');
  }
  // 닫기 (저장했으므로 확인 창 없어야 함)
  await waitFor(async () => (await closeButtons(page)).length > 0, 5000);
  let btn = (await closeButtons(page)).pop();
  if (btn) await page.mouse.click(btn.x, btn.y);
  await page.waitForTimeout(800);
  rec('excel', '저장 후 닫기 → 확인 창 없음', (await hasText(page, '저장하지 않고 닫')) ? 'FAIL' : 'OK');
  const cardHas = await waitFor(() => page.evaluate(() => (document.querySelector('.markdown-preview .excel-inline-viewer')?.textContent || '').includes('STEP6수정')), 10000);
  rec('excel', '카드 미리보기에 즉시 반영', cardHas ? 'OK' : 'FAIL');

  // 고치고 Enter 없이(값 입력만) 닫기 → 확인 창
  await page.locator('.markdown-preview .excel-inline-viewer .excel-btn-edit').first().click();
  await waitFor(() => page.evaluate(() => window.__excelIsReady && window.__excelIsReady()), 30000);
  await page.waitForTimeout(1500);
  await typeCell(page, 'B3', '미저장');
  await page.waitForTimeout(800);
  btn = (await closeButtons(page)).pop();
  if (btn) await page.mouse.click(btn.x, btn.y);
  await page.waitForTimeout(800);
  rec('excel', '고치고 닫기 → 확인 창', (await hasText(page, '저장하지 않고 닫')) ? 'OK' : 'FAIL');
  await discardAndClose(page);
  await page.waitForTimeout(800);

  // 원본으로 되돌리기
  await page.locator('.markdown-preview .excel-inline-viewer .excel-btn-edit').first().click();
  await waitFor(() => page.evaluate(() => window.__excelIsReady && window.__excelIsReady()), 30000);
  await page.waitForTimeout(1500);
  await clickText(page, '원본으로 되돌리기');
  await page.waitForTimeout(500);
  await clickText(page, '원본으로 되돌리기');
  const restored = await waitFor(async () => (await readAtt(page, 'att-x'))?.data === orig, 15000);
  rec('excel', '원본으로 되돌리기', restored ? 'OK' : 'FAIL');
  btn = (await closeButtons(page)).pop();
  if (btn) { await page.mouse.click(btn.x, btn.y); await page.waitForTimeout(500); await discardAndClose(page); }
  await page.waitForTimeout(800);

  log('\n[PDF] 주석 → 저장 → 카드 반영 → 원본 되돌리기');
  page.on('dialog', (d) => d.accept().catch(() => {}));
  const darkBefore = await pdfDark(page);
  await page.locator('.markdown-preview .pdf-btn-edit').first().click();
  await page.waitForTimeout(4000);
  const canvas = await page.evaluate(() => {
    const cs = Array.from(document.querySelectorAll('canvas')).filter((c) => !c.closest('.markdown-preview')).map((c) => { const r = c.getBoundingClientRect(); return { x: r.x, y: r.y, w: r.width, h: r.height }; }).filter((r) => r.w > 200);
    return cs[cs.length - 1];
  });
  if (canvas) {
    // 화면 안에 보이는 위쪽(제목 아래)에 지그재그 선을 긋는다
    const y0 = Math.min(canvas.y + 150, 600);
    await page.mouse.move(canvas.x + canvas.w * 0.2, y0);
    await page.mouse.down();
    for (let i = 1; i <= 20; i++) await page.mouse.move(canvas.x + canvas.w * (0.2 + i * 0.03), y0 + (i % 2) * 20, { steps: 3 });
    await page.mouse.up();
  }
  await page.waitForTimeout(500);
  await clickText(page, '저장 및 반영');
  const pSaved = await waitFor(async () => { const a = await readAtt(page, 'att-pdf'); return a && a.data !== pdfOrig; }, 20000);
  const pAtt = await readAtt(page, 'att-pdf');
  rec('pdf', '주석 저장', pSaved ? 'OK' : 'FAIL', canvas ? '' : '캔버스 못 찾음');
  rec('pdf', '처음 저장 시 원본 보관', pAtt?.originalData === pdfOrig ? 'OK' : 'FAIL');
  if (pSaved) fs.writeFileSync(path.join(outDir, 'step6_PDF_주석_저장결과.pdf'), b64(pAtt.data));
  const hasRestoreBtn = await page.locator('button', { hasText: '원본으로 되돌리기' }).count();
  rec('pdf', '편집기에 원본으로 되돌리기 버튼', hasRestoreBtn ? 'OK' : 'FAIL');
  btn = (await closeButtons(page)).pop();
  if (btn) await page.mouse.click(btn.x, btn.y);
  await page.waitForTimeout(800);
  rec('pdf', '저장 후 닫기 → 확인 창 없음', (await hasText(page, '저장하지 않고 닫')) ? 'FAIL' : 'OK');
  await page.waitForTimeout(1500);
  const darkAfter = await pdfDark(page);
  rec('pdf', '카드 미리보기에 즉시 반영', darkAfter > darkBefore ? 'OK' : 'FAIL', `어두운 픽셀 ${darkBefore}→${darkAfter}`);
  await page.locator('.markdown-preview .pdf-btn-edit').first().click();
  await page.waitForTimeout(3000);
  await clickText(page, '원본으로 되돌리기');
  const pRestored = await waitFor(async () => (await readAtt(page, 'att-pdf'))?.data === pdfOrig, 10000);
  rec('pdf', '원본으로 되돌리기', pRestored ? 'OK' : 'FAIL');
  btn = (await closeButtons(page)).pop();
  if (btn) await page.mouse.click(btn.x, btn.y);
  await page.waitForTimeout(1500);
  const darkRestored = await pdfDark(page);
  rec('pdf', '되돌린 뒤 카드 반영', darkRestored === darkBefore ? 'OK' : 'WARN', `어두운 픽셀 ${darkRestored} (원본 ${darkBefore})`);
  // 주석을 그리고 저장 없이 닫기 → 확인 창, 저장하지 않고 닫으면 첨부 그대로
  await page.locator('.markdown-preview .pdf-btn-edit').first().click();
  await page.waitForTimeout(4000);
  const cv2 = await page.evaluate(() => {
    const cs = Array.from(document.querySelectorAll('canvas')).filter((c) => !c.closest('.markdown-preview')).map((c) => { const r = c.getBoundingClientRect(); return { x: r.x, y: r.y, w: r.width }; }).filter((r) => r.w > 200);
    return cs[cs.length - 1];
  });
  if (cv2) {
    const y0 = Math.min(cv2.y + 150, 600);
    await page.mouse.move(cv2.x + 100, y0); await page.mouse.down();
    for (let i = 1; i <= 10; i++) await page.mouse.move(cv2.x + 100 + i * 15, y0 + (i % 2) * 15, { steps: 2 });
    await page.mouse.up();
  }
  await page.waitForTimeout(500);
  btn = (await closeButtons(page)).pop();
  if (btn) await page.mouse.click(btn.x, btn.y);
  await page.waitForTimeout(800);
  const askedPdf = await hasText(page, '저장하지 않고 닫');
  rec('pdf', '고치고 닫기 → 확인 창', askedPdf ? 'OK' : 'FAIL');
  if (askedPdf) await discardAndClose(page);
  await page.waitForTimeout(800);
  rec('pdf', '저장하지 않고 닫은 뒤 첨부 그대로', (await readAtt(page, 'att-pdf'))?.data === pdfOrig ? 'OK' : 'FAIL');
}

async function typeCell(page, a1, text) {
  // Univer 캔버스에 포커스를 준 다음 셀을 선택하고 입력
  const box = await page.evaluate(() => {
    const cs = Array.from(document.querySelectorAll('canvas')).filter((c) => !c.closest('.markdown-preview')).map((c) => c.getBoundingClientRect()).filter((r) => r.width > 300);
    const r = cs[0]; return r ? { x: r.x + r.width / 2, y: r.y + r.height / 2 } : null;
  });
  if (box) await page.mouse.click(box.x, box.y);
  await page.evaluate((a1) => window.__excelUniverAPI.getActiveWorkbook().getActiveSheet().getRange(a1).activate(), a1);
  await page.waitForTimeout(300);
  await page.keyboard.type(text, { delay: 30 });
  await page.keyboard.press('Enter');
  await page.waitForTimeout(500);
}

const pdfDark = (page) => page.evaluate(() => {
  const c = document.querySelector('.markdown-preview canvas.pdf-canvas');
  if (!c || !c.width) return -1;
  const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
  let k = 0; for (let i = 0; i < d.length; i += 16) if (d[i] < 128 || d[i + 1] < 128) k++;
  return k;
});

async function phaseK2(page) {
  log('\n[엑셀] 숨긴 행·열 (K2)');
  const file = S('step6', 'step6_가상_숨김행열.xlsx');
  await seed(page, { 'att-k2': { id: 'att-k2', name: path.basename(file), type: XLSX, data: dataUrl(file, XLSX) } });
  const card = await page.evaluate(() => {
    const t = document.querySelector('.markdown-preview .excel-inline-viewer')?.textContent || '';
    return { hiddenColShown: t.includes('숨김값1'), hiddenRowShown: t.includes('가상인물3') };
  });
  rec('excel', 'K2 카드: 숨긴 열/행 표시 안 함', !card.hiddenColShown && !card.hiddenRowShown ? 'OK' : 'WARN', JSON.stringify(card));
  await page.locator('.markdown-preview .excel-btn-edit').first().click();
  await waitFor(() => page.evaluate(() => window.__excelIsReady && window.__excelIsReady()), 30000);
  await page.waitForTimeout(1500);
  const snap = await page.evaluate(() => {
    const s = window.__excelUniverAPI.getActiveWorkbook().save();
    const sh = Object.values(s.sheets)[0];
    return { row3: sh.rowData?.[3]?.hd ?? null, row4: sh.rowData?.[4]?.hd ?? null, col2: sh.columnData?.[2]?.hd ?? null };
  });
  rec('excel', 'K2 편집기: 숨긴 행·열 유지', snap.row3 === 1 && snap.row4 === 1 && snap.col2 === 1 ? 'OK' : 'FAIL', JSON.stringify(snap));
  // 다른 칸을 고쳐 저장해도 숨김이 그대로인지
  const orig = (await readAtt(page, 'att-k2')).data;
  await typeCell(page, 'B2', 'K2수정');
  await page.keyboard.press('Control+s');
  await page.waitForTimeout(1500);
  await clickText(page, '그래도 저장');
  const saved = await waitFor(async () => (await readAtt(page, 'att-k2'))?.data !== orig, 20000);
  if (saved) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(b64((await readAtt(page, 'att-k2')).data));
    const ws = wb.worksheets[0];
    const h = { row4: ws.getRow(4).hidden, row5: ws.getRow(5).hidden, colC: ws.getColumn(3).hidden };
    rec('excel', 'K2 저장 후에도 숨김 유지', h.row4 && h.row5 && h.colC ? 'OK' : 'FAIL', JSON.stringify(h));
  } else rec('excel', 'K2 저장 후에도 숨김 유지', 'FAIL', '저장 안 됨');
  const btn = (await closeButtons(page)).pop();
  if (btn) await page.mouse.click(btn.x, btn.y);
  await page.waitForTimeout(1500);
  const card2 = await page.evaluate(() => (document.querySelector('.markdown-preview .excel-inline-viewer')?.textContent || ''));
  rec('excel', 'K2 저장 후 카드: 숨긴 내용 안 보임 + 수정 반영', !card2.includes('숨김값') && card2.includes('K2수정') ? 'OK' : 'FAIL', `숨김값 보임=${card2.includes('숨김값')}, 수정 반영=${card2.includes('K2수정')}`);
  fs.writeFileSync(path.join(outDir, 'step6_엑셀_숨김행열_수정저장결과.xlsx'), b64((await readAtt(page, 'att-k2')).data));
}

async function phaseExport(page) {
  log('\n[PDF 내보내기] 3가지 방식');
  const files = [
    ['att-x', S('real', '실파일B_견적서_MS엑셀.xlsx'), XLSX],
    ['att-d', S('real', '워드A_사업보고서_원본.docx'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    ['att-hx', S('real', '한글C_업무계획_원본.hwpx'), 'application/haansofthwpx'],
    ['att-p', S('real', 'PPT_A_사업보고_원본.pptx'), 'application/vnd.openxmlformats-officedocument.presentationml.presentation'],
    ['att-pdf', S('step6', 'step6_가상_PDF_3쪽.pdf'), 'application/pdf'],
  ];
  const atts = {};
  for (const [id, f, mime] of files) atts[id] = { id, name: path.basename(f), type: mime, data: dataUrl(f, mime) };
  await seed(page, atts);
  await page.evaluate(() => {
    window.__printFrames = [];
    const orig = document.body.appendChild.bind(document.body);
    document.body.appendChild = (n) => { if (n.tagName === 'IFRAME') window.__printFrames.push(n); return orig(n); };
  });
  for (const mode of ['expanded', 'all', 'none']) {
    const label = { expanded: '화면에서 펼친 것만', all: '모두 펼쳐서 포함', none: '첨부 내용 빼기' }[mode];
    await page.locator('button[title^="현재 메모를 깔끔한 PDF"]').first().click();
    await page.waitForTimeout(500);
    await page.locator('label', { hasText: label }).first().click();
    const before = await page.evaluate(() => window.__printFrames.length);
    await page.locator('button', { hasText: '내보내기' }).last().click();
    const got = await waitFor(() => page.evaluate((b) => {
      const f = window.__printFrames[b];
      return !!f && (f.contentDocument?.body?.innerText || '').length > 50;
    }, before), 60000, 500);
    await page.waitForTimeout(3000);
    const m = await page.evaluate((b) => {
      const doc = window.__printFrames[b]?.contentDocument;
      if (!doc) return null;
      const txt = doc.body.innerText;
      return {
        len: txt.length,
        excelTable: doc.querySelectorAll('table').length,
        imgs: doc.querySelectorAll('img').length,
        svg: doc.querySelectorAll('svg').length,
        canvases: doc.querySelectorAll('canvas').length,
        hasPpt: /사업|보고/.test(txt),
      };
    }, before);
    rec('export', `PDF 내보내기: ${label}`, got && m ? 'OK' : 'FAIL', JSON.stringify(m));
  }
}

async function focusDoc(page, kind) {
  if (kind === 'docx') {
    // SuperDoc은 쪽 화면 위의 글자를 눌러야 커서가 들어간다. '1. 개요' 아래 첫 문단을 누른다
    const box = await page.evaluate(() => {
      // 화면에 실제로 보이는 것 중 가장 왼쪽 것 (측정용 숨은 요소 제외)
      const rs = Array.from(document.querySelectorAll('span, p, div')).filter((e) => !e.closest('.markdown-preview') && e.children.length === 0 && (e.textContent || '').startsWith('본 보고서는'))
        .map((e) => e.getBoundingClientRect()).filter((q) => q.width > 50 && q.height > 5 && q.y > 0 && q.y < innerHeight).sort((p1, p2) => p1.x - p2.x);
      const r = rs[0]; return r ? { x: r.x + 40, y: r.y + r.height / 2 } : null;
    });
    log('  워드 클릭 위치:', JSON.stringify(box));
    if (box) await page.mouse.click(box.x, box.y);
  } else {
    const fr = page.locator('iframe').last();
    const bb = await fr.boundingBox();
    if (bb) await page.mouse.click(bb.x + bb.width / 2, bb.y + 200);
  }
}

async function phaseDoc(page) {
  const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const dFile = S('real', '워드A_사업보고서_원본.docx');
  const hFile = S('real', '한글C_업무계획_원본.hwpx');
  const dOrig = dataUrl(dFile, DOCX);
  const hOrig = dataUrl(hFile, 'application/haansofthwpx');
  const h5File = S('real', '한글C_업무계획_원본.hwp');
  const h5Orig = dataUrl(h5File, 'application/x-hwp');
  await seed(page, {
    'att-d': { id: 'att-d', name: path.basename(dFile), type: DOCX, data: dOrig },
    'att-hx': { id: 'att-hx', name: path.basename(hFile), type: 'application/haansofthwpx', data: hOrig },
    'att-h': { id: 'att-h', name: path.basename(h5File), type: 'application/x-hwp', data: h5Orig },
  });
  page.on('dialog', (d) => d.accept().catch(() => {}));

  for (const [kind, id, orig, editSel, outName] of [
    ['docx', 'att-d', dOrig, '.docx-btn-edit', 'step6_워드_수정저장결과.docx'],
    ['hwpx', 'att-hx', hOrig, '.hwp-btn-edit', 'step6_한글_수정저장결과.hwpx'],
    ['hwp', 'att-h', h5Orig, '.hwp-btn-edit', 'step6_한글_수정저장결과.hwp'],
  ]) {
    log(`
[${kind}] 편집 → 저장 → 카드 반영 → 원본 되돌리기`);
    await page.locator(`.markdown-preview [data-att-id="${id}"] ${editSel}`).first().click();
    await page.waitForTimeout(kind === 'docx' ? 14000 : 8000);
    await focusDoc(page, kind);
    await page.waitForTimeout(500);
    await page.keyboard.press('End');
    await page.keyboard.type('STEP6', { delay: 50 });
    await page.waitForTimeout(1500);
    await clickText(page, '저장 및 반영');
    await page.waitForTimeout(1500);
    await clickText(page, '그래도 저장');
    const saved = await waitFor(async () => { const a = await readAtt(page, id); return a && a.data !== orig; }, 30000);
    const att = await readAtt(page, id);
    rec(kind, '편집 후 저장', saved ? 'OK' : 'FAIL');
    rec(kind, '처음 저장 시 원본 보관', att?.originalData === orig ? 'OK' : 'FAIL');
    if (saved) fs.writeFileSync(path.join(outDir, outName), b64(att.data));
    let btn = (await closeButtons(page)).pop();
    if (btn) await page.mouse.click(btn.x, btn.y);
    await page.waitForTimeout(1000);
    rec(kind, '저장 후 닫기 → 확인 창 없음', (await hasText(page, '저장하지 않고 닫')) ? 'FAIL' : 'OK');
    if (await hasText(page, '저장하지 않고 닫')) await discardAndClose(page);
    await page.waitForTimeout(3000);
    const cardHas = await page.evaluate((id) => {
      const v = document.querySelector(`.markdown-preview [data-att-id="${id}"]`);
      return (v?.textContent || '').includes('STEP6') || Array.from(v?.querySelectorAll('text') || []).map((t) => t.textContent).join('').includes('STEP6');
    }, id);
    rec(kind, '카드 미리보기에 즉시 반영', cardHas ? 'OK' : 'FAIL');
    // 원본으로 되돌리기
    await page.locator(`.markdown-preview [data-att-id="${id}"] ${editSel}`).first().click();
    await page.waitForTimeout(7000);
    // 편집기마다 버튼 이름이 조금 다르다 (원본 되돌리기 / 최초 원본 되돌리기 → 확인 창의 원본으로 되돌리기)
    await page.locator('button', { hasText: /원본 되돌리기/ }).first().click().catch(() => {});
    await page.waitForTimeout(700);
    await clickText(page, '원본으로 되돌리기');
    await page.waitForTimeout(1500);
    const restored = await waitFor(async () => (await readAtt(page, id))?.data === orig, 15000);
    rec(kind, '원본으로 되돌리기', restored ? 'OK' : 'FAIL');
    btn = (await closeButtons(page)).pop();
    if (btn) { await page.mouse.click(btn.x, btn.y); await page.waitForTimeout(600); if (await hasText(page, '저장하지 않고 닫')) await discardAndClose(page); }
    await page.waitForTimeout(1000);
    // 고치고(저장 없이) 닫기 → 확인 창이 떠야 하고, 저장하지 않고 닫으면 첨부는 그대로
    await page.locator(`.markdown-preview [data-att-id="${id}"] ${editSel}`).first().click();
    await page.waitForTimeout(kind === 'docx' ? 9000 : 8000);
    await focusDoc(page, kind);
    await page.waitForTimeout(500);
    await page.keyboard.press('End');
    await page.keyboard.type('X', { delay: 50 });
    await page.waitForTimeout(1500);
    btn = (await closeButtons(page)).pop();
    if (btn) await page.mouse.click(btn.x, btn.y);
    await page.waitForTimeout(800);
    const asked = await hasText(page, '저장하지 않고 닫');
    rec(kind, '고치고 닫기 → 확인 창', asked ? 'OK' : 'FAIL');
    if (asked) await discardAndClose(page);
    await page.waitForTimeout(1000);
    rec(kind, '저장하지 않고 닫은 뒤 첨부 그대로', (await readAtt(page, id))?.data === orig ? 'OK' : 'FAIL');
  }
}

async function phaseOffline(page) {
  log('\n[오프라인] 빌드본 + 서비스 워커, 네트워크 끊고 열기·편집기 열기');
  const files = [
    ['att-x', S('real', '실파일B_견적서_MS엑셀.xlsx'), XLSX, '.excel-inline-viewer', '.excel-btn-edit'],
    ['att-d', S('real', '워드A_사업보고서_원본.docx'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.docx-inline-viewer', '.docx-btn-edit'],
    ['att-hx', S('real', '한글C_업무계획_원본.hwpx'), 'application/haansofthwpx', '.hwp-inline-viewer', '.hwp-btn-edit'],
    ['att-p', S('real', 'PPT_A_사업보고_원본.pptx'), 'application/vnd.openxmlformats-officedocument.presentationml.presentation', '.pptx-inline-viewer', '.pptx-btn-fullscreen'],
    ['att-pdf', S('step6', 'step6_가상_PDF_3쪽.pdf'), 'application/pdf', '.pdf-inline-viewer', '.pdf-btn-edit'],
  ];
  const atts = {};
  for (const [id, f, mime] of files) atts[id] = { id, name: path.basename(f), type: mime, data: dataUrl(f, mime) };
  await seed(page, atts);
  const swReady = await waitFor(() => page.evaluate(async () => !!(navigator.serviceWorker && navigator.serviceWorker.controller)), 60000, 1000);
  if (!swReady) { await page.reload(); await page.waitForTimeout(3000); }
  // 미리 캐시(precache) 내려받기가 끝날 때까지 기다린다
  await waitFor(() => page.evaluate(async () => { const r = await navigator.serviceWorker.getRegistration(); return !!r?.active && !r.installing && !r.waiting; }), 90000, 1000);
  await page.waitForTimeout(3000);
  rec('offline', '서비스 워커 활성', (await page.evaluate(() => !!navigator.serviceWorker.controller)) ? 'OK' : 'FAIL');
  await page.context().setOffline(true);
  await page.reload({ waitUntil: 'domcontentloaded' });
  const loaded = await waitFor(() => page.evaluate(() => !!document.querySelector('.markdown-preview')), 20000);
  rec('offline', '오프라인 새로고침 후 앱 열림', loaded ? 'OK' : 'FAIL');
  await page.waitForTimeout(4000);
  for (const [id, , , viewer, btnSel] of files) {
    const kind = id.replace('att-', '');
    const ok = await page.evaluate(({ id, viewer }) => {
      const v = document.querySelector(`.markdown-preview ${viewer}[data-att-id="${id}"]`);
      if (!v) return false;
      return !!(v.querySelector('td') || v.querySelector('svg') || v.querySelector('canvas:not(.hidden)') || (v.querySelector('.docx-container')?.innerText || '').length > 20 || v.querySelector('.pptx-container'));
    }, { id, viewer });
    rec('offline', `${kind} 카드 표시`, ok ? 'OK' : 'FAIL');
    const before = (await closeButtons(page)).length;
    await page.locator(`.markdown-preview ${viewer}[data-att-id="${id}"] ${btnSel}`).first().click().catch(() => {});
    const opened = await waitFor(async () => (await closeButtons(page)).length > before, 30000);
    await page.waitForTimeout(6000);
    const errText = await page.evaluate(() => /실패|오류|불러오지 못|Failed|error/i.test(Array.from(document.querySelectorAll('[class*="fixed"]')).map((e) => e.innerText).join(' ').slice(0, 5000)));
    rec('offline', `${kind} 편집기(PPT는 큰 화면) 열기`, opened && !errText ? 'OK' : 'FAIL', `열림=${opened}, 오류문구=${errText}`);
    await page.screenshot({ path: path.join(outDir, `offline_${kind}.png`) });
    const b = (await closeButtons(page)).pop();
    if (b) { await page.mouse.click(b.x, b.y); await page.waitForTimeout(600); if (await hasText(page, '저장하지 않고 닫')) await discardAndClose(page); }
    await page.waitForTimeout(800);
  }
  await page.context().setOffline(false);
}

async function phaseBig(page) {
  log('\n[엑셀] 큰 파일 (약 10MB, 14시트 × 15,000행)');
  const file = S('step6', 'step6_가상_큰엑셀.xlsx');
  if (!fs.existsSync(file)) { rec('excel', '큰 파일', 'WARN', '시험 파일 없음: node scripts/make_step6_samples.mjs 실행 필요'); return; }
  const t0 = Date.now();
  await seed(page, { 'att-big': { id: 'att-big', name: path.basename(file), type: XLSX, data: dataUrl(file, XLSX) } });
  const shown = await waitFor(() => page.evaluate(() => document.querySelectorAll('.markdown-preview .excel-inline-viewer td').length > 20), 60000, 500);
  rec('excel', '큰 파일 카드 표시', shown ? 'OK' : 'FAIL', `${((Date.now() - t0) / 1000).toFixed(1)}초`);
  const t1 = Date.now();
  await page.evaluate(() => 1);
  rec('excel', '큰 파일 표시 뒤 화면 응답', Date.now() - t1 < 2000 ? 'OK' : 'FAIL', `${Date.now() - t1}ms`);
  const t2 = Date.now();
  await page.locator('.markdown-preview .excel-btn-fullscreen').first().click();
  const opened = await waitFor(async () => (await closeButtons(page)).length > 0, 60000, 500);
  rec('excel', '큰 파일 크게 보기', opened ? 'OK' : 'FAIL', `${((Date.now() - t2) / 1000).toFixed(1)}초`);
}

async function run() {
  fs.mkdirSync(outDir, { recursive: true });
  let vite = null; let browser = null; let code = 1;
  try {
    vite = await startVite();
    browser = await chromium.launch({ executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe', headless: true });
    for (const ph of PHASES) {
      const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
      const page = await ctx.newPage();
      page.on('pageerror', (e) => log('  [pageerror]', String(e).slice(0, 150)));
      try {
        if (ph === 'edit') await phaseEdit(page);
        if (ph === 'k2') await phaseK2(page);
        if (ph === 'export') await phaseExport(page);
        if (ph === 'big') await phaseBig(page);
        if (ph === 'doc') await phaseDoc(page);
        if (ph === 'offline') await phaseOffline(page);
      } catch (e) { rec(ph, '예외', 'FAIL', String(e).slice(0, 200)); await page.screenshot({ path: path.join(outDir, 'dbg_exception_' + ph + '.png') }).catch(() => {}); }
      await ctx.close();
    }
    code = 0;
  } catch (e) {
    console.error('[오류]', e);
  } finally {
    fs.writeFileSync(path.join(outDir, `phase2_result_${PHASES.join('_')}.json`), JSON.stringify(results, null, 2));
    if (browser) await browser.close().catch(() => {});
    killTree(vite);
    process.exit(code);
  }
}
run();
