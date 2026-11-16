// STEP6 전반 점검: 실제 앱(NoteEditor + IndexedDB)에서 형식별 카드·버튼·뷰어·닫기·팝아웃을 눌러 보고 결과를 JSON으로 남기는 스크립트
// 실행: node scripts/run_step6_audit_playwright.mjs  (5분 제한, try/finally로 서버·브라우저 정리)
import { chromium } from 'playwright-core';
import { spawn, execSync } from 'child_process';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const outDir = path.join(root, 'docs', 'office-upgrade', 'out', 'step6');
const S = (...p) => path.join(root, 'docs', 'office-upgrade', 'samples', ...p);
const PORT = 5197;
const ONLY = process.env.STEP6_ONLY ? process.env.STEP6_ONLY.split(',') : null;

const timer = setTimeout(() => { console.error('[TIMEOUT] 5분 초과, 강제 종료'); process.exit(2); }, 300000);
timer.unref();

const log = (...a) => console.log(`[${new Date().toISOString().slice(11, 19)}]`, ...a);

const FILES = [
  { kind: 'excel', id: 'att-x', file: S('real', '실파일B_견적서_MS엑셀.xlsx'), mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
  { kind: 'docx', id: 'att-d', file: S('real', '워드A_사업보고서_원본.docx'), mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  { kind: 'hwpx', id: 'att-hx', file: S('real', '한글C_업무계획_원본.hwpx'), mime: 'application/haansofthwpx' },
  { kind: 'hwp', id: 'att-h', file: S('real', '한글C_업무계획_원본.hwp'), mime: 'application/x-hwp' },
  { kind: 'pptx', id: 'att-p', file: S('real', 'PPT_A_사업보고_원본.pptx'), mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
  { kind: 'pdf', id: 'att-pdf', file: S('step6', 'step6_가상_PDF_3쪽.pdf'), mime: 'application/pdf' },
];

const CARD = {
  excel: { viewer: '.excel-inline-viewer', full: '.excel-btn-fullscreen', edit: '.excel-btn-edit', imp: '.excel-btn-import' },
  docx: { viewer: '.docx-inline-viewer', full: '.docx-btn-fullscreen', edit: '.docx-btn-edit', imp: '.docx-btn-import' },
  hwpx: { viewer: '.hwp-inline-viewer', full: '.hwp-btn-fullscreen', edit: '.hwp-btn-edit', imp: '.hwp-btn-import' },
  hwp: { viewer: '.hwp-inline-viewer', full: '.hwp-btn-fullscreen', edit: '.hwp-btn-edit', imp: '.hwp-btn-import' },
  pptx: { viewer: '.pptx-inline-viewer', full: '.pptx-btn-fullscreen', edit: '.pptx-btn-edit', imp: '.pptx-btn-import' },
  pdf: { viewer: '.pdf-inline-viewer', full: '.pdf-btn-fullscreen', edit: '.pdf-btn-edit', imp: '.pdf-btn-import' },
};

function startVite() {
  const proc = spawn('npx', ['vite', '--port', String(PORT), '--strictPort'], { cwd: root, shell: true, stdio: ['ignore', 'pipe', 'pipe'] });
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error('Vite 30초 시간 초과')), 30000);
    const onData = (d) => { const t2 = d.toString().replace(/[[0-9;]*m/g, ""); if (t2.includes(`localhost:${PORT}`) || t2.includes("ready in")) { clearTimeout(t); resolve(proc); } };
    proc.stdout.on('data', onData);
    proc.stderr.on('data', onData);
    proc.on('exit', (c) => { clearTimeout(t); reject(new Error(`Vite 종료 code=${c}`)); });
  });
}

function killTree(proc) {
  if (!proc) return;
  try { execSync(`taskkill /pid ${proc.pid} /T /F`, { stdio: 'ignore' }); } catch { /* 이미 종료 */ }
}

const results = {};
const rec = (kind, item, verdict, detail = '') => {
  results[kind] = results[kind] || {};
  results[kind][item] = { verdict, detail };
  log(`  ${verdict === 'OK' ? '✅' : verdict === 'FAIL' ? '❌' : '⚠️'} [${kind}] ${item}: ${detail}`);
};

async function visibleCloseButtons(page) {
  return page.evaluate(() => Array.from(document.querySelectorAll('button[title*="닫기"]'))
    .filter((b) => { const r = b.getBoundingClientRect(); return r.width > 0 && r.height > 0 && !b.closest('.markdown-preview'); })
    .map((b) => { const r = b.getBoundingClientRect(); return { x: r.x, y: r.y, w: r.width, h: r.height, vw: innerWidth, vh: innerHeight }; }));
}

async function hasExitConfirm(page) {
  return page.evaluate(() => Array.from(document.querySelectorAll('button, h3, p')).some((el) => {
    const r = el.getBoundingClientRect();
    return r.width > 0 && /저장하지 않고 닫/.test(el.textContent || '');
  }));
}

async function waitModal(page, before, ms = 15000) {
  const start = Date.now();
  while (Date.now() - start < ms) {
    const n = (await visibleCloseButtons(page)).length;
    if (n > before) return true;
    await page.waitForTimeout(250);
  }
  return false;
}

async function waitNoModal(page, count, ms = 5000) {
  const start = Date.now();
  while (Date.now() - start < ms) {
    if ((await visibleCloseButtons(page)).length <= count) return true;
    await page.waitForTimeout(200);
  }
  return false;
}

function judgeCard(kind, m) {
  if (!m.exists) return ['FAIL', '카드 없음'];
  if (kind === 'excel') return m.filled > 5 ? ['OK', `칸 ${m.cells}개 중 내용 ${m.filled}, 테두리 style ${m.bordered}`] : ['FAIL', JSON.stringify(m)];
  if (kind === 'docx') return m.visible && m.textLen > 50 ? ['OK', `글자 ${m.textLen}, 표 ${m.tables}, 그림 ${m.imgs}`] : ['FAIL', JSON.stringify(m)];
  if (kind === 'hwp' || kind === 'hwpx') return m.visible && m.svg > 0 ? ['OK', `svg ${m.svg}, text ${m.textNodes}, ${m.page}`] : ['FAIL', JSON.stringify(m)];
  if (kind === 'pdf') return m.visible && m.darkPixels > 50 ? ['OK', `canvas ${m.cw}x${m.ch}, 글자 픽셀 ${m.darkPixels}, ${m.page}`] : ['FAIL', JSON.stringify(m)];
  if (kind === 'pptx') return m.slide ? ['OK', `첫 슬라이드 ${m.w}x${m.h}, 요소 ${m.elements}, 글자 ${m.textLen}, 전체 ${m.count}장`] : ['FAIL', '실제 첫 슬라이드가 그려지지 않음 (안내 그림)'];
  return ['WARN', JSON.stringify(m)];
}

async function clickIn(page, scope, sel) {
  const loc = page.locator(scope ? `${scope} ${sel}` : sel).first();
  if (!(await loc.count())) return false;
  await loc.scrollIntoViewIfNeeded().catch(() => {});
  await loc.click({ timeout: 5000 });
  return true;
}

async function checkCloseAtWidths(page, kind, label) {
  const sizes = [800, 1000, 1280];
  const bad = [];
  for (const w of sizes) {
    await page.setViewportSize({ width: w, height: 800 });
    await page.waitForTimeout(300);
    const btns = await visibleCloseButtons(page);
    const top = btns[btns.length - 1];
    if (!top || top.x < 0 || top.y < 0 || top.x + top.w > w + 1 || top.y + top.h > 801) bad.push(`${w}px:${JSON.stringify(top)}`);
  }
  await page.setViewportSize({ width: 1400, height: 900 });
  rec(kind, `${label} 닫기 버튼 800/1000/1280`, bad.length ? 'FAIL' : 'OK', bad.join(' ') || '세 너비 모두 화면 안');
}

async function auditMain(page, f) {
  const { kind } = f;
  const c = CARD[kind];
  const scope = `.markdown-preview ${c.viewer}[data-att-id="${f.id}"]`;
  log(`\n[${kind}] 메인 창 점검`);
  log('  DOM 카드:', JSON.stringify(await page.evaluate(() => Array.from(document.querySelectorAll('.markdown-preview [class*="-inline-viewer"]')).map((e) => `${(e.className.match(/\w+-inline-viewer/) || [''])[0]}:${e.dataset.attId}`))));
  await page.locator(scope).first().scrollIntoViewIfNeeded().catch(() => {});
  // 카드 렌더 대기 (최대 20초)
  let m;
  for (let i = 0; i < 40; i++) {
    m = await cardMetricsById(page, f).catch(() => ({ exists: false }));
    const [v0] = judgeCard(kind, m);
    if (v0 === 'OK') break;
    await page.waitForTimeout(500);
  }
  const [v, d] = judgeCard(kind, m);
  rec(kind, '카드 미리보기', v, d);
  await page.screenshot({ path: path.join(outDir, `card_${kind}.png`), clip: await page.locator(scope).first().boundingBox().catch(() => null) || undefined }).catch(() => {});

  // 크게 보기
  const before = (await visibleCloseButtons(page)).length;
  await clickIn(page, scope, c.full);
  const opened = await waitModal(page, before);
  rec(kind, '크게 보기 열림', opened ? 'OK' : 'FAIL');
  if (opened) {
    await page.waitForTimeout(2500);
    await page.screenshot({ path: path.join(outDir, `viewer_${kind}.png`) });
    await checkCloseAtWidths(page, kind, '큰 화면 뷰어');
    await page.keyboard.press('Escape');
    const closed = await waitNoModal(page, before);
    rec(kind, '큰 화면 뷰어 Esc 닫기', closed ? 'OK' : 'FAIL');
    if (!closed) { const b = await visibleCloseButtons(page); if (b.length) await page.mouse.click(b[b.length - 1].x + 5, b[b.length - 1].y + 5); await waitNoModal(page, before); }
  }

  // 직접 편집: 열기만 하고 닫기 → 확인 창이 없어야 함
  if (await page.locator(`${scope} ${c.edit}`).count()) {
    await clickIn(page, scope, c.edit);
    const eo = await waitModal(page, before, 30000);
    rec(kind, '직접 편집 열림', eo ? 'OK' : 'FAIL');
    if (eo) {
      await page.waitForTimeout(kind === 'excel' ? 6000 : 8000);
      await page.screenshot({ path: path.join(outDir, `editor_${kind}.png`) });
      await checkCloseAtWidths(page, kind, '편집기');
      const b = await visibleCloseButtons(page);
      const top = b[b.length - 1];
      await page.mouse.click(top.x + top.w / 2, top.y + top.h / 2);
      await page.waitForTimeout(800);
      const confirm = await hasExitConfirm(page);
      const closed = await waitNoModal(page, before, 3000);
      rec(kind, '열기만 하고 닫기 → 확인 창 없음', !confirm && closed ? 'OK' : 'FAIL', `확인창=${confirm}, 닫힘=${closed}`);
      if (!closed) {
        const disc = page.getByText('저장하지 않고 닫기', { exact: false }).last();
        if (await disc.count()) await disc.click().catch(() => {});
        await waitNoModal(page, before, 5000);
      }
    }
  } else {
    rec(kind, '직접 편집 열림', 'N/A', '편집 버튼 없음(비활성)');
  }

  // 가져오기: 본문 길이 증가
  const lenBefore = await noteContentLength(page);
  await clickIn(page, scope, c.imp);
  await page.waitForTimeout(4000);
  const lenAfter = await noteContentLength(page);
  rec(kind, '메모로 가져오기', lenAfter > lenBefore ? 'OK' : 'FAIL', `본문 ${lenBefore}→${lenAfter}자`);

  // 다운로드(저장) 링크
  const href = await page.locator(`${scope} a[download]`).first().getAttribute('href').catch(() => null);
  rec(kind, '저장(다운로드) 링크', href && (href.startsWith('data:') || href.startsWith('blob:')) ? 'OK' : 'FAIL', href ? href.slice(0, 30) : '없음');
}

async function cardMetricsById(page, f) {
  return page.evaluate(async ({ id, kind }) => {
    const v = document.querySelector(`.markdown-preview [data-att-id="${id}"]${kind === 'excel' ? '.excel-inline-viewer' : kind === 'docx' ? '.docx-inline-viewer' : kind === 'pptx' ? '.pptx-inline-viewer' : kind === 'pdf' ? '.pdf-inline-viewer' : '.hwp-inline-viewer'}`);
    if (!v) return { exists: false };
    const out = { exists: true };
    if (kind === 'excel') {
      const tds = v.querySelectorAll('td');
      out.cells = tds.length; out.filled = Array.from(tds).filter((t) => t.textContent.trim()).length;
      out.bordered = Array.from(tds).filter((t) => /border/.test(t.getAttribute('style') || '')).length;
    } else if (kind === 'docx') {
      const c = v.querySelector('.docx-container');
      out.visible = !!c && !c.classList.contains('hidden'); out.textLen = (c?.innerText || '').trim().length;
      out.tables = c?.querySelectorAll('table').length || 0; out.imgs = c?.querySelectorAll('img').length || 0;
    } else if (kind === 'hwp' || kind === 'hwpx') {
      const c = v.querySelector('.hwp-svg-container');
      out.visible = !!c && !c.classList.contains('hidden'); out.svg = c?.querySelectorAll('svg').length || 0;
      out.textNodes = c?.querySelectorAll('text').length || 0;
      out.page = v.querySelector('[id^="hwp-page-indicator"]')?.textContent?.trim() || '';
    } else if (kind === 'pdf') {
      const c = v.querySelector('canvas.pdf-canvas');
      out.visible = !!c && !c.classList.contains('hidden'); out.cw = c?.width || 0; out.ch = c?.height || 0;
      try { const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let k = 0; for (let i = 0; i < d.length; i += 16) if (d[i] < 128) k++; out.darkPixels = k; } catch (e) { out.pixErr = String(e); }
      out.page = v.querySelector('[id^="pdf-page-indicator"]')?.textContent?.trim() || '';
    } else if (kind === 'pptx') {
      const c = v.querySelector('.pptx-container');
      const sl = c?.querySelector('[data-testid="pptx-card-slide"]');
      out.slide = !!sl;
      if (sl) { const r = sl.parentElement.getBoundingClientRect(); out.w = Math.round(r.width); out.h = Math.round(r.height); out.elements = sl.children.length; out.textLen = (sl.innerText || '').trim().length; out.count = c.getAttribute('data-slide-count'); }
    }
    return out;
  }, { id: f.id, kind: f.kind });
}

async function noteContentLength(page) {
  // 디바운스 저장을 기다린 뒤 IndexedDB의 본문 길이를 읽는다
  await page.waitForTimeout(1200);
  return page.evaluate(() => new Promise((res) => {
    const r = indexedDB.open('MyMemo_DB');
    r.onsuccess = () => { const g = r.result.transaction('notes').objectStore('notes').get('note-step6'); g.onsuccess = () => res((g.result?.content || '').length); };
  }));
}

async function auditPopout(page, f) {
  const { kind } = f;
  const c = CARD[kind];
  const win = '.step6-floating';
  const scope = `${win} [data-att-id="${f.id}"]${c.viewer}`;
  log(`\n[${kind}] 창 분리(플로팅) 점검`);
  const exists = await page.locator(scope).count();
  if (!exists) { rec(kind, '팝아웃 카드', 'FAIL', '카드 없음'); return; }
  await page.locator(scope).first().scrollIntoViewIfNeeded().catch(() => {});
  await page.waitForTimeout(1500);
  const m = await page.evaluate(({ id, kind }) => {
    const v = document.querySelector(`.step6-floating [data-att-id="${id}"]`);
    if (!v) return {};
    if (kind === 'excel') return { filled: Array.from(v.querySelectorAll('td')).filter((t) => t.textContent.trim()).length };
    if (kind === 'docx') { const x = v.querySelector('.docx-container'); return { visible: x && !x.classList.contains('hidden'), len: (x?.innerText || '').length }; }
    if (kind === 'pdf') { const x = v.querySelector('canvas'); return { visible: x && !x.classList.contains('hidden'), w: x?.width }; }
    if (kind === 'pptx') return { slide: !!v.querySelector('[data-testid="pptx-card-slide"]') };
    const x = v.querySelector('.hwp-svg-container'); return { visible: x && !x.classList.contains('hidden'), svg: x?.querySelectorAll('svg').length };
  }, { id: f.id, kind });
  rec(kind, '팝아웃 카드 렌더', JSON.stringify(m).includes('true') || (m.filled > 5) || m.svg > 0 ? 'OK' : 'FAIL', JSON.stringify(m));

  for (const [label, sel] of [['크게 보기', c.full], ['직접 편집', c.edit], ['가져오기', c.imp]]) {
    if (!(await page.locator(`${scope} ${sel}`).count())) { rec(kind, `팝아웃 ${label}`, 'N/A', '버튼 없음'); continue; }
    const before = (await visibleCloseButtons(page)).length;
    const lenBefore = label === '가져오기' ? await noteContentLength(page) : 0;
    let alertMsg = '';
    const onDialog = (d) => { alertMsg = d.message(); d.accept().catch(() => {}); };
    page.on('dialog', onDialog);
    await clickIn(page, scope, sel);
    if (label === '가져오기') {
      await page.waitForTimeout(4000);
      const lenAfter = await noteContentLength(page);
      rec(kind, `팝아웃 ${label}`, lenAfter > lenBefore ? 'OK' : 'FAIL', `본문 ${lenBefore}→${lenAfter}자 ${alertMsg.slice(0, 40)}`);
    } else {
      const opened = await waitModal(page, before, label === '직접 편집' ? 20000 : 10000);
      rec(kind, `팝아웃 ${label}`, opened ? 'OK' : 'FAIL', alertMsg.slice(0, 60));
      if (opened) {
        await page.waitForTimeout(label === '직접 편집' ? 6000 : 1500);
        await page.keyboard.press('Escape');
        if (!(await waitNoModal(page, before, 3000))) {
          const b = await visibleCloseButtons(page); const t = b[b.length - 1];
          if (t) await page.mouse.click(t.x + t.w / 2, t.y + t.h / 2);
          await page.waitForTimeout(500);
          const disc = page.getByText('저장하지 않고 닫기', { exact: false }).last();
          if (await disc.count()) await disc.click().catch(() => {});
          await waitNoModal(page, before, 4000);
        }
      }
    }
    page.off('dialog', onDialog);
  }
}

async function run() {
  fs.mkdirSync(outDir, { recursive: true });
  let vite = null; let browser = null; let code = 1;
  try {
    log('[1] Vite 시작');
    vite = await startVite();
    log('[2] Chrome 시작');
    browser = await chromium.launch({ executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe', headless: true });
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(String(e).slice(0, 200)));
    page.on('dialog', (d) => { log(`  [dialog] ${d.type()}: ${d.message().slice(0, 80)}`); d.accept().catch(() => {}); });

    const files = FILES.filter((f) => !ONLY || ONLY.includes(f.kind));
    const atts = {};
    let content = '# STEP6 점검 메모\n\n본문 첫 줄입니다.\n\n';
    for (const f of files) {
      const name = path.basename(f.file);
      atts[f.id] = { id: f.id, name, type: f.mime, data: `data:${f.mime};base64,${fs.readFileSync(f.file).toString('base64')}` };
      content += `[${name}](attachment:${f.id})\n\n`;
    }
    log('[3] IndexedDB 시드');
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
          tx.objectStore('notes').put({ id: 'note-step6', projectId: 'proj-step6', title: 'STEP6 점검 메모', content, attachments: atts, isLocalOnly: true, syncStatus: 'synced', createdAt: '2026-11-16T00:00:00.000Z', updatedAt: '2026-11-16T00:00:00.000Z' });
          tx.oncomplete = resolve; tx.onerror = reject;
        };
        req.onerror = reject;
      });
      localStorage.setItem('mymemo_active_note_id', 'note-step6');
      localStorage.setItem('mymemo_active_project_id', 'ALL');
    }, { atts, content });
    await page.reload({ waitUntil: 'domcontentloaded' });
    await page.waitForSelector('.markdown-preview', { timeout: 20000 });
    await page.waitForTimeout(3000);

    log('[4] 메인 창 형식별 점검');
    for (const f of files) {
      try { await auditMain(page, f); } catch (e) { rec(f.kind, '메인 점검 예외', 'FAIL', String(e).slice(0, 200)); }
    }

    log('[5] 창 분리(플로팅) 점검');
    await page.reload({ waitUntil: 'domcontentloaded' });
    await page.waitForSelector('.markdown-preview', { timeout: 20000 });
    await page.locator('button[title^="이 메모를 화면 내 플로팅"]').first().click();
    await page.waitForTimeout(1500);
    // 플로팅 창 루트 표시 (미리보기 컨테이너 중 메인이 아닌 것)
    await page.evaluate(() => {
      const previews = Array.from(document.querySelectorAll('.markdown-preview'));
      const fl = previews[previews.length - 1];
      fl?.classList.add('step6-floating');
    });
    for (const f of files) {
      try { await auditPopout(page, f); } catch (e) { rec(f.kind, '팝아웃 점검 예외', 'FAIL', String(e).slice(0, 200)); }
    }
    results._pageErrors = errors.slice(0, 30);
    code = 0;
  } catch (e) {
    console.error('[오류]', e);
  } finally {
    fs.writeFileSync(path.join(outDir, 'audit_result.json'), JSON.stringify(results, null, 2));
    log(`결과 저장: ${path.join(outDir, 'audit_result.json')}`);
    if (browser) await browser.close().catch(() => {});
    killTree(vite);
    process.exit(code);
  }
}

run();
