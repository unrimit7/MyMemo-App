import { chromium } from 'playwright-core';
import { spawn } from 'child_process';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');
const outDir = path.join(projectRoot, 'docs', 'office-upgrade', 'out', 'excel_view');

// 5분 타임아웃 가드 (절대 무한 대기하지 않음)
const globalTimer = setTimeout(() => {
  console.error('\n[TIMEOUT] 5분 시간 제한 초과로 프로세스를 강제 종료합니다.');
  process.exit(1);
}, 300000);
globalTimer.unref();

function assert(condition, message) {
  if (!condition) {
    throw new Error(`[AssertionFailed] ${message}`);
  }
}

function fileToDataUrl(filePath, mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
  const buf = fs.readFileSync(filePath);
  return `data:${mime};base64,${buf.toString('base64')}`;
}

async function startViteServer() {
  const port = 5195;
  const proc = spawn('npx', ['vite', '--port', String(port), '--strictPort'], {
    cwd: projectRoot,
    shell: true,
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  const url = `http://localhost:${port}/viewer_test.html`;

  await new Promise((resolve, reject) => {
    let started = false;
    const timeout = setTimeout(() => {
      if (!started) reject(new Error('Vite server start timeout'));
    }, 20000);

    const onData = (data) => {
      const text = data.toString();
      if (text.includes('Local:') || text.includes('ready in') || text.includes(String(port))) {
        started = true;
        clearTimeout(timeout);
        resolve();
      }
    };

    proc.stdout.on('data', onData);
    proc.stderr.on('data', onData);
    proc.on('error', (err) => {
      clearTimeout(timeout);
      reject(err);
    });
  });

  return { proc, url };
}

async function main() {
  console.log('='.repeat(80));
  console.log('  [엑셀 인라인 카드 / 큰 화면 뷰어 / 팝아웃 시트 탭 전환, 표시형식, 서식 실측 검증]');
  console.log('='.repeat(80));

  fs.mkdirSync(outDir, { recursive: true });

  console.log('1. Vite 테스트 서버 가동...');
  let server = null;
  let browser = null;
  let exitCode = 1;

  try {
    server = await startViteServer();
    console.log(`  서버 준비 완료: ${server.url}`);

    browser = await chromium.launch({
      executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
      headless: true,
    });

    const context = await browser.newContext({
      viewport: { width: 1280, height: 900 },
    });
    const page = await context.newPage();

    page.on('console', (msg) => {
      if (msg.type() === 'error') {
        console.log('  [브라우저 에러]:', msg.text());
      }
    });

    await page.goto(server.url, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(600);

    const sampleMultiFile = path.join(projectRoot, 'docs', 'office-upgrade', 'samples', 'real', 'sample_다중시트_예시.xlsx');
    const sampleMultiDataUrl = fileToDataUrl(sampleMultiFile);

    const fileAMSFile = path.join(projectRoot, 'docs', 'office-upgrade', 'samples', 'real', '실파일A_인사평가_MS엑셀.xlsx');
    const fileAMSDataUrl = fileToDataUrl(fileAMSFile);

    const fileBMSFile = path.join(projectRoot, 'docs', 'office-upgrade', 'samples', 'real', '실파일B_견적서_MS엑셀.xlsx');
    const fileBMSDataUrl = fileToDataUrl(fileBMSFile);

    const fileAHanselFile = path.join(projectRoot, 'docs', 'office-upgrade', 'samples', 'real', '실파일A_인사평가_한셀.xlsx');
    const fileAHanselDataUrl = fileToDataUrl(fileAHanselFile);

    // -------------------------------------------------------------------------
    // [검증 1] 메모 본문 인라인 카드: 탭 전환, numFmt(A4/E4), 수식 계산(G4), 기본 글자색 rgb(0,0,0)
    // -------------------------------------------------------------------------
    console.log('\n' + '-'.repeat(80));
    console.log('[검증 1: 메모 본문 인라인 카드 시트 탭 전환, 표시 형식, 수식 동적 계산, 검정 기본 글자색]');
    console.log('-'.repeat(80));

    await page.evaluate(({ dataUrl, title }) => {
      window.__loadInlineCard(dataUrl, title, 'test_att_homepage');
    }, { dataUrl: sampleMultiDataUrl, title: 'sample_다중시트_예시.xlsx' });

    await page.waitForSelector('.excel-inline-viewer');
    await page.waitForTimeout(600);

    // 1-1. 초기 시트 (인증서 현황) 확인
    const initialA1 = await page.evaluate(() => {
      const td = document.querySelector('.excel-inline-viewer table tbody tr:first-child td.excel-td');
      return td ? td.textContent.trim() : '';
    });
    console.log(`  [단언 1-1] 인라인 카드 초기 A1 값: "${initialA1}" (기대값: "2026년 11월 정기 결제 예정" 또는 "SSL 인증서 관리 대장" 포함)`);
    assert(initialA1.includes('결제 예정') || initialA1.includes('SSL 인증서') || initialA1.includes('인증서'), '첫 번째 시트 데이터 A1 렌더링 확인');

    // 1-2. 셀 기본 글자색 검정 rgb(0, 0, 0) 및 배경 흰색 검증 (데이터 행 Row 4)
    const cellStyles = await page.evaluate(() => {
      const trs = document.querySelectorAll('.excel-inline-viewer table tbody tr');
      const dataTd = trs[3] ? trs[3].querySelector('td.excel-td') : null;
      if (!dataTd) return null;
      const cs = window.getComputedStyle(dataTd);
      return {
        color: cs.color,
        backgroundColor: cs.backgroundColor,
      };
    });
    console.log(`  [단언 1-2] 인라인 카드 데이터 셀 계산 스타일: 실제 color=${cellStyles?.color}, bg=${cellStyles?.backgroundColor} (기대값: color=rgb(0, 0, 0), bg=rgb(255, 255, 255))`);
    assert(cellStyles && cellStyles.color === 'rgb(0, 0, 0)', `셀 글자색이 정확한 검정 rgb(0, 0, 0)이어야 함 (실제: ${cellStyles?.color})`);
    assert(cellStyles && (cellStyles.backgroundColor === 'rgb(255, 255, 255)' || cellStyles.backgroundColor === 'rgba(0, 0, 0, 0)'), '셀 배경색이 흰색(#ffffff) 기본으로 보존 확인');

    // 스크린샷 1
    const shot1 = path.join(outDir, '01_인라인카드_초기시트_인증서현황.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shot1 });
    console.log(`  📸 스크린샷 저장: ${shot1}`);

    // 1-3. 시트 탭 클릭: [서비스 현황]
    console.log('  -> 시트 탭 [서비스 현황] 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="서비스 현황"]');
    await page.waitForTimeout(400);

    const switchedA1 = await page.evaluate(() => {
      const td = document.querySelector('.excel-inline-viewer table tbody tr:first-child td.excel-td');
      return td ? td.textContent.trim() : '';
    });
    console.log(`  [단언 1-3] 시트 탭 클릭 후 A1 값: "${switchedA1}" (기대값: "클라우드 인프라 서비스 현황" 포함)`);
    assert(switchedA1.includes('클라우드') || switchedA1.includes('인프라'), '시트 탭 클릭 시 표시 시트 데이터(A1)가 정상 변경됨');

    // 스크린샷 2
    const shot2 = path.join(outDir, '02_인라인카드_탭전환_서비스현황.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shot2 });
    console.log(`  📸 스크린샷 저장: ${shot2}`);

    // 1-4. 시트 탭 클릭: [결제 예정] -> A4 날짜, E4 금액, G4 수식 계산(남은 일수) 단언
    console.log('  -> 시트 탭 [결제 예정] 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="결제 예정"]');
    await page.waitForTimeout(500);

    const paymentValues = await page.evaluate(() => {
      const trs = document.querySelectorAll('.excel-inline-viewer table tbody tr');
      // 4행: Row 4
      const r4 = trs[3];
      if (!r4) return null;
      const tds = r4.querySelectorAll('td.excel-td');
      const a4Td = tds[0];
      const e4Td = tds[4];
      const g4Td = tds[6];
      return {
        a4: a4Td ? a4Td.textContent.trim() : '',
        a4Align: a4Td ? window.getComputedStyle(a4Td).textAlign : '',
        e4: e4Td ? e4Td.textContent.trim() : '',
        e4Align: e4Td ? window.getComputedStyle(e4Td).textAlign : '',
        g4: g4Td ? g4Td.textContent.trim() : '',
        g4Align: g4Td ? window.getComputedStyle(g4Td).textAlign : '',
      };
    });

    console.log(`  [단언 1-4] 결제 예정 A4 날짜: 실제="${paymentValues?.a4}", 기대값="2026-11-27"`);
    assert(paymentValues?.a4 === '2026-11-27', `A4 날짜는 '2026-11-27'이어야 함 (실제: ${paymentValues?.a4})`);
    console.log(`  [단언 1-4-1] 결제 예정 A4 날짜 가로 정렬: 실제="${paymentValues?.a4Align}", 기대값="center"`);
    assert(paymentValues?.a4Align === 'center', `A4 날짜 셀 가로 정렬은 center이어야 함 (실제: ${paymentValues?.a4Align})`);

    console.log(`  [단언 1-5] 결제 예정 E4 금액: 실제="${paymentValues?.e4}", 기대값="23,100원"`);
    assert(paymentValues?.e4 === '23,100원', `E4 금액은 '23,100원'이어야 함 (실제: ${paymentValues?.e4})`);
    console.log(`  [단언 1-5-1] 결제 예정 E4 금액 가로 정렬: 실제="${paymentValues?.e4Align}", 기대값="right"`);
    assert(paymentValues?.e4Align === 'right', `E4 금액 셀 가로 정렬은 right이어야 함 (실제: ${paymentValues?.e4Align})`);

    // G4 = A4 - TODAY()
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const target = new Date(2026, 10, 27);
    const expectedDiff = Math.round((target.getTime() - today.getTime()) / (24 * 3600 * 1000));
    const expectedG4 = expectedDiff >= 0 ? `${expectedDiff}일` : `지남 ${Math.abs(expectedDiff)}일`;
    console.log(`  [단언 1-6] 결제 예정 G4 남은 일수(=A4-TODAY()): 실제="${paymentValues?.g4}", 기대값="${expectedG4}"`);
    assert(paymentValues?.g4 === expectedG4, `G4 수식 계산값은 '${expectedG4}'이어야 함 (실제: ${paymentValues?.g4})`);
    console.log(`  [단언 1-6-1] 결제 예정 G4 가로 정렬: 실제="${paymentValues?.g4Align}", 기대값="right"`);
    assert(paymentValues?.g4Align === 'right', `G4 셀 가로 정렬은 right이어야 함 (실제: ${paymentValues?.g4Align})`);

    // -------------------------------------------------------------------------
    // [검증 2] 큰 화면 뷰어 (ExcelViewerModal): 탭 전환, 결제 예정 표시형식, 검정 기본 글자색
    // -------------------------------------------------------------------------
    console.log('\n' + '-'.repeat(80));
    console.log('[검증 2: 큰 화면 뷰어 (ExcelViewerModal) 시트 탭 전환, 표시형식 및 원본 서식 보존]');
    console.log('-'.repeat(80));

    await page.evaluate(({ dataUrl, title }) => {
      window.__loadExcelViewer(dataUrl, title, 'test_modal_homepage');
    }, { dataUrl: sampleMultiDataUrl, title: 'sample_다중시트_예시.xlsx' });

    await page.waitForSelector('table.min-w-full');
    await page.waitForTimeout(600);

    const modalStyles = await page.evaluate(() => {
      const trs = document.querySelectorAll('table.min-w-full tbody tr');
      const dataTd = trs[3] ? trs[3].querySelectorAll('td')[1] : null;
      if (!dataTd) return null;
      const cs = window.getComputedStyle(dataTd);
      return {
        color: cs.color,
        backgroundColor: cs.backgroundColor,
      };
    });
    console.log(`  [단언 2-1] 큰 화면 뷰어 셀 스타일: 실제 color=${modalStyles?.color}, bg=${modalStyles?.backgroundColor} (기대값: color=rgb(0, 0, 0))`);
    assert(modalStyles && modalStyles.color === 'rgb(0, 0, 0)', `큰 화면 뷰어 셀 글자색이 검정 rgb(0, 0, 0)이어야 함 (실제: ${modalStyles?.color})`);

    // 스크린샷 3
    const shot3 = path.join(outDir, '03_큰화면뷰어_초기시트_인증서현황.png');
    await page.screenshot({ path: shot3 });
    console.log(`  📸 스크린샷 저장: ${shot3}`);

    // 시트 탭 [결제 예정] 클릭
    console.log('  -> 큰 화면 뷰어 시트 탭 [결제 예정] 클릭...');
    await page.click('button:has-text("결제 예정")');
    await page.waitForTimeout(500);

    const modalPaymentVals = await page.evaluate(() => {
      const trs = document.querySelectorAll('table.min-w-full tbody tr');
      const r4 = trs[3];
      if (!r4) return null;
      const tds = r4.querySelectorAll('td');
      // index 0은 행번호(#), 1은 A, 5는 E, 7은 G
      const a4Td = tds[1];
      const e4Td = tds[5];
      const g4Td = tds[7];
      return {
        a4: a4Td ? a4Td.textContent.trim() : '',
        a4Align: a4Td ? window.getComputedStyle(a4Td).textAlign : '',
        e4: e4Td ? e4Td.textContent.trim() : '',
        e4Align: e4Td ? window.getComputedStyle(e4Td).textAlign : '',
        g4: g4Td ? g4Td.textContent.trim() : '',
        g4Align: g4Td ? window.getComputedStyle(g4Td).textAlign : '',
      };
    });
    console.log(`  [단언 2-2] 큰 화면 뷰어 결제 예정 A4: 실제="${modalPaymentVals?.a4}", 기대값="2026-11-27"`);
    assert(modalPaymentVals?.a4 === '2026-11-27', `A4 날짜 '2026-11-27' 확인`);
    console.log(`  [단언 2-2-1] 큰 화면 뷰어 결제 예정 A4 가로 정렬: 실제="${modalPaymentVals?.a4Align}", 기대값="center"`);
    assert(modalPaymentVals?.a4Align === 'center', `A4 날짜 가로 정렬 center 확인 (실제: ${modalPaymentVals?.a4Align})`);
    console.log(`  [단언 2-3] 큰 화면 뷰어 결제 예정 E4: 실제="${modalPaymentVals?.e4}", 기대값="23,100원"`);
    assert(modalPaymentVals?.e4 === '23,100원', `E4 금액 '23,100원' 확인`);
    console.log(`  [단언 2-3-1] 큰 화면 뷰어 결제 예정 E4 가로 정렬: 실제="${modalPaymentVals?.e4Align}", 기대값="right"`);
    assert(modalPaymentVals?.e4Align === 'right', `E4 금액 가로 정렬 right 확인 (실제: ${modalPaymentVals?.e4Align})`);
    console.log(`  [단언 2-4] 큰 화면 뷰어 결제 예정 G4: 실제="${modalPaymentVals?.g4}", 기대값="${expectedG4}"`);
    assert(modalPaymentVals?.g4 === expectedG4, `G4 수식 계산값 '${expectedG4}' 확인`);
    console.log(`  [단언 2-4-1] 큰 화면 뷰어 결제 예정 G4 가로 정렬: 실제="${modalPaymentVals?.g4Align}", 기대값="right"`);
    assert(modalPaymentVals?.g4Align === 'right', `G4 가로 정렬 right 확인 (실제: ${modalPaymentVals?.g4Align})`);

    // 스크린샷 4
    const shot4 = path.join(outDir, '04_큰화면뷰어_탭전환_결제예정.png');
    await page.screenshot({ path: shot4 });
    console.log(`  📸 스크린샷 저장: ${shot4}`);

    // -------------------------------------------------------------------------
    // [검증 3] 창 분리 (팝아웃, FloatingNoteWindow) 시트 탭 전환 검증
    // -------------------------------------------------------------------------
    console.log('\n' + '-'.repeat(80));
    console.log('[검증 3: 창 분리 (팝아웃, FloatingNoteWindow) 시트 탭 전환 검증]');
    console.log('-'.repeat(80));

    await page.evaluate(({ dataUrl, title }) => {
      window.__loadPopout(dataUrl, title, 'test_popout_homepage');
    }, { dataUrl: sampleMultiDataUrl, title: 'sample_다중시트_예시.xlsx' });

    await page.waitForSelector('.excel-inline-viewer');
    await page.waitForTimeout(600);

    const shot5 = path.join(outDir, '05_팝아웃창_초기시트_인증서현황.png');
    await page.screenshot({ path: shot5 });
    console.log(`  📸 스크린샷 저장: ${shot5}`);

    console.log('  -> 팝아웃 창 시트 탭 [연간 비용] 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="연간 비용"]');
    await page.waitForTimeout(400);

    const popoutA1_switched = await page.evaluate(() => {
      const td = document.querySelector('.excel-inline-viewer table tbody tr:first-child td.excel-td');
      return td ? td.textContent.trim() : '';
    });
    console.log(`  [단언 3-1] 팝아웃 창 탭 클릭 후 A1 값: "${popoutA1_switched}" (기대값: "연간 IT 유지비용 종합 요약" 포함)`);
    assert(popoutA1_switched.includes('연간') || popoutA1_switched.includes('유지비용'), '팝아웃 창에서 시트 탭 클릭 시 해당 시트(연간 비용)로 완벽히 전환됨');

    const shot6 = path.join(outDir, '06_팝아웃창_탭전환_연간비용.png');
    await page.screenshot({ path: shot6 });
    console.log(`  📸 스크린샷 저장: ${shot6}`);

    // -------------------------------------------------------------------------
    // [검증 4] 실파일A·B(MS/한셀): 견적서 C5, F11, B8(RichText), 인사평가 G4, 월별매출 D열(백분율)
    // -------------------------------------------------------------------------
    console.log('\n' + '-'.repeat(80));
    console.log('[검증 4: 실파일A·B 서식 검증 (견적서 C5/F11/B8 RichText, 인사평가 G4, 월별매출 백분율)]');
    console.log('-'.repeat(80));

    // 4-1. 실파일B 견적서 검증
    await page.evaluate(({ dataUrl, title }) => {
      window.__loadInlineCard(dataUrl, title, 'test_fileB_ms');
    }, { dataUrl: fileBMSDataUrl, title: '실파일B_견적서_MS엑셀.xlsx' });
    await page.waitForTimeout(600);

    const quoteVals = await page.evaluate(() => {
      const trs = document.querySelectorAll('.excel-inline-viewer table tbody tr');
      // C5: 5행 3열 (rIdx 4, cIdx 2)
      const r5 = trs[4];
      const c5 = r5 ? r5.querySelectorAll('td.excel-td')[2]?.textContent?.trim() : '';
      // F11: 11행 6열 (rIdx 10, cIdx 5)
      const r11 = trs[10];
      const f11Td = r11 ? r11.querySelectorAll('td.excel-td')[5] : null;
      const f11 = f11Td ? f11Td.textContent?.trim() : '';
      const f11Align = f11Td ? window.getComputedStyle(f11Td).textAlign : '';
      // B8 RichText: 8행 2열
      const r8 = trs[7];
      const b8Td = r8 ? r8.querySelectorAll('td.excel-td')[1] : null;
      const b8Html = b8Td ? b8Td.innerHTML : '';
      const b8Text = b8Td ? b8Td.textContent?.trim() : '';
      return { c5, f11, f11Align, b8Html, b8Text };
    });

    console.log(`  [단언 4-1] 견적서 C5 날짜 서식: 실제="${quoteVals?.c5}", 기대값="2026년 11월 11일"`);
    assert(quoteVals?.c5 === '2026년 11월 11일', `C5는 '2026년 11월 11일'이어야 함 (실제: ${quoteVals?.c5})`);

    console.log(`  [단언 4-2] 견적서 F11 금액 서식: 실제="${quoteVals?.f11}", 기대값="1,250,000"`);
    assert(quoteVals?.f11 === '1,250,000', `F11은 '1,250,000'이어야 함 (실제: ${quoteVals?.f11})`);
    console.log(`  [단언 4-2-1] 견적서 F11 금액 가로 정렬: 실제="${quoteVals?.f11Align}", 기대값="right"`);
    assert(quoteVals?.f11Align === 'right', `견적서 F11 셀 가로 정렬은 right이어야 함 (실제: ${quoteVals?.f11Align})`);

    console.log(`  [단언 4-3] 견적서 B8 서식 있는 텍스트(RichText): 실제 HTML="${quoteVals?.b8Html}"`);
    assert(quoteVals?.b8Html.includes('color:') && quoteVals?.b8Html.includes('부가세 포함'), 'B8 셀에 부분 빨간 굵은 글씨 RichText 런이 포함되어야 함');

    const shot9 = path.join(outDir, '09_실파일B_MS엑셀_견적서_인라인카드.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shot9 });
    console.log(`  📸 실파일B MS엑셀 견적서 스크린샷: ${shot9}`);

    // 견적서 [월별매출] 시트 탭 클릭 -> D열 백분율 단언
    console.log('  -> 견적서 시트 탭 [월별매출] 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="월별매출"]');
    await page.waitForTimeout(500);

    const salesD2 = await page.evaluate(() => {
      const trs = document.querySelectorAll('.excel-inline-viewer table tbody tr');
      const r2 = trs[1];
      return r2 ? r2.querySelectorAll('td.excel-td')[3]?.textContent?.trim() : '';
    });
    console.log(`  [단언 4-4] 월별매출 D2 달성률 백분율 서식: 실제="${salesD2}", 기대값="80.0%"`);
    assert(salesD2 === '80.0%', `D2 달성률은 '80.0%'이어야 함 (실제: ${salesD2})`);

    // 4-2. 실파일A 인사평가 검증
    await page.evaluate(({ dataUrl, title }) => {
      window.__loadInlineCard(dataUrl, title, 'test_fileA_ms');
    }, { dataUrl: fileAMSDataUrl, title: '실파일A_인사평가_MS엑셀.xlsx' });
    await page.waitForTimeout(600);

    const shot7 = path.join(outDir, '07_실파일A_MS엑셀_인라인카드.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shot7 });
    console.log(`  📸 실파일A MS엑셀 스크린샷: ${shot7}`);

    const evalG4 = await page.evaluate(() => {
      const trs = document.querySelectorAll('.excel-inline-viewer table tbody tr');
      // 4행 7열(G4)
      const r4 = trs[3];
      const g4Td = r4 ? r4.querySelectorAll('td.excel-td')[6] : null;
      return {
        text: g4Td ? g4Td.textContent?.trim() : '',
        align: g4Td ? window.getComputedStyle(g4Td).textAlign : '',
      };
    });
    console.log(`  [단언 4-5] 인사평가 G4 인센티브 서식: 실제="${evalG4?.text}", 기대값="3,000,000원"`);
    assert(evalG4?.text === '3,000,000원', `G4 인센티브는 '3,000,000원'이어야 함 (실제: ${evalG4?.text})`);
    console.log(`  [단언 4-5-1] 인사평가 G4 가로 정렬: 실제="${evalG4?.align}", 기대값="right"`);
    assert(evalG4?.align === 'right', `인사평가 G4 셀 가로 정렬은 right이어야 함 (실제: ${evalG4?.align})`);

    // 부서코드 탭 클릭
    console.log('  -> 실파일A 시트 탭 [부서코드] 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="부서코드"]');
    await page.waitForTimeout(400);

    const shot8 = path.join(outDir, '08_실파일A_MS엑셀_탭전환_부서코드.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shot8 });
    console.log(`  📸 실파일A MS엑셀 부서코드 탭 스크린샷: ${shot8}`);

    // 실파일A 한셀 렌더링
    await page.evaluate(({ dataUrl, title }) => {
      window.__loadInlineCard(dataUrl, title, 'test_fileA_hansel');
    }, { dataUrl: fileAHanselDataUrl, title: '실파일A_인사평가_한셀.xlsx' });
    await page.waitForTimeout(500);
    const fileA_hansel_shot = path.join(outDir, '10_실파일A_한셀_인라인카드.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: fileA_hansel_shot });
    console.log(`  📸 실파일A 한셀 스크린샷: ${fileA_hansel_shot}`);

    // -------------------------------------------------------------------------
    // [검증 5] 워드 / PPT / 한글 인라인 카드 테마 영향 점검 (not-prose 격리)
    // -------------------------------------------------------------------------
    console.log('\n' + '-'.repeat(80));
    console.log('[검증 5: 워드/PPT/한글 인라인 카드 테마 영향 점검 (격리성 검증)]');
    console.log('-'.repeat(80));

    const isolationCheck = await page.evaluate(({ multiSheetUrl }) => {
      const renderFn = window.__renderAttachmentsInHtml;
      if (!renderFn) return { error: '__renderAttachmentsInHtml is not defined' };
      const dummyAtts = {
        'docx_1': { id: 'docx_1', name: '보고서.docx', type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', data: 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,AAA=' },
        'pptx_1': { id: 'pptx_1', name: '발표.pptx', type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', data: 'data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,AAA=' },
        'hwp_1': { id: 'hwp_1', name: '기획서.hwp', type: 'application/x-hwp', data: 'data:application/x-hwp;base64,AAA=' },
        'xlsx_1': { id: 'xlsx_1', name: '표.xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', data: multiSheetUrl },
      };
      const htmlInput = '<a href="attachment:docx_1">워드</a>\n<a href="attachment:pptx_1">PPT</a>\n<a href="attachment:hwp_1">한글</a>\n<a href="attachment:xlsx_1">엑셀</a>';
      const rendered = renderFn(htmlInput, dummyAtts);
      return {
        hasDocxNotProse: rendered.includes('docx-inline-viewer not-prose'),
        hasPptxNotProse: rendered.includes('pptx-inline-viewer not-prose'),
        hasHwpNotProse: rendered.includes('hwp-inline-viewer not-prose'),
        hasExcelNotProse: rendered.includes('excel-inline-viewer not-prose'),
      };
    }, { multiSheetUrl: sampleMultiDataUrl });

    assert(isolationCheck.hasDocxNotProse, '워드 인라인 뷰어 not-prose 확인');
    assert(isolationCheck.hasPptxNotProse, 'PPT 인라인 뷰어 not-prose 확인');
    assert(isolationCheck.hasHwpNotProse, '한글 인라인 뷰어 not-prose 확인');
    assert(isolationCheck.hasExcelNotProse, '엑셀 인라인 뷰어 not-prose 확인');
    console.log('  ✓ 워드/PPT/한글/엑셀 인라인 카드 모두 not-prose 클래스로 앱 테마로부터 완벽 격리 확인');

    console.log('\n' + '='.repeat(80));
    console.log('  [모든 엑셀 시트 탭 전환, 표시 형식, 수식 계산, 서식 검증 완료: 100% 통과]');
    console.log('='.repeat(80));

    exitCode = 0;
  } catch (err) {
    console.error('\n검증 실패:', err);
    exitCode = 1;
  } finally {
    if (browser) {
      try {
        await browser.close();
      } catch {}
    }
    if (server?.proc) {
      try {
        server.proc.kill();
      } catch {}
    }
    clearTimeout(globalTimer);
    process.exit(exitCode);
  }
}

main();
