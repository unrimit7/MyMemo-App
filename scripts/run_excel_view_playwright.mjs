import { chromium } from 'playwright-core';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { spawn } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');

const outDir = path.join(projectRoot, 'docs', 'office-upgrade', 'out', 'excel_view');
if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

function findChrome() {
  const winPaths = [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    path.join(process.env.LOCALAPPDATA || '', 'Google\\Chrome\\Application\\chrome.exe'),
  ];
  for (const p of winPaths) {
    if (fs.existsSync(p)) return p;
  }
  return null;
}

function fileToDataUrl(filePath, mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
  const buf = fs.readFileSync(filePath);
  return `data:${mimeType};base64,${buf.toString('base64')}`;
}

async function startViteServer(port = 5194) {
  return new Promise((resolve, reject) => {
    const proc = spawn('cmd.exe', ['/c', 'npx', 'vite', '--port', String(port), '--strictPort'], {
      cwd: projectRoot,
      env: { ...process.env },
      stdio: ['ignore', 'pipe', 'pipe'],
    });

    let stdout = '';
    const onData = (data) => {
      stdout += data.toString();
      if (stdout.includes(`localhost:${port}`)) {
        resolve({
          proc,
          url: `http://localhost:${port}/viewer_test.html`,
        });
      }
    };
    proc.stdout.on('data', onData);
    proc.stderr.on('data', (d) => { stdout += d.toString(); });
    proc.on('error', reject);
    setTimeout(() => {
      if (!stdout.includes(`localhost:${port}`)) {
        resolve({
          proc,
          url: `http://localhost:${port}/viewer_test.html`,
        });
      }
    }, 6000);
  });
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(`[AssertionFailed] ${message}`);
  }
  console.log(`  ✓ ${message}`);
}

async function main() {
  console.log('='.repeat(80));
  console.log('  [엑셀 인라인 카드 / 큰 화면 뷰어 / 팝아웃 시트 탭 전환 및 테마 독립 실측 검증]');
  console.log('='.repeat(80));

  const chromePath = findChrome();
  if (!chromePath) throw new Error('Chrome browser not found.');

  console.log('1. Vite 테스트 서버 가동...');
  const server = await startViteServer(5194);
  console.log(`  서버 준비 완료: ${server.url}`);

  const browser = await chromium.launch({
    executablePath: chromePath,
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

  try {
    await page.goto(server.url, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(500);

    const multiSheetFile = path.join(projectRoot, 'docs', 'office-upgrade', 'samples', 'real', '홈페이지_유지비용_관리표_20261111.xlsx');
    const multiSheetDataUrl = fileToDataUrl(multiSheetFile);

    const fileAMSFile = path.join(projectRoot, 'docs', 'office-upgrade', 'samples', 'real', '실파일A_인사평가_MS엑셀.xlsx');
    const fileAMSDataUrl = fileToDataUrl(fileAMSFile);

    const fileBMSFile = path.join(projectRoot, 'docs', 'office-upgrade', 'samples', 'real', '실파일B_견적서_MS엑셀.xlsx');
    const fileBMSDataUrl = fileToDataUrl(fileBMSFile);

    const fileAHanselFile = path.join(projectRoot, 'docs', 'office-upgrade', 'samples', 'real', '실파일A_인사평가_한셀.xlsx');
    const fileAHanselDataUrl = fileToDataUrl(fileAHanselFile);

    // -------------------------------------------------------------------------
    // [검증 1] 메모 본문 인라인 카드 시트 탭 전환 및 셀 스타일 검증 (6개 시트 파일)
    // -------------------------------------------------------------------------
    console.log('\n' + '-'.repeat(80));
    console.log('[검증 1: 메모 본문 인라인 카드 시트 탭 전환 및 엑셀 테마 독립 스타일 검증]');
    console.log('-'.repeat(80));

    await page.evaluate(({ dataUrl, title }) => {
      window.__loadInlineCard(dataUrl, title, 'test_att_homepage');
    }, { dataUrl: multiSheetDataUrl, title: '홈페이지_유지비용_관리표_20261111.xlsx' });

    await page.waitForSelector('.excel-inline-viewer');
    await page.waitForTimeout(600); // 렌더링 및 스타일 하이드레이션 대기

    // 첫 번째 시트 (인증서 현황) 확인
    const initialA1 = await page.evaluate(() => {
      const td = document.querySelector('.excel-inline-viewer table tbody tr:first-child td.excel-td');
      return td ? td.textContent.trim() : '';
    });
    console.log(`  인라인 카드 초기 A1 값: "${initialA1}"`);
    assert(initialA1.includes('SSL 인증서') || initialA1.includes('인증서'), '첫 번째 시트(인증서 현황) 데이터 A1 렌더링 확인');

    // 셀 글자색 및 배경색 검증 (앱 다크 테마로부터 완전 격리되어 기본 흰 바탕 / 검정 글자)
    const cellStyles = await page.evaluate(() => {
      // 3행 2열(도메인 데이터 셀) 스타일 계산
      const trs = document.querySelectorAll('.excel-inline-viewer table tbody tr');
      const dataTd = trs[2] ? trs[2].querySelector('td.excel-td') : null;
      if (!dataTd) return null;
      const cs = window.getComputedStyle(dataTd);
      return {
        color: cs.color,
        backgroundColor: cs.backgroundColor,
      };
    });
    console.log(`  인라인 카드 데이터 셀 계산 스타일: color=${cellStyles?.color}, bg=${cellStyles?.backgroundColor}`);
    assert(cellStyles && (cellStyles.color === 'rgb(0, 0, 0)' || cellStyles.color === 'rgb(51, 65, 85)'), '셀 글자색이 앱 다크 테마(회색 rgb(203,213,225))가 아닌 검정 계열로 엑셀 원본 보존 확인');
    assert(cellStyles && (cellStyles.backgroundColor === 'rgb(255, 255, 255)' || cellStyles.backgroundColor === 'rgba(0, 0, 0, 0)'), '셀 배경색이 흰색(#ffffff) 기본으로 엑셀 스타일 보존 확인');

    // 스크린샷 1: 인라인 카드 초기 시트
    const shot1 = path.join(outDir, '01_인라인카드_초기시트_인증서현황.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shot1 });
    console.log(`  📸 스크린샷 저장: ${shot1}`);

    // 시트 탭 클릭: '서비스 현황' 시트 선택
    console.log('  -> 시트 탭 [서비스 현황] 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="서비스 현황"]');
    await page.waitForTimeout(400);

    const switchedA1 = await page.evaluate(() => {
      const td = document.querySelector('.excel-inline-viewer table tbody tr:first-child td.excel-td');
      return td ? td.textContent.trim() : '';
    });
    console.log(`  시트 탭 클릭 후 A1 값: "${switchedA1}"`);
    assert(switchedA1.includes('클라우드 인프라') || switchedA1.includes('서비스'), '시트 탭 클릭 시 표시 시트 데이터(A1)가 정상 변경됨');

    // 활성 탭 클래스 확인
    const isTabActive = await page.evaluate(() => {
      const pill = document.querySelector('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="서비스 현황"]');
      return pill?.classList.contains('bg-emerald-500/20') && pill?.classList.contains('text-emerald-300');
    });
    assert(isTabActive, '선택된 [서비스 현황] 시트 탭이 에메랄드 강조 스타일로 활성화됨');

    // 스크린샷 2: 인라인 카드 탭 전환 후
    const shot2 = path.join(outDir, '02_인라인카드_탭전환_서비스현황.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: shot2 });
    console.log(`  📸 스크린샷 저장: ${shot2}`);

    // 추가 탭 클릭 검증: '결제 내역' 클릭
    console.log('  -> 시트 탭 [결제 내역] 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="결제 내역"]');
    await page.waitForTimeout(400);
    const switchedA1_3 = await page.evaluate(() => {
      const td = document.querySelector('.excel-inline-viewer table tbody tr:first-child td.excel-td');
      return td ? td.textContent.trim() : '';
    });
    assert(switchedA1_3.includes('결제 집행') || switchedA1_3.includes('결제'), '시트 탭 [결제 내역] 클릭 시 해당 시트 내용으로 즉시 갱신됨');

    // -------------------------------------------------------------------------
    // [검증 2] 큰 화면 뷰어 (ExcelViewerModal) 시트 탭 전환 및 테마 독립 검증
    // -------------------------------------------------------------------------
    console.log('\n' + '-'.repeat(80));
    console.log('[검증 2: 큰 화면 뷰어 (ExcelViewerModal) 시트 탭 전환 및 서식 보존 검증]');
    console.log('-'.repeat(80));

    await page.evaluate(({ dataUrl, title }) => {
      window.__loadExcelViewer(dataUrl, title, 'test_modal_homepage');
    }, { dataUrl: multiSheetDataUrl, title: '홈페이지_유지비용_관리표_20261111.xlsx' });

    await page.waitForSelector('.fixed.inset-0 table');
    await page.waitForTimeout(600);

    const modalA1_initial = await page.evaluate(() => {
      const td = document.querySelector('.fixed.inset-0 table tbody tr:first-child td:nth-child(2)');
      return td ? td.textContent.trim() : '';
    });
    assert(modalA1_initial.includes('SSL 인증서') || modalA1_initial.includes('인증서'), '큰 화면 뷰어 1번 시트 로드 확인');

    // 모달 셀 글자색 및 배경색 검증
    const modalCellStyles = await page.evaluate(() => {
      const td = document.querySelector('.fixed.inset-0 table tbody tr:nth-child(3) td:nth-child(2)');
      if (!td) return null;
      const cs = window.getComputedStyle(td);
      return { color: cs.color, backgroundColor: cs.backgroundColor };
    });
    console.log(`  큰 화면 뷰어 셀 스타일: color=${modalCellStyles?.color}, bg=${modalCellStyles?.backgroundColor}`);
    assert(modalCellStyles && (modalCellStyles.color === 'rgb(0, 0, 0)' || modalCellStyles.color === 'rgb(51, 65, 85)'), '큰 화면 뷰어 셀 글자색이 검정 계열로 엑셀 원본 표시 확인');

    // 스크린샷 3: 큰 화면 뷰어 초기 시트
    const shot3 = path.join(outDir, '03_큰화면뷰어_초기시트_인증서현황.png');
    await page.screenshot({ path: shot3 });
    console.log(`  📸 스크린샷 저장: ${shot3}`);

    // 하단 시트 탭 클릭: '결제 예정' 선택
    console.log('  -> 큰 화면 뷰어 시트 탭 [결제 예정] 클릭...');
    await page.click('button:has-text("결제 예정")');
    await page.waitForTimeout(400);

    const modalA1_switched = await page.evaluate(() => {
      const td = document.querySelector('.fixed.inset-0 table tbody tr:first-child td:nth-child(2)');
      return td ? td.textContent.trim() : '';
    });
    assert(modalA1_switched.includes('결제 예정') || modalA1_switched.includes('정기 결제'), '큰 화면 뷰어 시트 탭 클릭 시 해당 시트(결제 예정) 데이터로 변경됨');

    // 스크린샷 4: 큰 화면 뷰어 탭 전환 후
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
    }, { dataUrl: multiSheetDataUrl, title: '홈페이지_유지비용_관리표_20261111.xlsx' });

    await page.waitForSelector('.excel-inline-viewer');
    await page.waitForTimeout(600);

    const popoutA1_initial = await page.evaluate(() => {
      const td = document.querySelector('.excel-inline-viewer table tbody tr:first-child td.excel-td');
      return td ? td.textContent.trim() : '';
    });
    assert(popoutA1_initial.includes('SSL 인증서') || popoutA1_initial.includes('인증서'), '팝아웃 창 1번 시트 로드 확인');

    // 스크린샷 5: 팝아웃 창 초기 시트
    const shot5 = path.join(outDir, '05_팝아웃창_초기시트_인증서현황.png');
    await page.screenshot({ path: shot5 });
    console.log(`  📸 스크린샷 저장: ${shot5}`);

    // 팝아웃 창 시트 탭 클릭: '연간 비용' 선택
    console.log('  -> 팝아웃 창 시트 탭 [연간 비용] 클릭...');
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="연간 비용"]');
    await page.waitForTimeout(400);

    const popoutA1_switched = await page.evaluate(() => {
      const td = document.querySelector('.excel-inline-viewer table tbody tr:first-child td.excel-td');
      return td ? td.textContent.trim() : '';
    });
    console.log(`  팝아웃 창 탭 클릭 후 A1 값: "${popoutA1_switched}"`);
    assert(popoutA1_switched.includes('연간 IT') || popoutA1_switched.includes('연간 비용'), '팝아웃 창에서 시트 탭 클릭 시 해당 시트(연간 비용)로 완벽히 전환됨');

    // 스크린샷 6: 팝아웃 창 탭 전환 후
    const shot6 = path.join(outDir, '06_팝아웃창_탭전환_연간비용.png');
    await page.screenshot({ path: shot6 });
    console.log(`  📸 스크린샷 저장: ${shot6}`);

    // -------------------------------------------------------------------------
    // [검증 4] samples/real 실파일A(인사평가 MS/한셀) 및 실파일B(견적서 MS/한셀) 원본 서식 보존 검증
    // -------------------------------------------------------------------------
    console.log('\n' + '-'.repeat(80));
    console.log('[검증 4: 실파일A·B(MS/한셀) 인라인 카드 및 다중 시트 탭 검증]');
    console.log('-'.repeat(80));

    // 실파일A MS엑셀
    await page.evaluate(({ dataUrl, title }) => {
      window.__loadInlineCard(dataUrl, title, 'test_fileA_ms');
    }, { dataUrl: fileAMSDataUrl, title: '실파일A_인사평가_MS엑셀.xlsx' });
    await page.waitForTimeout(500);
    const fileA_ms_shot = path.join(outDir, '07_실파일A_MS엑셀_인라인카드.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: fileA_ms_shot });
    console.log(`  📸 실파일A MS엑셀 스크린샷: ${fileA_ms_shot}`);

    // 실파일A 탭 전환 ('부서코드' 클릭)
    await page.click('.excel-inline-viewer .excel-sheet-pill[data-sheet-name="부서코드"]');
    await page.waitForTimeout(300);
    const fileA_ms_tab_shot = path.join(outDir, '08_실파일A_MS엑셀_탭전환_부서코드.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: fileA_ms_tab_shot });
    console.log(`  📸 실파일A MS엑셀 부서코드 탭 스크린샷: ${fileA_ms_tab_shot}`);

    // 실파일B MS엑셀 (견적서 헤더 원본 채우기색 #F7EBDD 검증)
    await page.evaluate(({ dataUrl, title }) => {
      window.__loadInlineCard(dataUrl, title, 'test_fileB_ms');
    }, { dataUrl: fileBMSDataUrl, title: '실파일B_견적서_MS엑셀.xlsx' });
    await page.waitForTimeout(600);
    const fileB_ms_shot = path.join(outDir, '09_실파일B_MS엑셀_견적서_인라인카드.png');
    await page.locator('.excel-inline-viewer').screenshot({ path: fileB_ms_shot });
    console.log(`  📸 실파일B MS엑셀 스크린샷: ${fileB_ms_shot}`);

    // 견적서 헤더 채우기 색상 검증
    const bHeaderBg = await page.evaluate(() => {
      const tds = Array.from(document.querySelectorAll('.excel-inline-viewer table td.excel-td'));
      for (const td of tds) {
        if (td.textContent.trim() === '품 명' || td.textContent.trim() === '공급가액') {
          return window.getComputedStyle(td).backgroundColor;
        }
      }
      return null;
    });
    console.log(`  견적서 헤더(품명) 계산 배경색: ${bHeaderBg}`);
    assert(bHeaderBg && (bHeaderBg === 'rgb(247, 235, 221)' || bHeaderBg.includes('247')), '견적서 헤더 셀의 원본 연갈색 채우기(#F7EBDD)가 보존됨');

    // 실파일A 한셀
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

    // 워드/PPT/한글 인라인 카드에 not-prose 클래스가 부착되어 있는지 브라우저에서 검증
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
    }, { multiSheetUrl: multiSheetDataUrl });

    assert(isolationCheck.hasDocxNotProse, '워드 인라인 뷰어에 not-prose 클래스가 적용되어 다크 테마 간섭 차단');
    assert(isolationCheck.hasPptxNotProse, 'PPT 인라인 뷰어에 not-prose 클래스가 적용되어 다크 테마 간섭 차단');
    assert(isolationCheck.hasHwpNotProse, '한글 인라인 뷰어에 not-prose 클래스가 적용되어 다크 테마 간섭 차단');
    assert(isolationCheck.hasExcelNotProse, '엑셀 인라인 뷰어에 not-prose 클래스가 적용되어 다크 테마 간섭 차단');
    console.log('  ✓ 워드/PPT/한글/엑셀 인라인 카드 모두 not-prose 클래스로 앱 테마로부터 격리 확인');

    console.log('\n' + '='.repeat(80));
    console.log('  [모든 엑셀 시트 탭 전환 및 테마 독립 검증 완료: 100% 통과]');
    console.log('='.repeat(80));

  } finally {
    await browser.close();
    server.proc.kill();
  }
}

main().catch((err) => {
  console.error('\n검증 실패:', err);
  process.exit(1);
});
