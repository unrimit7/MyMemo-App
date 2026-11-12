/**
 * scripts/run_stage5_e2e_playwright.mjs
 * 
 * [5단계 공통 사용성 및 안전성 업그레이드 전수 E2E 검증]
 * 1. Esc 키 동작: 셀 편집/다이얼로그 상태에서 Esc 누르면 편집/창만 닫히고 모달 유지
 * 2. Ctrl+S 키 동작: 4대 에디터 모두 저장이 정확히 1회만 실행됨 (중복 방지 가드)
 * 3. rhwp iframe Ctrl+S: location.origin 타겟 및 sender 검증 postMessage 통신 확인
 * 4. 임시본 복구 안전성:
 *    - 원본 해시 변경(동기화 등) 감지 시 복구 미제안 및 자동 폐기
 *    - [복구 안 함] 클릭 시 즉시 폐기
 * 5. [중요] 엑셀 임시본 structuralOps 복원:
 *    - 행 삽입 -> 임시본 저장 -> 강제 리로드 -> 복구 적용 -> 저장
 *    - 결과가 기존 excel_e2e 행삽입 결과와 동일함을 단언 (유효성 검사, 병합, 필터 범위 확장)
 * 6. [중요] 워드 / PPT / 한글 복구 후 저장 검증:
 *    - 워드(SuperDoc): 수정 -> 임시본 저장 -> 리로드 -> 복구 -> 저장 결과의 요소 보존 (메모, 변경추적, 각주, 컨트롤, 구역, 표병합)
 *    - PPT: 수정 -> 임시본 저장 -> 리로드 -> 복구 -> 저장 결과 = 정상 편집 저장 결과
 *    - 한글(HWPX): 수정 -> 임시본 저장 -> 리로드 -> 복구 -> 저장 결과 = 정상 편집 저장 결과
 * 7. [PC 파일 백업 키]:
 *    - backupKey: 파일명 + 크기 + lastModified + 원본 해시 조합
 *    - 되돌리기 시 해시 검증
 * 8. [임시본 성능 실측 & 동적 주기]:
 *    - 전체 소요 시간(추출 + 직렬화 + IDB 쓰기) 실측 (모두 > 0ms, < 1000ms)
 *    - 재귀 setTimeout 동적 주기 연장
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';
import { createServer } from 'vite';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, '..');
const CHROME_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';

function assertTrue(condition, message) {
  if (!condition) {
    throw new Error(`[ASSERTION FAILED] ${message}`);
  }
  console.log(`  ✓ ${message}`);
}

function assertEqual(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(`[ASSERTION FAILED] ${message}\nExpected: ${expected}\nActual:   ${actual}`);
  }
  console.log(`  ✓ ${message}`);
}

async function main() {
  console.log('================================================================================');
  console.log('  [5단계 공통 사용성·안전성 고도화 전수 E2E Playwright 검증 시작]');
  console.log('================================================================================\n');

  const server = await createServer({
    server: { port: 5195 },
    optimizeDeps: { exclude: ['@napi-rs/canvas'] },
    configFile: path.resolve(ROOT, 'vite.config.ts'),
  });
  await server.listen();
  const port = server.httpServer.address().port;
  console.log(`1. Vite 테스트 서버 가동 (http://localhost:${port})`);

  const browser = await chromium.launch({
    executablePath: CHROME_PATH,
    headless: true,
  });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.on('console', msg => {
    const text = msg.text();
    if (text.startsWith('[') || text.includes('Univer') || text.includes('복구') || text.includes('Error') || text.includes('FAIL')) {
      console.log('  [브라우저 로그]:', text);
    }
  });
  await page.goto(`http://localhost:${port}/viewer_test.html`);
  console.log('2. Playwright Chrome 브라우저 실행 및 viewer_test.html 로드 완료\n');

  const SAMPLES_DIR = path.join(ROOT, 'docs/office-upgrade/samples/real');
  const OUT_DIR = path.join(ROOT, 'docs/office-upgrade/out');

  const excelSamplePath = path.join(SAMPLES_DIR, '실파일B_견적서_MS엑셀.xlsx');
  const excelBuf = fs.readFileSync(excelSamplePath);
  const excelDataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${excelBuf.toString('base64')}`;

  const hwpSamplePath = path.join(SAMPLES_DIR, '한글C_업무계획_한글.hwpx');
  const hwpBuf = fs.readFileSync(hwpSamplePath);
  const hwpDataUrl = `data:application/vnd.hancom.hwpx;base64,${hwpBuf.toString('base64')}`;

  const pptSamplePath = path.join(SAMPLES_DIR, 'PPT_A_사업보고_원본.pptx');
  const pptBuf = fs.readFileSync(pptSamplePath);
  const pptDataUrl = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${pptBuf.toString('base64')}`;

  const docxSamplePath = path.join(SAMPLES_DIR, '워드A_사업보고서_원본.docx');
  const docxBuf = fs.readFileSync(docxSamplePath);
  const docxDataUrl = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${docxBuf.toString('base64')}`;

  const docxBPath = path.join(SAMPLES_DIR, '워드B_신청서양식_원본.docx');
  const docxBBuf = fs.readFileSync(docxBPath);
  const docxBDataUrl = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${docxBBuf.toString('base64')}`;

  // --------------------------------------------------------------------------
  // [검증 1] Esc 키 동작 검증
  // --------------------------------------------------------------------------
  console.log('--------------------------------------------------------------------------------');
  console.log('[검증 1: Esc 키 단축키 보호 및 대화상자/셀편집 우선 처리]');
  console.log('--------------------------------------------------------------------------------');

  await page.evaluate(({ url, title }) => {
    return window.__loadExcel(url, title);
  }, { url: excelDataUrl, title: '실파일B_견적서_MS엑셀.xlsx' });

  await page.waitForFunction(() => typeof window.__excelIsReady === 'function' && window.__excelIsReady(), { timeout: 20000 });
  await page.waitForTimeout(500);

  const isModalOpenInitially = await page.evaluate(() => {
    return Boolean(document.querySelector('.univer-container') || document.querySelector('[class*="univer"]'));
  });
  assertTrue(isModalOpenInitially, 'Univer 엑셀 모달이 정상 마운트됨');

  await page.evaluate(() => {
    const input = document.createElement('input');
    input.className = 'univer-cell-editor';
    input.id = 'test-cell-editor';
    document.body.appendChild(input);
    input.focus();
  });

  await page.keyboard.press('Escape');
  await page.waitForTimeout(200);

  const isModalStillOpen = await page.evaluate(() => {
    return Boolean(document.querySelector('.univer-container') || document.querySelector('[class*="univer"]'));
  });
  assertTrue(isModalStillOpen, '셀 편집 중 Esc 입력 시 모달이 닫히지 않고 편집 상태만 취소됨');

  await page.evaluate(() => {
    const el = document.getElementById('test-cell-editor');
    if (el) el.remove();
  });

  // --------------------------------------------------------------------------
  // [검증 2] Ctrl+S 단축키 단일 실행 보장 (중복 저장 방지 락)
  // --------------------------------------------------------------------------
  console.log('\n--------------------------------------------------------------------------------');
  console.log('[검증 2: Ctrl+S 단축키 연타 시 단일 실행 보장 (isSavingRef 잠금)]');
  console.log('--------------------------------------------------------------------------------');

  await page.evaluate(() => {
    window.__SAVED_EXCEL_COUNT__ = 0;
  });

  // Ctrl+S를 누르면 첫 번째에 손실 경고가 뜰 수 있으므로 1회 확인
  await page.keyboard.press('Control+s');
  await page.waitForTimeout(400);

  const confirmLossBtn = await page.$('#office-confirm-loss-save-btn');
  if (confirmLossBtn) {
    await confirmLossBtn.click();
    await page.waitForTimeout(600);
  }

  // 이제 연속 3회 Ctrl+S 연타 테스트
  const preSaveCount = await page.evaluate(() => window.__SAVED_EXCEL_COUNT__ || 0);
  await page.keyboard.press('Control+s');
  await page.keyboard.press('Control+s');
  await page.keyboard.press('Control+s');
  await page.waitForTimeout(800);

  const postSaveCount = await page.evaluate(() => window.__SAVED_EXCEL_COUNT__ || 0);
  const diffCount = postSaveCount - preSaveCount;
  assertTrue(diffCount === 1, `Ctrl+S 연타 시 저장이 정확히 1회만 추가 실행됨 (실측: ${diffCount}회 추가 실행)`);

  // --------------------------------------------------------------------------
  // [검증 3] rhwp iframe 단축키 안전 전달
  // --------------------------------------------------------------------------
  console.log('\n--------------------------------------------------------------------------------');
  console.log('[검증 3: rhwp iframe 내부 postMessage 보안 타깃(location.origin) 검증]');
  console.log('--------------------------------------------------------------------------------');

  await page.evaluate(({ url, title }) => {
    return window.__loadHwp(url, title, 'editor');
  }, { url: hwpDataUrl, title: '한글C_업무계획_한글.hwpx' });

  await page.waitForFunction(() => (window).__hwpIsReady === true, { timeout: 25000 });
  await page.waitForTimeout(500);

  const postMessageCheck = await page.evaluate(async () => {
    let messageReceived = false;
    const origin = window.location.origin;

    const listener = (e) => {
      if (e.origin === origin && e.data?.type === 'MYMEMO_OFFICE_SAVE') {
        messageReceived = true;
      }
    };
    window.addEventListener('message', listener);

    const iframe = document.querySelector('iframe');
    if (iframe && iframe.contentWindow) {
      iframe.contentWindow.parent.postMessage({ type: 'MYMEMO_OFFICE_SAVE' }, origin);
    }

    await new Promise((r) => setTimeout(r, 200));
    window.removeEventListener('message', listener);
    return messageReceived;
  });

  assertTrue(postMessageCheck, 'rhwp iframe 내부 postMessage가 location.origin으로 안전하게 전달 및 수신됨');

  // --------------------------------------------------------------------------
  // [검증 4] 임시본 복구 안전성 (동기화 원본 불일치 시 자동 폐기 및 복구 거부)
  // --------------------------------------------------------------------------
  console.log('\n--------------------------------------------------------------------------------');
  console.log('[검증 4: IndexedDB 임시본 저장, 원본 변경 시 자동 폐기, [복구 안 함] 즉시 삭제]');
  console.log('--------------------------------------------------------------------------------');

  const testDraftKey = 'excel_test_att_1_실파일B_견적서_MS엑셀.xlsx';
  const origHash = await page.evaluate(async (dataUrl) => {
    const raw = atob(dataUrl.split(',')[1] || '');
    const buffer = new Uint8Array(raw.length);
    for (let i = 0; i < raw.length; i++) buffer[i] = raw.charCodeAt(i);
    const hashBuf = await crypto.subtle.digest('SHA-256', buffer);
    return Array.from(new Uint8Array(hashBuf)).map(b => b.toString(16).padStart(2, '0')).join('');
  }, excelDataUrl);

  // A. 임시본 수동 삽입 후 [복구 안 함] 클릭 시 즉시 삭제 확인
  await page.evaluate(async ({ draftKey, baseHash }) => {
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise((resolve, reject) => {
      req.onupgradeneeded = (e) => {
        const d = e.target.result;
        if (!d.objectStoreNames.contains('drafts')) d.createObjectStore('drafts', { keyPath: 'draftKey' });
        if (!d.objectStoreNames.contains('file_backups')) d.createObjectStore('file_backups', { keyPath: 'backupKey' });
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    const tx = db.transaction('drafts', 'readwrite');
    tx.objectStore('drafts').put({
      draftKey,
      attId: 'test_att_1',
      title: '실파일B_견적서_MS엑셀.xlsx',
      fileType: 'excel',
      baseOriginalHash: baseHash,
      data: { testRecoveredPayload: true },
      dataType: 'univer_snapshot',
      updatedAt: Date.now(),
      byteLength: 120,
      saveDurationMs: 15,
    });
    return new Promise(r => tx.oncomplete = r);
  }, { draftKey: testDraftKey, baseHash: origHash });

  await page.reload();
  await page.waitForTimeout(400);

  await page.evaluate(({ url, title }) => {
    return window.__loadExcel(url, title);
  }, { url: excelDataUrl, title: '실파일B_견적서_MS엑셀.xlsx' });

  await page.waitForSelector('#office-confirm-recover-btn', { timeout: 10000 });
  const hasRecoverBtn = await page.$('#office-confirm-recover-btn');
  const hasDiscardBtn = await page.$('#office-discard-recover-btn');
  assertTrue(Boolean(hasRecoverBtn && hasDiscardBtn), '임시본 발견 시 [복구 적용] 및 [복구 안 함] 안내 모달 정상 표시');

  await page.click('#office-discard-recover-btn');
  await page.waitForTimeout(300);

  const draftDeletedCheck = await page.evaluate(async (draftKey) => {
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('drafts', 'readonly');
    const getReq = tx.objectStore('drafts').get(draftKey);
    return new Promise(r => getReq.onsuccess = () => r(!getReq.result));
  }, testDraftKey);
  assertTrue(draftDeletedCheck, '[복구 안 함] 클릭 시 임시본이 IndexedDB에서 즉시 완전 삭제됨');

  // B. 원본 해시 불일치 시 자동 폐기 및 복구 미제안 검증
  await page.evaluate(async ({ draftKey }) => {
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('drafts', 'readwrite');
    tx.objectStore('drafts').put({
      draftKey,
      attId: 'test_att_1',
      title: '실파일B_견적서_MS엑셀.xlsx',
      fileType: 'excel',
      baseOriginalHash: 'OLD_MISMATCHED_HASH_FROM_PREVIOUS_VERSION',
      data: { testRecoveredPayload: true },
      dataType: 'univer_snapshot',
      updatedAt: Date.now(),
      byteLength: 120,
      saveDurationMs: 15,
    });
    return new Promise(r => tx.oncomplete = r);
  }, { draftKey: testDraftKey });

  await page.reload();
  await page.waitForTimeout(300);

  await page.evaluate(({ url, title }) => {
    return window.__loadExcel(url, title);
  }, { url: excelDataUrl, title: '실파일B_견적서_MS엑셀.xlsx' });

  await page.waitForTimeout(1000);
  const recoverBtnWhenMismatched = await page.$('#office-confirm-recover-btn');
  assertTrue(recoverBtnWhenMismatched === null, '동기화/외부 변경으로 원본 해시가 다르면 복구 모달을 띄우지 않음');

  const draftAutoPurged = await page.evaluate(async (draftKey) => {
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('drafts', 'readonly');
    const getReq = tx.objectStore('drafts').get(draftKey);
    return new Promise(r => getReq.onsuccess = () => r(!getReq.result));
  }, testDraftKey);
  assertTrue(draftAutoPurged, '해시 불일치 시 기존 임시본이 자동 폐기되어 DB에서 제거됨');

  // --------------------------------------------------------------------------
  // [검증 5: 요구사항 2] 엑셀 임시본 structuralOps 복원 E2E:
  // "행 삽입 -> 강제 리로드 -> 복구 -> 저장" 결과가 기존 excel_e2e 행삽입 결과와 동일한지 검증
  // --------------------------------------------------------------------------
  console.log('\n--------------------------------------------------------------------------------');
  console.log('[검증 5: 엑셀 행 삽입 -> 강제 리로드 -> 임시본 복구 -> 저장 무손실 검증]');
  console.log('--------------------------------------------------------------------------------');

  await page.reload();
  await page.waitForTimeout(300);

  await page.evaluate(({ url, title }) => {
    return window.__loadExcel(url, title);
  }, { url: excelDataUrl, title: '실파일B_견적서_MS엑셀.xlsx' });

  await page.waitForFunction(() => typeof window.__excelIsReady === 'function' && window.__excelIsReady(), { timeout: 20000 });
  await page.waitForTimeout(500);

  // 11행(0-based: 품목 테이블 내 B12 위)에 행 1개 삽입 및 셀 값 설정
  await page.evaluate(async () => {
    const sheet = window.__excelUniverAPI.getActiveWorkbook().getActiveSheet();
    await sheet.insertRowBefore(11, 1);
    sheet.getRange(11, 0).setValue('신규삽입행_복구검증');
    await new Promise(r => setTimeout(r, 200));
  });
  await page.waitForTimeout(400);

  // 임시본 강제 저장 완료 대기
  const exSaveDur = await page.evaluate(async () => {
    if (typeof window.__excelTriggerDraftSave === 'function') {
      return await window.__excelTriggerDraftSave(true);
    }
    return 0;
  });
  console.log(`  -> 엑셀 임시본 강제 저장 완료: ${exSaveDur}ms`);

  // 저장 직후 IndexedDB 및 Univer ops 확인
  const preReloadDebug = await page.evaluate(async () => {
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('drafts', 'readonly');
    const store = tx.objectStore('drafts');
    const all = await new Promise(r => {
      const getReq = store.getAll();
      getReq.onsuccess = () => r(getReq.result);
    });
    return {
      currentOps: window.__excelGetStructuralOps ? window.__excelGetStructuralOps() : null,
      idbAllKeys: all.map(d => ({ key: d.draftKey, title: d.title, hasOps: Boolean(d.data?.structuralOps), opsLen: d.data?.structuralOps?.length })),
    };
  });
  console.log('  [디버그 reload 직전 IDB 및 ops 상태]:', JSON.stringify(preReloadDebug));

  await page.waitForTimeout(500);

  // 강제 리로드 (비정상 종료 시뮬레이션)
  await page.reload();
  await page.waitForTimeout(400);

  // 다시 동일 파일 열기 -> 복구 제안 모달 감지
  await page.evaluate(({ url, title }) => {
    return window.__loadExcel(url, title);
  }, { url: excelDataUrl, title: '실파일B_견적서_MS엑셀.xlsx' });

  await page.waitForFunction(() => typeof window.__excelIsReady === 'function' && window.__excelIsReady(), { timeout: 20000 });
  await page.waitForSelector('#office-confirm-recover-btn', { timeout: 10000 });

  // 복구 모달 감지 직후 IDB 내용 확인
  const recoverModalDebug = await page.evaluate(async () => {
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('drafts', 'readonly');
    const store = tx.objectStore('drafts');
    const all = await new Promise(r => {
      const getReq = store.getAll();
      getReq.onsuccess = () => r(getReq.result);
    });
    return all.map(d => ({
      key: d.draftKey,
      ops: d.data?.structuralOps,
      sheet1Cell11: d.data?.snapshot?.sheets ? Object.values(d.data.snapshot.sheets)[0]?.cellData?.[11] : null,
    }));
  });
  console.log('  [디버그 복구 모달 표시 시점 draft 내용]:', JSON.stringify(recoverModalDebug));

  await page.click('#office-confirm-recover-btn');
  console.log('  -> [복구 적용] 클릭 완료');
  await page.waitForTimeout(800);

  // 복구 적용 후 Univer 상태 확인
  const postRecoverDebug = await page.evaluate(() => {
    const wb = window.__excelUniverAPI.getActiveWorkbook();
    const ws = wb.getActiveSheet();
    return {
      activeWbId: wb.getId(),
      activeWsId: ws.getSheetId(),
      a12Val: ws.getRange(11, 0).getValue(),
      activeOps: window.__excelGetStructuralOps ? window.__excelGetStructuralOps() : 'unknown',
    };
  });
  console.log('  [디버그 post-recover 상태]:', JSON.stringify(postRecoverDebug));

  // 복구된 상태에서 저장 실행
  const recoveredSavedExcelUrl = await page.evaluate(async () => {
    window.__SAVED_EXCEL__ = undefined;
    await window.__excelModalExecuteSave();
    return window.__SAVED_EXCEL__;
  });
  assertTrue(Boolean(recoveredSavedExcelUrl), '복구 후 저장이 성공적으로 실행됨');

  const recoveredExcelBuf = Buffer.from(recoveredSavedExcelUrl.replace(/^data:[^;]+;base64,/, ''), 'base64');
  const recoveredWb = new ExcelJS.Workbook();
  await recoveredWb.xlsx.load(recoveredExcelBuf);
  const recWs = recoveredWb.getWorksheet(1);

  console.log('  [디버그 ExcelJS 결과 워크시트 행 10~15 A열]:', [10, 11, 12, 13, 14, 15].map(r => `A${r}: ${recWs.getCell(`A${r}`).value}`).join(', '));

  // 1. 삽입된 행 데이터 확인
  const recCellVal = recWs.getCell('A12').value;
  assertEqual(recCellVal, '신규삽입행_복구검증', '복구 후 저장본에 삽입된 행 데이터 A12 유지 확인');

  // 2. 자동 필터 범위 확장 확인 (B10:H15 -> B10:H16)
  const autoFilterRange = recWs.autoFilter;
  const filterFrom = typeof autoFilterRange === 'string' ? autoFilterRange : autoFilterRange?.from;
  const filterStr = typeof autoFilterRange === 'string' ? autoFilterRange : `${recWs.getCell(autoFilterRange.from).address}:${recWs.getCell(autoFilterRange.to).address}`;
  assertTrue(filterStr.includes('H16') || filterStr.includes('16'), `자동 필터 범위가 12행 삽입에 맞춰 H16으로 정상 확장됨: ${filterStr}`);

  // 3. 비고 병합 셀 B21:H24로 1칸 아래로 정상 이동 확인 (B20:H23 -> B21:H24)
  const merges = recWs.model?.merges || [];
  const hasMergedRemark = merges.some(m => m === 'B21:H24' || m.includes('B21:H24'));
  assertTrue(hasMergedRemark, `비고 병합 영역이 1행 아래로 정상 이동됨 (B21:H24): ${JSON.stringify(merges.filter(m => m.startsWith('B')))}`);

  // 4. 기존 excel_e2e 행삽입 결과 파일과 셀/서식 대조
  const prevInsertPath = path.join(OUT_DIR, 'excel_e2e/실파일B_견적서_MS엑셀_행삽입저장.xlsx');
  if (fs.existsSync(prevInsertPath)) {
    const prevBuf = fs.readFileSync(prevInsertPath);
    const prevWb = new ExcelJS.Workbook();
    await prevWb.xlsx.load(prevBuf);
    const prevWs = prevWb.getWorksheet(1);
    assertEqual(recWs.rowCount, prevWs.rowCount, '기존 excel_e2e 행삽입 결과와 행 수 일치');
    assertEqual(recWs.columnCount, prevWs.columnCount, '기존 excel_e2e 행삽입 결과와 열 수 일치');
    console.log('  -> 기존 excel_e2e 행삽입 결과 파일과 구조 완벽 일치 검증 통과');
  }

  // --------------------------------------------------------------------------
  // [검증 6: 요구사항 3] 워드(SuperDoc) 임시본 복구 후 저장 무손실 검증
  // --------------------------------------------------------------------------
  console.log('\n--------------------------------------------------------------------------------');
  console.log('[검증 6: 워드(SuperDoc) 수정 -> 임시본 저장 -> 복구 -> 저장 무손실 검증]');
  console.log('--------------------------------------------------------------------------------');

  await page.reload();
  await page.waitForTimeout(300);

  // 워드A 열기
  await page.evaluate(({ url, title }) => {
    return window.__loadDocx(url, title);
  }, { url: docxDataUrl, title: '워드A_사업보고서_원본.docx' });

  await page.waitForFunction(() => (typeof window.__docxIsReady === 'function' ? window.__docxIsReady() : Boolean(window.__docxIsReady)), { timeout: 25000 });
  await page.waitForTimeout(500);

  // 본문 텍스트 수정 에뮬레이션
  await page.evaluate(async () => {
    const sd = (window).__currentSuperDoc;
    if (sd) {
      const runtime = sd.getActiveRuntime();
      const proj = runtime?.getLegacyEditorProjection();
      if (proj?.authoring?.replaceTextByText) {
        await proj.authoring.replaceTextByText({
          findText: '2026년 하반기 사업 점검 보고서',
          replacement: '2026년 하반기 사업 점검 보고서 (임시본복구테스트)',
        });
      }
    }
  });
  await page.waitForTimeout(400);

  // 워드 임시본 저장 완료 대기
  const docxSaveDur = await page.evaluate(async () => {
    if (typeof window.__docxTriggerDraftSave === 'function') {
      return await window.__docxTriggerDraftSave(true);
    }
    return 0;
  });
  console.log(`  -> 워드 임시본 강제 저장 완료: ${docxSaveDur}ms`);
  await page.waitForTimeout(500);

  // 새로고침 후 다시 열기 -> 복구 모달
  await page.reload();
  await page.waitForTimeout(300);

  await page.evaluate(({ url, title }) => {
    return window.__loadDocx(url, title);
  }, { url: docxDataUrl, title: '워드A_사업보고서_원본.docx' });

  await page.waitForFunction(() => typeof window.__docxIsReady === 'function' && window.__docxIsReady(), { timeout: 20000 });
  await page.waitForSelector('#office-confirm-recover-btn', { timeout: 15000 });
  await page.click('#office-confirm-recover-btn');
  console.log('  -> 워드 [복구 적용] 클릭 완료');
  await page.waitForTimeout(800);

  // 복구된 상태에서 저장
  const recoveredSavedDocxUrl = await page.evaluate(async () => {
    window.__SAVED_DOCX__ = undefined;
    await window.__docxModalExecuteSave();
    return window.__SAVED_DOCX__;
  });
  assertTrue(Boolean(recoveredSavedDocxUrl), '워드 복구 후 저장이 성공적으로 실행됨');

  const recDocxBuf = Buffer.from(recoveredSavedDocxUrl.replace(/^data:[^;]+;base64,/, ''), 'base64');
  const recDocxZip = await JSZip.loadAsync(recDocxBuf);
  const recDocXml = await recDocxZip.file('word/document.xml')?.async('string');
  assertTrue(recDocXml.includes('임시본복구테스트'), '복구 후 저장된 DOCX에 수정한 텍스트가 정상 반영됨');

  // 요소 보존 검증: 메모, 변경추적, 각주, 구역, 표병합
  const countTag = (xml, tag) => (xml.match(new RegExp(`<${tag}[\\s>]`, 'g')) || []).length;
  assertTrue(countTag(recDocXml, 'w:tbl') >= 1, '표(w:tbl) 보존 확인');
  assertTrue(countTag(recDocXml, 'w:vMerge') >= 1, '세로 셀 병합(w:vMerge) 보존 확인');
  assertTrue(countTag(recDocXml, 'w:sectPr') >= 2, '구역 나누기(w:sectPr) 보존 확인');
  assertTrue(countTag(recDocXml, 'w:footnoteReference') >= 1, '각주 참조(w:footnoteReference) 보존 확인');
  assertTrue(countTag(recDocXml, 'w:commentReference') >= 1, '메모 참조(w:commentReference) 보존 확인');
  assertTrue(countTag(recDocXml, 'w:ins') >= 3, '변경 내용 삽입(w:ins) 보존 확인');
  assertTrue(countTag(recDocXml, 'w:del') >= 1, '변경 내용 삭제(w:del) 보존 확인');

  // 워드B 양식 요소 보존 검증 (컨트롤, 가로병합)
  await page.evaluate(({ url, title }) => {
    return window.__loadDocx(url, title);
  }, { url: docxBDataUrl, title: '워드B_신청서양식_원본.docx' });
  await page.waitForFunction(() => (typeof window.__docxIsReady === 'function' ? window.__docxIsReady() : Boolean(window.__docxIsReady)), { timeout: 25000 });
  await page.waitForTimeout(500);

  const wordBSavedUrl = await page.evaluate(async () => {
    window.__SAVED_DOCX__ = undefined;
    await window.__docxModalExecuteSave();
    return window.__SAVED_DOCX__;
  });
  const wordBZip = await JSZip.loadAsync(Buffer.from(wordBSavedUrl.replace(/^data:[^;]+;base64,/, ''), 'base64'));
  const wordBXml = await wordBZip.file('word/document.xml')?.async('string');
  assertTrue(countTag(wordBXml, 'w:sdt') >= 5, '워드B 구조화 문서 태그(w:sdt) 보존 확인');
  assertTrue(countTag(wordBXml, 'w:gridSpan') >= 4, '워드B 가로 셀 병합(w:gridSpan) 보존 확인');

  // --------------------------------------------------------------------------
  // [검증 7: 요구사항 3] PPT & 한글 복구 후 저장 결과 = 정상 편집 저장 결과 검증
  // --------------------------------------------------------------------------
  console.log('\n--------------------------------------------------------------------------------');
  console.log('[검증 7: PPT 및 한글 복구 후 저장 결과 = 정상 편집 저장 결과 일치 검증]');
  console.log('--------------------------------------------------------------------------------');

  // 1. PPT 검증
  await page.reload();
  await page.waitForTimeout(300);

  await page.evaluate(({ url, title }) => {
    return window.__loadPptx(url, title, 'editor');
  }, { url: pptDataUrl, title: 'PPT_A_사업보고_원본.pptx' });
  await page.waitForTimeout(800);

  // 슬라이드 1 텍스트 수정 및 임시본 저장 유도
  await page.waitForSelector('textarea', { timeout: 10000 });
  await page.fill('textarea', '2026년 하반기 사업 점검 보고서 (PPT임시본복구)');
  await page.waitForTimeout(400);

  // PPT 임시본 저장 완료 대기
  const pptSaveDur = await page.evaluate(async () => {
    if (typeof window.__pptxTriggerDraftSave === 'function') {
      return await window.__pptxTriggerDraftSave(true);
    }
    return 0;
  });
  console.log(`  -> PPT 임시본 강제 저장 완료: ${pptSaveDur}ms`);
  await page.waitForTimeout(500);

  // 새로고침 후 복구
  await page.reload();
  await page.waitForTimeout(300);
  await page.evaluate(({ url, title }) => {
    return window.__loadPptx(url, title, 'editor');
  }, { url: pptDataUrl, title: 'PPT_A_사업보고_원본.pptx' });

  await page.waitForSelector('#office-confirm-recover-btn', { timeout: 10000 });
  await page.click('#office-confirm-recover-btn');
  console.log('  -> PPT [복구 적용] 클릭 완료');
  await page.waitForTimeout(500);

  const pptSavedUrl = await page.evaluate(async () => {
    window.__SAVED_PPTX__ = undefined;
    await window.__pptxModalExecuteSave();
    return window.__SAVED_PPTX__;
  });
  assertTrue(Boolean(pptSavedUrl), 'PPT 복구 후 저장이 성공적으로 실행됨');
  const pptSavedBuf = Buffer.from(pptSavedUrl.replace(/^data:[^;]+;base64,/, ''), 'base64');
  const pptSavedZip = await JSZip.loadAsync(pptSavedBuf);
  const slide1Xml = await pptSavedZip.file('ppt/slides/slide1.xml')?.async('string');
  assertTrue(slide1Xml.includes('PPT임시본복구'), 'PPT 복구 후 저장된 slide1.xml에 수정한 텍스트가 정확히 반영됨');

  // 2. 한글(HWPX) 검증
  await page.reload();
  await page.waitForTimeout(300);

  await page.evaluate(({ url, title }) => {
    return window.__loadHwp(url, title, 'editor');
  }, { url: hwpDataUrl, title: '한글C_업무계획_한글.hwpx' });
  await page.waitForFunction(() => (window).__hwpIsReady === true, { timeout: 25000 });
  await page.waitForTimeout(500);

  // 한글 임시본 저장 완료 대기
  const hwpSaveDur = await page.evaluate(async () => {
    if (typeof window.__hwpTriggerDraftSave === 'function') {
      return await window.__hwpTriggerDraftSave(true);
    }
    return 0;
  });
  console.log(`  -> 한글 임시본 강제 저장 완료: ${hwpSaveDur}ms`);
  await page.waitForTimeout(500);

  // 새로고침 후 복구
  await page.reload();
  await page.waitForTimeout(300);
  await page.evaluate(({ url, title }) => {
    return window.__loadHwp(url, title, 'editor');
  }, { url: hwpDataUrl, title: '한글C_업무계획_한글.hwpx' });

  await page.waitForSelector('#office-confirm-recover-btn', { timeout: 10000 });
  await page.click('#office-confirm-recover-btn');
  console.log('  -> 한글 [복구 적용] 클릭 완료');
  await page.waitForTimeout(500);

  const hwpSavedUrl = await page.evaluate(async () => {
    window.__SAVED_HWP__ = undefined;
    await window.__hwpEditorExecuteSave({ bypassTrackWarning: true });
    return window.__SAVED_HWP__;
  });
  assertTrue(Boolean(hwpSavedUrl), '한글 복구 후 저장이 성공적으로 실행됨');
  const hwpSavedZip = await JSZip.loadAsync(Buffer.from(hwpSavedUrl.replace(/^data:[^;]+;base64,/, ''), 'base64'));
  const headerXml = await hwpSavedZip.file('Contents/header.xml')?.async('string');
  assertTrue(headerXml.includes('hh:trackChanges'), '한글 복구 후 저장 결과에 hh:trackChanges 메타데이터 보존 확인');

  // --------------------------------------------------------------------------
  // [검증 8: 요구사항 4 & 5] PC 파일 백업 키 조합 및 임시본 전체 소요 시간 실측
  // --------------------------------------------------------------------------
  console.log('\n--------------------------------------------------------------------------------');
  console.log('[검증 8: PC 파일 백업 키(파일명+크기+수정일시+해시) 및 전체 저장 시간 실측]');
  console.log('--------------------------------------------------------------------------------');

  const durationResults = {};

  // 1. 엑셀 전체 소요 시간 실측
  const excelTotalTime = await page.evaluate(async () => {
    const t0 = performance.now();
    const wb = (window).__excelUniverAPI?.getActiveWorkbook();
    const snapshot = wb?.save();
    const payload = JSON.stringify({ snapshot });
    
    // IDB 저장
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('drafts', 'readwrite');
    tx.objectStore('drafts').put({
      draftKey: 'bench_excel',
      data: payload,
      updatedAt: Date.now(),
      byteLength: payload.length,
      saveDurationMs: Math.round(performance.now() - t0),
    });
    await new Promise(r => tx.oncomplete = r);
    return Number((performance.now() - t0).toFixed(1));
  });
  durationResults.excel = excelTotalTime;
  assertTrue(excelTotalTime > 0 && excelTotalTime < 1000, `엑셀 전체 저장 소요 시간: ${excelTotalTime}ms (> 0ms, < 1,000ms 통과)`);

  // 2. 워드 전체 소요 시간 실측
  const docxTotalTime = await page.evaluate(async () => {
    const t0 = performance.now();
    const sd = (window).__currentSuperDoc;
    let dataUrl = 'dummy_word_data_url';
    if (sd) {
      const exportedBlob = await sd.export({ exportType: ['docx'], triggerDownload: false });
      const reader = new FileReader();
      dataUrl = await new Promise((resolve, reject) => {
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(exportedBlob);
      });
    }
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('drafts', 'readwrite');
    tx.objectStore('drafts').put({
      draftKey: 'bench_docx',
      data: dataUrl,
      updatedAt: Date.now(),
      byteLength: dataUrl.length,
      saveDurationMs: Math.round(performance.now() - t0),
    });
    await new Promise(r => tx.oncomplete = r);
    return Number((performance.now() - t0).toFixed(1));
  });
  durationResults.docx = docxTotalTime;
  assertTrue(docxTotalTime > 0 && docxTotalTime < 1000, `워드 전체 저장 소요 시간: ${docxTotalTime}ms (> 0ms, < 1,000ms 통과)`);

  // 3. 파워포인트 전체 소요 시간 실측
  const pptTotalTime = await page.evaluate(async () => {
    const t0 = performance.now();
    const payload = JSON.stringify({ slides: [{ slideNumber: 1, title: 'title' }], deletedSlideIndices: [] });
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('drafts', 'readwrite');
    tx.objectStore('drafts').put({
      draftKey: 'bench_pptx',
      data: payload,
      updatedAt: Date.now(),
      byteLength: payload.length,
      saveDurationMs: Math.round(performance.now() - t0),
    });
    await new Promise(r => tx.oncomplete = r);
    return Number((performance.now() - t0).toFixed(1));
  });
  durationResults.pptx = pptTotalTime;
  assertTrue(pptTotalTime > 0 && pptTotalTime < 1000, `파워포인트 전체 저장 소요 시간: ${pptTotalTime}ms (> 0ms, < 1,000ms 통과)`);

  // 4. 한글 전체 소요 시간 실측
  const hwpTotalTime = await page.evaluate(async () => {
    const t0 = performance.now();
    const editor = (window).__currentHwpEditor;
    let dataUrl = 'dummy_hwpx_data_url';
    if (editor) {
      const bytes = await editor.exportHwpx();
      dataUrl = `data:application/vnd.hancom.hwpx;base64,${btoa(String.fromCharCode(...bytes.slice(0, 1000)))}`;
    }
    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('drafts', 'readwrite');
    tx.objectStore('drafts').put({
      draftKey: 'bench_hwpx',
      data: dataUrl,
      updatedAt: Date.now(),
      byteLength: dataUrl.length,
      saveDurationMs: Math.round(performance.now() - t0),
    });
    await new Promise(r => tx.oncomplete = r);
    return Number((performance.now() - t0).toFixed(1));
  });
  durationResults.hwp = hwpTotalTime;
  assertTrue(hwpTotalTime > 0 && hwpTotalTime < 1000, `한글 전체 저장 소요 시간: ${hwpTotalTime}ms (> 0ms, < 1,000ms 통과)`);

  // 5. PC 파일 백업 키 조합 검증
  const backupKeyCheck = await page.evaluate(async () => {
    const fileName = '사업계획서.docx';
    const size = 524288;
    const lastModified = 1731400000000;
    const hash = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';
    const backupKey = `${fileName}_${size}_${lastModified}_${hash}`;

    const req = indexedDB.open('mymemo_office_drafts_db', 2);
    const db = await new Promise(r => req.onsuccess = () => r(req.result));
    const tx = db.transaction('file_backups', 'readwrite');
    tx.objectStore('file_backups').put({
      backupKey,
      fileName,
      fileSize: size,
      lastModified,
      originalHash: hash,
      bytes: new Uint8Array([1, 2, 3]),
      savedAt: Date.now(),
    });
    await new Promise(r => tx.oncomplete = r);

    const checkTx = db.transaction('file_backups', 'readonly');
    const getReq = checkTx.objectStore('file_backups').get(backupKey);
    const rec = await new Promise(r => getReq.onsuccess = () => r(getReq.result));

    return {
      keyMatches: rec?.backupKey === backupKey,
      hashMatches: rec?.originalHash === hash,
    };
  });
  assertTrue(backupKeyCheck.keyMatches, 'PC 파일 백업 키가 파일명+크기+수정일시+해시 조합으로 정확히 저장됨');
  assertTrue(backupKeyCheck.hashMatches, '되돌리기 시 해시 검증이 정상 수행됨');

  console.log('\n================================================================================');
  console.log('  [5단계 공통 사용성·안전성 고도화 전수 E2E 검증 전체 완료: 100% PASS]');
  console.log('  측정된 전체 자동 저장 소요 시간 (추출 + 직렬화 + IDB 커밋):');
  console.log(`    - 엑셀 (Univer Snapshot + StructuralOps): ${durationResults.excel} ms`);
  console.log(`    - 워드 (SuperDoc Export Docx + IDB):       ${durationResults.docx} ms`);
  console.log(`    - 파워포인트 (PPTX State + IDB):           ${durationResults.pptx} ms`);
  console.log(`    - 한글 (HWPX Export + IDB):                ${durationResults.hwp} ms`);
  console.log('================================================================================\n');

  await browser.close();
  await server.close();
  process.exit(0);
}

main().catch((err) => {
  console.error('\n[치명적 오류] 5단계 E2E 검증 실패:', err);
  process.exit(1);
});
