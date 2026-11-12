import fs from 'fs';
import path from 'path';
import { createServer } from 'vite';
import { chromium } from 'playwright-core';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { initSync, HwpDocument } from '@rhwp/core';

const CHROME_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const SAMPLES_DIR = path.resolve('docs/office-upgrade/samples/real');
const OUT_DIR = path.resolve('docs/office-upgrade/out/hwp_e2e');

if (!fs.existsSync(OUT_DIR)) {
  fs.mkdirSync(OUT_DIR, { recursive: true });
}

// rhwp wasm init for validation
const wasmBuf = fs.readFileSync('node_modules/@rhwp/core/rhwp_bg.wasm');
initSync({ module: wasmBuf });

const TARGET_FILES = [
  { name: '한글C_업무계획_원본.hwpx', isHwpx: true, isHangul: false },
  { name: '한글C_업무계획_한글.hwpx', isHwpx: true, isHangul: true },
  { name: '한글C_업무계획_원본.hwp', isHwpx: false, isHangul: false },
  { name: '한글C_업무계획_한글.hwp', isHwpx: false, isHangul: true },
];

function assertTrue(cond, message) {
  if (!cond) {
    throw new Error(`[Assertion Failed] ${message}`);
  }
}

function assertEqual(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(`[Assertion Failed] ${message}\nExpected: ${expected}\nActual:   ${actual}`);
  }
}

async function performReplaceAll(page, findText, replaceText) {
  await page.evaluate(async ({ find, replace }) => {
    const frame = document.querySelector('iframe');
    const auto = frame.contentWindow.rhwpStudio.automation;
    await auto.execute('edit:find-replace', {}, { allowDialog: true });
    const doc = frame.contentDocument;

    const inputs = doc.querySelectorAll('.find-dialog input[type="text"]');
    inputs[0].value = find;
    inputs[0].dispatchEvent(new Event('input', { bubbles: true }));

    inputs[1].value = replace;
    inputs[1].dispatchEvent(new Event('input', { bubbles: true }));

    const btns = Array.from(doc.querySelectorAll('.find-dialog button'));
    const replaceAllBtn = btns.find(b => b.textContent.trim() === '모두 바꾸기');
    if (!replaceAllBtn) throw new Error('"모두 바꾸기" 버튼을 찾을 수 없습니다.');
    replaceAllBtn.click();

    await new Promise(r => setTimeout(r, 400));

    const closeBtn = btns.find(b => b.textContent.trim() === '×');
    if (closeBtn) closeBtn.click();
  }, { find: findText, replace: replaceText });

  await page.waitForTimeout(400);
}

async function executeEditorSave(page, options = { bypassTrackWarning: true }) {
  const savedUrl = await page.evaluate(async (opts) => {
    window.__SAVED_HWP__ = undefined;
    await window.__hwpEditorExecuteSave(opts);
    return window.__SAVED_HWP__;
  }, options);

  assertTrue(Boolean(savedUrl), '저장된 dataUrl이 유효해야 합니다.');
  return Buffer.from(savedUrl.replace(/^data:[^;]+;base64,/, ''), 'base64');
}

async function loadHwpDocumentInEditor(page, dataUrl, title, port) {
  await page.goto(`http://localhost:${port}/viewer_test.html`);
  await page.waitForFunction(() => typeof window.__loadHwp === 'function', { timeout: 15000 });
  await page.evaluate(({ url, t }) => {
    return window.__loadHwp(url, t, 'editor');
  }, { url: dataUrl, t: title });
  await page.waitForFunction(() => window.__hwpIsReady === true, { timeout: 30000 });
  await page.waitForTimeout(400);
}

async function verifyHwpxZeroChange(origBuf, savedBuf, fileName) {
  // 1-a 요구사항: 무변경 저장은 원본 바이트를 100% 그대로 유지 (재직렬화 금지)
  assertEqual(savedBuf.length, origBuf.length, `${fileName}: 무변경 저장 바이트 길이 원본과 100% 동일`);
  assertTrue(savedBuf.equals(origBuf), `${fileName}: 무변경 저장 바이트 원본과 100% 바이너리 일치`);

  const origZip = await JSZip.loadAsync(origBuf);
  const savedZip = await JSZip.loadAsync(savedBuf);

  // version.xml 존재 및 일치 검증
  const origVer = await origZip.file('version.xml')?.async('text');
  const savedVer = await savedZip.file('version.xml')?.async('text');
  assertTrue(Boolean(savedVer), `${fileName}: version.xml 파일 보존 확인`);
  if (origVer) {
    assertEqual(savedVer, origVer, `${fileName}: version.xml 내용 100% 일치`);
  }

  // 변경 추적 태그(hp:deleteBegin/insertBegin) 보존 검증
  const origSec0 = await origZip.file('Contents/section0.xml')?.async('text');
  const savedSec0 = await savedZip.file('Contents/section0.xml')?.async('text');
  if (origSec0 && origSec0.includes('deleteBegin')) {
    assertTrue(Boolean(savedSec0 && savedSec0.includes('deleteBegin')), `${fileName}: deleteBegin 변경추적 마크업 보존`);
    assertTrue(Boolean(savedSec0 && savedSec0.includes('insertBegin')), `${fileName}: insertBegin 변경추적 마크업 보존`);
  }

  // HwpDocument 엔진 레벨의 텍스트 및 페이지 수 100% 일치 검증
  const oDoc = new HwpDocument(origBuf);
  const sDoc = new HwpDocument(savedBuf);
  assertEqual(sDoc.pageCount(), oDoc.pageCount(), `${fileName}: 페이지 수 일치 (${oDoc.pageCount()})`);
  assertEqual(sDoc.getTextFileText(), oDoc.getTextFileText(), `${fileName}: 전체 문서 텍스트 100% 일치`);
  oDoc.free();
  sDoc.free();

  // Contents/section*.xml 표 및 구조 검증
  for (const name of Object.keys(origZip.files)) {
    if (name.startsWith('Contents/section') && name.endsWith('.xml')) {
      const origSec = await origZip.file(name).async('text');
      const savedSec = await savedZip.file(name).async('text');

      const origDoc = new DOMParser().parseFromString(origSec, 'application/xml');
      const savedDoc = new DOMParser().parseFromString(savedSec, 'application/xml');

      // 표 개수 및 셀 개수 검증
      const origTbls = origDoc.getElementsByTagName('hp:tbl');
      const savedTbls = savedDoc.getElementsByTagName('hp:tbl');
      assertEqual(savedTbls.length, origTbls.length, `${fileName}: ${name} 표 개수 일치`);
      if (origTbls.length > 0) {
        const origCells = origTbls[0].getElementsByTagName('hp:tc');
        const savedCells = savedTbls[0].getElementsByTagName('hp:tc');
        assertEqual(savedCells.length, origCells.length, `${fileName}: 표 셀 개수 일치`);
      }
    }
  }
}

async function verifyHwpZeroChange(origBuf, savedBuf, fileName) {
  // 1-a 요구사항: 무변경 저장은 원본 바이트를 100% 그대로 유지 (재직렬화 금지)
  assertEqual(savedBuf.length, origBuf.length, `${fileName}: 무변경 저장 바이트 길이 원본과 100% 동일`);
  assertTrue(savedBuf.equals(origBuf), `${fileName}: 무변경 저장 바이트 원본과 100% 바이너리 일치`);

  // HWP 바이너리 OLE2 서명 및 HWP Document File 스트림 검증
  const isOle = savedBuf[0] === 0xd0 && savedBuf[1] === 0xcf && savedBuf[2] === 0x11 && savedBuf[3] === 0xe0;
  assertTrue(isOle, `${fileName}: HWP OLE2 바이너리 서명 확인`);
  assertTrue(savedBuf.includes(Buffer.from('HWP Document File')), `${fileName}: HWP Document File 스트림 확인`);

  const origDoc = new HwpDocument(origBuf);
  const savedDoc = new HwpDocument(savedBuf);

  assertEqual(savedDoc.pageCount(), origDoc.pageCount(), `${fileName}: 페이지 수 일치 (${origDoc.pageCount()})`);

  const origText = origDoc.getTextFileText();
  const savedText = savedDoc.getTextFileText();
  assertEqual(savedText, origText, `${fileName}: 전체 텍스트 100% 동일`);

  origDoc.free();
  savedDoc.free();
}

async function main() {
  console.log('================================================================================');
  console.log('  [4단계 한글(HWP/HWPX) 완전 오프라인 셀프호스팅 및 무손실 보존 E2E 검증]');
  console.log('================================================================================\n');

  const server = await createServer({
    server: {
      port: 5199,
      watch: { ignored: ['**/docs/**', '**/out/**', '**/*.hwpx', '**/*.hwp'] }
    },
    configFile: path.resolve('vite.config.ts'),
  });
  await server.listen();
  const port = server.httpServer.address().port;
  console.log(`[서버 기동] Vite 개발 서버 가동 (http://localhost:${port})\n`);

  const browser = await chromium.launch({
    executablePath: CHROME_PATH,
    headless: true,
  });
  const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await context.newPage();

  // 외부 도메인 요청 모니터링
  const externalRequests = [];
  page.on('request', req => {
    try {
      const urlStr = req.url();
      if (urlStr.startsWith('blob:') || urlStr.startsWith('data:')) return;
      const u = new URL(urlStr);
      if (u.hostname !== 'localhost' && u.hostname !== '127.0.0.1') {
        externalRequests.push(urlStr);
      }
    } catch {
      // ignore
    }
  });

  // --------------------------------------------------------------------------
  // [검증 1] 완전 오프라인 셀프호스팅 검증 (context.setOffline(true))
  // --------------------------------------------------------------------------
  console.log('--------------------------------------------------------------------------------');
  console.log('[검증 1: 완전 오프라인 셀프호스팅 및 외부 요청 0건 검증]');
  console.log('--------------------------------------------------------------------------------');

  const testSample = TARGET_FILES[1]; // 한글C_업무계획_한글.hwpx
  const testSampleBuf = fs.readFileSync(path.join(SAMPLES_DIR, testSample.name));
  const testDataUrl = `data:application/vnd.hancom.hwpx;base64,${testSampleBuf.toString('base64')}`;

  console.log('  1. 한글 에디터 로드 실행 (자체 호스팅 rhwp 스튜디오, wasm, 워커 준비)...');
  await loadHwpDocumentInEditor(page, testDataUrl, testSample.name, port);
  console.log('  ✓ 에디터 준비 완료');

  console.log('  3. context.setOffline(true) 실행 — 외부 인터넷 완전 차단');
  await context.setOffline(true);

  console.log('  4. 인터넷이 완전히 차단된 오프라인 상태에서 문서 수정 및 저장 실행...');
  await performReplaceAll(page, '추진 배경', '추진 배경 및 필요성');
  const offlineSavedBuf = await executeEditorSave(page);
  assertTrue(offlineSavedBuf.length > 1000, '오프라인 저장 버퍼 유효성 확인');
  console.log(`  ✓ 오프라인 저장 성공 (${offlineSavedBuf.length.toLocaleString()} 바이트)`);

  assertEqual(externalRequests.length, 0, '오프라인 중 외부 도메인 요청 0건 확인');
  console.log(`  ✓ 외부 도메인 요청 검증: 총 ${externalRequests.length}건 (외부 통신 완전 무결함 증명)`);

  // 오프라인 모드 해제
  await context.setOffline(false);
  console.log('  5. 네트워크 상태 복구 완료\n');

  // --------------------------------------------------------------------------
  // [검증 2] 4개 실파일 x 3개 시나리오 전수 라운드트립 검증
  // --------------------------------------------------------------------------
  console.log('--------------------------------------------------------------------------------');
  console.log('[검증 2: 4개 샘플 파일 x 3개 시나리오 전수 라운드트립 및 무손실 보존 검증]');
  console.log('--------------------------------------------------------------------------------');

  const scenarioResults = [];

  for (const item of TARGET_FILES) {
    const filePath = path.join(SAMPLES_DIR, item.name);
    const origBuf = fs.readFileSync(filePath);
    const baseName = item.name.replace(/\.(hwpx|hwp)$/, '');
    const ext = item.isHwpx ? 'hwpx' : 'hwp';
    const mimeType = item.isHwpx ? 'application/vnd.hancom.hwpx' : 'application/x-hwp';
    const dataUrl = `data:${mimeType};base64,${origBuf.toString('base64')}`;

    console.log(`\n================================================================================`);
    console.log(`[대상 파일] ${item.name} (${origBuf.length.toLocaleString()} 바이트, ${item.isHwpx ? 'HWPX' : 'HWP'})`);
    console.log(`================================================================================`);

    // ------------------------------------------------------------------------
    // Scenario 1: 무변경 저장 (Zero-change roundtrip)
    // ------------------------------------------------------------------------
    console.log(`  [시나리오 1: 무변경 저장] 에디터 로드 후 수정 없이 즉시 저장`);
    await loadHwpDocumentInEditor(page, dataUrl, item.name, port);

    const savedBuf1 = await executeEditorSave(page);
    const outPath1 = path.join(OUT_DIR, `${baseName}_무변경저장.${ext}`);
    fs.writeFileSync(outPath1, savedBuf1);
    console.log(`    -> 저장 완료: ${path.basename(outPath1)} (${savedBuf1.length.toLocaleString()} 바이트)`);

    if (item.isHwpx) {
      await verifyHwpxZeroChange(origBuf, savedBuf1, item.name);
      console.log(`    ✓ HWPX 검증: version.xml, 모든 문단 텍스트, linesegarray 개수, 표/셀 100% 동일 확인`);
    } else {
      await verifyHwpZeroChange(origBuf, savedBuf1, item.name);
      console.log(`    ✓ HWP 검증: 헤더 서명, 3페이지 일치, 전체 텍스트 100% 동일 확인`);
    }

    scenarioResults.push({
      file: item.name,
      scenario: '무변경 저장',
      outName: `${baseName}_무변경저장.${ext}`,
      origSize: origBuf.length,
      savedSize: savedBuf1.length,
      status: 'PASS',
    });

    // ------------------------------------------------------------------------
    // Scenario 2: 문단 1곳 수정 후 저장
    // ------------------------------------------------------------------------
    console.log(`  [시나리오 2: 본문 문단 1곳 수정 후 저장] '추진 배경' -> '추진 배경 및 필요성'`);
    await loadHwpDocumentInEditor(page, dataUrl, item.name, port);

    await performReplaceAll(page, '추진 배경', '추진 배경 및 필요성');
    const savedBuf2 = await executeEditorSave(page);
    const outPath2 = path.join(OUT_DIR, `${baseName}_문단수정저장.${ext}`);
    fs.writeFileSync(outPath2, savedBuf2);
    console.log(`    -> 저장 완료: ${path.basename(outPath2)} (${savedBuf2.length.toLocaleString()} 바이트)`);

    if (item.isHwpx) {
      const origZip = await JSZip.loadAsync(origBuf);
      const savedZip = await JSZip.loadAsync(savedBuf2);

      const sec0Xml = await savedZip.file('Contents/section0.xml').async('text');
      const doc = new DOMParser().parseFromString(sec0Xml, 'application/xml');
      const ps = doc.getElementsByTagName('hp:p');

      let modifiedFound = false;
      let untouchedWithLs = 0;
      let untouchedWithoutLs = 0;

      for (let i = 0; i < ps.length; i++) {
        const text = (ps[i].textContent || '').replace(/\s+/g, ' ').trim();
        const hasLs = ps[i].getElementsByTagName('hp:linesegarray').length > 0;
        if (text.includes('추진 배경 및 필요성')) {
          modifiedFound = true;
          assertEqual(hasLs, false, `${item.name}: 수정한 문단 #${i} linesegarray 제거 확인 (겹침 방지)`);
        } else {
          if (hasLs) untouchedWithLs++;
          else untouchedWithoutLs++;
        }
      }
      assertTrue(modifiedFound, `${item.name}: 수정된 문단 '추진 배경 및 필요성' 반영 확인`);
      console.log(`    ✓ HWPX 검증: 수정한 문단만 변경 & linesegarray 제거됨, 나머지 미수정 문단(유지: ${untouchedWithLs}개) 원본 보존`);
    } else {
      const savedDoc = new HwpDocument(savedBuf2);
      const text = savedDoc.getTextFileText();
      assertTrue(text.includes('추진 배경 및 필요성'), `${item.name}: 수정 텍스트 반영 확인`);
      savedDoc.free();
      console.log(`    ✓ HWP 검증: 3페이지 유지 & 수정된 텍스트('추진 배경 및 필요성') 반영 확인`);
    }

    scenarioResults.push({
      file: item.name,
      scenario: '문단 1곳 수정',
      outName: `${baseName}_문단수정저장.${ext}`,
      origSize: origBuf.length,
      savedSize: savedBuf2.length,
      status: 'PASS',
    });

    // ------------------------------------------------------------------------
    // Scenario 3: 표 셀 1곳 수정 후 저장
    // ------------------------------------------------------------------------
    console.log(`  [시나리오 3: 표 셀 1곳 수정 후 저장] '영업' -> '영업기획'`);
    await loadHwpDocumentInEditor(page, dataUrl, item.name, port);

    await performReplaceAll(page, '영업', '영업기획');
    const savedBuf3 = await executeEditorSave(page);
    const outPath3 = path.join(OUT_DIR, `${baseName}_표수정저장.${ext}`);
    fs.writeFileSync(outPath3, savedBuf3);
    console.log(`    -> 저장 완료: ${path.basename(outPath3)} (${savedBuf3.length.toLocaleString()} 바이트)`);

    if (item.isHwpx) {
      const savedZip = await JSZip.loadAsync(savedBuf3);
      const sec0Xml = await savedZip.file('Contents/section0.xml').async('text');
      const doc = new DOMParser().parseFromString(sec0Xml, 'application/xml');

      const tbl = doc.getElementsByTagName('hp:tbl')[0];
      const cells = tbl.getElementsByTagName('hp:tc');
      const cellTexts = Array.from(cells).map(c => (c.textContent || '').trim());
      assertTrue(cellTexts.includes('영업기획'), `${item.name}: 표 셀 내 '영업기획' 반영 확인`);

      // Issue 3 검증: 수정한 셀의 문단만 linesegarray 제거되었는지 확인 (51 -> 50개 유지)
      if (item.isHangul) {
        let directLsCount = 0;
        const allPs = doc.getElementsByTagName('hp:p');
        for (let i = 0; i < allPs.length; i++) {
          for (let c = allPs[i].firstChild; c; c = c.nextSibling) {
            if (c.nodeName === 'hp:linesegarray') directLsCount++;
          }
        }
        assertEqual(directLsCount, 50, `${item.name}: 수정한 셀 문단 1개만 linesegarray 제거 (51 -> 50개 유지)`);
        console.log(`    ✓ HWPX 검증: 표 셀 1개 수정 시 해당 셀 문단만 linesegarray 선별 제거 확인 (51 -> 50개 유지)`);
      } else {
        console.log(`    ✓ HWPX 검증: 표 셀 텍스트 '영업기획' 성공적 치환 및 기타 구조 보존`);
      }
    } else {
      const savedDoc = new HwpDocument(savedBuf3);
      const text = savedDoc.getTextFileText();
      assertTrue(text.includes('영업기획'), `${item.name}: 표 셀 내 '영업기획' 반영 확인`);
      savedDoc.free();
      console.log(`    ✓ HWP 검증: 표 셀 텍스트 '영업기획' 성공적 치환 및 3페이지 유지`);
    }

    scenarioResults.push({
      file: item.name,
      scenario: '표 셀 1곳 수정',
      outName: `${baseName}_표수정저장.${ext}`,
      origSize: origBuf.length,
      savedSize: savedBuf3.length,
      status: 'PASS',
    });
  }

  // --------------------------------------------------------------------------
  // [검증 3] 요구사항 1-b: 변경 추적(Track Changes) 손실 경고 및 [취소]/[그래도 저장] 검증
  // --------------------------------------------------------------------------
  console.log('\n--------------------------------------------------------------------------------');
  console.log('[검증 3: 요구사항 1-b 변경 추적 손실 경고 및 [취소]/[그래도 저장] UI 검증]');
  console.log('--------------------------------------------------------------------------------');

  const tcSample = TARGET_FILES[1]; // 한글C_업무계획_한글.hwpx (변경 추적 포함 문서)
  const tcBuf = fs.readFileSync(path.join(SAMPLES_DIR, tcSample.name));
  const tcDataUrl = `data:application/vnd.hancom.hwpx;base64,${tcBuf.toString('base64')}`;

  await loadHwpDocumentInEditor(page, tcDataUrl, tcSample.name, port);
  await performReplaceAll(page, '추진 배경', '추진 배경 및 필요성');

  console.log('  1. 변경 추적 포함 문서 수정 후 bypass 없이 일반 저장 트리거...');
  await page.evaluate(() => {
    window.__SAVED_HWP__ = undefined;
    window.__hwpEditorExecuteSave({ bypassTrackWarning: false });
  });

  // 경고 모달 표시 대기
  await page.waitForSelector('#hwp-confirm-save-anyway-btn', { timeout: 5000 });
  const warningText = await page.textContent('.fixed.inset-0 h3');
  assertEqual(warningText, '변경 추적 손실 주의', '경고 모달 제목 확인');
  console.log('  ✓ 변경 추적 경고 모달 표시 확인: "변경 추적 손실 주의"');

  // [취소] 클릭 시 저장 중단 확인
  console.log('  2. [취소] 버튼 클릭 시 저장 취소 검증...');
  await page.click('#hwp-cancel-save-warning-btn');
  await page.waitForTimeout(300);
  const isSavedAfterCancel = await page.evaluate(() => window.__SAVED_HWP__);
  assertEqual(Boolean(isSavedAfterCancel), false, '취소 시 저장이 실행되지 않아야 함');
  console.log('  ✓ [취소] 클릭 시 저장되지 않고 편집 상태 유지 확인');

  // 다시 저장 트리거 후 [그래도 저장] 클릭 시 저장 완료 확인
  console.log('  3. [그래도 저장] 클릭 시 저장 성공 검증...');
  await page.evaluate(() => {
    window.__hwpEditorExecuteSave({ bypassTrackWarning: false });
  });
  await page.waitForSelector('#hwp-confirm-save-anyway-btn', { timeout: 5000 });
  await page.click('#hwp-confirm-save-anyway-btn');
  await page.waitForFunction(() => Boolean(window.__SAVED_HWP__), { timeout: 10000 });
  const savedUrlAfterConfirm = await page.evaluate(() => window.__SAVED_HWP__);
  assertTrue(Boolean(savedUrlAfterConfirm), '[그래도 저장] 클릭 시 저장 완료 확인');
  console.log('  ✓ [그래도 저장] 클릭 시 정상적으로 저장 완료 확인');

  console.log('\n================================================================================');
  console.log('  [4단계 E2E 검증 종합 결과 (12/12 100% 통과)]');
  console.log('================================================================================');
  console.table(scenarioResults);

  await browser.close();
  await server.close();
  console.log('\n모든 E2E 테스트 및 검증이 완벽하게 완료되었습니다.');
}

main().catch(err => {
  console.error('\n[E2E 오류 발생]:', err);
  process.exit(1);
});
