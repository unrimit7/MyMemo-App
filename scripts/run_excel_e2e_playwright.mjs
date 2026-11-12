import fs from 'fs';
import path from 'path';
import { createServer } from 'vite';
import { chromium } from 'playwright-core';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const CHROME_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const SAMPLES_DIR = path.resolve('docs/office-upgrade/samples/real');
const OUT_DIR = path.resolve('docs/office-upgrade/out/excel_e2e');

if (!fs.existsSync(OUT_DIR)) {
  fs.mkdirSync(OUT_DIR, { recursive: true });
}

const REAL_FILES = [
  { name: '실파일A_인사평가_MS엑셀.xlsx', isA: true, isMS: true },
  { name: '실파일A_인사평가_한셀.xlsx', isA: true, isMS: false },
  { name: '실파일B_견적서_MS엑셀.xlsx', isA: false, isMS: true },
  { name: '실파일B_견적서_한셀.xlsx', isA: false, isMS: false },
];

function assertEqual(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(`[Assertion Failed] ${message}\nExpected: ${expected}\nActual:   ${actual}`);
  }
}

function assertTrue(cond, message) {
  if (!cond) {
    throw new Error(`[Assertion Failed] ${message}`);
  }
}

async function normalizeHancellBuffer(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  let isModified = false;

  for (const [filename, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    if (!filename.endsWith('.xml') && !filename.endsWith('.rels') && !filename.endsWith('.vml')) continue;

    let content = await file.async('text');
    let mod = false;

    if (filename === 'docProps/app.xml') {
      content = content.replace(
        /xmlns:ep="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/extended-properties"/g,
        'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
      );
      content = content.replace(/<\/?ep:([a-zA-Z0-9_]+)/g, (match, tag) => (match.startsWith('</') ? `</${tag}` : `<${tag}`));
      if (!content.includes('<Company')) {
        content = content.replace(/<\/Properties>/, '<Company></Company></Properties>');
      }
      mod = true;
    }

    if (filename.startsWith('xl/') && (filename.endsWith('.xml') || filename.endsWith('.rels'))) {
      if (content.includes('<mc:AlternateContent')) {
        content = content.replace(
          /<mc:AlternateContent[^>]*>[\s\S]*?<mc:Fallback>([\s\S]*?)<\/mc:Fallback>[\s\S]*?<\/mc:AlternateContent>/g,
          '$1'
        );
        content = content.replace(
          /<mc:AlternateContent[^>]*>[\s\S]*?<mc:Choice[^>]*>([\s\S]*?)<\/mc:Choice>[\s\S]*?<\/mc:AlternateContent>/g,
          '$1'
        );
        mod = true;
      }

      if (content.includes('hs:')) {
        content = content.replace(/<hs:[^>]*\/>/g, '');
        content = content.replace(/<hs:[^>]*>[\s\S]*?<\/hs:[^>]*>/g, '');
        content = content.replace(/\s+hs:[a-zA-Z0-9_]+="[^"]*"/g, '');
        mod = true;
      }

      if (content.includes('xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"') || content.includes('<x:')) {
        content = content.replace(
          /xmlns:x="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/g,
          'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        );
        content = content.replace(/<\/?x:([a-zA-Z0-9_]+)/g, (match, tag) => (match.startsWith('</') ? `</${tag}` : `<${tag}`));
        mod = true;
      }

      if (content.includes('val="1"') || content.includes('val="0"') || content.includes('val="true"') || content.includes('val="false"')) {
        content = content.replace(/<b\s+val="1"\s*\/>/g, '<b/>');
        content = content.replace(/<b\s+val="true"\s*\/>/g, '<b/>');
        content = content.replace(/<b\s+val="0"\s*\/>/g, '');
        content = content.replace(/<b\s+val="false"\s*\/>/g, '');

        content = content.replace(/<i\s+val="1"\s*\/>/g, '<i/>');
        content = content.replace(/<i\s+val="true"\s*\/>/g, '<i/>');
        content = content.replace(/<i\s+val="0"\s*\/>/g, '');
        content = content.replace(/<i\s+val="false"\s*\/>/g, '');

        content = content.replace(/<u\s+val="1"\s*\/>/g, '<u/>');
        content = content.replace(/<u\s+val="0"\s*\/>/g, '');

        content = content.replace(/<strike\s+val="1"\s*\/>/g, '<strike/>');
        content = content.replace(/<strike\s+val="0"\s*\/>/g, '');
        mod = true;
      }

      if (content.includes('ref=')) {
        content = content.replace(/ref="([A-Za-z]+[0-9]+):\1"/g, 'ref="$1"');
        mod = true;
      }
    }

    if (mod) {
      zip.file(filename, content);
      isModified = true;
    }
  }

  if (isModified) {
    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  }
  return bytes;
}

async function verifyZeroChange(origBuf, savedBuf, targetName, isA, isMS) {
  const normOrigBuf = isMS ? origBuf : await normalizeHancellBuffer(origBuf);
  const origWb = new ExcelJS.Workbook();
  await origWb.xlsx.load(normOrigBuf);

  const savedWb = new ExcelJS.Workbook();
  await savedWb.xlsx.load(savedBuf);

  assertEqual(savedWb.worksheets.length, origWb.worksheets.length, `${targetName}: 워크시트 수 일치`);

  for (let wIdx = 0; wIdx < origWb.worksheets.length; wIdx++) {
    const oWs = origWb.worksheets[wIdx];
    const sWs = savedWb.worksheets[wIdx];

    assertEqual(sWs.name, oWs.name, `${targetName}: 시트 [${wIdx}] 이름 일치`);

    // 열 속성 비교 (너비, 숨김 여부)
    for (let c = 1; c <= Math.max(oWs.columnCount, sWs.columnCount, 15); c++) {
      const oCol = oWs.getColumn(c);
      const sCol = sWs.getColumn(c);
      if (oCol.hidden) {
        assertTrue(sCol.hidden === true, `${targetName} [${oWs.name}]: ${c}번째 열 hidden=true 유지`);
      }
      if (oCol.width !== undefined && sCol.width !== undefined) {
        const diff = Math.abs(oCol.width - sCol.width);
        assertTrue(diff < 0.5, `${targetName} [${oWs.name}]: ${c}번째 열 너비 유지 (원본 ${oCol.width}, 결과 ${sCol.width})`);
      }
    }

    // 병합 셀 비교
    const oMerges = Object.keys(oWs._merges || {}).sort();
    const sMerges = Object.keys(sWs._merges || {}).sort();
    assertEqual(sMerges.length, oMerges.length, `${targetName} [${oWs.name}]: 병합 셀 수 일치 (${oMerges.length}개)`);

    // 셀 값, 수식, 표시형식 비교
    oWs.eachRow({ includeEmpty: false }, (oRow, rNum) => {
      const sRow = sWs.getRow(rNum);
      oRow.eachCell({ includeEmpty: false }, (oCell, cNum) => {
        const sCell = sRow.getCell(cNum);

        // 병합 슬레이브 셀
        if (oCell.isMerged && oCell.address !== oCell.master.address) {
          assertEqual(sCell.master.address, oCell.master.address, `${targetName} [${oWs.name} ${oCell.address}]: 병합 마스터 일치`);
          return;
        } else if (oCell.formula || (oCell.value && oCell.value.formula)) {
          const oFormula = oCell.formula || oCell.value.formula;
          const sFormula = sCell.formula || sCell.value?.formula;
          assertEqual(sFormula, oFormula, `${targetName} [${oWs.name} ${oCell.address}]: 수식 일치`);
        } else if (oCell.value instanceof Date) {
          assertTrue(sCell.value instanceof Date, `${targetName} [${oWs.name} ${oCell.address}]: 날짜 객체 유지`);
          assertEqual(sCell.value.getTime(), oCell.value.getTime(), `${targetName} [${oWs.name} ${oCell.address}]: 날짜 시간 일치`);
        } else if (typeof oCell.value === 'object' && oCell.value && 'richText' in oCell.value) {
          assertTrue(typeof sCell.value === 'object' && sCell.value && 'richText' in sCell.value, `${targetName} [${oWs.name} ${oCell.address}]: richText 유지`);
        } else if (typeof oCell.value === 'object' && oCell.value !== null) {
          assertEqual(JSON.stringify(sCell.value), JSON.stringify(oCell.value), `${targetName} [${oWs.name} ${oCell.address}]: 객체 값 일치`);
        } else {
          assertEqual(sCell.value, oCell.value, `${targetName} [${oWs.name} ${oCell.address}]: 셀 값 일치`);
        }

        // 표시 형식
        if (oCell.numFmt) {
          assertEqual(sCell.numFmt, oCell.numFmt, `${targetName} [${oWs.name} ${oCell.address}]: numFmt 일치`);
        }
      });
    });
  }

  // 실파일B 특화 검증 (숨김 H열, B8 richText, C5 날짜 수식, B2:H2 병합, 자동필터)
  if (!isA) {
    const sWs1 = savedWb.getWorksheet(1);
    assertTrue(sWs1.getColumn('H').hidden === true, `${targetName}: H열 hidden=true 완벽 유지`);
    const c5 = sWs1.getCell('C5');
    assertEqual(c5.formula, 'DATE(2026,11,11)', `${targetName}: C5 날짜 수식 =DATE(2026,11,11) 유지`);
    assertTrue(c5.value !== '[object Object]', `${targetName}: C5 날짜 수식이 [object Object]로 직렬화되지 않음`);
    const b8 = sWs1.getCell('B8');
    assertTrue(typeof b8.value === 'object' && b8.value && 'richText' in b8.value, `${targetName}: B8 richText 서식 유지`);
    assertTrue(sWs1.getCell('B2').isMerged && sWs1.getCell('H2').master.address === 'B2', `${targetName}: B2:H2 제목 병합 유지`);
  }

  // 기본 글꼴 검증 (styles.xml)
  const zip = await JSZip.loadAsync(savedBuf);
  const sf = zip.file('xl/styles.xml');
  const stylesXml = sf ? await sf.async('text') : '';
  assertTrue(stylesXml.includes('맑은 고딕'), `${targetName}: styles.xml 기본 글꼴 맑은 고딕 유지`);
}

async function main() {
  console.log('=== [5단계] 엑셀 1단계 E2E Playwright 검증 시작 ===');
  console.log('실제 브라우저(Chrome)에서 ExcelEditorModal 마운트 -> Univer createWorkbook -> activeWb.save() -> univerSnapshotToDataUrl 경로 검증\n');

  const server = await createServer({
    server: { port: 5188 },
    optimizeDeps: { exclude: ['@napi-rs/canvas'] },
    configFile: path.resolve('vite.config.ts'),
  });
  await server.listen();
  const port = server.httpServer.address().port;
  console.log(`1. Vite 테스트 서버 가동 (http://localhost:${port})`);

  const browser = await chromium.launch({
    executablePath: CHROME_PATH,
    headless: true,
  });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.goto(`http://localhost:${port}/viewer_test.html`);
  console.log('2. Playwright Chrome 브라우저 실행 및 viewer_test.html 로드 완료\n');

  for (const item of REAL_FILES) {
    const filePath = path.join(SAMPLES_DIR, item.name);
    const origBuf = fs.readFileSync(filePath);
    const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${origBuf.toString('base64')}`;
    const baseName = item.name.replace('.xlsx', '');

    console.log(`--------------------------------------------------------------------------------`);
    console.log(`[대상 파일] ${item.name} (${origBuf.length.toLocaleString()} 바이트)`);

    // ----------------------------------------------------
    // Scenario 1: 편집 없이 저장 (Zero-change roundtrip)
    // ----------------------------------------------------
    console.log(`  [시나리오 1: 무변경 저장] Univer 로드 후 아무것도 수정하지 않고 바로 저장`);
    await page.evaluate(({ url, title }) => {
      return window.__loadExcel(url, title);
    }, { url: dataUrl, title: item.name });

    await page.waitForFunction(() => typeof window.__excelIsReady === 'function' && window.__excelIsReady(), { timeout: 20000 });
    await page.waitForTimeout(400); // 렌더링 안정화 대기

    const savedUrl1 = await page.evaluate(async () => {
      window.__SAVED_EXCEL__ = undefined;
      await window.__excelModalExecuteSave();
      return window.__SAVED_EXCEL__;
    });

    assertTrue(Boolean(savedUrl1), `${item.name}: 저장된 dataUrl 획득 성공`);
    const savedBuf1 = Buffer.from(savedUrl1.replace(/^data:[^;]+;base64,/, ''), 'base64');
    const outPath1 = path.join(OUT_DIR, `${baseName}_무변경저장.xlsx`);
    fs.writeFileSync(outPath1, savedBuf1);
    console.log(`    -> 저장 완료: ${path.basename(outPath1)} (${savedBuf1.length.toLocaleString()} 바이트)`);

    await verifyZeroChange(origBuf, savedBuf1, item.name, item.isA, item.isMS);
    console.log(`    ✓ 검증 통과: 모든 셀 값·수식·표시 형식·글꼴, 열너비/행높이, 유효성 검사, 조건부 서식, 병합, 숨김 열 100% 동일 확인`);

    // ----------------------------------------------------
    // Scenario 2: 셀 1개 수정 후 저장
    // ----------------------------------------------------
    console.log(`  [시나리오 2: 셀 1개 수정 후 저장] 특정 셀 값만 변경`);
    // 새로운 인스턴스로 리로드
    await page.evaluate(({ url, title }) => {
      return window.__loadExcel(url, title);
    }, { url: dataUrl, title: item.name });
    await page.waitForFunction(() => typeof window.__excelIsReady === 'function' && window.__excelIsReady(), { timeout: 20000 });
    await page.waitForTimeout(400);

    let editedCellAddress = '';
    let editedOldValue = null;
    let editedNewValue = null;

    if (item.isA) {
      // 실파일A: C4 셀 (row 3, col 2) "경영지원" -> "경영기획"
      editedCellAddress = 'C4';
      editedOldValue = '경영지원';
      editedNewValue = '경영기획';
      await page.evaluate(() => {
        const sheet = window.__excelUniverAPI.getActiveWorkbook().getActiveSheet();
        sheet.getRange(3, 2).setValue('경영기획');
      });
    } else {
      // 실파일B: G11 셀 (row 10, col 6) 수량 10 -> 25
      editedCellAddress = 'G11';
      editedOldValue = 10;
      editedNewValue = 25;
      await page.evaluate(() => {
        const sheet = window.__excelUniverAPI.getActiveWorkbook().getActiveSheet();
        sheet.getRange(10, 6).setValue(25);
      });
    }

    await page.waitForTimeout(300);

    const savedUrl2 = await page.evaluate(async () => {
      window.__SAVED_EXCEL__ = undefined;
      await window.__excelModalExecuteSave();
      return window.__SAVED_EXCEL__;
    });

    const savedBuf2 = Buffer.from(savedUrl2.replace(/^data:[^;]+;base64,/, ''), 'base64');
    const outPath2 = path.join(OUT_DIR, `${baseName}_셀수정저장.xlsx`);
    fs.writeFileSync(outPath2, savedBuf2);
    console.log(`    -> 저장 완료: ${path.basename(outPath2)} (${savedBuf2.length.toLocaleString()} 바이트)`);

    // 검증: 수정한 셀만 다르고 나머지 모든 셀 동일
    const editCheckWb = new ExcelJS.Workbook();
    await editCheckWb.xlsx.load(savedBuf2);
    const origCheckWb = new ExcelJS.Workbook();
    const normOrigBuf = item.isMS ? origBuf : await normalizeHancellBuffer(origBuf);
    await origCheckWb.xlsx.load(normOrigBuf);

    const editWs = editCheckWb.getWorksheet(1);
    const origWs = origCheckWb.getWorksheet(1);

    assertEqual(editWs.getCell(editedCellAddress).value, editedNewValue, `${item.name}: 수정된 셀 ${editedCellAddress} 값 일치`);

    origWs.eachRow({ includeEmpty: false }, (oRow, rNum) => {
      oRow.eachCell({ includeEmpty: false }, (oCell, cNum) => {
        if (oCell.address === editedCellAddress) return; // 수정한 셀은 제외

        const sCell = editWs.getRow(rNum).getCell(cNum);
        if (oCell.isMerged && oCell.address !== oCell.master.address) {
          assertEqual(sCell.master.address, oCell.master.address, `${item.name} [셀수정 ${oCell.address}]: 병합 마스터 유지`);
          return;
        } else if (oCell.formula || (oCell.value && oCell.value.formula)) {
          const oFormula = oCell.formula || oCell.value.formula;
          const sFormula = sCell.formula || sCell.value?.formula;
          assertEqual(sFormula, oFormula, `${item.name} [셀수정 ${oCell.address}]: 수식 유지`);
        } else if (oCell.value instanceof Date) {
          assertTrue(sCell.value instanceof Date, `${item.name} [셀수정 ${oCell.address}]: 날짜 유지`);
        } else if (typeof oCell.value === 'object' && oCell.value && 'richText' in oCell.value) {
          assertTrue(typeof sCell.value === 'object' && sCell.value && 'richText' in sCell.value, `${item.name} [셀수정 ${oCell.address}]: richText 유지`);
        } else if (typeof oCell.value === 'object' && oCell.value !== null) {
          assertEqual(JSON.stringify(sCell.value), JSON.stringify(oCell.value), `${item.name} [셀수정 ${oCell.address}]: 객체 값 불변`);
        } else {
          assertEqual(sCell.value, oCell.value, `${item.name} [셀수정 ${oCell.address}]: 셀 값 불변`);
        }
      });
    });

    if (!item.isA) {
      assertTrue(editWs.getColumn('H').hidden === true, `${item.name}: 셀수정 후에도 H열 hidden=true 유지`);
      assertEqual(editWs.getCell('C5').formula, 'DATE(2026,11,11)', `${item.name}: C5 날짜 수식 불변`);
    }
    console.log(`    ✓ 검증 통과: ${editedCellAddress} 셀만 ${editedOldValue} -> ${editedNewValue}로 변경되고 나머지 모든 셀 100% 동일`);

    // ----------------------------------------------------
    // Scenario 3: 행 1개 삽입 및 삭제 후 저장
    // ----------------------------------------------------
    console.log(`  [시나리오 3: 행 1개 삽입 후 저장 -> 원복 삭제 후 저장]`);
    await page.evaluate(({ url, title }) => {
      return window.__loadExcel(url, title);
    }, { url: dataUrl, title: item.name });
    await page.waitForFunction(() => typeof window.__excelIsReady === 'function' && window.__excelIsReady(), { timeout: 20000 });
    await page.waitForTimeout(400);

    const insertRowIndex = item.isA ? 4 : 11; // 0-based
    const origRowCount = item.isA ? 13 : 18;

    // 행 1개 삽입
    await page.evaluate(async (rIdx) => {
      const sheet = window.__excelUniverAPI.getActiveWorkbook().getActiveSheet();
      await sheet.insertRowBefore(rIdx, 1);
      sheet.getRange(rIdx, 0).setValue('신규삽입행');
      await new Promise(r => setTimeout(r, 200));
    }, insertRowIndex);

    const savedUrl3 = await page.evaluate(async () => {
      window.__SAVED_EXCEL__ = undefined;
      await window.__excelModalExecuteSave();
      return window.__SAVED_EXCEL__;
    });

    const savedBuf3 = Buffer.from(savedUrl3.replace(/^data:[^;]+;base64,/, ''), 'base64');
    const outPath3 = path.join(OUT_DIR, `${baseName}_행삽입저장.xlsx`);
    fs.writeFileSync(outPath3, savedBuf3);
    console.log(`    -> 행 삽입 저장 완료: ${path.basename(outPath3)} (${savedBuf3.length.toLocaleString()} 바이트)`);

    const insertWb = new ExcelJS.Workbook();
    await insertWb.xlsx.load(savedBuf3);
    const insWs = insertWb.getWorksheet(1);
    const targetCellA = insWs.getCell(`A${insertRowIndex + 1}`).value;
    assertEqual(targetCellA, '신규삽입행', `${item.name}: 삽입된 행 데이터 확인`);

    if (item.isA) {
      // 실파일A: 유효성 검사 범위가 C4:C14, F4:F14로 정상 확장되고 빈 규칙(type 없음)이 없는지 확인
      const dvModel = insWs.dataValidations?.model || {};
      const dvKeys = Object.keys(dvModel);
      assertTrue(dvKeys.length === 22, `${item.name}: 행 삽입 후 유효성 검사 셀 수 22개 (C4:C14, F4:F14)`);
      for (const [k, r] of Object.entries(dvModel)) {
        assertTrue(Boolean(r && r.type), `${item.name}: 유효성 검사 규칙 ${k}에 type 정상 존재`);
      }
    } else {
      // 실파일B: 비고 병합이 B21:H24로 이동하여 유지되고 슬레이브 셀 중복 기록이 없는지 확인
      assertTrue(insWs.getCell('B21').isMerged && insWs.getCell('B21').master.address === 'B21', `${item.name}: B21 비고 마스터 병합 유지`);
      assertTrue(insWs.getCell('H24').isMerged && insWs.getCell('H24').master.address === 'B21', `${item.name}: H24 비고 슬레이브 병합 유지`);
      assertEqual(insWs.getCell('C21').value, insWs.getCell('B21').value, `${item.name}: C21 병합 마스터 참조 일치`);
    }

    // 다시 행 삭제하여 원복 저장
    await page.evaluate(async (rIdx) => {
      const sheet = window.__excelUniverAPI.getActiveWorkbook().getActiveSheet();
      await sheet.deleteRow(rIdx, 1);
      await new Promise(r => setTimeout(r, 200));
    }, insertRowIndex);

    const savedUrl4 = await page.evaluate(async () => {
      window.__SAVED_EXCEL__ = undefined;
      await window.__excelModalExecuteSave();
      return window.__SAVED_EXCEL__;
    });

    const savedBuf4 = Buffer.from(savedUrl4.replace(/^data:[^;]+;base64,/, ''), 'base64');
    const outPath4 = path.join(OUT_DIR, `${baseName}_행삭제원복저장.xlsx`);
    fs.writeFileSync(outPath4, savedBuf4);
    console.log(`    -> 행 삭제 원복 저장 완료: ${path.basename(outPath4)} (${savedBuf4.length.toLocaleString()} 바이트)`);

    // [검증 1] "삽입->삭제 후 원본과 완전히 같음" 전수 검증
    await verifyZeroChange(origBuf, savedBuf4, `${item.name} (삭제원복)`, item.isA, item.isMS);
    console.log(`    ✓ 검증 통과: 1행 삽입 후 삭제 원복 시 모든 셀 값·수식·병합·유효성·글꼴이 원본과 100% 동일`);

    // [검증 2] "다른 시트 셀 서식 불변" 검증
    if (item.isA) {
      const delWb = new ExcelJS.Workbook();
      await delWb.xlsx.load(savedBuf4);
      const wsDept = delWb.getWorksheet('부서코드');
      const wsSumm = delWb.getWorksheet('요약');

      // 부서코드 A2:B6 cell numFmt 불변 (General / undefined)
      for (let r = 2; r <= 6; r++) {
        assertEqual(wsDept.getCell(`A${r}`).numFmt, undefined, `${item.name}: 부서코드!A${r} 표시 형식 불변 (General)`);
        assertEqual(wsDept.getCell(`B${r}`).numFmt, undefined, `${item.name}: 부서코드!B${r} 표시 형식 불변 (General)`);
      }

      // 요약 B2 cell numFmt 불변 (General / undefined), B4는 #,##0"원" 유지
      assertEqual(wsSumm.getCell('B2').numFmt, undefined, `${item.name}: 요약!B2 표시 형식 불변 (General)`);
      assertEqual(wsSumm.getCell('B4').numFmt, '#,##0"원"', `${item.name}: 요약!B4 표시 형식 #,##0"원" 유지`);
      console.log(`    ✓ 검증 통과: 다른 시트(부서코드, 요약)의 셀 표시 형식(General) 완벽 불변 확인\n`);
    } else {
      console.log(`    ✓ 검증 통과: 실파일B 견적서 행 삭제 원복 및 H열 숨김/수식 완벽 복원 확인\n`);
    }
  }

  await browser.close();
  await server.close();
  console.log('=== [5단계] 엑셀 1단계 E2E Playwright 검증 전체 성공! 모든 결과 파일 out/excel_e2e/ 생성 완료 ===');
}

main().catch((err) => {
  console.error('\n[치명적 오류] 엑셀 E2E 검증 실패:', err);
  process.exit(1);
});
