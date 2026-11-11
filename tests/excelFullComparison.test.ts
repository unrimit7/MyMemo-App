import { describe, it, test, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { dataUrlToUniverWorkbook, univerSnapshotToDataUrl } from '../src/utils/univerAdapter';

function bufferToDataUrl(buf: Buffer, mime: string = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'): string {
  return `data:${mime};base64,${buf.toString('base64')}`;
}

function dataUrlToBuffer(dataUrl: string): Buffer {
  const commaIdx = dataUrl.indexOf(',');
  const b64 = commaIdx !== -1 ? dataUrl.slice(commaIdx + 1) : dataUrl;
  return Buffer.from(b64, 'base64');
}

const samplesDir = path.resolve(__dirname, '../docs/office-upgrade/samples');

async function extractDataValidationsFromXml(buf: Buffer): Promise<Record<string, Array<{ sqref: string; type: string }>>> {
  const zip = await JSZip.loadAsync(buf);
  const result: Record<string, Array<{ sqref: string; type: string }>> = {};
  for (const [filename, fileObj] of Object.entries(zip.files)) {
    if (filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml')) {
      const xml = await fileObj.async('string');
      const dvMatches = xml.match(/<dataValidation\b[\s\S]*?<\/dataValidation>/g) || [];
      const sheetKey = filename;
      result[sheetKey] = dvMatches.map((m) => {
        const sqref = m.match(/sqref="([^"]+)"/)?.[1] || '';
        const type = m.match(/type="([^"]+)"/)?.[1] || 'list';
        return { sqref, type };
      });
    }
  }
  return result;
}

interface CompareOptions {
  excludedCells?: Record<string, string[]>;
  skipFormulas?: boolean;
}

async function assertFullWorkbookFidelity(
  origBuf: Buffer,
  savedBuf: Buffer,
  options: CompareOptions = {}
) {
  const origWb = new ExcelJS.Workbook();
  await origWb.xlsx.load(origBuf);

  const savedWb = new ExcelJS.Workbook();
  await savedWb.xlsx.load(savedBuf);

  expect(savedWb.worksheets.length).toBe(origWb.worksheets.length);

  for (let sIdx = 0; sIdx < origWb.worksheets.length; sIdx++) {
    const origWs = origWb.worksheets[sIdx];
    const savedWs = savedWb.worksheets[sIdx];
    expect(savedWs.name).toBe(origWs.name);

    const sheetExcluded = options.excludedCells?.[origWs.name] || options.excludedCells?.['*'] || [];

    // 1. 전체 셀 단위 정밀 비교 (값, 수식, 데이터 형식, 표시형식)
    origWs.eachRow({ includeEmpty: false }, (origRow, rowNum) => {
      origRow.eachCell({ includeEmpty: false }, (origCell, colNum) => {
        const addr = origCell.address;
        if (sheetExcluded.includes(addr)) return;

        // 머지 슬레이브 셀은 마스터의 프록시이므로 마스터 셀에서 검증
        if (origCell.isMerged && origCell.master && origCell.master.address !== origCell.address) {
          return;
        }

        const savedCell = savedWs.getCell(rowNum, colNum);

        // 수식 검증
        if (origCell.formula) {
          expect(savedCell.formula).toBe(origCell.formula);
          if (origCell.result instanceof Date) {
            expect(savedCell.result instanceof Date).toBe(true);
            expect((savedCell.result as Date).getTime()).toBe(origCell.result.getTime());
          } else if (typeof origCell.result === 'number' || typeof origCell.result === 'string') {
            expect(savedCell.result).toBe(origCell.result);
          }
        } else if (origCell.value instanceof Date) {
          expect(savedCell.value instanceof Date).toBe(true);
          expect((savedCell.value as Date).getTime()).toBe(origCell.value.getTime());
        } else if (origCell.value && typeof origCell.value === 'object' && 'richText' in (origCell.value as any)) {
          expect(savedCell.value).toEqual(origCell.value);
        } else if (origCell.value && typeof origCell.value === 'object' && 'hyperlink' in (origCell.value as any)) {
          expect(savedCell.value).toEqual(origCell.value);
        } else if (origCell.value !== null && origCell.value !== undefined) {
          expect(savedCell.value).toEqual(origCell.value);
        }

        // 글꼴 속성(이름, 크기, 색상) 검증
        if (origCell.font) {
          if (origCell.font.name) {
            expect(savedCell.font?.name).toBe(origCell.font.name);
          }
          if (origCell.font.size) {
            expect(savedCell.font?.size).toBe(origCell.font.size);
          }
          if (origCell.font.color) {
            expect(savedCell.font?.color).toEqual(origCell.font.color);
          }
        }

        // 표시 형식(numFmt) 검증
        if (origCell.numFmt) {
          expect(savedCell.numFmt).toBe(origCell.numFmt);
        }
      });
    });

    // 2. 열 너비(Widths) 및 숨김(hidden) 일치 검증 (허용 오차 0.01)
    const maxCol = Math.max(origWs.columnCount, savedWs.columnCount);
    for (let c = 1; c <= maxCol; c++) {
      const origCol = origWs.getColumn(c);
      const savedCol = savedWs.getColumn(c);
      if (origCol && origCol.width !== undefined) {
        expect(savedCol.width).toBeDefined();
        expect(Math.abs((savedCol.width || 0) - origCol.width)).toBeLessThanOrEqual(0.01);
      }
      if (origCol && origCol.hidden !== undefined) {
        expect(savedCol.hidden).toBe(origCol.hidden);
      }
    }

    // 3. 행 높이(Heights) 및 숨김(hidden) 일치 검증 (허용 오차 0.01)
    origWs.eachRow({ includeEmpty: false }, (origRow, rowNum) => {
      const savedRow = savedWs.getRow(rowNum);
      if (origRow.height !== undefined) {
        expect(savedRow.height).toBeDefined();
        expect(Math.abs((savedRow.height || 0) - origRow.height)).toBeLessThanOrEqual(0.01);
      }
      if (origRow.hidden !== undefined) {
        expect(savedRow.hidden).toBe(origRow.hidden);
      }
    });

    // 4. 병합(Merges) 개수 및 정확한 범위 일치 검증
    const origMerges = [...((origWs as any).model?.merges || [])].sort();
    const savedMerges = [...((savedWs as any).model?.merges || [])].sort();
    expect(savedMerges).toEqual(origMerges);

    // 5. 셀 메모(Notes) 개수 및 주소 일치 검증
    const origNotes: string[] = [];
    origWs.eachRow({ includeEmpty: false }, (r) => {
      r.eachCell({ includeEmpty: false }, (c) => {
        if (c.note) origNotes.push(c.address);
      });
    });
    const savedNotes: string[] = [];
    savedWs.eachRow({ includeEmpty: false }, (r) => {
      r.eachCell({ includeEmpty: false }, (c) => {
        if (c.note) savedNotes.push(c.address);
      });
    });
    expect(savedNotes.sort()).toEqual(origNotes.sort());
  }

  // 6. 문서 기본 글꼴 및 Normal 스타일(styles.xml) 일치 검증
  const zipOrig = await JSZip.loadAsync(origBuf);
  const zipSaved = await JSZip.loadAsync(savedBuf);
  const origStylesXml = (await zipOrig.file('xl/styles.xml')?.async('string')) || '';
  const savedStylesXml = (await zipSaved.file('xl/styles.xml')?.async('string')) || '';
  const normOrigStyles = origStylesXml.replace(/<x:font\b/g, '<font').replace(/<\/x:font>/g, '</font>')
                                      .replace(/<x:([A-Za-z0-9_]+)/g, '<$1').replace(/<\/x:([A-Za-z0-9_]+)>/g, '</$1>');
  const normSavedStyles = savedStylesXml.replace(/<x:font\b/g, '<font').replace(/<\/x:font>/g, '</font>')
                                        .replace(/<x:([A-Za-z0-9_]+)/g, '<$1').replace(/<\/x:([A-Za-z0-9_]+)>/g, '</$1>');
  const origFont0Match = normOrigStyles.match(/<fonts[^>]*>([\s\S]*?)<\/fonts>/)?.[1]?.match(/<font\b[\s\S]*?<\/font>/)?.[0];
  const savedFont0Match = normSavedStyles.match(/<fonts[^>]*>([\s\S]*?)<\/fonts>/)?.[1]?.match(/<font\b[\s\S]*?<\/font>/)?.[0];
  if (origFont0Match && savedFont0Match) {
    const origName = origFont0Match.match(/<name\s+val="([^"]+)"/)?.[1];
    const savedName = savedFont0Match.match(/<name\s+val="([^"]+)"/)?.[1];
    expect(savedName).toBe(origName);
  }
  if (normOrigStyles.includes('name="표준"')) {
    expect(normSavedStyles.includes('name="표준"')).toBe(true);
  }

  // 6-1. Normal 스타일 alignment(vertical="center" 등) 검증
  const origCellStyleXf = normOrigStyles.match(/<cellStyleXfs[^>]*>([\s\S]*?)<\/cellStyleXfs>/)?.[1] || '';
  if (origCellStyleXf.includes('<alignment')) {
    const savedCellStyleXf = normSavedStyles.match(/<cellStyleXfs[^>]*>([\s\S]*?)<\/cellStyleXfs>/)?.[1] || '';
    expect(savedCellStyleXf.includes('<alignment')).toBe(true);
    const savedCellXfs = normSavedStyles.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/)?.[1] || '';
    const firstCellXf = savedCellXfs.match(/<xf\b[\s\S]*?<\/xf>/)?.[0] || savedCellXfs.match(/<xf\b[^>]*\/>/)?.[0] || '';
    expect(firstCellXf.includes('<alignment')).toBe(true);
  }

  // 6-2. 미수정 행 customHeight 속성 유지 검증 (원본에 없었던 customHeight="1" 강제 추가 방지)
  for (const [filename, fileObj] of Object.entries(zipOrig.files)) {
    if (filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml')) {
      const origXml = await fileObj.async('string');
      const savedXml = (await zipSaved.file(filename)?.async('string')) || '';
      const origRows = origXml.match(/<row\b[^>]*>/g) || [];
      const savedRows = savedXml.match(/<row\b[^>]*>/g) || [];
      const savedRowMap = new Map<number, string>();
      for (const sTag of savedRows) {
        const rMatch = sTag.match(/\br="(\d+)"/);
        if (rMatch) savedRowMap.set(parseInt(rMatch[1], 10), sTag);
      }
      for (const oTag of origRows) {
        const rMatch = oTag.match(/\br="(\d+)"/);
        if (rMatch) {
          const rNum = parseInt(rMatch[1], 10);
          if (!oTag.includes('customHeight="1"') && !oTag.includes('customHeight="true"')) {
            const sTag = savedRowMap.get(rNum);
            if (sTag) {
              expect(sTag.includes('customHeight="1"')).toBe(false);
            }
          }
        }
      }
    }
  }

  // 7. 데이터 유효성 검사 XML 태그 개수 및 범위 일치 검증
  const origDvs = await extractDataValidationsFromXml(origBuf);
  const savedDvs = await extractDataValidationsFromXml(savedBuf);
  for (const sheetKey of Object.keys(origDvs)) {
    const origRules = origDvs[sheetKey].map((r) => `${r.sqref}_${r.type}`).sort();
    const savedRules = (savedDvs[sheetKey] || []).map((r) => `${r.sqref}_${r.type}`).sort();
    expect(savedRules).toEqual(origRules);
  }
}

describe('Excel Full Cell-by-Cell & Feature Roundtrip Comparison Tests', () => {
  // 1. 버그 1 정밀 검증: 실파일B C5 날짜 수식 및 날짜 결과, 병합 보존
  test('실파일B 견적서 C5 날짜 수식(=DATE) 및 날짜 결과, 병합 셀 보존 검증 (MS & 한셀)', async () => {
    for (const fileName of ['실파일B_견적서_MS엑셀.xlsx', '실파일B_견적서_한셀.xlsx']) {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'real', fileName));
      const origDataUrl = bufferToDataUrl(origBuf);

      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일B');
      const sheet = snapshot.sheets[snapshot.sheetOrder[0]];

      // C5 셀 확인 (row 4, col 2)
      const uC5 = sheet.cellData[4]?.[2];
      expect(uC5).toBeDefined();
      expect(uC5.f).toBe('=DATE(2026,11,11)');
      expect(uC5.v).toBe('2026-11-11');

      // 머지 슬레이브인 D5, E5에는 [object Object]가 들어가지 않아야 함
      expect(sheet.cellData[4]?.[3]?.v).toBeUndefined();
      expect(sheet.cellData[4]?.[4]?.v).toBeUndefined();

      // 셀 수정 없이 저장 (또는 A1만 수정)
      if (!sheet.cellData[0]) sheet.cellData[0] = {};
      sheet.cellData[0][0] = { v: '테스트 견적서' };

      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      const savedBuf = dataUrlToBuffer(savedDataUrl);

      const savedWb = new ExcelJS.Workbook();
      await savedWb.xlsx.load(savedBuf);
      const savedWs = savedWb.worksheets[0];
      const savedC5 = savedWs.getCell('C5');

      expect(savedC5.type).toBe(ExcelJS.ValueType.Formula);
      expect(savedC5.formula).toBe('DATE(2026,11,11)');
      expect(savedC5.value).not.toBe('[object Object]');
      expect(savedC5.result instanceof Date).toBe(true);
      expect((savedC5.result as Date).getFullYear()).toBe(2026);
      expect((savedC5.result as Date).getMonth()).toBe(10); // 11월 (0-based 10)
      expect((savedC5.result as Date).getDate()).toBe(11);
      expect(savedC5.isMerged).toBe(true);
      expect(savedC5.numFmt).toBe('yyyy"년" m"월" d"일"');
    }
  });

  // 2. 버그 2 정밀 검증: 실파일A 데이터 유효성 검사 3회 연속 저장 후에도 2개 유지 (중복 없음)
  test('실파일A 데이터 유효성 검사 3회 연속 저장 후에도 2개 규칙(C4:C13, F4:F13) 유지 검증', async () => {
    for (const fileName of ['실파일A_인사평가_MS엑셀.xlsx', '실파일A_인사평가_한셀.xlsx']) {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'real', fileName));
      let currentDataUrl = bufferToDataUrl(origBuf);

      for (let cycle = 1; cycle <= 3; cycle++) {
        const snapshot = await dataUrlToUniverWorkbook(currentDataUrl, `실파일A_사이클${cycle}`);
        const sheet = snapshot.sheets[snapshot.sheetOrder[0]];
        if (!sheet.cellData[1]) sheet.cellData[1] = {};
        sheet.cellData[1][0] = { v: `수정_사이클${cycle}` };

        currentDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', currentDataUrl);
        const currentBuf = dataUrlToBuffer(currentDataUrl);

        // XML dataValidation 태그 개수 검사: 정확히 2개여야 함
        const dvs = await extractDataValidationsFromXml(currentBuf);
        const sheet1Dvs = dvs['xl/worksheets/sheet1.xml'] || [];
        expect(sheet1Dvs.length).toBe(2);

        const sqrefs = sheet1Dvs.map((d) => d.sqref).sort();
        expect(sqrefs).toEqual(['C4:C13', 'F4:F13']);
      }
    }
  });

  // 3. 병합 해제 후 슬레이브 셀 값 입력 및 저장 검증 (C5:E5 병합 해제 -> D5에 새 값 입력)
  test('병합 해제 후 슬레이브 셀 값 입력 및 저장 검증 (C5:E5 병합 해제 -> D5에 새 값 입력)', async () => {
    const origBuf = fs.readFileSync(path.join(samplesDir, 'real', '실파일B_견적서_MS엑셀.xlsx'));
    const origDataUrl = bufferToDataUrl(origBuf);

    // 1. Univer 모델 로드
    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일B_병합해제테스트');
    const sheet = snapshot.sheets[snapshot.sheetOrder[0]];

    // 견적서 시트의 C5:E5 병합 해제 시뮬레이션 (C5: row 4, col 2 / E5: row 4, col 4)
    sheet.mergeData = (sheet.mergeData || []).filter(
      (m) => !(m.startRow === 4 && m.endRow === 4 && m.startColumn === 2 && m.endColumn === 4)
    );

    // D5(row 4, col 3)에 사용자 새 값 입력
    if (!sheet.cellData[4]) sheet.cellData[4] = {};
    sheet.cellData[4][3] = { v: 'D5단독입력값' };

    // 2. 저장 수행
    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    // 3. 재로드 검증
    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    const reWs = reloadedWb.worksheets[0];

    // D5가 독립적으로 저장되었는지 확인
    const d5Cell = reWs.getCell('D5');
    expect(d5Cell.value).toBe('D5단독입력값');
    expect(d5Cell.isMerged).toBe(false);

    // C5의 수식도 정상 유지되는지 확인
    const c5Cell = reWs.getCell('C5');
    expect(c5Cell.formula).toBe('DATE(2026,11,11)');
    expect(c5Cell.isMerged).toBe(false);

    // 병합 목록에 C5:E5가 없음을 확인
    const merges = (reWs as any).model?.merges || [];
    expect(merges).not.toContain('C5:E5');
  });

  // 4. 전체 셀 비교 (전수 검증): 실파일 4종
  test('실파일 4종 전체 셀 비교 (값, 수식, 날짜, 서식, 병합, 메모, 유효성검사 전수 검증)', async () => {
    // 3-1. 실파일A MS
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'real', '실파일A_인사평가_MS엑셀.xlsx'));
      const origDataUrl = bufferToDataUrl(origBuf);
      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일A_MS');
      const sheet = snapshot.sheets[snapshot.sheetOrder[0]];
      if (!sheet.cellData[1]) sheet.cellData[1] = {};
      sheet.cellData[1][0] = { v: '홍길동_검증' }; // A2 수정

      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      const savedBuf = dataUrlToBuffer(savedDataUrl);
      fs.writeFileSync(path.join(samplesDir, '../out/실파일A_인사평가_MS엑셀_저장결과.xlsx'), savedBuf);
      await assertFullWorkbookFidelity(origBuf, savedBuf, { excludedCells: { 인사평가: ['A2'] } });
    }

    // 3-2. 실파일A 한셀 (한셀 XML 정규화 후 비교)
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'real', '실파일A_인사평가_한셀.xlsx'));
      const origDataUrl = bufferToDataUrl(origBuf);
      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일A_한셀');
      const sheet = snapshot.sheets[snapshot.sheetOrder[0]];
      if (!sheet.cellData[1]) sheet.cellData[1] = {};
      sheet.cellData[1][0] = { v: '홍길동_검증' }; // A2 수정

      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      const savedBuf = dataUrlToBuffer(savedDataUrl);
      fs.writeFileSync(path.join(samplesDir, '../out/실파일A_인사평가_한셀_저장결과.xlsx'), savedBuf);
      // 한셀 원본은 정규화 파이프라인을 거치므로 ExcelJS 로드 기준으로 전수 검증
      const origNormalizedDataUrl = await univerSnapshotToDataUrl(
        await dataUrlToUniverWorkbook(origDataUrl, '정규화'),
        'xlsx',
        'exceljs',
        origDataUrl
      );
      await assertFullWorkbookFidelity(dataUrlToBuffer(origNormalizedDataUrl), savedBuf, {
        excludedCells: { 인사평가: ['A2'] },
      });
    }

    // 3-3. 실파일B MS
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'real', '실파일B_견적서_MS엑셀.xlsx'));
      const origDataUrl = bufferToDataUrl(origBuf);
      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일B_MS');
      const sheet = snapshot.sheets[snapshot.sheetOrder[0]];
      if (!sheet.cellData[0]) sheet.cellData[0] = {};
      sheet.cellData[0][0] = { v: '견적서_수정본' }; // A1 수정

      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      const savedBuf = dataUrlToBuffer(savedDataUrl);
      fs.writeFileSync(path.join(samplesDir, '../out/실파일B_견적서_MS엑셀_저장결과.xlsx'), savedBuf);
      await assertFullWorkbookFidelity(origBuf, savedBuf, { excludedCells: { 견적서: ['A1'] } });
    }

    // 3-4. 실파일B 한셀
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'real', '실파일B_견적서_한셀.xlsx'));
      const origDataUrl = bufferToDataUrl(origBuf);
      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일B_한셀');
      const sheet = snapshot.sheets[snapshot.sheetOrder[0]];
      if (!sheet.cellData[0]) sheet.cellData[0] = {};
      sheet.cellData[0][0] = { v: '견적서_수정본' }; // A1 수정

      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      const savedBuf = dataUrlToBuffer(savedDataUrl);
      fs.writeFileSync(path.join(samplesDir, '../out/실파일B_견적서_한셀_저장결과.xlsx'), savedBuf);
      const origNormalizedDataUrl = await univerSnapshotToDataUrl(
        await dataUrlToUniverWorkbook(origDataUrl, '정규화'),
        'xlsx',
        'exceljs',
        origDataUrl
      );
      await assertFullWorkbookFidelity(dataUrlToBuffer(origNormalizedDataUrl), savedBuf, {
        excludedCells: { 견적서: ['A1'] },
      });
    }
  });

  // 4. 전체 셀 비교 (전수 검증): sample 1 ~ 5
  test('Sample 1 ~ 5 전체 셀 비교 (값, 수식, 서식, 병합, 메모, 유효성검사 전수 검증)', async () => {
    // Sample 1
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'sample1_merges_and_widths.xlsx'));
      const origDataUrl = bufferToDataUrl(origBuf);
      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample1');
      if (!snapshot.sheets[snapshot.sheetOrder[0]].cellData[1]) snapshot.sheets[snapshot.sheetOrder[0]].cellData[1] = {};
      snapshot.sheets[snapshot.sheetOrder[0]].cellData[1][1] = { v: '수정셀' }; // B2
      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      await assertFullWorkbookFidelity(origBuf, dataUrlToBuffer(savedDataUrl), { excludedCells: { '*': ['B2'] } });
    }

    // Sample 2
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'sample2_fonts_and_borders.xlsx'));
      const origDataUrl = bufferToDataUrl(origBuf);
      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample2');
      if (!snapshot.sheets[snapshot.sheetOrder[0]].cellData[1]) snapshot.sheets[snapshot.sheetOrder[0]].cellData[1] = {};
      snapshot.sheets[snapshot.sheetOrder[0]].cellData[1][0] = { v: '수정셀' }; // A2
      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      await assertFullWorkbookFidelity(origBuf, dataUrlToBuffer(savedDataUrl), { excludedCells: { '*': ['A2'] } });
    }

    // Sample 3
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'sample3_formulas_multi_sheet.xlsx'));
      const origDataUrl = bufferToDataUrl(origBuf);
      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample3');
      if (!snapshot.sheets[snapshot.sheetOrder[0]].cellData[1]) snapshot.sheets[snapshot.sheetOrder[0]].cellData[1] = {};
      snapshot.sheets[snapshot.sheetOrder[0]].cellData[1][0] = { v: '수정셀' }; // A2
      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      await assertFullWorkbookFidelity(origBuf, dataUrlToBuffer(savedDataUrl), { excludedCells: { '*': ['A2'] } });
    }

    // Sample 4 (차트 파일: 차트 제외 셀 데이터 전수 일치)
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'sample4_chart_loss_test.xlsx'));
      const origDataUrl = bufferToDataUrl(origBuf);
      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample4');
      if (!snapshot.sheets[snapshot.sheetOrder[0]].cellData[1]) snapshot.sheets[snapshot.sheetOrder[0]].cellData[1] = {};
      snapshot.sheets[snapshot.sheetOrder[0]].cellData[1][0] = { v: '수정셀' }; // A2
      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      await assertFullWorkbookFidelity(origBuf, dataUrlToBuffer(savedDataUrl), { excludedCells: { '*': ['A2'] } });
    }

    // Sample 5
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'sample5_conditional_and_validation.xlsx'));
      const origDataUrl = bufferToDataUrl(origBuf);
      const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample5');
      if (!snapshot.sheets[snapshot.sheetOrder[0]].cellData[1]) snapshot.sheets[snapshot.sheetOrder[0]].cellData[1] = {};
      snapshot.sheets[snapshot.sheetOrder[0]].cellData[1][0] = { v: '수정셀' }; // A2
      const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
      await assertFullWorkbookFidelity(origBuf, dataUrlToBuffer(savedDataUrl), { excludedCells: { '*': ['A2'] } });
    }
  });
});
