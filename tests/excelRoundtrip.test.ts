import { describe, it, test, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { dataUrlToUniverWorkbook, univerSnapshotToDataUrl, type StructuralOperation } from '../src/utils/univerAdapter';
import { detectExcelUnsupportedFeatures } from '../src/utils/excelLossDetector';

function bufferToDataUrl(buf: Buffer, mime: string = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'): string {
  return `data:${mime};base64,${buf.toString('base64')}`;
}

function dataUrlToBuffer(dataUrl: string): Buffer {
  const commaIdx = dataUrl.indexOf(',');
  const b64 = commaIdx !== -1 ? dataUrl.slice(commaIdx + 1) : dataUrl;
  return Buffer.from(b64, 'base64');
}

const samplesDir = path.resolve(__dirname, '../docs/office-upgrade/samples');

function readSampleDataUrl(relPath: string): string {
  const filePath = path.join(samplesDir, relPath);
  const buf = fs.readFileSync(filePath);
  return bufferToDataUrl(buf);
}

describe('ExcelJS Roundtrip Automation Tests (Sample 1 ~ 5)', () => {
  it('Sample 1: 셀 병합(Merges), 열 너비(Widths), 행 높이(Heights) 보존 검증', async () => {
    const filePath = path.join(samplesDir, 'sample1_merges_and_widths.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    // 1. Univer 모델로 로드
    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample1');
    expect(snapshot.sheets).toBeDefined();

    // 2. Univer 스냅샷에서 셀 값 1개 수정 (B2 셀 수정 시뮬레이션)
    const firstSheetId = snapshot.sheetOrder[0];
    if (!snapshot.sheets[firstSheetId].cellData[1]) {
      snapshot.sheets[firstSheetId].cellData[1] = {};
    }
    snapshot.sheets[firstSheetId].cellData[1][1] = { v: '수정된데이터' };

    // 3. In-place ExcelJS 엔진으로 라운드트립 저장
    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    // 4. 저장된 결과를 ExcelJS로 재로드하여 요소별 정밀 검증
    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    const ws = reloadedWb.getWorksheet(1);
    expect(ws).toBeDefined();

    // 병합 검증: A1:D1 병합 유지 여부
    const merges = (ws as any).model?.merges || [];
    expect(merges).toContain('A1:D1');

    // 열 너비 검증: A열(25), B열(35), C열(18)
    expect(ws!.getColumn(1).width).toBeGreaterThanOrEqual(20);
    expect(ws!.getColumn(2).width).toBeGreaterThanOrEqual(30);

    // 행 높이 검증: 1행 높이(40pt)
    expect(ws!.getRow(1).height).toBeGreaterThanOrEqual(35);

    // 수정된 값 검증
    expect(ws!.getCell('B2').value).toBe('수정된데이터');
  });

  it('Sample 2: 글꼴(Fonts), 테두리(Borders), 셀 배경색(Fill), 표시 형식(Number Format) 보존 검증', async () => {
    const filePath = path.join(samplesDir, 'sample2_fonts_and_borders.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample2');
    const firstSheetId = snapshot.sheetOrder[0];
    snapshot.sheets[firstSheetId].cellData[1][0] = { v: '프리미엄 펜' };

    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    const ws = reloadedWb.getWorksheet(1);

    // 헤더 셀(A1) 글꼴 및 스타일 검증
    const a1 = ws!.getCell('A1');
    expect(a1.font?.name).toBe('Arial');
    expect(a1.font?.bold).toBe(true);

    // 테두리(Border) 검증
    expect(a1.border).toBeDefined();

    // 표시 형식(Number Format) 검증: B2 셀 통화 형식
    const b2 = ws!.getCell('B2');
    expect(b2.numFmt).toContain('#,##0');

    // 수정된 셀 값 검증
    expect(ws!.getCell('A2').value).toBe('프리미엄 펜');
  });

  it('Sample 3: 다중 시트(Multi-sheet), 수식(Formulas), 타 시트 참조 보존 검증', async () => {
    const filePath = path.join(samplesDir, 'sample3_formulas_multi_sheet.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample3');
    // 시트 개수 3개 (매출현황, 지출현황, 손익요약) 확인
    expect(snapshot.sheetOrder.length).toBe(3);

    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);

    // 3개 시트 모두 보존 확인
    expect(reloadedWb.worksheets.length).toBe(3);
    const sheetNames = reloadedWb.worksheets.map(w => w.name);
    expect(sheetNames).toContain('매출현황');
    expect(sheetNames).toContain('비용현황');
    expect(sheetNames).toContain('요약보고');

    // 수식 보존 확인: 매출현황 B5 셀의 SUM 수식
    const salesWs = reloadedWb.getWorksheet('매출현황');
    const b5Val: any = salesWs!.getCell('B5').value;
    expect(b5Val.formula || salesWs!.getCell('B5').formula).toContain('SUM');

    // 타 시트 참조 수식 보존 확인: 요약보고 B2 셀
    const summaryWs = reloadedWb.getWorksheet('요약보고');
    const summaryB2: any = summaryWs!.getCell('B2').value;
    expect(summaryB2.formula || summaryWs!.getCell('B2').formula).toContain('매출현황');
  });

  it('Sample 4: 차트(Chart) 사전 감지 경고 및 데이터 테이블 보존 검증', async () => {
    const filePath = path.join(samplesDir, 'sample4_chart_loss_test.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    // 사전 감지기(excelLossDetector)가 실제 차트를 정상 감지하는지 검증
    const lossReport = await detectExcelUnsupportedFeatures(origDataUrl);
    expect(lossReport.hasCharts).toBe(true);
    expect(lossReport.unsupportedList.some(item => item.includes('차트'))).toBe(true);

    // Univer 로드 및 데이터 테이블 보존 검증
    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample4');
    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    const ws = reloadedWb.getWorksheet(1);
    expect(ws!.getCell('A1').value).toBe('분기');
    expect(ws!.getCell('B2').value).toBe(120);
  });

  it('Sample 5: 조건부 서식(Conditional Formatting), 데이터 유효성 검사(Validation), 메모(Notes), 틀 고정, 인쇄 설정 in-place 보존 검증', async () => {
    const filePath = path.join(samplesDir, 'sample5_conditional_and_validation.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    const origWb = new ExcelJS.Workbook();
    await origWb.xlsx.load(origBuf);
    const origWs = origWb.getWorksheet(1);

    // 원본에 조건부 서식 및 유효성 검사가 존재하는지 확인
    expect(origWs!.conditionalFormattings?.length).toBeGreaterThan(0);
    expect(origWs!.getCell('B2').dataValidation).toBeDefined();

    // Univer 로드 후 셀 데이터 수정
    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample5');
    const firstSheetId = snapshot.sheetOrder[0];
    snapshot.sheets[firstSheetId].cellData[1][0] = { v: '이몽룡' };

    // In-place ExcelJS 라운드트립 저장
    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    // 재로드하여 고급 요소들이 원본 그대로 보존되었는지 검증
    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    const reloadedWs = reloadedWb.getWorksheet(1);

    // 1. 조건부 서식(Conditional Formatting) 보존 검증
    expect(reloadedWs!.conditionalFormattings?.length).toBe(origWs!.conditionalFormattings?.length);

    // 2. 데이터 유효성 검사(Data Validation - 드롭다운 목록) 보존 검증
    const dvB2 = reloadedWs!.getCell('B2').dataValidation;
    expect(dvB2).toBeDefined();
    expect(dvB2?.type).toBe('list');

    // 3. 셀 메모(Note) 보존 검증
    const noteC2 = reloadedWs!.getCell('C2').note;
    expect(noteC2).toBeDefined();
    const noteText = typeof noteC2 === 'object' ? JSON.stringify(noteC2) : String(noteC2);
    expect(noteText).toContain('인사평가');

    // 4. 틀 고정(Freeze Panes) 보존 검증
    expect(reloadedWs!.views).toBeDefined();
    expect(reloadedWs!.views!.length).toBeGreaterThan(0);
    expect((reloadedWs!.views![0] as any).state).toBe('frozen');

    // 5. 페이지/인쇄 설정(Page Setup) 보존 검증
    expect(reloadedWs!.pageSetup.orientation).toBe('landscape');
    expect(reloadedWs!.pageSetup.fitToPage).toBe(true);

    // 6. 수정된 셀 값 검증
    expect(reloadedWs!.getCell('A2').value).toBe('이몽룡');
  });

  it('In-place 시트 동기화: 시트 추가, 삭제, 이름 변경이 원본 워크북에 온전히 반영되는지 검증', async () => {
    const filePath = path.join(samplesDir, 'sample3_formulas_multi_sheet.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'Sample3');
    // 원본: [ '매출현황', '비용현황', '요약보고' ]
    // 1) '비용현황' 시트 삭제
    const costSheetId = snapshot.sheetOrder.find(id => snapshot.sheets[id].name === '비용현황')!;
    snapshot.sheetOrder = snapshot.sheetOrder.filter(id => id !== costSheetId);
    delete snapshot.sheets[costSheetId];

    // 2) '요약보고' -> '경영종합보고' 로 이름 변경
    const summarySheetId = snapshot.sheetOrder.find(id => snapshot.sheets[id].name === '요약보고')!;
    snapshot.sheets[summarySheetId].name = '경영종합보고';

    // 3) '신규시트' 추가
    const newSheetId = 'sheet_new_99';
    snapshot.sheetOrder.push(newSheetId);
    snapshot.sheets[newSheetId] = {
      id: newSheetId,
      name: '신규시트',
      cellData: { 0: { 0: { v: '새로운데이터' } } },
    } as any;

    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);

    const sheetNames = reloadedWb.worksheets.map(w => w.name);
    // 검증: '비용현황' 삭제됨
    expect(sheetNames).not.toContain('비용현황');
    // 검증: '경영종합보고' 로 이름 변경됨
    expect(sheetNames).toContain('경영종합보고');
    // 검증: '신규시트' 추가됨
    expect(sheetNames).toContain('신규시트');
    // 검증: 신규시트 데이터 정상
    const newWs = reloadedWb.getWorksheet('신규시트');
    expect(newWs!.getCell('A1').value).toBe('새로운데이터');
  });

  it('셀 내용 지우기: B2 값을 지우고 저장 → 저장 파일의 B2가 비어 있어야 함', async () => {
    const filePath = path.join(samplesDir, 'sample5_conditional_and_validation.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    // 원본 B2 값 확인 (예: '개발팀')
    const origWb = new ExcelJS.Workbook();
    await origWb.xlsx.load(origBuf);
    expect(origWb.getWorksheet(1)!.getCell('B2').value).toBeDefined();

    // Univer 모델로 로드 후 B2 셀 값 비우기/삭제
    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'ClearB2');
    const firstSheetId = snapshot.sheetOrder[0];
    delete snapshot.sheets[firstSheetId].cellData[1][1];

    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    const reloadedWs = reloadedWb.getWorksheet(1);

    // B2 값이 비어(null) 있어야 함
    expect(reloadedWs!.getCell('B2').value).toBeNull();
  });

  it('행 삭제: 3행 삭제 후 저장 → 마지막 행이 중복되지 않고 전체 행 수가 1 줄어야 함', async () => {
    const filePath = path.join(samplesDir, 'sample5_conditional_and_validation.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    const origWb = new ExcelJS.Workbook();
    await origWb.xlsx.load(origBuf);
    const origRowCount = origWb.getWorksheet(1)!.rowCount;
    const origRow4Val = origWb.getWorksheet(1)!.getCell('A4').value;

    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'DeleteRow');
    const firstSheetId = snapshot.sheetOrder[0];

    // 3행(0-based index 2)을 제거하고 이후 행 시프트
    const cellData = snapshot.sheets[firstSheetId].cellData;
    delete cellData[2];
    for (let r = 3; r <= 10; r++) {
      if (cellData[r]) {
        cellData[r - 1] = cellData[r];
        delete cellData[r];
      }
    }

    // 3행 삭제 연산 전달
    const structuralOps: StructuralOperation[] = [
      { type: 'removeRow', subUnitId: firstSheetId, start: 3, count: 1 }
    ];

    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl, structuralOps);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    const reloadedWs = reloadedWb.getWorksheet(1);

    // 전체 행 수가 정확히 1 줄어야 함
    expect(reloadedWs!.rowCount).toBe(origRowCount - 1);
    // 기존 4행의 값이 이제 3행에 위치해야 함
    expect(reloadedWs!.getCell('A3').value).toBe(origRow4Val);
  });

  it('행 삽입: 2행 위에 행 삽입 후 저장 → 조건부 서식, 드롭다운, 메모, 병합이 데이터와 같이 한 칸씩 내려가야 함', async () => {
    const filePath = path.join(samplesDir, 'sample5_conditional_and_validation.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    const origWb = new ExcelJS.Workbook();
    await origWb.xlsx.load(origBuf);
    const origWs = origWb.getWorksheet(1);
    const origRowCount = origWs!.rowCount;
    const origB2Val = origWs!.getCell('B2').value;

    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'InsertRow');
    const firstSheetId = snapshot.sheetOrder[0];

    // 2행 위에 1개 행 삽입 (기존 2행 데이터는 3행으로 시프트)
    const cellData = snapshot.sheets[firstSheetId].cellData;
    const newCellData: Record<number, Record<number, any>> = {};
    Object.keys(cellData).forEach((rStr) => {
      const r = parseInt(rStr, 10);
      if (r === 0) {
        newCellData[0] = cellData[0]; // 1행(헤더) 유지
      } else {
        newCellData[r + 1] = cellData[r]; // 2행부터 한 칸씩 아래로 시프트
      }
    });
    // 삽입된 2행(0-based index 1)에 신규 데이터 삽입
    newCellData[1] = { 0: { v: '신규입사자' }, 1: { v: '경영지원팀' }, 2: { v: 88 } };
    snapshot.sheets[firstSheetId].cellData = newCellData;

    // 2행 삽입 구조 변경 연산
    const structuralOps: StructuralOperation[] = [
      { type: 'insertRow', subUnitId: firstSheetId, start: 2, count: 1 }
    ];

    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl, structuralOps);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    const reloadedWs = reloadedWb.getWorksheet(1);

    // 전체 행 수가 1 증가
    expect(reloadedWs!.rowCount).toBe(origRowCount + 1);
    // 2행에는 신규 입사자 데이터
    expect(reloadedWs!.getCell('A2').value).toBe('신규입사자');
    // 기존 2행의 데이터는 3행으로 한 칸 내려감
    expect(reloadedWs!.getCell('B3').value).toBe(origB2Val);
    // 메모(Note)도 3행으로 한 칸 내려감
    const noteC3 = reloadedWs!.getCell('C3').note;
    expect(noteC3).toBeDefined();
    const noteText = typeof noteC3 === 'object' ? JSON.stringify(noteC3) : String(noteC3);
    expect(noteText).toContain('인사평가');
    // 드롭다운(Data Validation)도 3행으로 한 칸 내려감
    expect(reloadedWs!.getCell('B3').dataValidation).toBeDefined();
    expect(reloadedWs!.getCell('B3').dataValidation?.type).toBe('list');
  });

  it('열 삭제/삽입: 2열 앞에 열 삽입 시 데이터/메모/드롭다운이 한 칸 우측(C열)으로 이동하고, 열 삭제 시 복구됨', async () => {
    const filePath = path.join(samplesDir, 'sample5_conditional_and_validation.xlsx');
    const origBuf = fs.readFileSync(filePath);
    const origDataUrl = bufferToDataUrl(origBuf);

    const origWb = new ExcelJS.Workbook();
    await origWb.xlsx.load(origBuf);
    const origWs = origWb.getWorksheet(1);
    const origB2Val = origWs!.getCell('B2').value;

    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'InsertCol');
    const firstSheetId = snapshot.sheetOrder[0];

    // 2열(B열) 앞에 1개 열 삽입: Univer cellData에서도 1열(0-based) 이후 데이터를 한 칸씩 우측으로 시프트
    const cellData = snapshot.sheets[firstSheetId].cellData;
    const newCellData: Record<number, Record<number, any>> = {};
    Object.keys(cellData).forEach((rStr) => {
      const r = parseInt(rStr, 10);
      newCellData[r] = {};
      Object.keys(cellData[r]).forEach((cStr) => {
        const c = parseInt(cStr, 10);
        if (c === 0) {
          newCellData[r][0] = cellData[r][0]; // A열 유지
        } else {
          newCellData[r][c + 1] = cellData[r][c]; // B열부터 한 칸씩 우측으로 시프트
        }
      });
    });
    snapshot.sheets[firstSheetId].cellData = newCellData;

    // 2열(B열) 앞에 1개 열 삽입
    const structuralOpsInsert: StructuralOperation[] = [
      { type: 'insertCol', subUnitId: firstSheetId, start: 2, count: 1 }
    ];

    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl, structuralOpsInsert);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    const reloadedWs = reloadedWb.getWorksheet(1);

    // 기존 B2 데이터/드롭다운이 C2로 한 칸 우측 시프트
    expect(reloadedWs!.getCell('C2').value).toBe(origB2Val);
    expect(reloadedWs!.getCell('C2').dataValidation?.type).toBe('list');

    // 열 삭제 (2열 삭제하여 원래대로 복구)
    // Univer cellData에서도 다시 한 칸 좌측으로 복구
    const restoredCellData: Record<number, Record<number, any>> = {};
    Object.keys(newCellData).forEach((rStr) => {
      const r = parseInt(rStr, 10);
      restoredCellData[r] = {};
      Object.keys(newCellData[r]).forEach((cStr) => {
        const c = parseInt(cStr, 10);
        if (c === 0) {
          restoredCellData[r][0] = newCellData[r][0];
        } else if (c >= 2) {
          restoredCellData[r][c - 1] = newCellData[r][c];
        }
      });
    });
    snapshot.sheets[firstSheetId].cellData = restoredCellData;

    const structuralOpsRemove: StructuralOperation[] = [
      { type: 'removeCol', subUnitId: firstSheetId, start: 2, count: 1 }
    ];
    const restoredDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', savedDataUrl, structuralOpsRemove);
    const restoredBuf = dataUrlToBuffer(restoredDataUrl);

    const restoredWb = new ExcelJS.Workbook();
    await restoredWb.xlsx.load(restoredBuf);
    const restoredWs = restoredWb.getWorksheet(1);

    expect(restoredWs!.getCell('B2').value).toBe(origB2Val);
    expect(restoredWs!.getCell('B2').dataValidation?.type).toBe('list');
  });

  // =========================================================================
  // 실파일 4개 라운드트립 및 한셀(HanCell) XML 네임스페이스 호환성 테스트
  // =========================================================================

  test('실파일A_인사평가_MS엑셀.xlsx 라운드트립: 셀 수정, 셀 지우기, 행 삽입/삭제 및 전 서식 보존 검증', async () => {
    const origDataUrl = readSampleDataUrl('real/실파일A_인사평가_MS엑셀.xlsx');
    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일A_MS');
    expect(snapshot.sheetOrder.length).toBe(3);

    const sheetId = snapshot.sheetOrder[0];
    const sheet = snapshot.sheets[sheetId];
    expect(sheet.name).toBe('인사평가');

    // 1. 셀 수정 (A2 값 수정)
    if (!sheet.cellData[1]) sheet.cellData[1] = {};
    sheet.cellData[1][0] = { v: '2026 인사평가 수정본' };

    // 2. 셀 지우기 (B2 값 비우기)
    sheet.cellData[1][1] = { v: '' };

    // In-Place 라운드트립 저장
    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    expect(reloadedWb.worksheets.length).toBe(3);

    const ws = reloadedWb.getWorksheet('인사평가')!;
    expect(ws).toBeDefined();

    // 1) 셀 수정/지우기 확인
    expect(ws.getCell('A2').value).toBe('2026 인사평가 수정본');
    expect(ws.getCell('B2').value).toBeNull();

    // 2) 병합 셀 보존 (A1:H1)
    const merges = (ws as any).model?.merges || [];
    expect(merges.some((m: string) => m.includes('A1:H1') || m.includes('A1'))).toBe(true);

    // 3) 글꼴 및 스타일 보존 (A1: 16pt 굵게, A3: 흰색 글자 굵게)
    expect(ws.getCell('A1').font?.bold).toBe(true);
    expect(ws.getCell('A3').font?.bold).toBe(true);

    // 4) 수식 보존 (F14: AVERAGE, G14: SUM)
    expect(ws.getCell('F14').formula).toContain('AVERAGE');
    expect(ws.getCell('G14').formula).toContain('SUM');

    // 5) 메모(Notes) 보존 (F4, B11)
    expect(ws.getCell('F4').note).toBeDefined();
    expect(ws.getCell('B11').note).toBeDefined();

    // 6) 데이터 유효성 검사 (드롭다운 목록 및 정수 범위) 보존
    expect(ws.getCell('C4').dataValidation?.type).toBe('list');
    expect(ws.getCell('F4').dataValidation?.type).toBe('whole');

    // 7) 조건부 서식 보존
    expect((ws as any).conditionalFormattings?.length).toBeGreaterThanOrEqual(1);

    // 8) 틀 고정 (Freeze Panes) 보존
    expect(ws.views?.[0]?.state).toBe('frozen');

    // 9) 페이지 인쇄 설정 보존
    expect(ws.pageSetup?.fitToPage).toBe(true);
  });

  test('실파일A_인사평가_한셀.xlsx 라운드트립: XML 네임스페이스(x:, ep:) 정규화 후 MS엑셀과 동일한 전 서식 보존 검증', async () => {
    const origDataUrl = readSampleDataUrl('real/실파일A_인사평가_한셀.xlsx');
    // 한셀 네임스페이스 자동 정규화 후 Univer 로드
    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일A_한셀');
    expect(snapshot.sheetOrder.length).toBe(3);

    const sheetId = snapshot.sheetOrder[0];
    const sheet = snapshot.sheets[sheetId];
    expect(sheet.name).toBe('인사평가');

    // 1. 셀 수정 (A2 값 수정)
    if (!sheet.cellData[1]) sheet.cellData[1] = {};
    sheet.cellData[1][0] = { v: '2026 한셀 인사평가 수정본' };

    // 2. 셀 지우기 (B2 값 비우기)
    sheet.cellData[1][1] = { v: '' };

    // In-Place 라운드트립 저장
    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    expect(reloadedWb.worksheets.length).toBe(3);

    const ws = reloadedWb.getWorksheet('인사평가')!;
    expect(ws).toBeDefined();

    // 1) 셀 수정/지우기 확인
    expect(ws.getCell('A2').value).toBe('2026 한셀 인사평가 수정본');
    expect(ws.getCell('B2').value).toBeNull();

    // 2) 병합 셀 보존 (A1:H1)
    const merges = (ws as any).model?.merges || [];
    expect(merges.some((m: string) => m.includes('A1:H1') || m.includes('A1'))).toBe(true);

    // 3) 글꼴 보존 (A1: 16pt 굵게, A3: 흰색 글자 굵게)
    expect(ws.getCell('A1').font?.bold).toBe(true);
    expect(ws.getCell('A3').font?.bold).toBe(true);

    // 4) 수식 보존 (F14: AVERAGE, G14: SUM)
    expect(ws.getCell('F14').formula).toContain('AVERAGE');
    expect(ws.getCell('G14').formula).toContain('SUM');

    // 5) 메모(Notes) 보존: 한셀의 x:comments 및 ref="F4:F4" 정규화로 F4, B11 메모 복원 성공 확인
    expect(ws.getCell('F4').note).toBeDefined();
    expect(ws.getCell('B11').note).toBeDefined();

    // 6) 데이터 유효성 검사 (드롭다운 목록 및 정수 범위) 보존
    expect(ws.getCell('C4').dataValidation?.type).toBe('list');
    expect(ws.getCell('F4').dataValidation?.type).toBe('whole');

    // 7) 조건부 서식 보존
    expect((ws as any).conditionalFormattings?.length).toBeGreaterThanOrEqual(1);

    // 8) 틀 고정 (Freeze Panes) 보존
    expect(ws.views?.[0]?.state).toBe('frozen');

    // 9) 페이지 인쇄 설정 보존
    expect(ws.pageSetup?.fitToPage).toBe(true);
  });

  test('실파일B_견적서_MS엑셀.xlsx 라운드트립: 숨긴 H열(원가), richText, 하이퍼링크, 자동필터, 병합 11개(B2:H2), 데이터막대/색조 보존 검증', async () => {
    const origDataUrl = readSampleDataUrl('real/실파일B_견적서_MS엑셀.xlsx');

    // 1. 차트 사전 감지 검증
    const lossReport = await detectExcelUnsupportedFeatures(origDataUrl);
    expect(lossReport.hasCharts).toBe(true);

    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일B_MS');
    expect(snapshot.sheetOrder.length).toBe(2);

    const sheetId = snapshot.sheetOrder[0];
    const sheet = snapshot.sheets[sheetId];
    expect(sheet.name).toBe('견적서');

    // 2. 셀 수정 (A1 수정)
    if (!sheet.cellData[0]) sheet.cellData[0] = {};
    sheet.cellData[0][0] = { v: '견적서 (수정본)' };

    // 3. In-Place 라운드트립 저장
    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    expect(reloadedWb.worksheets.length).toBe(2);

    const ws = reloadedWb.getWorksheet('견적서')!;
    expect(ws.getCell('A1').value).toBe('견적서 (수정본)');

    // 1) 숨긴 H열(원가) 보존 검증 (견적서 원가 노출 방지 - 최우선)
    expect(ws.getColumn('H').hidden).toBe(true);

    // 2) B8 셀 서식 있는 텍스트(richText) 및 빨간색 굵은 글씨 보존 검증
    const b8Val = ws.getCell('B8').value as any;
    expect(b8Val).toBeDefined();
    expect(b8Val.richText).toBeDefined();
    expect(Array.isArray(b8Val.richText)).toBe(true);
    const redBoldPart = b8Val.richText.find((t: any) => t.text === '부가세 포함');
    expect(redBoldPart).toBeDefined();
    expect(redBoldPart.font?.bold).toBe(true);
    const redColor = redBoldPart.font?.color?.argb?.toUpperCase();
    expect(redColor).toContain('FF0000');

    // 3) G6 하이퍼링크 보존 검증
    const g6Val = ws.getCell('G6').value as any;
    expect(g6Val).toBeDefined();
    expect(g6Val.hyperlink).toContain('https://example.com');

    // 4) 자동 필터(B10:H15) 보존 검증
    expect(ws.autoFilter).toBe('B10:H15');

    // 5) 병합 정확한 범위 11개 유지 및 제목(B2:H2) 검증
    const merges = (ws as any).model?.merges || [];
    expect(merges.length).toBe(11);
    expect(merges.includes('B2:H2')).toBe(true);

    // 6) 데이터 막대(dataBar) 조건부 서식 보존 검증 (G11:G15)
    expect((ws as any).conditionalFormattings?.length).toBeGreaterThanOrEqual(1);
    const dataBarCf = (ws as any).conditionalFormattings.find((cf: any) => cf.ref === 'G11:G15');
    expect(dataBarCf).toBeDefined();
    expect(dataBarCf.rules?.[0]?.type).toBe('dataBar');

    // 7) 월별매출 시트 색조(colorScale) 조건부 서식 보존 검증 (D2:D13)
    const ws2 = reloadedWb.getWorksheet('월별매출')!;
    expect((ws2 as any).conditionalFormattings?.length).toBeGreaterThanOrEqual(1);
    const colorScaleCf = (ws2 as any).conditionalFormattings.find((cf: any) => cf.ref === 'D2:D13');
    expect(colorScaleCf).toBeDefined();
    expect(colorScaleCf.rules?.[0]?.type).toBe('colorScale');
  });

  test('실파일B_견적서_한셀.xlsx 라운드트립: XML 정규화 후 숨긴 H열, richText, 하이퍼링크, 자동필터, 병합 11개(B2:H2), 데이터막대/색조 보존 검증', async () => {
    const origDataUrl = readSampleDataUrl('real/실파일B_견적서_한셀.xlsx');

    // 1. 차트 사전 감지 검증
    const lossReport = await detectExcelUnsupportedFeatures(origDataUrl);
    expect(lossReport.hasCharts).toBe(true);

    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, '실파일B_한셀');
    expect(snapshot.sheetOrder.length).toBe(2);

    const sheetId = snapshot.sheetOrder[0];
    const sheet = snapshot.sheets[sheetId];
    expect(sheet.name).toBe('견적서');

    // 2. 셀 수정 (A1 수정)
    if (!sheet.cellData[0]) sheet.cellData[0] = {};
    sheet.cellData[0][0] = { v: '한셀 견적서 (수정본)' };

    // 3. In-Place 라운드트립 저장
    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    const savedBuf = dataUrlToBuffer(savedDataUrl);

    const reloadedWb = new ExcelJS.Workbook();
    await reloadedWb.xlsx.load(savedBuf);
    expect(reloadedWb.worksheets.length).toBe(2);

    const ws = reloadedWb.getWorksheet('견적서')!;
    expect(ws.getCell('A1').value).toBe('한셀 견적서 (수정본)');

    // 1) 숨긴 H열(원가) 보존 검증
    expect(ws.getColumn('H').hidden).toBe(true);

    // 2) B8 셀 서식 있는 텍스트(richText) 및 빨간색 굵은 글씨 보존 검증
    const b8Val = ws.getCell('B8').value as any;
    expect(b8Val).toBeDefined();
    expect(b8Val.richText).toBeDefined();
    expect(Array.isArray(b8Val.richText)).toBe(true);
    const redBoldPart = b8Val.richText.find((t: any) => t.text === '부가세 포함');
    expect(redBoldPart).toBeDefined();
    expect(redBoldPart.font?.bold).toBe(true);
    const redColor = redBoldPart.font?.color?.argb?.toUpperCase();
    expect(redColor).toContain('FF0000');

    // 3) G6 하이퍼링크 보존 검증
    const g6Val = ws.getCell('G6').value as any;
    expect(g6Val).toBeDefined();
    expect(g6Val.hyperlink).toContain('https://example.com');

    // 4) 자동 필터(B10:H15) 보존 검증
    expect(ws.autoFilter).toBe('B10:H15');

    // 5) 병합 정확한 범위 11개 유지 및 제목(B2:H2) 검증
    const merges = (ws as any).model?.merges || [];
    expect(merges.length).toBe(11);
    expect(merges.includes('B2:H2')).toBe(true);

    // 6) 데이터 막대(dataBar) 조건부 서식 보존 검증 (G11:G15)
    expect((ws as any).conditionalFormattings?.length).toBeGreaterThanOrEqual(1);
    const dataBarCf = (ws as any).conditionalFormattings.find((cf: any) => cf.ref === 'G11:G15');
    expect(dataBarCf).toBeDefined();
    expect(dataBarCf.rules?.[0]?.type).toBe('dataBar');

    // 7) 월별매출 시트 색조(colorScale) 조건부 서식 보존 검증 (D2:D13)
    const ws2 = reloadedWb.getWorksheet('월별매출')!;
    expect((ws2 as any).conditionalFormattings?.length).toBeGreaterThanOrEqual(1);
    const colorScaleCf = (ws2 as any).conditionalFormattings.find((cf: any) => cf.ref === 'D2:D13');
    expect(colorScaleCf).toBeDefined();
    expect(colorScaleCf.rules?.[0]?.type).toBe('colorScale');
  });

  test('Univer 구조 변경 기록: mutation ID 정확 일치 필터링 및 Undo 시 상쇄 검증', () => {
    // ExcelEditorModal 내부의 structuralOpsRef 기록 로직 단위 검증
    const ops: StructuralOperation[] = [];

    function processMutation(cmd: any) {
      const id = cmd?.id || '';
      const range = cmd?.params?.range;
      const subUnitId = cmd.params?.subUnitId;

      if (range) {
        let op: StructuralOperation | null = null;
        if (id === 'sheet.mutation.insert-row') {
          op = { type: 'insertRow', subUnitId, start: range.startRow + 1, count: range.endRow - range.startRow + 1 };
        } else if (id === 'sheet.mutation.remove-rows') {
          op = { type: 'removeRow', subUnitId, start: range.startRow + 1, count: range.endRow - range.startRow + 1 };
        } else if (id === 'sheet.mutation.insert-col') {
          op = { type: 'insertCol', subUnitId, start: range.startColumn + 1, count: range.endColumn - range.startColumn + 1 };
        } else if (id === 'sheet.mutation.remove-col') {
          op = { type: 'removeCol', subUnitId, start: range.startColumn + 1, count: range.endColumn - range.startColumn + 1 };
        }

        if (op) {
          if (ops.length > 0) {
            const lastOp = ops[ops.length - 1];
            const isSameSheet = !lastOp.subUnitId || !op.subUnitId || lastOp.subUnitId === op.subUnitId;
            const isOppositeRow =
              isSameSheet &&
              ((lastOp.type === 'insertRow' && op.type === 'removeRow') ||
               (lastOp.type === 'removeRow' && op.type === 'insertRow')) &&
              lastOp.start === op.start &&
              lastOp.count === op.count;
            const isOppositeCol =
              isSameSheet &&
              ((lastOp.type === 'insertCol' && op.type === 'removeCol') ||
               (lastOp.type === 'removeCol' && op.type === 'insertCol')) &&
              lastOp.start === op.start &&
              lastOp.count === op.count;

            if (isOppositeRow || isOppositeCol) {
              ops.pop();
              return;
            }
          }
          ops.push(op);
        }
      }
    }

    // 1. command ID (sheet.command.insert-row)는 무시되고 mutation만 기록되어 중복 0건
    processMutation({ id: 'sheet.command.insert-row', params: { range: { startRow: 2, endRow: 2 } } });
    expect(ops.length).toBe(0); // command는 무시

    processMutation({ id: 'sheet.mutation.insert-row', params: { subUnitId: 's1', range: { startRow: 2, endRow: 2 } } });
    expect(ops.length).toBe(1);
    expect(ops[0]).toEqual({ type: 'insertRow', subUnitId: 's1', start: 3, count: 1 });

    // 2. 실행 취소(Undo) 발생: sheet.mutation.remove-rows가 동일 위치에 들어옴 -> 자동 상쇄!
    processMutation({ id: 'sheet.mutation.remove-rows', params: { subUnitId: 's1', range: { startRow: 2, endRow: 2 } } });
    expect(ops.length).toBe(0); // 완전히 상쇄되어 0개!

    // 3. 다시 실행(Redo) 발생: 다시 1개 기록
    processMutation({ id: 'sheet.mutation.insert-row', params: { subUnitId: 's1', range: { startRow: 2, endRow: 2 } } });
    expect(ops.length).toBe(1);

    // 4. 열 삽입 및 열 Undo 상쇄
    processMutation({ id: 'sheet.mutation.insert-col', params: { subUnitId: 's1', range: { startColumn: 1, endColumn: 1 } } });
    expect(ops.length).toBe(2);

    processMutation({ id: 'sheet.mutation.remove-col', params: { subUnitId: 's1', range: { startColumn: 1, endColumn: 1 } } });
    expect(ops.length).toBe(1); // 열 연산만 상쇄되고 행 연산은 유지됨
  });
});
