import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { dataUrlToUniverWorkbook, univerSnapshotToDataUrl } from '../src/utils/univerAdapter';
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
});
