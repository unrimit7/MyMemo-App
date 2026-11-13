import ExcelJS from 'exceljs';

console.log('[1/4] sample_다중시트_예시.xlsx 로드 및 새 워크북 생성 시작...');
const filePath = 'docs/office-upgrade/samples/real/sample_다중시트_예시.xlsx';
const oldWb = new ExcelJS.Workbook();
await oldWb.xlsx.readFile(filePath);

const newWb = new ExcelJS.Workbook();

// 1. '결제 예정' 시트를 첫 번째 시트로 생성 (페이지 로드 직후 첫 화면 검증 가능)
const sheet = newWb.addWorksheet('결제 예정');
console.log('[2/4] 결제 예정 시트 데이터 및 서식 작성 중...');

// Row 1: Header title
sheet.mergeCells('A1:H1');
const titleCell = sheet.getCell('A1');
titleCell.value = '2026년 11월 정기 결제 예정';
titleCell.font = { name: '맑은 고딕', size: 14, bold: true, color: { argb: 'FF1E293B' } };
titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

// Row 2: 빈 줄
sheet.getRow(2).values = [];

// Row 3: Column headers
const headers = ['결제일', '구분', '서비스명', '결제주기', '결제금액', '결제수단', '남은 일수', '상태'];
const hRow = sheet.getRow(3);
headers.forEach((h, idx) => {
  const cell = hRow.getCell(idx + 1);
  cell.value = h;
  cell.font = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FF334155' } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };
});

// Row 4: 사용자 검증 대상 행
// A4 = 2026-11-27 (Date, numFmt: yyyy-mm-dd, 가운데 정렬)
// E4 = 23100 (numFmt: #,##0"원";-#,##0"원";"-")
// G4 = =A4-TODAY() (numFmt: 0"일";"지남 "0"일") - 캐시값 없음
const row4 = sheet.getRow(4);
row4.getCell(1).value = new Date(Date.UTC(2026, 10, 27));
row4.getCell(1).numFmt = 'yyyy-mm-dd';
row4.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

row4.getCell(2).value = '도메인';
row4.getCell(3).value = 'kr 도메인 연장';
row4.getCell(4).value = '연간';

row4.getCell(5).value = 23100;
row4.getCell(5).numFmt = '#,##0"원";-#,##0"원";"-"';

row4.getCell(6).value = '신용카드';

row4.getCell(7).value = { formula: 'A4-TODAY()' }; // 캐시값 없음
row4.getCell(7).numFmt = '0"일";"지남 "0"일"';

row4.getCell(8).value = '결제대기';

// Row 5
const row5 = sheet.getRow(5);
row5.getCell(1).value = new Date(Date.UTC(2026, 10, 20));
row5.getCell(1).numFmt = 'yyyy-mm-dd';
row5.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
row5.getCell(2).value = 'SSL 인증서';
row5.getCell(3).value = '와일드카드 인증서';
row5.getCell(4).value = '연간';
row5.getCell(5).value = 110000;
row5.getCell(5).numFmt = '#,##0"원";-#,##0"원";"-"';
row5.getCell(6).value = '법인카드';
row5.getCell(7).value = { formula: 'A5-TODAY()' };
row5.getCell(7).numFmt = '0"일";"지남 "0"일"';
row5.getCell(8).value = '승인완료';

// Row 6 (과거 결제 건)
const row6 = sheet.getRow(6);
row6.getCell(1).value = new Date(Date.UTC(2026, 10, 10));
row6.getCell(1).numFmt = 'yyyy-mm-dd';
row6.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
row6.getCell(2).value = '호스팅';
row6.getCell(3).value = '기본 웹 서버';
row6.getCell(4).value = '월간';
row6.getCell(5).value = 65000;
row6.getCell(5).numFmt = '#,##0"원";-#,##0"원";"-"';
row6.getCell(6).value = '계좌이체';
row6.getCell(7).value = { formula: 'A6-TODAY()' };
row6.getCell(7).numFmt = '0"일";"지남 "0"일"';
row6.getCell(8).value = '결제완료';

// 2. 기존 다른 시트들 순서대로 복사
console.log('[3/4] 기존 다른 시트들 복사 중...');
for (const oldWs of oldWb.worksheets) {
  if (oldWs.name === '결제 예정') continue;
  const newWs = newWb.addWorksheet(oldWs.name);
  oldWs.eachRow((r, rNum) => {
    const newR = newWs.getRow(rNum);
    r.eachCell((c, cNum) => {
      const newC = newR.getCell(cNum);
      newC.value = c.value;
      if (c.numFmt) newC.numFmt = c.numFmt;
      if (c.font) newC.font = { ...c.font };
      if (c.fill) newC.fill = { ...c.fill };
      if (c.alignment) newC.alignment = { ...c.alignment };
      if (c.border) newC.border = { ...c.border };
    });
  });
}

console.log('[4/4] 파일 저장 중... 최종 시트 순서:', newWb.worksheets.map(s => s.name));
await newWb.xlsx.writeFile(filePath);
console.log('✓ sample_다중시트_예시.xlsx 저장 성공 완료!');
process.exit(0);
