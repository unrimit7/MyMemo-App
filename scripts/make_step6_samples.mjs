// STEP6 점검용 가상 시험 파일(PDF, 숨김 행/열 엑셀, 큰 엑셀)을 만드는 스크립트
// 모든 내용은 가상 데이터다. 실행: node scripts/make_step6_samples.mjs
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.resolve(__dirname, '..', 'docs', 'office-upgrade', 'samples', 'step6');
fs.mkdirSync(outDir, { recursive: true });

async function makePdf() {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let p = 1; p <= 3; p++) {
    const page = doc.addPage([595, 842]);
    page.drawText(`STEP6 Sample PDF - Page ${p}`, { x: 60, y: 760, size: 22, font, color: rgb(0.1, 0.1, 0.4) });
    for (let i = 0; i < 12; i++) {
      page.drawText(`Line ${i + 1}: virtual test data ${p * 100 + i}`, { x: 60, y: 700 - i * 24, size: 12, font });
    }
    page.drawRectangle({ x: 60, y: 300, width: 300, height: 60, borderColor: rgb(0.8, 0.2, 0.2), borderWidth: 2 });
  }
  const bytes = await doc.save();
  fs.writeFileSync(path.join(outDir, 'step6_가상_PDF_3쪽.pdf'), bytes);
}

async function makeHiddenRowsCols() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('숨김시험');
  ws.addRow(['번호', '이름', '숨김열', '금액']);
  for (let i = 1; i <= 10; i++) ws.addRow([i, `가상인물${i}`, `숨김값${i}`, i * 1000]);
  ws.getColumn(3).hidden = true; // C열 숨김
  ws.getRow(4).hidden = true; // 4행 숨김 (번호 3)
  ws.getRow(5).hidden = true; // 5행 숨김 (번호 4)
  ws.getCell('D12').value = { formula: 'SUM(D2:D11)', result: 55000 };
  ['A1', 'B1', 'C1', 'D1'].forEach((a) => {
    ws.getCell(a).border = { top: { style: 'thin' }, bottom: { style: 'medium' }, left: { style: 'thin' }, right: { style: 'thin' } };
    ws.getCell(a).font = { bold: true };
  });
  ws.getColumn(2).width = 20;
  ws.getRow(1).height = 30;
  await wb.xlsx.writeFile(path.join(outDir, 'step6_가상_숨김행열.xlsx'));
}

async function makeBigXlsx() {
  // 약 10MB 안팎의 큰 파일: 시트 14개 × 15,000행 × 8열
  const wb = new ExcelJS.Workbook();
  for (let s = 1; s <= 14; s++) {
    const ws = wb.addWorksheet(`시트${s}`);
    ws.addRow(['번호', '거래처', '품목', '수량', '단가', '금액', '비고', '코드']);
    for (let r = 1; r <= 15000; r++) {
      ws.addRow([r, `가상거래처${r % 97}`, `품목-${s}-${r}`, r % 50, 1000 + (r % 300), (r % 50) * (1000 + (r % 300)), `비고 텍스트 ${r} 입니다`, `C${s}${String(r).padStart(6, '0')}`]);
    }
  }
  await wb.xlsx.writeFile(path.join(outDir, 'step6_가상_큰엑셀.xlsx'));
}

await makePdf();
await makeHiddenRowsCols();
await makeBigXlsx();
for (const f of fs.readdirSync(outDir)) {
  console.log(f, fs.statSync(path.join(outDir, f)).size);
}
