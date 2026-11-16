// 숨긴 행·열(K2)이 미리보기와 편집기 변환에서 숨김으로 유지되는지 확인하는 회귀 테스트
import { describe, test, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import { parseExcelWorkbook, parseExcelWorkbookWithStyles, renderExcelTableHtml } from '../src/utils/excelViewer';
import { dataUrlToUniverWorkbook } from '../src/utils/univerAdapter';

const file = path.join(__dirname, '..', 'docs', 'office-upgrade', 'samples', 'step6', 'step6_가상_숨김행열.xlsx');
const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${fs.readFileSync(file).toString('base64')}`;

describe('숨긴 행·열 (K2)', () => {
  test('빠른 미리보기(SheetJS)에서 숨긴 C열과 4·5행 내용이 보이지 않는다', () => {
    const parsed = parseExcelWorkbook(dataUrl, 'k2-fast');
    const { tableBodyHtml, tableHeaderHtml } = renderExcelTableHtml(parsed, '숨김시험', 50);
    expect(tableBodyHtml).not.toContain('숨김값1');
    expect(tableBodyHtml).not.toContain('가상인물3');
    expect(tableBodyHtml).not.toContain('가상인물4');
    expect(tableBodyHtml).toContain('가상인물5');
    expect(tableHeaderHtml).not.toMatch(/>C<\/th>/);
  });

  test('서식 미리보기(ExcelJS)에서도 숨긴 행·열이 보이지 않는다', async () => {
    const parsed = await parseExcelWorkbookWithStyles(dataUrl, 'k2-styled');
    expect(parsed.hiddenCols?.['숨김시험']).toEqual([2]);
    expect(parsed.hiddenRows?.['숨김시험']).toEqual([3, 4]);
    const { tableBodyHtml } = renderExcelTableHtml(parsed, '숨김시험', 50);
    expect(tableBodyHtml).not.toContain('숨김값');
    expect(tableBodyHtml).not.toContain('가상인물3');
    expect(tableBodyHtml).toContain('가상인물10');
  });

  test('편집기(Univer)로 불러올 때 숨긴 행·열 표시(hd)를 넘긴다', async () => {
    const wb = await dataUrlToUniverWorkbook(dataUrl, 'k2');
    const sheet = wb.sheets[wb.sheetOrder[0]] as any;
    expect(sheet.rowData[3]?.hd).toBe(1);
    expect(sheet.rowData[4]?.hd).toBe(1);
    expect(sheet.columnData[2]?.hd).toBe(1);
    expect(sheet.rowData[2]?.hd).toBeUndefined();
  });
});
