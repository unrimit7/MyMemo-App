// 셀 일부를 고쳐 값(v) 없이 서식 있는 글(p)만 남은 칸이 저장 시 지워지지 않는지 확인하는 회귀 테스트
import { describe, test, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { dataUrlToUniverWorkbook, univerSnapshotToDataUrl, withPlainTextFromRichCells } from '../src/utils/univerAdapter';

const file = path.join(__dirname, '..', 'docs', 'office-upgrade', 'samples', 'step6', 'step6_가상_숨김행열.xlsx');
const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${fs.readFileSync(file).toString('base64')}`;

describe('서식 있는 글로만 남은 칸 저장', () => {
  test('withPlainTextFromRichCells: v가 비고 p에 글이 있으면 v를 채운다', () => {
    const out = withPlainTextFromRichCells({ 0: { 0: { p: { body: { dataStream: '첫줄\r둘째줄\r\n' } } } as any, 1: { v: 5 } } });
    expect(out[0][0].v).toBe('첫줄\n둘째줄');
    expect(out[0][1].v).toBe(5);
  });

  test('A1을 셀 안에서 고친 상태(p만 있음)로 저장해도 입력한 글이 남는다', async () => {
    const snapshot = await dataUrlToUniverWorkbook(dataUrl, 'rich');
    const sheet = snapshot.sheets[snapshot.sheetOrder[0]] as any;
    const prev = sheet.cellData[0][0];
    sheet.cellData[0][0] = {
      s: prev?.s,
      p: { id: '__INTERNAL_EDITOR__DOCS_NORMAL', body: { dataStream: '번호K2수정\r\n', textRuns: [{ st: 0, ed: 6, ts: { bl: 1 } }] } },
    };
    const saved = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', dataUrl);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(saved.slice(saved.indexOf(',') + 1), 'base64'));
    const ws = wb.worksheets[0];
    expect(ws.getCell('A1').value).toBe('번호K2수정');
    expect(ws.getCell('B1').value).toBe('이름');
    // 숨긴 행·열도 그대로
    expect(ws.getRow(4).hidden).toBe(true);
    expect(ws.getColumn(3).hidden).toBe(true);
  });
});
