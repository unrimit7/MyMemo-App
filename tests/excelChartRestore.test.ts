// 엑셀 저장 시 원본 차트·그리기를 다시 넣어 보존하는지, 안전하지 않을 때는 넣지 않는지 확인하는 회귀 테스트 (STEP7)
import { describe, test, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { dataUrlToUniverWorkbook, univerSnapshotToDataUrl } from '../src/utils/univerAdapter';
import { chartRestoreBlockReason } from '../src/utils/excelChartRestore';

const SAMPLES = path.join(__dirname, '..', 'docs', 'office-upgrade', 'samples', 'real');
const OUT = path.join(__dirname, '..', 'docs', 'office-upgrade', 'out', 'step7');
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const readDataUrl = (name: string) => `data:${XLSX_MIME};base64,${fs.readFileSync(path.join(SAMPLES, name)).toString('base64')}`;
const toBytes = (dataUrl: string) => new Uint8Array(Buffer.from(dataUrl.slice(dataUrl.indexOf(',') + 1), 'base64'));

async function saveWithA1(name: string, a1: string, mutate?: (snap: any) => void, ops: any[] = []) {
  const orig = readDataUrl(name);
  const snap = await dataUrlToUniverWorkbook(orig, name);
  const sheet = snap.sheets[snap.sheetOrder[0]] as any;
  sheet.cellData[0] = sheet.cellData[0] || {};
  sheet.cellData[0][0] = { ...(sheet.cellData[0][0] || {}), v: a1 };
  mutate?.(snap);
  const saved = await univerSnapshotToDataUrl(snap, 'xlsx', 'exceljs', orig, ops);
  return { orig: toBytes(orig), saved: toBytes(saved) };
}

async function sheetPathByName(zip: JSZip, name: string) {
  const wb = await zip.file('xl/workbook.xml')!.async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  const rid = (wb.match(new RegExp(`<sheet\\b[^>]*name="${name}"[^>]*r:id="([^"]+)"`)) || [])[1];
  const target = (rels.match(new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`)) || rels.match(new RegExp(`Target="([^"]+)"[^>]*Id="${rid}"`)) || [])[1];
  return `xl/${target}`;
}

fs.mkdirSync(OUT, { recursive: true });

describe('엑셀 저장 시 원본 차트 보존 (STEP7)', () => {
  for (const [name, outName] of [
    ['실파일B_견적서_MS엑셀.xlsx', 'step7_견적서_MS엑셀_A1수정_차트보존.xlsx'],
    ['실파일B_견적서_한셀.xlsx', 'step7_견적서_한셀_A1수정_차트보존.xlsx'],
  ]) {
    test(`${name}: 칸을 고쳐 저장해도 차트·그리기·rels·Content_Types가 원본대로 들어간다`, async () => {
      const { orig, saved } = await saveWithA1(name, 'STEP7 차트 보존 시험');
      fs.writeFileSync(path.join(OUT, outName), saved);
      const oz = await JSZip.loadAsync(orig);
      const sz = await JSZip.loadAsync(saved);

      // 차트·그리기 파일은 원본과 바이트까지 같다
      for (const part of Object.keys(oz.files).filter((n) => /^xl\/(charts|drawings)\/.+/.test(n) && !oz.files[n].dir)) {
        const a = await oz.file(part)!.async('uint8array');
        const b = await sz.file(part)?.async('uint8array');
        expect(b, part).toBeTruthy();
        expect(Buffer.compare(Buffer.from(a), Buffer.from(b!)), part).toBe(0);
      }

      // 차트가 있던 시트(월별매출)에 drawing 연결
      const sheetPath = await sheetPathByName(sz, '월별매출');
      const sheetXml = await sz.file(sheetPath)!.async('string');
      const rid = (sheetXml.match(/<drawing r:id="([^"]+)"\/>/) || [])[1];
      expect(rid).toBeTruthy();
      expect(sheetXml.indexOf('<drawing')).toBeGreaterThan(sheetXml.indexOf('<pageMargins'));
      const relsPath = sheetPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
      const rels = await sz.file(relsPath)!.async('string');
      expect(rels).toContain(`Id="${rid}"`);
      expect(rels).toMatch(/Target="\.\.\/drawings\/drawing\d+\.xml"/);

      const ct = await sz.file('[Content_Types].xml')!.async('string');
      expect(ct).toMatch(/PartName\s*=\s*"\/xl\/charts\/chart\d+\.xml"/);
      expect(ct).toMatch(/PartName\s*=\s*"\/xl\/drawings\/drawing\d+\.xml"/);

      // 고친 칸과 나머지 값은 그대로, ExcelJS로 다시 열린다
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(Buffer.from(saved));
      expect(wb.getWorksheet('견적서')!.getCell('A1').value).toBe('STEP7 차트 보존 시험');
      expect(wb.getWorksheet('월별매출')!.getCell('A2').value).toBeTruthy();
    });
  }

  test('행을 삽입한 저장에서는 차트를 넣지 않는다 (참조 범위가 어긋날 수 있음)', async () => {
    const { saved } = await saveWithA1('실파일B_견적서_MS엑셀.xlsx', '행삽입', undefined, [
      { type: 'insertRow', sheetName: '월별매출', start: 3, count: 1 },
    ]);
    fs.writeFileSync(path.join(OUT, 'step7_견적서_MS엑셀_행삽입_차트미보존.xlsx'), saved);
    const sz = await JSZip.loadAsync(saved);
    expect(Object.keys(sz.files).some((n) => n.startsWith('xl/charts/'))).toBe(false);
  });

  test('차트가 참조하는 시트 이름을 바꾸면 차트를 넣지 않는다', async () => {
    const { saved } = await saveWithA1('실파일B_견적서_MS엑셀.xlsx', '이름변경', (snap) => {
      const id = snap.sheetOrder.find((sid: string) => snap.sheets[sid].name === '월별매출');
      snap.sheets[id].name = '매출2';
    });
    const sz = await JSZip.loadAsync(saved);
    expect(Object.keys(sz.files).some((n) => n.startsWith('xl/charts/'))).toBe(false);
  });

  test('chartRestoreBlockReason: 보존 가능하면 null, 행 삽입·이름 변경이면 이유', async () => {
    const orig = toBytes(readDataUrl('실파일B_견적서_MS엑셀.xlsx'));
    expect(await chartRestoreBlockReason(orig, ['견적서', '월별매출'], 0)).toBeNull();
    expect(await chartRestoreBlockReason(orig, ['견적서', '월별매출'], 1)).toMatch(/삽입하거나 삭제/);
    expect(await chartRestoreBlockReason(orig, ['견적서', '매출2'], 0)).toMatch(/월별매출/);
  });
});
