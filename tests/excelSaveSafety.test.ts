import { describe, it, expect, vi } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {
  dataUrlToUniverWorkbook,
  univerSnapshotToDataUrl,
  validateExcelIntegrity,
  extractXlsxIntegrityMetrics,
} from '../src/utils/univerAdapter';

const sampleFile = path.resolve(__dirname, '../docs/office-upgrade/samples/real/sample_다중시트_예시.xlsx');

function bufferToDataUrl(buf: Buffer, mime: string = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'): string {
  return `data:${mime};base64,${buf.toString('base64')}`;
}

describe('Excel 저장 안전성 및 무결성 검증 테스트 (무음 손실 저장 방지)', () => {
  it('1. ExcelJS writeBuffer 강제 실패 시 SheetJS로 대체 저장하지 않고 예외를 즉시 throw해야 한다', async () => {
    const origBuf = fs.readFileSync(sampleFile);
    const origDataUrl = bufferToDataUrl(origBuf);
    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'sample');

    // ExcelJS writeBuffer 강제 실패 모의
    // 주의: Workbook.prototype.xlsx는 getter라서 prototype에서 직접 읽으면 모든 인스턴스가 공유하는
    // 잘못된 XLSX 객체가 만들어져 이후 모든 load()가 실패한다. 실제 XLSX 클래스의 prototype을 모의한다.
    const xlsxProto = Object.getPrototypeOf(new ExcelJS.Workbook().xlsx);
    const spy = vi.spyOn(xlsxProto, 'writeBuffer').mockRejectedValueOnce(
      new Error('FORCED_EXCELJS_WRITE_DISK_ERROR')
    );

    let errorThrown: any = null;
    let savedResult: string | null = null;
    try {
      savedResult = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    } catch (err: any) {
      errorThrown = err;
    } finally {
      spy.mockRestore();
    }

    // 예외가 반드시 던져져야 하며, SheetJS fallback Data URL이 반환되어서는 안 됨
    expect(savedResult).toBeNull();
    expect(errorThrown).toBeDefined();
    expect(errorThrown.message).toContain('FORCED_EXCELJS_WRITE_DISK_ERROR');
  });

  it('2. 저장 결과에 JavaScript 날짜 문자열("GMT+")이 포함된 경우 무결성 검증에서 저장을 차단해야 한다', async () => {
    const origBuf = fs.readFileSync(sampleFile);
    const zip = await JSZip.loadAsync(origBuf);

    // 날짜가 JS 문자열("Fri Nov 27 2026 00:00:00 GMT+0900 (한국 표준시)")로 변질된 시나리오 주입
    const sheet1 = zip.file('xl/worksheets/sheet1.xml');
    expect(sheet1).toBeDefined();
    let sheetXml = await sheet1!.async('string');
    sheetXml = sheetXml.replace(/<v>[^<]*<\/v>/, '<is><t>Fri Nov 27 2026 00:00:00 GMT+0900 (한국 표준시)</t></is>');
    zip.file('xl/worksheets/sheet1.xml', sheetXml);

    const corruptBytes = await zip.generateAsync({ type: 'uint8array' });
    const origBytes = new Uint8Array(origBuf);

    await expect(validateExcelIntegrity(origBytes, corruptBytes)).rejects.toThrow(
      'GMT+'
    );
  });

  it('3. 원본 수식이 존재하는데 저장 결과에서 전량 소실된 경우 무결성 검증에서 저장을 차단해야 한다', async () => {
    const origBuf = fs.readFileSync(sampleFile);
    const zip = await JSZip.loadAsync(origBuf);

    // 수식(<f>...</f>) 전량 제거 시나리오 주입
    const sheetFiles = Object.keys(zip.files).filter((n) => n.startsWith('xl/worksheets/sheet') && n.endsWith('.xml'));
    for (const sf of sheetFiles) {
      const file = zip.file(sf);
      if (file) {
        let xml = await file.async('string');
        xml = xml.replace(/<f\b[^>]*>[\s\S]*?<\/f>/g, '');
        zip.file(sf, xml);
      }
    }

    const strippedBytes = await zip.generateAsync({ type: 'uint8array' });
    const origBytes = new Uint8Array(origBuf);

    await expect(validateExcelIntegrity(origBytes, strippedBytes)).rejects.toThrow(
      '수식'
    );
  });

  it('4. 원본 서식(cellXfs)이 풍부한데 SheetJS처럼 1~2개로 축소된 경우 무결성 검증에서 저장을 차단해야 한다', async () => {
    const origBuf = fs.readFileSync(sampleFile);
    const zip = await JSZip.loadAsync(origBuf);

    // styles.xml의 cellXfs를 1개로 축소
    const stylesFile = zip.file('xl/styles.xml');
    expect(stylesFile).toBeDefined();
    let stylesXml = await stylesFile!.async('string');
    stylesXml = stylesXml.replace(/<cellXfs\b[^>]*count="(\d+)"[\s\S]*?<\/cellXfs>/, '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>');
    zip.file('xl/styles.xml', stylesXml);

    const flattenedBytes = await zip.generateAsync({ type: 'uint8array' });
    const origBytes = new Uint8Array(origBuf);

    await expect(validateExcelIntegrity(origBytes, flattenedBytes)).rejects.toThrow(
      'cellXfs'
    );
  });

  it('5. 정상 라운드트립 시 무결성 검증 통과 및 빈 행/열이 추가되지 않고 원본 구조가 보존되어야 한다', async () => {
    const origBuf = fs.readFileSync(sampleFile);
    const origDataUrl = bufferToDataUrl(origBuf);
    const snapshot = await dataUrlToUniverWorkbook(origDataUrl, 'sample');

    // 1000행, 52열로 확장되었는지 확인
    const firstSheet = snapshot.sheets[snapshot.sheetOrder[0]];
    expect(firstSheet.rowCount).toBeGreaterThanOrEqual(1000);
    expect(firstSheet.columnCount).toBeGreaterThanOrEqual(52);

    // 정상 In-Place 저장 실행
    const savedDataUrl = await univerSnapshotToDataUrl(snapshot, 'xlsx', 'exceljs', origDataUrl);
    expect(savedDataUrl).toBeDefined();

    // 저장된 파일 재분석
    const savedBase64 = savedDataUrl.split(',')[1];
    const savedBytes = new Uint8Array(Buffer.from(savedBase64, 'base64'));
    const metrics = await extractXlsxIntegrityMetrics(savedBytes);

    expect(metrics.hasGmtString).toBe(false);
    expect(metrics.formulaCount).toBeGreaterThan(0);
    expect(metrics.cellXfsCount).toBeGreaterThan(5);

    // ExcelJS로 열어서 실제 유효 시트 크기 및 행 개수가 무한정 불어나지 않았는지 검증
    const checkWb = new ExcelJS.Workbook();
    await checkWb.xlsx.load(savedBytes as any);
    const checkWs = checkWb.getWorksheet(firstSheet.name || 1)!;
    expect(checkWs.actualRowCount).toBeLessThanOrEqual(100); // 원본 실제 데이터 범위 유지
  });

  // 셀 하나의 값을 인라인 문자열로 바꾼 복사본 생성 (테스트용)
  async function withInlineText(src: Buffer | Uint8Array, text: string): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(src);
    const sheet1 = zip.file('xl/worksheets/sheet1.xml')!;
    let xml = await sheet1.async('string');
    xml = xml.replace(/<v>[^<]*<\/v>/, `<is><t>${text}</t></is>`);
    zip.file('xl/worksheets/sheet1.xml', xml);
    return await zip.generateAsync({ type: 'uint8array' });
  }

  it('6. 원본에 이미 JS 날짜 문자열이 있던 손상 파일은 같은 개수면 저장을 막지 않아야 한다', async () => {
    const origBuf = fs.readFileSync(sampleFile);
    const damaged = await withInlineText(origBuf, 'Fri Nov 27 2026 00:00:00 GMT+0900 (한국 표준시)');
    // 원본과 저장본 모두 같은 손상 문자열 1개 → 새로 오염된 것이 아니므로 통과
    await expect(validateExcelIntegrity(damaged, damaged)).resolves.toBeUndefined();
    const m = await extractXlsxIntegrityMetrics(damaged);
    expect(m.gmtStringCount).toBe(1);
  });

  it('7. 셀에 "한국 표준시"나 "GMT+0900"을 일반 문장으로 쓴 경우는 손상으로 판정하지 않아야 한다', async () => {
    const origBuf = fs.readFileSync(sampleFile);
    const normalText = await withInlineText(origBuf, '회의 시간은 한국 표준시(GMT+0900) 기준');
    const m = await extractXlsxIntegrityMetrics(normalText);
    expect(m.gmtStringCount).toBe(0);
    await expect(validateExcelIntegrity(new Uint8Array(origBuf), normalText)).resolves.toBeUndefined();
  });
});
