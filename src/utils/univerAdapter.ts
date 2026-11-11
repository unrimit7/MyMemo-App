import * as XLSX from 'xlsx';
import type { IWorkbookData, IWorksheetData, ICellData, IRange } from '@univerjs/core';
import { LocaleType } from '@univerjs/core';

// Data URL 또는 Base64 문자열을 안전하게 Uint8Array로 변환
function dataUrlToUint8Array(dataOrUrl: string): Uint8Array {
  const commaIdx = dataOrUrl.indexOf(',');
  const base64 = commaIdx !== -1 ? dataOrUrl.slice(commaIdx + 1) : dataOrUrl;
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/**
 * XLSX/CSV Base64 또는 Data URL을 Univer의 IWorkbookData 모델로 변환합니다.
 */
export function dataUrlToUniverWorkbook(dataOrUrl: string, defaultName: string = 'Workbook'): IWorkbookData {
  const workbookData: IWorkbookData = {
    id: `workbook_${Date.now()}`,
    name: defaultName,
    appVersion: '3.0.0-alpha',
    locale: LocaleType.KO_KR,
    sheets: {},
    sheetOrder: [],
    styles: {},
  };

  try {
    const bytes = dataUrlToUint8Array(dataOrUrl);
    const workbook = XLSX.read(bytes, { type: 'array', cellDates: true, cellFormula: true });
    const sheetNames = workbook.SheetNames || [];

    if (sheetNames.length === 0) {
      sheetNames.push('Sheet1');
    }

    sheetNames.forEach((sheetName, index) => {
      const sheetId = `sheet_${index + 1}`;
      const ws = workbook.Sheets[sheetName];

      const cellData: Record<number, Record<number, ICellData>> = {};
      const mergeData: IRange[] = [];
      let maxRow = 29;
      let maxCol = 15;

      if (ws) {
        const refStr = ws['!ref'];
        if (refStr) {
          const range = XLSX.utils.decode_range(refStr);
          maxRow = Math.max(maxRow, range.e.r);
          maxCol = Math.max(maxCol, range.e.c);

          for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
              const cellAddress = XLSX.utils.encode_cell({ r, c });
              const cell = ws[cellAddress];
              if (cell) {
                if (!cellData[r]) cellData[r] = {};
                const cItem: ICellData = {};

                if (cell.f) {
                  cItem.f = cell.f.startsWith('=') ? cell.f : `=${cell.f}`;
                }
                if (cell.v !== undefined && cell.v !== null) {
                  cItem.v = cell.v;
                }
                cellData[r][c] = cItem;
              }
            }
          }
        }

        // 셀 병합(Merges) 복원
        if (ws['!merges'] && Array.isArray(ws['!merges'])) {
          ws['!merges'].forEach((m) => {
            mergeData.push({
              startRow: m.s.r,
              endRow: m.e.r,
              startColumn: m.s.c,
              endColumn: m.e.c,
            });
          });
        }
      }

      const worksheet: IWorksheetData = {
        id: sheetId,
        name: sheetName,
        tabColor: '',
        hidden: 0,
        freeze: { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 },
        rowCount: Math.max(100, maxRow + 20),
        columnCount: Math.max(26, maxCol + 10),
        zoomRatio: 1,
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 88,
        defaultRowHeight: 24,
        mergeData,
        cellData,
        rowData: {},
        columnData: {},
        rowHeader: { width: 46 },
        columnHeader: { height: 26 },
        showGridlines: 1,
        rightToLeft: 0,
      };

      workbookData.sheets[sheetId] = worksheet;
      workbookData.sheetOrder.push(sheetId);
    });
  } catch (err) {
    console.warn('dataUrlToUniverWorkbook 파싱 실패, 기본 빈 시트로 생성합니다:', err);
    const fallbackSheetId = 'sheet_default';
    workbookData.sheets[fallbackSheetId] = {
      id: fallbackSheetId,
      name: 'Sheet1',
      tabColor: '',
      hidden: 0,
      freeze: { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 },
      rowCount: 100,
      columnCount: 26,
      zoomRatio: 1,
      scrollTop: 0,
      scrollLeft: 0,
      defaultColumnWidth: 88,
      defaultRowHeight: 24,
      mergeData: [],
      cellData: {},
      rowData: {},
      columnData: {},
      rowHeader: { width: 46 },
      columnHeader: { height: 26 },
      showGridlines: 1,
      rightToLeft: 0,
    };
    workbookData.sheetOrder.push(fallbackSheetId);
  }

  return workbookData;
}

/**
 * Univer의 스냅샷(IWorkbookData)을 다시 XLSX 또는 CSV Data URL 문자열로 생성합니다.
 */
export function univerSnapshotToDataUrl(
  snapshot: IWorkbookData,
  format: 'xlsx' | 'csv' = 'xlsx'
): string {
  const wb = XLSX.utils.book_new();
  const sheetIds = snapshot.sheetOrder || Object.keys(snapshot.sheets || {});

  if (sheetIds.length === 0) {
    const emptyWs = XLSX.utils.aoa_to_sheet([['']]);
    XLSX.utils.book_append_sheet(wb, emptyWs, 'Sheet1');
  } else {
    sheetIds.forEach((sheetId) => {
      const sheet = snapshot.sheets[sheetId];
      if (!sheet) return;

      const ws: XLSX.WorkSheet = {};
      let maxR = 0;
      let maxC = 0;
      let hasData = false;

      const cellData = sheet.cellData || {};
      const rowKeys = Object.keys(cellData);

      for (let i = 0; i < rowKeys.length; i++) {
        const r = parseInt(rowKeys[i], 10);
        const rowObj = cellData[r];
        if (!rowObj) continue;

        const colKeys = Object.keys(rowObj);
        for (let j = 0; j < colKeys.length; j++) {
          const c = parseInt(colKeys[j], 10);
          const cell = rowObj[c];
          if (!cell) continue;

          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, c);
          hasData = true;

          const addr = XLSX.utils.encode_cell({ r, c });
          const val = cell.v;
          const formula = cell.f;

          if (formula) {
            ws[addr] = {
              t: typeof val === 'number' ? 'n' : 's',
              f: formula.replace(/^=/, ''),
              v: val !== undefined ? val : 0,
            };
          } else if (val !== undefined && val !== null && val !== '') {
            const t = typeof val === 'number' ? 'n' : typeof val === 'boolean' ? 'b' : 's';
            ws[addr] = { t, v: val };
          }
        }
      }

      if (sheet.mergeData && Array.isArray(sheet.mergeData) && sheet.mergeData.length > 0) {
        ws['!merges'] = sheet.mergeData.map((m) => ({
          s: { r: m.startRow, c: m.startColumn },
          e: { r: m.endRow, c: m.endColumn },
        }));
      }

      ws['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: hasData ? maxR : 0, c: hasData ? maxC : 0 },
      });

      const sheetTitle = (sheet.name || 'Sheet').slice(0, 31);
      XLSX.utils.book_append_sheet(wb, ws, sheetTitle);
    });
  }

  if (format === 'csv') {
    const firstSheetName = wb.SheetNames[0] || 'Sheet1';
    const csvStr = XLSX.utils.sheet_to_csv(wb.Sheets[firstSheetName]);
    const b64 = btoa(unescape(encodeURIComponent(csvStr)));
    return `data:text/csv;charset=utf-8;base64,${b64}`;
  }

  const base64Out = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
  return `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64Out}`;
}
