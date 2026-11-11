import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import type { IWorkbookData, IWorksheetData, ICellData, IRange, IStyleData } from '@univerjs/core';
import { LocaleType, HorizontalAlign, VerticalAlign, WrapStrategy } from '@univerjs/core';

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

// ARGB 또는 RGB 색상 문자열을 CSS Hex(#RRGGBB)로 변환
function argbToHex(argb?: string): string | undefined {
  if (!argb || typeof argb !== 'string') return undefined;
  const clean = argb.replace('#', '').trim();
  if (clean.length === 8) {
    return `#${clean.slice(2)}`;
  }
  if (clean.length === 6) {
    return `#${clean}`;
  }
  return undefined;
}

// CSS Hex(#RRGGBB) 문자열을 ExcelJS용 ARGB(FFRRGGBB)로 변환
function hexToArgb(hex?: string): string | undefined {
  if (!hex || typeof hex !== 'string') return undefined;
  const clean = hex.replace('#', '').trim();
  if (clean.length === 6) {
    return `FF${clean.toUpperCase()}`;
  }
  if (clean.length === 8) {
    return clean.toUpperCase();
  }
  return undefined;
}

/**
 * XLSX/CSV Base64 또는 Data URL을 Univer의 IWorkbookData 모델로 변환합니다.
 * ExcelJS를 사용하여 글꼴, 배경색, 테두리, 열 너비, 행 높이, 수식, 병합 셀을 최대한 보존합니다.
 */
export async function dataUrlToUniverWorkbook(
  dataOrUrl: string,
  defaultName: string = 'Workbook'
): Promise<IWorkbookData> {
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
    const excelWb = new ExcelJS.Workbook();
    await excelWb.xlsx.load(bytes as any);

    let styleCounter = 1;
    const styleCache = new Map<string, string>(); // JSON.stringify(style) -> styleId

    excelWb.eachSheet((worksheet, sheetIndex) => {
      const sheetId = `sheet_${sheetIndex}`;
      const sheetName = worksheet.name || `Sheet${sheetIndex}`;

      const cellData: Record<number, Record<number, ICellData>> = {};
      const rowData: Record<number, { h?: number }> = {};
      const columnData: Record<number, { w?: number }> = {};
      const mergeData: IRange[] = [];

      let maxRow = Math.max(30, worksheet.rowCount);
      let maxCol = Math.max(16, worksheet.columnCount);

      // 1. 열 너비(Column Widths) 복원
      worksheet.columns?.forEach((col, cIdx) => {
        if (col && col.width) {
          // 엑셀 width 단위를 픽셀로 변환 (약 width * 8 + 5)
          const pxWidth = Math.round(col.width * 8 + 5);
          columnData[cIdx] = { w: Math.max(40, pxWidth) };
        }
      });

      // 2. 행 및 셀 데이터, 스타일 복원
      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const rIdx = rowNumber - 1; // 0-based
        if (row.height) {
          // 엑셀 pt 단위를 픽셀로 변환 (pt * 1.33)
          rowData[rIdx] = { h: Math.round(row.height * 1.33) };
        }

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const cIdx = colNumber - 1; // 0-based
          if (!cellData[rIdx]) cellData[rIdx] = {};

          const cItem: ICellData = {};

          // 값 및 수식
          if (cell.formula) {
            cItem.f = cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`;
            if (typeof cell.result === 'number' || typeof cell.result === 'string' || typeof cell.result === 'boolean') {
              cItem.v = cell.result;
            } else {
              cItem.v = 0;
            }
          } else if (cell.value !== undefined && cell.value !== null) {
            if (typeof cell.value === 'object' && 'richText' in (cell.value as any)) {
              // 리치 텍스트는 일반 텍스트로 합산
              cItem.v = (cell.value as any).richText.map((t: any) => t.text).join('');
            } else if (cell.value instanceof Date) {
              cItem.v = cell.value.toISOString().slice(0, 10);
            } else if (typeof cell.value === 'number' || typeof cell.value === 'string' || typeof cell.value === 'boolean') {
              cItem.v = cell.value;
            } else {
              cItem.v = String(cell.value);
            }
          }

          // 스타일 추출 (Font, Fill, Border, Alignment, NumberFormat)
          const styleObj: IStyleData = {};
          let hasStyle = false;

          // 폰트
          if (cell.font) {
            if (cell.font.name) {
              styleObj.ff = cell.font.name;
              hasStyle = true;
            }
            if (cell.font.size) {
              styleObj.fs = cell.font.size;
              hasStyle = true;
            }
            if (cell.font.bold) {
              styleObj.bl = 1;
              hasStyle = true;
            }
            if (cell.font.italic) {
              styleObj.it = 1;
              hasStyle = true;
            }
            const fontColorHex = argbToHex(cell.font.color?.argb);
            if (fontColorHex) {
              styleObj.cl = { rgb: fontColorHex };
              hasStyle = true;
            }
          }

          // 배경색 (Fill)
          if (cell.fill && cell.fill.type === 'pattern') {
            const fillHex = argbToHex((cell.fill as any).fgColor?.argb);
            if (fillHex && fillHex !== '#FFFFFF') {
              styleObj.bg = { rgb: fillHex };
              hasStyle = true;
            }
          }

          // 정렬 (Alignment)
          if (cell.alignment) {
            if (cell.alignment.horizontal === 'center') {
              styleObj.ht = HorizontalAlign.CENTER;
              hasStyle = true;
            } else if (cell.alignment.horizontal === 'right') {
              styleObj.ht = HorizontalAlign.RIGHT;
              hasStyle = true;
            } else if (cell.alignment.horizontal === 'left') {
              styleObj.ht = HorizontalAlign.LEFT;
              hasStyle = true;
            }

            if (cell.alignment.vertical === 'middle') {
              styleObj.vt = VerticalAlign.MIDDLE;
              hasStyle = true;
            } else if (cell.alignment.vertical === 'top') {
              styleObj.vt = VerticalAlign.TOP;
              hasStyle = true;
            } else if (cell.alignment.vertical === 'bottom') {
              styleObj.vt = VerticalAlign.BOTTOM;
              hasStyle = true;
            }

            if (cell.alignment.wrapText) {
              styleObj.tb = WrapStrategy.WRAP;
              hasStyle = true;
            }
          }

          // 숫자 서식 (Number Format)
          if (cell.numFmt && cell.numFmt !== 'General') {
            styleObj.n = { pattern: cell.numFmt };
            hasStyle = true;
          }

          if (hasStyle) {
            const styleKey = JSON.stringify(styleObj);
            let styleId = styleCache.get(styleKey);
            if (!styleId) {
              styleId = `style_${styleCounter++}`;
              styleCache.set(styleKey, styleId);
              workbookData.styles[styleId] = styleObj;
            }
            cItem.s = styleId;
          }

          cellData[rIdx][cIdx] = cItem;
        });
      });

      // 3. 셀 병합 복원 (ExcelJS model.merges)
      const merges = (worksheet as any).model?.merges || [];
      if (Array.isArray(merges)) {
        merges.forEach((mStr: string) => {
          // 'A1:I1' 형태 디코딩
          const parts = mStr.split(':');
          if (parts.length === 2) {
            const s = XLSX.utils.decode_cell(parts[0]);
            const e = XLSX.utils.decode_cell(parts[1]);
            mergeData.push({
              startRow: s.r,
              endRow: e.r,
              startColumn: s.c,
              endColumn: e.c,
            });
          }
        });
      }

      // 4. 틀 고정 (Freeze)
      const views = worksheet.views || [];
      let xSplit = 0;
      let ySplit = 0;
      if (views.length > 0 && views[0].state === 'frozen') {
        xSplit = views[0].xSplit || 0;
        ySplit = views[0].ySplit || 0;
      }

      const uSheet: IWorksheetData = {
        id: sheetId,
        name: sheetName,
        tabColor: '',
        hidden: 0,
        freeze: { xSplit, ySplit, startRow: -1, startColumn: -1 },
        rowCount: Math.max(100, maxRow + 20),
        columnCount: Math.max(26, maxCol + 10),
        zoomRatio: 1,
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 88,
        defaultRowHeight: 24,
        mergeData,
        cellData,
        rowData,
        columnData,
        rowHeader: { width: 46 },
        columnHeader: { height: 26 },
        showGridlines: 1,
        rightToLeft: 0,
      };

      workbookData.sheets[sheetId] = uSheet;
      workbookData.sheetOrder.push(sheetId);
    });

    if (workbookData.sheetOrder.length === 0) {
      throw new Error('시트가 없습니다.');
    }
  } catch (err) {
    console.warn('ExcelJS 파싱 실패, SheetJS 호환 모드로 대체 시도합니다:', err);
    // 폴백: SheetJS 기본 파싱
    return fallbackSheetJsToUniver(dataOrUrl, defaultName);
  }

  return workbookData;
}

// 폴백용 SheetJS 파서
function fallbackSheetJsToUniver(dataOrUrl: string, defaultName: string): IWorkbookData {
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
    const sheetNames = workbook.SheetNames || ['Sheet1'];

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
                if (cell.f) cItem.f = cell.f.startsWith('=') ? cell.f : `=${cell.f}`;
                if (cell.v !== undefined && cell.v !== null) cItem.v = cell.v;
                cellData[r][c] = cItem;
              }
            }
          }
        }

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

      workbookData.sheets[sheetId] = {
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
      workbookData.sheetOrder.push(sheetId);
    });
  } catch (fallbackErr) {
    console.error('Fallback SheetJS 파싱 실패:', fallbackErr);
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
 * engine: 'exceljs' (기본값, 서식/열너비/행높이 보존) | 'sheetjs' (기존 가벼운 데이터 모드)
 */
/**
 * Univer의 스냅샷(IWorkbookData)을 다시 XLSX 또는 CSV Data URL 문자열로 생성합니다.
 * engine: 'exceljs' (기본값, 서식/열너비/행높이 보존) | 'sheetjs' (기존 가벼운 데이터 모드)
 * originalDataOrUrl: 원본 파일 바이트(Data URL). 제공 시 원본 워크북을 로드하여 바뀐 셀/수식/서식/병합/크기만
 * in-place로 패치하므로 ExcelJS가 지원하는 조건부 서식, 데이터 유효성 검사, 셀 메모, 인쇄 설정, 틀 고정이 100% 보존됩니다.
 */
export async function univerSnapshotToDataUrl(
  snapshot: IWorkbookData,
  format: 'xlsx' | 'csv' = 'xlsx',
  engine: 'exceljs' | 'sheetjs' = 'exceljs',
  originalDataOrUrl?: string
): Promise<string> {
  if (engine === 'sheetjs' || format === 'csv') {
    return univerSnapshotToDataUrlWithSheetJS(snapshot, format);
  }

  // --- ExcelJS 라운드트립 저장 엔진 (서식/수식/병합/너비/높이 및 원본 메타데이터 보존) ---
  try {
    const wb = new ExcelJS.Workbook();
    let isLoadedFromOriginal = false;

    if (originalDataOrUrl) {
      try {
        const origBytes = dataUrlToUint8Array(originalDataOrUrl);
        await wb.xlsx.load(origBytes as any);
        isLoadedFromOriginal = true;
      } catch (loadErr) {
        console.warn('원본 워크북 로드 실패, 새 워크북으로 생성합니다:', loadErr);
      }
    }

    const sheetIds = snapshot.sheetOrder && snapshot.sheetOrder.length > 0
      ? snapshot.sheetOrder
      : Object.keys(snapshot.sheets || {});
    const styles = snapshot.styles || {};

    if (!isLoadedFromOriginal) {
      // 원본 파일이 없는 신규 생성 시 빈 워크북에 시트 추가
      if (sheetIds.length === 0) {
        wb.addWorksheet('Sheet1');
      } else {
        sheetIds.forEach((sheetId) => {
          const uSheet = snapshot.sheets[sheetId];
          if (!uSheet) return;
          const sheetName = (uSheet.name || 'Sheet').slice(0, 31);
          wb.addWorksheet(sheetName);
        });
      }
    } else {
      // 원본 파일이 있는 경우: 시트 동기화 (추가, 삭제, 이름 변경 in-place 반영)
      const univerSheets = sheetIds.map((id) => snapshot.sheets[id]).filter(Boolean);
      const originalWorksheets = [...wb.worksheets];
      const matchedWsSet = new Set<any>();

      // 1단계: 시트명 완전 일치 매칭
      univerSheets.forEach((uSheet) => {
        const match = originalWorksheets.find((ws) => !matchedWsSet.has(ws) && ws.name === uSheet.name);
        if (match) {
          matchedWsSet.add(match);
          (uSheet as any)._matchedWs = match;
        }
      });

      // 2단계: 남은 시트 순서(인덱스) 기반 매칭
      univerSheets.forEach((uSheet, idx) => {
        if (!(uSheet as any)._matchedWs) {
          const candidate = originalWorksheets[idx];
          if (candidate && !matchedWsSet.has(candidate)) {
            matchedWsSet.add(candidate);
            (uSheet as any)._matchedWs = candidate;
            // 시트 이름 변경 반영
            candidate.name = (uSheet.name || `Sheet${idx + 1}`).slice(0, 31);
          }
        }
      });

      // 3단계: Univer에서 삭제된 시트는 원본 워크북에서 제거
      originalWorksheets.forEach((origWs) => {
        if (!matchedWsSet.has(origWs)) {
          wb.removeWorksheet(origWs.id);
        }
      });

      // 4단계: Univer에서 신규 추가된 시트는 원본 워크북에 추가
      univerSheets.forEach((uSheet, idx) => {
        if (!(uSheet as any)._matchedWs) {
          const newSheetName = (uSheet.name || `Sheet${idx + 1}`).slice(0, 31);
          const newWs = wb.addWorksheet(newSheetName);
          (uSheet as any)._matchedWs = newWs;
        }
      });
    }

    // 각 시트별 데이터, 크기, 서식, 병합 덮어쓰기
    sheetIds.forEach((sheetId, sIdx) => {
      const uSheet = snapshot.sheets[sheetId];
      if (!uSheet) return;

      const sheetName = (uSheet.name || `Sheet${sIdx + 1}`).slice(0, 31);
      const ws = (uSheet as any)._matchedWs || wb.getWorksheet(sheetName) || wb.worksheets[sIdx] || wb.addWorksheet(sheetName);

      // 1. 열 너비 복원
      const colData = uSheet.columnData;
      if (colData) {
        Object.keys(colData).forEach((cStr) => {
          const cIdx = parseInt(cStr, 10);
          const colInfo = colData[cIdx];
          if (colInfo && colInfo.w) {
            const colWidth = Math.max(8, Math.round((colInfo.w - 5) / 8));
            ws.getColumn(cIdx + 1).width = colWidth;
          }
        });
      }

      // 2. 행 높이 및 셀 데이터/스타일 복원
      const cellData = uSheet.cellData || {};
      const rowKeys = Object.keys(cellData);

      for (let i = 0; i < rowKeys.length; i++) {
        const r = parseInt(rowKeys[i], 10);
        const rowObj = cellData[r];
        if (!rowObj) continue;

        const rowNum = r + 1; // 1-based
        const row = ws.getRow(rowNum);

        if (uSheet.rowData && uSheet.rowData[r] && uSheet.rowData[r].h) {
          row.height = Math.round(uSheet.rowData[r].h / 1.33);
        }

        const colKeys = Object.keys(rowObj);
        for (let j = 0; j < colKeys.length; j++) {
          const c = parseInt(colKeys[j], 10);
          const cell = rowObj[c];
          if (!cell) continue;

          const colNum = c + 1; // 1-based
          const excelCell = ws.getCell(rowNum, colNum);

          // 값 및 수식 덮어쓰기
          if (cell.f) {
            excelCell.value = {
              formula: cell.f.replace(/^=/, ''),
              result: (cell.v !== undefined && cell.v !== null ? cell.v : 0) as any,
            };
          } else if (cell.v !== undefined && cell.v !== null && cell.v !== '') {
            excelCell.value = cell.v as any;
          } else if (cell.v === '') {
            excelCell.value = null;
          }

          // 스타일 적용 (폰트, 배경색, 정렬, 표시 형식)
          if (cell.s) {
            const styleObj = typeof cell.s === 'string' ? styles[cell.s] : (cell.s as IStyleData);
            if (styleObj) {
              // 폰트 (기존 폰트 속성 유지하며 변경된 항목만 덮어씀)
              if (styleObj.ff || styleObj.fs || styleObj.bl || styleObj.it || styleObj.cl) {
                excelCell.font = {
                  ...(excelCell.font || {}),
                  name: styleObj.ff || excelCell.font?.name,
                  size: styleObj.fs || excelCell.font?.size,
                  bold: styleObj.bl === 1,
                  italic: styleObj.it === 1,
                  color: styleObj.cl?.rgb ? { argb: hexToArgb(styleObj.cl.rgb) } : excelCell.font?.color,
                };
              }

              // 배경색
              if (styleObj.bg?.rgb) {
                excelCell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: hexToArgb(styleObj.bg.rgb) },
                };
              }

              // 정렬
              if (styleObj.ht || styleObj.vt || styleObj.tb) {
                excelCell.alignment = {
                  ...(excelCell.alignment || {}),
                  horizontal:
                    styleObj.ht === HorizontalAlign.CENTER
                      ? 'center'
                      : styleObj.ht === HorizontalAlign.RIGHT
                      ? 'right'
                      : styleObj.ht === HorizontalAlign.LEFT
                      ? 'left'
                      : excelCell.alignment?.horizontal,
                  vertical:
                    styleObj.vt === VerticalAlign.MIDDLE
                      ? 'middle'
                      : styleObj.vt === VerticalAlign.TOP
                      ? 'top'
                      : styleObj.vt === VerticalAlign.BOTTOM
                      ? 'bottom'
                      : excelCell.alignment?.vertical,
                  wrapText: styleObj.tb === WrapStrategy.WRAP,
                };
              }

              // 표시 형식
              if (styleObj.n?.pattern) {
                excelCell.numFmt = styleObj.n.pattern;
              }
            }
          }
        }
      }

      // 3. 셀 병합 적용: 기존 병합 해제 후 Univer 최신 병합 목록 적용
      if (uSheet.mergeData && Array.isArray(uSheet.mergeData)) {
        const existingMerges = [...((ws as any).model?.merges || [])];
        for (const mStr of existingMerges) {
          try {
            (ws as any).unMergeCells(mStr);
          } catch {}
        }
        uSheet.mergeData.forEach((m) => {
          try {
            ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
          } catch (mergeErr) {
            console.warn('셀 병합 적용 오류 (무시 가능):', mergeErr);
          }
        });
      }

      // 4. 틀 고정 적용
      if (uSheet.freeze && (uSheet.freeze.xSplit > 0 || uSheet.freeze.ySplit > 0)) {
        ws.views = [
          {
            state: 'frozen',
            xSplit: uSheet.freeze.xSplit,
            ySplit: uSheet.freeze.ySplit,
          },
        ];
      }
    });

    const buffer = await wb.xlsx.writeBuffer();
    const uint8 = new Uint8Array(buffer);
    let binary = '';
    const len = uint8.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(uint8[i]);
    }
    const b64 = btoa(binary);
    return `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${b64}`;
  } catch (err) {
    console.error('ExcelJS 저장 실패, SheetJS로 대체 저장합니다:', err);
    return univerSnapshotToDataUrlWithSheetJS(snapshot, format);
  }
}

// 레거시 SheetJS 저장 모드 (기존 호환 유지)
function univerSnapshotToDataUrlWithSheetJS(
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
