import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
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
 * 한셀(Hancom HanCell) 등 비표준 오피스에서 생성된 XLSX 파일의 XML 네임스페이스 및 비표준 태그를 정규화합니다.
 * ExcelJS가 오류 없이 파싱할 수 있도록 x:, ep: 접두어 제거, mc:AlternateContent Fallback 추출,
 * ref 단일 셀 범위(F4:F4 -> F4), boolean 태그(<b val="1"/> -> <b/>)를 표준 OpenXML 형식으로 변환합니다.
 */
export async function normalizeHancellXlsx(bytesOrDataUrl: Uint8Array | string): Promise<Uint8Array> {
  const bytes = typeof bytesOrDataUrl === 'string' ? dataUrlToUint8Array(bytesOrDataUrl) : bytesOrDataUrl;

  try {
    const zip = await JSZip.loadAsync(bytes);
    let isModified = false;

    for (const [filename, file] of Object.entries(zip.files)) {
      if (file.dir) continue;
      if (!filename.endsWith('.xml') && !filename.endsWith('.rels') && !filename.endsWith('.vml')) continue;

      let content = await file.async('text');
      let mod = false;

      // 1. docProps/app.xml: ep: 접두어 및 Company 태그 보정
      if (filename === 'docProps/app.xml') {
        content = content.replace(
          /xmlns:ep="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/extended-properties"/g,
          'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
        );
        content = content.replace(/<\/?ep:([a-zA-Z0-9_]+)/g, (match, tag) => (match.startsWith('</') ? `</${tag}` : `<${tag}`));
        if (!content.includes('<Company')) {
          content = content.replace(/<\/Properties>/, '<Company></Company></Properties>');
        }
        mod = true;
      }

      // 2. xl/ 내부의 모든 xml/rels 파일 (styles, comments, worksheets, workbook, sharedStrings 등) - .vml 파일 제외
      if (filename.startsWith('xl/') && (filename.endsWith('.xml') || filename.endsWith('.rels'))) {
        // mc:AlternateContent Fallback 우선 채택 (한셀 xf 확장태그 호환)
        if (content.includes('<mc:AlternateContent')) {
          content = content.replace(
            /<mc:AlternateContent[^>]*>[\s\S]*?<mc:Fallback>([\s\S]*?)<\/mc:Fallback>[\s\S]*?<\/mc:AlternateContent>/g,
            '$1'
          );
          content = content.replace(
            /<mc:AlternateContent[^>]*>[\s\S]*?<mc:Choice[^>]*>([\s\S]*?)<\/mc:Choice>[\s\S]*?<\/mc:AlternateContent>/g,
            '$1'
          );
          mod = true;
        }

        // 한셀 전용 확장 태그 및 속성 제거 (hs:parashape 등)
        if (content.includes('hs:')) {
          content = content.replace(/<hs:[^>]*\/>/g, '');
          content = content.replace(/<hs:[^>]*>[\s\S]*?<\/hs:[^>]*>/g, '');
          content = content.replace(/\s+hs:[a-zA-Z0-9_]+="[^"]*"/g, '');
          mod = true;
        }

        // spreadsheetml 네임스페이스 및 x: 접두어 정규화
        if (content.includes('xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"') || content.includes('<x:')) {
          content = content.replace(
            /xmlns:x="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/g,
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
          );
          content = content.replace(/<\/?x:([a-zA-Z0-9_]+)/g, (match, tag) => (match.startsWith('</') ? `</${tag}` : `<${tag}`));
          mod = true;
        }

        // boolean 서식 속성 정규화 (<b val="1"/> -> <b/>, <b val="0"/> -> 삭제)
        if (content.includes('val="1"') || content.includes('val="0"') || content.includes('val="true"') || content.includes('val="false"')) {
          content = content.replace(/<b\s+val="1"\s*\/>/g, '<b/>');
          content = content.replace(/<b\s+val="true"\s*\/>/g, '<b/>');
          content = content.replace(/<b\s+val="0"\s*\/>/g, '');
          content = content.replace(/<b\s+val="false"\s*\/>/g, '');

          content = content.replace(/<i\s+val="1"\s*\/>/g, '<i/>');
          content = content.replace(/<i\s+val="true"\s*\/>/g, '<i/>');
          content = content.replace(/<i\s+val="0"\s*\/>/g, '');
          content = content.replace(/<i\s+val="false"\s*\/>/g, '');

          content = content.replace(/<u\s+val="1"\s*\/>/g, '<u/>');
          content = content.replace(/<u\s+val="0"\s*\/>/g, '');

          content = content.replace(/<strike\s+val="1"\s*\/>/g, '<strike/>');
          content = content.replace(/<strike\s+val="0"\s*\/>/g, '');
          mod = true;
        }

        // 단일 셀 범위 정규화 ref="F4:F4" -> ref="F4"
        if (content.includes('ref=')) {
          content = content.replace(/ref="([A-Za-z]+[0-9]+):\1"/g, 'ref="$1"');
          mod = true;
        }
      }

      if (mod) {
        zip.file(filename, content);
        isModified = true;
      }
    }

    if (isModified) {
      return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    }
  } catch (err) {
    console.warn('한셀/OpenXML 네임스페이스 정규화 중 예외 (원본 바이너리 유지):', err);
  }

  return bytes;
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
    const rawBytes = dataUrlToUint8Array(dataOrUrl);
    const bytes = await normalizeHancellXlsx(rawBytes);
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
          columnData[cIdx] = { w: Math.max(10, pxWidth) };
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

          // 머지 슬레이브 셀(isMerged && master.address !== address)인 경우:
          // Univer에서 마스터 셀만 값을 가져야 하므로 슬레이브 셀에는 값이나 수식을 채우지 않음 (스타일만 보존)
          const isSlaveCell = Boolean(cell.isMerged && cell.master && cell.master.address !== cell.address);

          if (!isSlaveCell) {
            // 값 및 수식
            if (cell.formula) {
              cItem.f = cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`;
              if (typeof cell.result === 'number' || typeof cell.result === 'string' || typeof cell.result === 'boolean') {
                cItem.v = cell.result;
              } else if (cell.result instanceof Date) {
                cItem.v = cell.result.toISOString().slice(0, 10);
              } else {
                cItem.v = 0;
              }
            } else if (cell.value !== undefined && cell.value !== null) {
              if (typeof cell.value === 'object' && 'richText' in (cell.value as any)) {
                // 리치 텍스트는 일반 텍스트로 합산
                cItem.v = (cell.value as any).richText.map((t: any) => t.text).join('');
              } else if (typeof cell.value === 'object' && 'hyperlink' in (cell.value as any)) {
                // 하이퍼링크 객체는 표시 텍스트 또는 링크 URL 추출
                cItem.v = (cell.value as any).text || (cell.value as any).hyperlink;
              } else if (cell.value instanceof Date) {
                cItem.v = cell.value.toISOString().slice(0, 10);
              } else if (typeof cell.value === 'number' || typeof cell.value === 'string' || typeof cell.value === 'boolean') {
                cItem.v = cell.value;
              } else if (typeof cell.value === 'object' && 'formula' in (cell.value as any)) {
                const fObj = cell.value as any;
                cItem.f = fObj.formula.startsWith('=') ? fObj.formula : `=${fObj.formula}`;
                if (fObj.result instanceof Date) {
                  cItem.v = fObj.result.toISOString().slice(0, 10);
                } else if (typeof fObj.result === 'number' || typeof fObj.result === 'string' || typeof fObj.result === 'boolean') {
                  cItem.v = fObj.result;
                } else {
                  cItem.v = 0;
                }
              } else {
                cItem.v = String(cell.value);
              }
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
  (workbookData as any).isFallbackMode = true;

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
export interface StructuralOperation {
  type: 'insertRow' | 'removeRow' | 'insertCol' | 'removeCol';
  sheetName?: string;
  subUnitId?: string;
  sheetIndex?: number;
  start: number; // 1-based index
  count: number;
}

// 엑셀 열 문자(A, B, ..., Z, AA, ...)와 1-based 인덱스 변환
function colToNum(colStr: string): number {
  let num = 0;
  for (let i = 0; i < colStr.length; i++) {
    num = num * 26 + (colStr.charCodeAt(i) - 64);
  }
  return num;
}

function numToCol(num: number): string {
  let s = '';
  while (num > 0) {
    const m = (num - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    num = Math.floor((num - m) / 26);
  }
  return s;
}

// 조건부 서식 참조 범위('A2:B5' 등) 행 번호 시프트
function shiftRangeRows(ref: string, startRow: number, count: number): string {
  try {
    const parts = ref.split(':');
    const shiftedParts = parts.map((cellStr) => {
      const match = cellStr.match(/^([A-Za-z]+)(\d+)$/);
      if (!match) return cellStr;
      const col = match[1];
      const row = parseInt(match[2], 10);
      if (row >= startRow) {
        const newRow = Math.max(1, row + count);
        return `${col}${newRow}`;
      }
      return cellStr;
    });
    return shiftedParts.join(':');
  } catch {
    return ref;
  }
}

// 조건부 서식 참조 범위('A2:B5' 등) 열 번호 시프트
function shiftRangeCols(ref: string, startCol: number, count: number): string {
  try {
    const parts = ref.split(':');
    const shiftedParts = parts.map((cellStr) => {
      const match = cellStr.match(/^([A-Za-z]+)(\d+)$/);
      if (!match) return cellStr;
      const colLetter = match[1].toUpperCase();
      const rowNum = match[2];
      const colNum = colToNum(colLetter);
      if (colNum >= startCol) {
        const newColNum = Math.max(1, colNum + count);
        return `${numToCol(newColNum)}${rowNum}`;
      }
      return cellStr;
    });
    return shiftedParts.join(':');
  } catch {
    return ref;
  }
}

// 셀 주소 기준 데이터 유효성 검사 행 시프트 함수 (배열 분할/타입누락 방지)
function shiftDataValidations(model: any, start: number, count: number): any {
  const newModel: any = {};
  if (!model) return newModel;

  const colRules: Record<string, Array<{ row: number; rule: any }>> = {};
  for (const [addr, rule] of Object.entries(model)) {
    if (!rule || !(rule as any).type) continue;
    const match = addr.match(/^([A-Z]+)([0-9]+)$/);
    if (!match) continue;
    const col = match[1];
    const row = parseInt(match[2], 10);
    if (!colRules[col]) colRules[col] = [];
    colRules[col].push({ row, rule });
  }

  for (const [col, items] of Object.entries(colRules)) {
    items.sort((a, b) => a.row - b.row);

    if (count > 0) {
      const prevRule = items.find((it) => it.row === start - 1)?.rule;
      const nextRule = items.find((it) => it.row === start)?.rule;
      const shouldFill = prevRule && nextRule && prevRule.type === nextRule.type;

      for (const it of items) {
        if (it.row < start) {
          newModel[`${col}${it.row}`] = it.rule;
        } else {
          newModel[`${col}${it.row + count}`] = it.rule;
        }
      }

      if (shouldFill) {
        for (let r = 0; r < count; r++) {
          newModel[`${col}${start + r}`] = prevRule;
        }
      }
    } else if (count < 0) {
      const delCount = -count;
      const delEnd = start + delCount - 1;

      for (const it of items) {
        if (it.row < start) {
          newModel[`${col}${it.row}`] = it.rule;
        } else if (it.row > delEnd) {
          newModel[`${col}${it.row - delCount}`] = it.rule;
        }
      }
    } else {
      for (const it of items) {
        newModel[`${col}${it.row}`] = it.rule;
      }
    }
  }

  return newModel;
}

// 셀 주소 기준 데이터 유효성 검사 열 시프트 함수
function shiftDataValidationsCols(model: any, startCol: number, count: number): any {
  const newModel: any = {};
  if (!model) return newModel;

  const rowRules: Record<number, Array<{ col: number; rule: any }>> = {};
  for (const [addr, rule] of Object.entries(model)) {
    if (!rule || !(rule as any).type) continue;
    const match = addr.match(/^([A-Z]+)([0-9]+)$/);
    if (!match) continue;
    const col = colToNum(match[1]);
    const row = parseInt(match[2], 10);
    if (!rowRules[row]) rowRules[row] = [];
    rowRules[row].push({ col, rule });
  }

  for (const [rowStr, items] of Object.entries(rowRules)) {
    const row = parseInt(rowStr, 10);
    items.sort((a, b) => a.col - b.col);

    if (count > 0) {
      const prevRule = items.find((it) => it.col === startCol - 1)?.rule;
      const nextRule = items.find((it) => it.col === startCol)?.rule;
      const shouldFill = prevRule && nextRule && prevRule.type === nextRule.type;

      for (const it of items) {
        if (it.col < startCol) {
          newModel[`${numToCol(it.col)}${row}`] = it.rule;
        } else {
          newModel[`${numToCol(it.col + count)}${row}`] = it.rule;
        }
      }

      if (shouldFill) {
        for (let c = 0; c < count; c++) {
          newModel[`${numToCol(startCol + c)}${row}`] = prevRule;
        }
      }
    } else if (count < 0) {
      const delCount = -count;
      const delEnd = startCol + delCount - 1;

      for (const it of items) {
        if (it.col < startCol) {
          newModel[`${numToCol(it.col)}${row}`] = it.rule;
        } else if (it.col > delEnd) {
          newModel[`${numToCol(it.col - delCount)}${row}`] = it.rule;
        }
      }
    } else {
      for (const it of items) {
        newModel[`${numToCol(it.col)}${row}`] = it.rule;
      }
    }
  }

  return newModel;
}

// 자동 필터(autoFilter) 및 _xlnm._FilterDatabase 이름 정의 범위 행 시프트 함수
function shiftAutoFilter(ws: ExcelJS.Worksheet, start: number, nExpand: number) {
  if (!ws.autoFilter || nExpand === 0) return;
  const afStr = typeof ws.autoFilter === 'string' ? ws.autoFilter : (ws.autoFilter as any).ref;
  if (!afStr) return;
  const m = afStr.match(/^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/);
  if (!m) return;
  const c1 = m[1], r1 = parseInt(m[2], 10);
  const c2 = m[3], r2 = parseInt(m[4], 10);

  // 헤더 다음 행(r1 < start)부터 데이터 끝 및 직후 행(start <= r2 + 1) 사이 삽입/삭제 시 범위 조정
  if (start > r1 && start <= r2 + (nExpand > 0 ? 1 : 0)) {
    const newR2 = Math.max(r1, r2 + nExpand);
    const newRef = `${c1}${r1}:${c2}${newR2}`;
    ws.autoFilter = newRef;

    // 워크북 definedNames _xlnm._FilterDatabase 동기화 (단일 범위만 유지)
    const wb = ws.workbook;
    if (wb && (wb as any).definedNames && Array.isArray((wb as any).definedNames.model)) {
      (wb as any).definedNames.model.forEach((entry: any) => {
        if (entry.name === '_xlnm._FilterDatabase') {
          const sheetPrefix = (Array.isArray(entry.ranges) && entry.ranges[0]?.match(/^(.*!)/)?.[1]) || `'${ws.name}'!`;
          entry.ranges = [`${sheetPrefix}$${c1}$${r1}:$${c2}$${newR2}`];
        }
      });
    }
  }
}

// 셀 메모, 데이터 유효성 검사, 조건부 서식 위치를 함께 보존·이동시키는 행 삽입/삭제 함수
function safeSpliceRows(ws: ExcelJS.Worksheet, start: number, count: number, ...inserts: any[]) {
  const nInserts = inserts.length;
  const nExpand = nInserts - count;
  const nKeep = start + count;
  const nEnd = (ws as any)._rows?.length || ws.rowCount;

  // 1. 데이터 유효성 검사 범위 이동 (배열 쪼개짐/타입 누락 방지)
  if ((ws as any).dataValidations && (ws as any).dataValidations.model) {
    (ws as any).dataValidations.model = shiftDataValidations((ws as any).dataValidations.model, start, nExpand);
  }

  // 1-1. 자동 필터(autoFilter) 및 _FilterDatabase 범위 확장/축소
  shiftAutoFilter(ws, start, nExpand);

  // 2. 셀 메모(Note) 이동
  if (nExpand > 0) {
    const rowMeta: Array<{ toRow: number; notes: any[] }> = [];
    for (let i = nKeep; i <= nEnd; i++) {
      const r = (ws as any)._rows?.[i - 1];
      if (r) {
        const notes: any[] = [];
        r.eachCell({ includeEmpty: true }, (cell: any, cNum: number) => {
          if (cell.note) {
            notes.push({ cNum, note: cell.note });
            cell.note = undefined;
            delete (cell as any)._comment;
          }
        });
        if (notes.length > 0) rowMeta.push({ toRow: i + nExpand, notes });
      }
    }

    ws.spliceRows(start, count, ...inserts);

    rowMeta.forEach(({ toRow, notes }) => {
      const dstRow = ws.getRow(toRow);
      notes.forEach(({ cNum, note }) => {
        dstRow.getCell(cNum).note = note;
      });
    });

    if ((ws as any).conditionalFormattings) {
      (ws as any).conditionalFormattings.forEach((cf: any) => {
        if (cf.ref) cf.ref = shiftRangeRows(cf.ref, start, nExpand);
      });
    }
  } else if (nExpand < 0) {
    const delCount = -nExpand;
    const rowMeta: Array<{ toRow: number; notes: any[] }> = [];
    for (let i = nKeep; i <= nEnd; i++) {
      const r = (ws as any)._rows?.[i - 1];
      if (r) {
        const notes: any[] = [];
        r.eachCell({ includeEmpty: true }, (cell: any, cNum: number) => {
          if (cell.note) {
            notes.push({ cNum, note: cell.note });
            cell.note = undefined;
            delete (cell as any)._comment;
          }
        });
        if (notes.length > 0) rowMeta.push({ toRow: i - delCount, notes });
      }
    }

    ws.spliceRows(start, count, ...inserts);

    rowMeta.forEach(({ toRow, notes }) => {
      const dstRow = ws.getRow(toRow);
      notes.forEach(({ cNum, note }) => {
        dstRow.getCell(cNum).note = note;
      });
    });

    if ((ws as any).conditionalFormattings) {
      (ws as any).conditionalFormattings.forEach((cf: any) => {
        if (cf.ref) cf.ref = shiftRangeRows(cf.ref, start, -delCount);
      });
    }
  } else {
    ws.spliceRows(start, count, ...inserts);
  }
}

// 셀 메모, 데이터 유효성 검사, 조건부 서식 위치를 함께 보존·이동시키는 열 삽입/삭제 함수
function safeSpliceColumns(ws: ExcelJS.Worksheet, start: number, count: number, ...inserts: any[]) {
  const nInserts = inserts.length;
  const nExpand = nInserts - count;
  const nKeep = start + count;

  // 1. 데이터 유효성 검사 범위 이동
  if ((ws as any).dataValidations && (ws as any).dataValidations.model) {
    (ws as any).dataValidations.model = shiftDataValidationsCols((ws as any).dataValidations.model, start, nExpand);
  }

  // 2. 셀 메모 이동
  if (nExpand > 0) {
    const colMeta: any[] = [];
    ws.eachRow({ includeEmpty: true }, (row, rNum) => {
      row.eachCell({ includeEmpty: true }, (cell: any, cNum: number) => {
        if (cNum >= start && cell.note) {
          colMeta.push({
            rNum,
            toCol: cNum + nExpand,
            note: cell.note,
          });
          cell.note = undefined;
          delete (cell as any)._comment;
        }
      });
    });

    ws.spliceColumns(start, count, ...inserts);

    colMeta.forEach(({ rNum, toCol, note }) => {
      const cell = ws.getCell(rNum, toCol);
      if (note) cell.note = note;
    });

    if ((ws as any).conditionalFormattings) {
      (ws as any).conditionalFormattings.forEach((cf: any) => {
        if (cf.ref) cf.ref = shiftRangeCols(cf.ref, start, nExpand);
      });
    }
  } else if (nExpand < 0) {
    const delCount = -nExpand;
    const colMeta: any[] = [];
    ws.eachRow({ includeEmpty: true }, (row, rNum) => {
      row.eachCell({ includeEmpty: true }, (cell: any, cNum: number) => {
        if (cNum >= nKeep && cell.note) {
          colMeta.push({
            rNum,
            toCol: cNum - delCount,
            note: cell.note,
          });
        }
      });
    });

    ws.spliceColumns(start, count, ...inserts);

    colMeta.forEach(({ rNum, toCol, note }) => {
      const cell = ws.getCell(rNum, toCol);
      if (note) cell.note = note;
    });

    if ((ws as any).conditionalFormattings) {
      (ws as any).conditionalFormattings.forEach((cf: any) => {
        if (cf.ref) cf.ref = shiftRangeCols(cf.ref, start, -delCount);
      });
    }
  } else {
    ws.spliceColumns(start, count, ...inserts);
  }
}

// 유효성 검사 범위(sqref) 파싱 및 중복 제거 헬퍼
function parseSqrefRange(sqref: string): { col1: number; row1: number; col2: number; row2: number } | null {
  const parts = sqref.split(':');
  const m1 = parts[0]?.match(/^([A-Za-z]+)(\d+)$/);
  if (!m1) return null;
  const col1 = colToNum(m1[1].toUpperCase());
  const row1 = parseInt(m1[2], 10);
  if (parts.length === 1) {
    return { col1, row1, col2: col1, row2: row1 };
  }
  const m2 = parts[1]?.match(/^([A-Za-z]+)(\d+)$/);
  if (!m2) return null;
  const col2 = colToNum(m2[1].toUpperCase());
  const row2 = parseInt(m2[2], 10);
  return {
    col1: Math.min(col1, col2),
    row1: Math.min(row1, row2),
    col2: Math.max(col1, col2),
    row2: Math.max(row1, row2),
  };
}

function isSqrefSubRange(sub: string, parent: string): boolean {
  const rSub = parseSqrefRange(sub);
  const rParent = parseSqrefRange(parent);
  if (!rSub || !rParent) return false;
  return (
    rSub.col1 >= rParent.col1 &&
    rSub.col2 <= rParent.col2 &&
    rSub.row1 >= rParent.row1 &&
    rSub.row2 <= rParent.row2
  );
}

// ExcelJS 4.4.0의 optimiseDataValidations 문자열 정렬 버그로 인해 분할/중복 생성된 dataValidation 태그 정리
function deduplicateDataValidationsXml(xml: string): string {
  const dvBlockMatch = xml.match(/<dataValidations[^>]*>([\s\S]*?)<\/dataValidations>/);
  if (!dvBlockMatch) return xml;

  const fullBlock = dvBlockMatch[0];
  const inner = dvBlockMatch[1];
  const dvTagRegex = /<dataValidation\b[\s\S]*?<\/dataValidation>/g;
  const tags = inner.match(dvTagRegex) || [];

  const parsed = tags.map((tag) => {
    const sqrefMatch = tag.match(/sqref="([^"]+)"/);
    const sqref = sqrefMatch ? sqrefMatch[1] : '';
    // sqref 속성만 제외하고 규칙 내용(type, allowBlank, formulae 등) 동일성 비교
    const ruleKey = tag.replace(/\s*sqref="[^"]+"/, '').trim();
    return { tag, sqref, ruleKey };
  });

  const kept: string[] = [];
  for (let i = 0; i < parsed.length; i++) {
    const current = parsed[i];
    let isRedundant = false;
    for (let j = 0; j < parsed.length; j++) {
      if (i === j) continue;
      const other = parsed[j];
      if (current.ruleKey === other.ruleKey && isSqrefSubRange(current.sqref, other.sqref)) {
        if (current.sqref !== other.sqref || i > j) {
          isRedundant = true;
          break;
        }
      }
    }
    if (!isRedundant) {
      kept.push(current.tag);
    }
  }

  if (kept.length === 0) return xml.replace(fullBlock, '');
  const newBlock = `<dataValidations count="${kept.length}">${kept.join('')}</dataValidations>`;
  return xml.replace(fullBlock, newBlock);
}

interface OriginalRowMeta {
  hasCustomHeight: boolean;
  thickBot: boolean;
  thickTop: boolean;
}

interface DefaultFontInfo {
  name: string;
  sz?: string | null;
  charset?: string | null;
  family?: string | null;
  scheme?: string | null;
  isPyoJun?: boolean;
  normalAlignment?: string | null;
  sheetRowAttrs?: Map<string, Map<number, OriginalRowMeta>>;
}

async function extractOriginalDocDefaults(buffer: Uint8Array): Promise<DefaultFontInfo | null> {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const stylesFile = zip.file('xl/styles.xml');
    let name: string | null = null;
    let sz: string | null = null;
    let charset: string | null = null;
    let family: string | null = null;
    let scheme: string | null = null;
    let isPyoJun = false;
    let normalAlignment: string | null = null;

    if (stylesFile) {
      let xml = await stylesFile.async('string');
      // x: 네임스페이스 접두어 정규화
      xml = xml.replace(/<x:font\b/g, '<font').replace(/<\/x:font>/g, '</font>')
               .replace(/<x:([A-Za-z0-9_]+)/g, '<$1').replace(/<\/x:([A-Za-z0-9_]+)>/g, '</$1>');
      const fontsBlock = xml.match(/<fonts[^>]*>([\s\S]*?)<\/fonts>/);
      if (fontsBlock) {
        const firstFont = fontsBlock[1].match(/<font\b[\s\S]*?<\/font>/)?.[0];
        if (firstFont) {
          name = firstFont.match(/<name\s+val="([^"]+)"/)?.[1] || null;
          sz = firstFont.match(/<sz\s+val="([^"]+)"/)?.[1] || null;
          charset = firstFont.match(/<charset\s+val="([^"]+)"/)?.[1] || null;
          family = firstFont.match(/<family\s+val="([^"]+)"/)?.[1] || null;
          scheme = firstFont.match(/<scheme\s+val="([^"]+)"/)?.[1] || null;
        }
      }
      isPyoJun = xml.includes('name="표준"');

      // 원본 Normal(표준) 스타일의 alignment 추출
      const cellStyleXfs = xml.match(/<cellStyleXfs[^>]*>([\s\S]*?)<\/cellStyleXfs>/)?.[1] || '';
      const firstStyleXf = cellStyleXfs.match(/<xf\b[\s\S]*?<\/xf>/)?.[0] || cellStyleXfs.match(/<xf\b[^>]*\/>/)?.[0];
      const alignMatch = firstStyleXf ? firstStyleXf.match(/<alignment\b[^>]*\/>/) : null;
      if (alignMatch) {
        normalAlignment = alignMatch[0];
      } else {
        const cellXfs = xml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/)?.[1] || '';
        const firstCellXf = cellXfs.match(/<xf\b[\s\S]*?<\/xf>/)?.[0] || cellXfs.match(/<xf\b[^>]*\/>/)?.[0];
        const alignMatch2 = firstCellXf ? firstCellXf.match(/<alignment\b[^>]*\/>/) : null;
        if (alignMatch2) {
          normalAlignment = alignMatch2[0];
        }
      }
    }

    // 시트별 원본 행 속성(customHeight 여부, thickBot, thickTop) 추출
    const sheetRowAttrs = new Map<string, Map<number, OriginalRowMeta>>();
    for (const [fname, fileObj] of Object.entries(zip.files)) {
      if (fname.startsWith('xl/worksheets/sheet') && fname.endsWith('.xml')) {
        const sXml = await fileObj.async('string');
        const rowMap = new Map<number, OriginalRowMeta>();
        const rowTags = sXml.match(/<row\b[^>]*>/g) || [];
        for (const tag of rowTags) {
          const rMatch = tag.match(/\br="(\d+)"/);
          if (rMatch) {
            const rNum = parseInt(rMatch[1], 10);
            rowMap.set(rNum, {
              hasCustomHeight: tag.includes('customHeight="1"') || tag.includes('customHeight="true"'),
              thickBot: tag.includes('thickBot="1"') || tag.includes('thickBot="true"'),
              thickTop: tag.includes('thickTop="1"') || tag.includes('thickTop="true"'),
            });
          }
        }
        sheetRowAttrs.set(fname, rowMap);
      }
    }

    if (!name && !normalAlignment && sheetRowAttrs.size === 0) return null;
    return { name: name || '맑은 고딕', sz, charset, family, scheme, isPyoJun, normalAlignment, sheetRowAttrs };
  } catch {
    return null;
  }
}

function patchStylesXmlDefaultFont(stylesXml: string, fontInfo: DefaultFontInfo): string {
  if (!fontInfo) return stylesXml;
  let newXml = stylesXml;

  if (fontInfo.name) {
    const fontsBlockMatch = newXml.match(/<fonts[^>]*>([\s\S]*?)<\/fonts>/);
    if (fontsBlockMatch) {
      const inner = fontsBlockMatch[1];
      const firstFontMatch = inner.match(/<font\b[\s\S]*?<\/font>/);
      if (firstFontMatch) {
        let font0 = firstFontMatch[0];
        if (font0.includes('<name')) {
          font0 = font0.replace(/<name\s+val="[^"]*"\s*\/>/, `<name val="${fontInfo.name}"/>`);
        } else {
          font0 = font0.replace('</font>', `<name val="${fontInfo.name}"/></font>`);
        }
        if (fontInfo.sz && font0.includes('<sz')) {
          font0 = font0.replace(/<sz\s+val="[^"]*"\s*\/>/, `<sz val="${fontInfo.sz}"/>`);
        }
        if (fontInfo.charset && !font0.includes('<charset')) {
          font0 = font0.replace('</font>', `<charset val="${fontInfo.charset}"/></font>`);
        }
        if (fontInfo.family && font0.includes('<family')) {
          font0 = font0.replace(/<family\s+val="[^"]*"\s*\/>/, `<family val="${fontInfo.family}"/>`);
        }
        if (fontInfo.scheme && font0.includes('<scheme')) {
          font0 = font0.replace(/<scheme\s+val="[^"]*"\s*\/>/, `<scheme val="${fontInfo.scheme}"/>`);
        }
        newXml = newXml.replace(firstFontMatch[0], font0);
      }
    }
  }

  if (fontInfo.isPyoJun && newXml.includes('name="Normal"')) {
    newXml = newXml.replace('name="Normal"', 'name="표준"');
  }

  // 원본 Normal(표준) 스타일 정렬(alignment) 복원 (한국어 엑셀 vertical="center" 등)
  if (fontInfo.normalAlignment) {
    // 1. cellStyleXfs 첫 번째 xf
    newXml = newXml.replace(
      /(<cellStyleXfs[^>]*>\s*<xf\b[^>]*?)(\/?>)/,
      (match, pre) => {
        if (pre.includes('<alignment')) return match;
        return `${pre}>${fontInfo.normalAlignment}</xf>`;
      }
    );
    // 2. cellXfs xfId="0" (기본 셀 xf)
    newXml = newXml.replace(
      /(<cellXfs[^>]*>\s*<xf\b[^>]*?xfId="0"[^>]*?)(\/?>)/,
      (match, pre) => {
        if (pre.includes('<alignment')) return match;
        return `${pre}>${fontInfo.normalAlignment}</xf>`;
      }
    );
  }

  return newXml;
}

async function cleanAndNormalizeWorkbookBuffer(
  buffer: ArrayBuffer | Uint8Array,
  origDefaultFont?: DefaultFontInfo | null,
  modifiedRowsBySheet?: Map<string, Set<number>>
): Promise<Uint8Array> {
  try {
    const zip = await JSZip.loadAsync(buffer);
    let changed = false;

    // 1. 유효성 검사 중복 정리 및 미수정 행 customHeight="1" 속성 정리
    const sheetFilterRefs = new Map<string, { c1: string; r1: string; c2: string; r2: string }>();
    for (const [filename, fileObj] of Object.entries(zip.files)) {
      if (filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml')) {
        let xml = await fileObj.async('string');
        let sheetChanged = false;

        const afMatch = xml.match(/<autoFilter\b[^>]*\bref="([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)"/);
        if (afMatch) {
          sheetFilterRefs.set(filename, { c1: afMatch[1], r1: afMatch[2], c2: afMatch[3], r2: afMatch[4] });
        }

        if (xml.includes('<dataValidations')) {
          const cleanedXml = deduplicateDataValidationsXml(xml);
          if (cleanedXml !== xml) {
            xml = cleanedXml;
            sheetChanged = true;
          }
        }

        // 사용자가 높이를 바꾸지 않은 행은 원본 customHeight 속성 유지 (ExcelJS 강제 추가 방지)
        const origSheetRowMeta = origDefaultFont?.sheetRowAttrs?.get(filename);
        if (origSheetRowMeta) {
          const modifiedRows = modifiedRowsBySheet?.get(filename) || new Set<number>();
          const cleanedXml = xml.replace(/<row\b([^>]*)>/g, (match, attrs) => {
            const rMatch = attrs.match(/\br="(\d+)"/);
            if (!rMatch) return match;
            const rNum = parseInt(rMatch[1], 10);
            if (modifiedRows.has(rNum)) return match; // 사용자가 직접 조정한 행은 건드리지 않음

            const orig = origSheetRowMeta.get(rNum);
            if (!orig) return match;

            let newAttrs = attrs;
            // 원본에 customHeight가 없었던 행은 customHeight="1" 제거
            if (!orig.hasCustomHeight) {
              newAttrs = newAttrs.replace(/\s+customHeight="1"/g, '').replace(/\s+customHeight="true"/g, '');
            }
            // 원본의 thickBot / thickTop 속성 유지
            if (orig.thickBot && !newAttrs.includes('thickBot')) {
              newAttrs += ' thickBot="1"';
            }
            if (orig.thickTop && !newAttrs.includes('thickTop')) {
              newAttrs += ' thickTop="1"';
            }
            return `<row${newAttrs}>`;
          });

          if (cleanedXml !== xml) {
            xml = cleanedXml;
            sheetChanged = true;
          }
        }

        if (sheetChanged) {
          zip.file(filename, xml);
          changed = true;
        }
      }
    }

    // 2. styles.xml 기본 글꼴(font 0), Normal 스타일명 및 alignment 복원
    if (origDefaultFont) {
      const stylesFile = zip.file('xl/styles.xml');
      if (stylesFile) {
        const stylesXml = await stylesFile.async('string');
        const patchedXml = patchStylesXmlDefaultFont(stylesXml, origDefaultFont);
        if (patchedXml !== stylesXml) {
          zip.file('xl/styles.xml', patchedXml);
          changed = true;
        }
      }
    }

    // 3. workbook.xml _xlnm._FilterDatabase 이름 정의를 sheet의 autoFilter ref와 동기화
    if (sheetFilterRefs.size > 0) {
      const wbFile = zip.file('xl/workbook.xml');
      if (wbFile) {
        let wbXml = await wbFile.async('string');
        let wbChanged = false;

        const f1 = sheetFilterRefs.get('xl/worksheets/sheet1.xml');
        if (f1) {
          const newTarget = `$${f1.c1}$${f1.r1}:$${f1.c2}$${f1.r2}`;
          const dnRegex = /(<definedName\b[^>]*name="_xlnm\._FilterDatabase"[^>]*>)([\s\S]*?)(<\/definedName>)/;
          if (dnRegex.test(wbXml)) {
            const patchedWbXml = wbXml.replace(dnRegex, (_m, p1, innerText, p3) => {
              const prefixMatch = innerText.match(/^([^!]+!)/);
              const prefix = prefixMatch ? prefixMatch[1] : '';
              return `${p1}${prefix}${newTarget}${p3}`;
            });
            if (patchedWbXml !== wbXml) {
              wbXml = patchedWbXml;
              wbChanged = true;
            }
          }
        }

        if (wbChanged) {
          zip.file('xl/workbook.xml', wbXml);
          changed = true;
        }
      }
    }

    if (!changed) {
      return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    }
    return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  } catch (err) {
    console.warn('저장 후처리 정규화 실패 (무시 가능):', err);
    return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  }
}

/**
 * Univer의 스냅샷(IWorkbookData)을 다시 XLSX 또는 CSV Data URL 문자열로 생성합니다.
 * engine: 'exceljs' (기본값, 서식/열너비/행높이 보존) | 'sheetjs' (기존 가벼운 데이터 모드)
 * originalDataOrUrl: 원본 파일 바이트(Data URL). 제공 시 원본 워크북을 로드하여 바뀐 셀/수식/서식/병합/크기만
 * in-place로 패치하므로 ExcelJS가 지원하는 조건부 서식, 데이터 유효성 검사, 셀 메모, 인쇄 설정, 틀 고정이 100% 보존됩니다.
 * structuralOperations: 행/열 삽입·삭제 명령 목록.
 */
export async function univerSnapshotToDataUrl(
  snapshot: IWorkbookData,
  format: 'xlsx' | 'csv' = 'xlsx',
  engine: 'exceljs' | 'sheetjs' = 'exceljs',
  originalDataOrUrl?: string,
  structuralOperations?: StructuralOperation[]
): Promise<string> {
  if (engine === 'sheetjs' || format === 'csv') {
    return univerSnapshotToDataUrlWithSheetJS(snapshot, format);
  }

  // --- ExcelJS 라운드트립 저장 엔진 (서식/수식/병합/너비/높이 및 원본 메타데이터 보존) ---
  try {
    const wb = new ExcelJS.Workbook();
    let isLoadedFromOriginal = false;
    let origDefaultFont: DefaultFontInfo | null = null;
    const modifiedRowsBySheet = new Map<string, Set<number>>();

    if (originalDataOrUrl) {
      try {
        const origRawBytes = dataUrlToUint8Array(originalDataOrUrl);
        origDefaultFont = await extractOriginalDocDefaults(origRawBytes);
        const origBytes = await normalizeHancellXlsx(origRawBytes);
        await wb.xlsx.load(origBytes as any);
        isLoadedFromOriginal = true;
      } catch (loadErr) {
        console.warn('원본 워크북 로드 실패:', loadErr);
        throw new Error('ORIGINAL_WORKBOOK_LOAD_FAILED');
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
      const sheetFileName = `xl/worksheets/sheet${sIdx + 1}.xml`;
      const modifiedRows = new Set<number>();
      modifiedRowsBySheet.set(sheetFileName, modifiedRows);

      // 0-0. 기존 병합을 먼저 모두 해제하여 각 셀이 독립적인 상태로 연산 및 데이터를 쓸 수 있도록 준비
      const existingMerges = [...((ws as any).model?.merges || [])];
      for (const mStr of existingMerges) {
        try {
          (ws as any).unMergeCells(mStr);
        } catch {}
      }

      // 0. 행/열 구조 변경(삽입/삭제) In-Place 적용 (메모/유효성검사/조건부서식 함께 시프트)
      const sheetOps = (structuralOperations || []).filter(
        (op) =>
          (!op.subUnitId && !op.sheetName && op.sheetIndex === undefined) ||
          op.subUnitId === sheetId ||
          op.sheetName === ws.name ||
          op.sheetIndex === sIdx
      );

      sheetOps.forEach((op) => {
        if (op.type === 'insertRow') {
          safeSpliceRows(ws, op.start, 0, ...Array(op.count).fill([]));
        } else if (op.type === 'removeRow') {
          safeSpliceRows(ws, op.start, op.count);
        } else if (op.type === 'insertCol') {
          safeSpliceColumns(ws, op.start, 0, ...Array(op.count).fill([]));
        } else if (op.type === 'removeCol') {
          safeSpliceColumns(ws, op.start, op.count);
        }
      });

      // 0-1. Univer 스냅샷의 현재 병합 목록(uSheet.mergeData)을 기준으로 슬레이브 셀 좌표 Set 구축
      const currentSlaveCells = new Set<string>();
      if (uSheet.mergeData && Array.isArray(uSheet.mergeData)) {
        uSheet.mergeData.forEach((m) => {
          for (let r = m.startRow; r <= m.endRow; r++) {
            for (let c = m.startColumn; c <= m.endColumn; c++) {
              if (r !== m.startRow || c !== m.startColumn) {
                currentSlaveCells.add(`${r},${c}`);
              }
            }
          }
        });
      }

      const cellData = uSheet.cellData || {};

      // 0-2. 원본 셀 값 비우기: 원본 워크시트 셀 중 Univer 스냅샷에 없거나 비어있는 셀의 값을 null로 초기화
      ws.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
        const rIdx = rowNumber - 1; // 0-based
        row.eachCell({ includeEmpty: false }, (excelCell: any, colNumber: number) => {
          const cIdx = colNumber - 1; // 0-based
          // Univer 현재 병합의 슬레이브 셀은 마스터의 영역이므로 임의로 비우지 않음
          if (currentSlaveCells.has(`${rIdx},${cIdx}`)) {
            return;
          }
          const uCell = cellData[rIdx]?.[cIdx];
          const isUniverEmpty =
            !uCell ||
            ((uCell.v === undefined || uCell.v === null || uCell.v === '') && !uCell.f);

          if (isUniverEmpty) {
            excelCell.value = null;
          }
        });
      });

      // 1. 열 너비 복원: 사용자가 실제로 크기를 변경한 열만 덮어쓰고, 나머지는 원본 정밀 float 너비 유지
      const colData = uSheet.columnData;
      if (colData) {
        Object.keys(colData).forEach((cStr) => {
          const cIdx = parseInt(cStr, 10);
          const colInfo = colData[cIdx];
          if (colInfo && colInfo.w) {
            const col = ws.getColumn(cIdx + 1);
            if (isLoadedFromOriginal && col.width !== undefined) {
              const initPx = Math.round(col.width * 8 + 5);
              if (Math.abs(colInfo.w - initPx) > 2) {
                col.width = Math.max(1, Math.round(((colInfo.w - 5) / 8) * 100) / 100);
              }
            } else if (!isLoadedFromOriginal) {
              col.width = Math.max(1, Math.round(((colInfo.w - 5) / 8) * 100) / 100);
            }
          }
        });
      }

      // 2. 행 높이 및 셀 데이터/스타일 복원
      const rowKeys = Object.keys(cellData);

      for (let i = 0; i < rowKeys.length; i++) {
        const r = parseInt(rowKeys[i], 10);
        const rowObj = cellData[r];
        if (!rowObj) continue;

        const rowNum = r + 1; // 1-based
        const row = ws.getRow(rowNum);

        // 행 높이: 사용자가 실제로 높이를 변경한 행만 덮어쓰고, 나머지는 원본 정밀 float 높이 유지
        if (uSheet.rowData && uSheet.rowData[r] && uSheet.rowData[r].h) {
          const origH = row.height;
          if (isLoadedFromOriginal && origH !== undefined) {
            const initPxH = Math.round(origH * 1.33);
            if (Math.abs(uSheet.rowData[r].h - initPxH) > 2) {
              row.height = Math.round((uSheet.rowData[r].h / 1.33) * 100) / 100;
              modifiedRows.add(rowNum);
            }
          } else if (!isLoadedFromOriginal) {
            row.height = Math.round(uSheet.rowData[r].h / 1.33);
            modifiedRows.add(rowNum);
          }
        }

        const colKeys = Object.keys(rowObj);
        for (let j = 0; j < colKeys.length; j++) {
          const c = parseInt(colKeys[j], 10);
          const cell = rowObj[c];
          if (!cell) continue;

          // Univer 현재 병합 목록의 슬레이브 위치인 셀은 마스터의 영역이므로 쓰기 건너뜀
          if (currentSlaveCells.has(`${r},${c}`)) {
            continue;
          }

          // [object Object] 방어: 객체가 문자열로 오염 직렬화된 값 방지
          if (cell.v === '[object Object]') {
            continue;
          }

          const colNum = c + 1; // 1-based
          const excelCell = ws.getCell(rowNum, colNum);

          // 값 및 수식 덮어쓰기
          if (cell.f) {
            let resultVal = (cell.v !== undefined && cell.v !== null ? cell.v : 0) as any;
            if (excelCell.result instanceof Date) {
              if (typeof resultVal === 'string' && /^\d{4}-\d{2}-\d{2}/.test(resultVal)) {
                resultVal = new Date(resultVal);
              } else if (resultVal === 0 || resultVal === '') {
                resultVal = excelCell.result;
              }
            }
            excelCell.value = {
              formula: cell.f.replace(/^=/, ''),
              result: resultVal,
            };
          } else if (cell.v !== undefined && cell.v !== null && cell.v !== '') {
            // 원본 셀이 richText 또는 hyperlink 객체이고 내용이 변경되지 않은 경우: 원본 richText/hyperlink 유지
            const origVal = excelCell.value;
            if (origVal && typeof origVal === 'object' && 'richText' in (origVal as any)) {
              const origText = (origVal as any).richText.map((t: any) => t.text).join('');
              if (origText === cell.v) {
                // 텍스트 미변경 -> 원본 richText 서식 유지
              } else {
                excelCell.value = cell.v as any;
              }
            } else if (origVal && typeof origVal === 'object' && 'hyperlink' in (origVal as any)) {
              const origText = (origVal as any).text || (origVal as any).hyperlink;
              if (origText === cell.v) {
                // 하이퍼링크 미변경 -> 원본 hyperlink 객체 유지
              } else {
                excelCell.value = cell.v as any;
              }
            } else if (excelCell.type === ExcelJS.ValueType.Date || origVal instanceof Date) {
              if (typeof cell.v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(cell.v)) {
                excelCell.value = new Date(cell.v);
              } else {
                excelCell.value = cell.v as any;
              }
            } else {
              excelCell.value = cell.v as any;
            }
          } else if (cell.v === '') {
            excelCell.value = null;
          }

          // 스타일 적용 (폰트, 배경색, 정렬, 표시 형식)
          if (cell.s) {
            const styleObj = typeof cell.s === 'string' ? styles[cell.s] : (cell.s as IStyleData);
            if (styleObj) {
              // 폰트 (실제로 변경된 경우에만 덮어써 cellXfs 인덱스 불필요 변경 방지)
              if (styleObj.ff || styleObj.fs || styleObj.bl || styleObj.it || styleObj.cl) {
                const curFont = excelCell.font || {};
                const fontSame =
                  (!styleObj.ff || curFont.name === styleObj.ff) &&
                  (!styleObj.fs || curFont.size === styleObj.fs) &&
                  (Boolean(curFont.bold) === (styleObj.bl === 1)) &&
                  (Boolean(curFont.italic) === (styleObj.it === 1));
                if (!fontSame || (styleObj.cl?.rgb && curFont.color?.argb !== hexToArgb(styleObj.cl.rgb))) {
                  excelCell.font = {
                    ...curFont,
                    name: styleObj.ff || curFont.name,
                    size: styleObj.fs || curFont.size,
                    bold: styleObj.bl === 1,
                    italic: styleObj.it === 1,
                    color: styleObj.cl?.rgb ? { argb: hexToArgb(styleObj.cl.rgb) } : curFont.color,
                  };
                }
              }

              // 배경색
              if (styleObj.bg?.rgb) {
                const curFill = excelCell.fill as any;
                const newArgb = hexToArgb(styleObj.bg.rgb);
                if (curFill?.fgColor?.argb !== newArgb) {
                  excelCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: newArgb },
                  };
                }
              }

              // 정렬
              if (styleObj.ht || styleObj.vt || styleObj.tb) {
                const curAlign = excelCell.alignment || {};
                const newH =
                  styleObj.ht === HorizontalAlign.CENTER
                    ? 'center'
                    : styleObj.ht === HorizontalAlign.RIGHT
                    ? 'right'
                    : styleObj.ht === HorizontalAlign.LEFT
                    ? 'left'
                    : curAlign.horizontal;
                const newV =
                  styleObj.vt === VerticalAlign.MIDDLE
                    ? 'middle'
                    : styleObj.vt === VerticalAlign.TOP
                    ? 'top'
                    : styleObj.vt === VerticalAlign.BOTTOM
                    ? 'bottom'
                    : curAlign.vertical;
                const newWrap = styleObj.tb === WrapStrategy.WRAP;
                if (curAlign.horizontal !== newH || curAlign.vertical !== newV || Boolean(curAlign.wrapText) !== newWrap) {
                  excelCell.alignment = {
                    ...curAlign,
                    horizontal: newH,
                    vertical: newV,
                    wrapText: newWrap,
                  };
                }
              }

              // 표시 형식
              if (styleObj.n?.pattern) {
                // 원본에 표시 형식이 없었던(General) 수식 셀은 Univer 수식 엔진이 참조 셀로부터 유추한 임의의 서식을 덮어쓰지 않음
                const isFormulaWithoutOrigFmt = isLoadedFromOriginal && (!excelCell.numFmt || excelCell.numFmt === 'General') && (cell.f || !excelCell.value);
                if (!isFormulaWithoutOrigFmt && excelCell.numFmt !== styleObj.n.pattern) {
                  excelCell.numFmt = styleObj.n.pattern;
                }
              }
            }
          }
        }
      }

      // 3. 셀 병합 적용: Univer 최신 병합 목록 적용
      if (uSheet.mergeData && Array.isArray(uSheet.mergeData)) {
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

    const rawBuffer = await wb.xlsx.writeBuffer();
    const uint8 = await cleanAndNormalizeWorkbookBuffer(rawBuffer, origDefaultFont, modifiedRowsBySheet);
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
