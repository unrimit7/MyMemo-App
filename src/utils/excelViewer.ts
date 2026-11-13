import * as XLSX from 'xlsx';

export interface ParsedSheetData {
  name: string;
  rows: (string | number)[][];
}

export interface CellMergeRange {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

export interface CellMergeInfo {
  isOrigin: boolean;
  isCovered: boolean;
  rowSpan: number;
  colSpan: number;
}

export interface RichTextRun {
  text: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  size?: number;
}

export interface CellVisualInfo {
  value: string | number;
  formattedText?: string;
  richTextRuns?: RichTextRun[];
  isFormulaFallback?: boolean;
  numFmt?: string;
  textColor?: string;      // 원본 지정 글자색 (예: '#ff0000'), 미지정 시 undefined -> 기본 #000000
  bgColor?: string;        // 원본 지정 배경색 (예: '#f7ebdd'), 미지정 시 undefined -> 기본 #ffffff
  bold?: boolean;
  italic?: boolean;
  align?: 'left' | 'center' | 'right';
  border?: {
    top?: boolean;
    bottom?: boolean;
    left?: boolean;
    right?: boolean;
    borderColor?: string;
  };
}

export interface ParsedWorkbook {
  sheetNames: string[];
  sheets: Record<string, (string | number)[][]>;
  cellStyles?: Record<string, CellVisualInfo[][]>;
  merges?: Record<string, CellMergeRange[]>;
}

/**
 * 특정 시트의 셀 병합(merges) 목록을 빠른 O(1) 좌표 조회가 가능한 맵으로 변환합니다.
 * key: `${r},${c}`
 */
export function buildMergeMap(merges: CellMergeRange[] = []): Map<string, CellMergeInfo> {
  const map = new Map<string, CellMergeInfo>();
  if (!merges || !Array.isArray(merges)) return map;

  for (const m of merges) {
    if (!m || !m.s || !m.e) continue;
    const minR = Math.min(m.s.r, m.e.r);
    const maxR = Math.max(m.s.r, m.e.r);
    const minC = Math.min(m.s.c, m.e.c);
    const maxC = Math.max(m.s.c, m.e.c);

    const rowSpan = maxR - minR + 1;
    const colSpan = maxC - minC + 1;

    // 단일 셀 병합(1x1)인 경우는 스킵
    if (rowSpan <= 1 && colSpan <= 1) continue;

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const key = `${r},${c}`;
        if (r === minR && c === minC) {
          map.set(key, { isOrigin: true, isCovered: false, rowSpan, colSpan });
        } else {
          map.set(key, { isOrigin: false, isCovered: true, rowSpan: 1, colSpan: 1 });
        }
      }
    }
  }

  return map;
}

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

// 엑셀 / CSV 캐시
const excelCache = new Map<string, ParsedWorkbook>();

export function invalidateExcelCache(attId?: string) {
  if (attId) {
    for (const key of Array.from(excelCache.keys())) {
      if (key === attId || key.startsWith(`${attId}_`)) {
        excelCache.delete(key);
      }
    }
  } else {
    excelCache.clear();
  }
}

/**
 * Excel 일련번호(serial date) 계산: 1899-12-30 기준 일수
 */
export function excelSerialFromDate(d: Date): number {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  const epoch = Date.UTC(1899, 11, 30);
  return Math.round((utc - epoch) / 86400000);
}

/**
 * 셀 표시 형식(numFmt) 포맷터
 * 섹션(양수;음수;0), 문자열 리터럴("원", "일"), 날짜/시간, 백분율, 천 단위 쉼표 지원
 * Date 객체의 toString() 노출 경로 원천 차단
 */
export function formatCellValue(val: any, numFmt?: string): string {
  if (val === undefined || val === null || val === '') return '';

  const cleanNumFmt = numFmt && numFmt.trim() !== '' && numFmt !== 'General' ? numFmt : undefined;

  // 1. Date 인스턴스인 경우
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    if (cleanNumFmt) {
      try {
        return XLSX.SSF.format(cleanNumFmt, val);
      } catch {
        // SSF 포맷 실패 시 아래 기본 포맷으로 폴백
      }
    }
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // 2. 숫자(number)인 경우
  if (typeof val === 'number') {
    if (isNaN(val)) return '';
    if (cleanNumFmt) {
      try {
        return XLSX.SSF.format(cleanNumFmt, val);
      } catch {
        // 폴백
      }
    }
    return String(val);
  }

  // 3. 문자열(string)인 경우
  if (typeof val === 'string') {
    const trimmed = val.trim();

    // 날짜 문자열(ISO 또는 YYYY-MM-DD 등)
    if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(trimmed)) {
      const parsedD = new Date(trimmed);
      if (!isNaN(parsedD.getTime())) {
        if (cleanNumFmt) {
          try {
            return XLSX.SSF.format(cleanNumFmt, parsedD);
          } catch {
            // 폴백
          }
        }
        const y = parsedD.getFullYear();
        const m = String(parsedD.getMonth() + 1).padStart(2, '0');
        const d = String(parsedD.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
    }

    // "Fri Nov 27 2026..." 형태의 문자열이 유입된 경우에도 자동 보정
    if (/^[A-Za-z]{3}\s+[A-Za-z]{3}\s+\d{1,2}\s+\d{4}/.test(trimmed)) {
      const parsedD = new Date(trimmed);
      if (!isNaN(parsedD.getTime())) {
        if (cleanNumFmt) {
          try {
            return XLSX.SSF.format(cleanNumFmt, parsedD);
          } catch {
            // 폴백
          }
        }
        const y = parsedD.getFullYear();
        const m = String(parsedD.getMonth() + 1).padStart(2, '0');
        const d = String(parsedD.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
    }

    // 숫자형 문자열이고 numFmt가 지정된 경우
    if (cleanNumFmt && /^-?\d+(\.\d+)?$/.test(trimmed)) {
      try {
        return XLSX.SSF.format(cleanNumFmt, Number(trimmed));
      } catch {
        // 폴백
      }
    }

    return val;
  }

  return String(val);
}

/**
 * 캐시값이 없거나 TODAY()/NOW() 등 휘발성 함수가 포함된 수식을 뷰어 화면 표시용으로 계산합니다.
 * 저장 파일의 수식/캐시값은 건드리지 않습니다.
 */
export function evaluateFormula(
  formula: string,
  resolveCell: (ref: string) => any
): { success: boolean; result?: any; formulaText: string } {
  const formulaText = formula.startsWith('=') ? formula : `=${formula}`;
  let f = formula.trim();
  if (f.startsWith('=')) f = f.slice(1).trim();

  const today = new Date();
  const todaySerial = excelSerialFromDate(today);

  // DATE(y, m, d)
  f = f.replace(/DATE\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/gi, (_, y, m, d) => {
    const dt = new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
    return String(excelSerialFromDate(dt));
  });

  // TODAY() / NOW()
  f = f.replace(/TODAY\s*\(\s*\)/gi, String(todaySerial));
  f = f.replace(/NOW\s*\(\s*\)/gi, String(todaySerial));

  // IF(cond, a, b)
  f = f.replace(/IF\s*\(([^,]+),([^,]+),([^)]+)\)/gi, (_, cond, a, b) => {
    return `((${cond}) ? (${a}) : (${b}))`;
  });

  // 셀 참조 (A4, B2 등)
  f = f.replace(/\b([A-Z]{1,3}[0-9]{1,7})\b/g, (match) => {
    const rawVal = resolveCell(match.toUpperCase());
    if (rawVal instanceof Date) {
      return String(excelSerialFromDate(rawVal));
    }
    if (typeof rawVal === 'number') {
      return String(rawVal);
    }
    if (typeof rawVal === 'string') {
      const d = new Date(rawVal);
      if (!isNaN(d.getTime()) && /^\d{4}[-/.]/.test(rawVal.trim())) {
        return String(excelSerialFromDate(d));
      }
      const n = Number(rawVal.replace(/,/g, '').trim());
      if (!isNaN(n)) return String(n);
    }
    return '0';
  });

  try {
    const sanitized = f.replace(/[^0-9+\-*/().?:\s><=!]/g, '');
    if (!sanitized) throw new Error('Empty sanitized formula');
    const computed = Function(`"use strict"; return (${sanitized});`)();
    if (typeof computed === 'number' && !isNaN(computed)) {
      return { success: true, result: computed, formulaText };
    }
    return { success: false, formulaText };
  } catch {
    return { success: false, formulaText };
  }
}

/**
 * XLSX, XLS, CSV Data URL 또는 Base64 문자열을 워크북 객체로 파싱합니다 (동기 파서).
 */
export function parseExcelWorkbook(dataOrUrl: string, cacheKey?: string): ParsedWorkbook {
  const dataSignature = cacheKey
    ? `${cacheKey}_${(dataOrUrl || '').slice(0, 100)}_${(dataOrUrl || '').length}_${(dataOrUrl || '').slice(-100)}`
    : undefined;

  if (dataSignature && excelCache.has(dataSignature)) {
    return excelCache.get(dataSignature)!;
  }

  try {
    const bytes = dataUrlToUint8Array(dataOrUrl);
    const workbook = XLSX.read(bytes, { type: 'array', cellDates: true });

    const sheets: Record<string, (string | number)[][]> = {};
    const merges: Record<string, CellMergeRange[]> = {};
    const sheetNames = workbook.SheetNames || [];

    sheetNames.forEach((name) => {
      const sheet = workbook.Sheets[name];
      if (sheet && sheet['!ref']) {
        const range = XLSX.utils.decode_range(sheet['!ref']);
        const rows: (string | number)[][] = [];

        for (let R = range.s.r; R <= range.e.r; ++R) {
          const rowVals: (string | number)[] = [];
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const addr = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = sheet[addr];
            let cellText = '';
            if (cell) {
              if (cell.w !== undefined && cell.w !== null) {
                cellText = cell.w;
              } else if (cell.t === 'd' && cell.v instanceof Date) {
                cellText = formatCellValue(cell.v, cell.z);
              } else if (cell.v !== undefined && cell.v !== null) {
                cellText = formatCellValue(cell.v, cell.z);
              }
            }
            rowVals.push(cellText);
          }
          rows.push(rowVals);
        }
        sheets[name] = rows;

        const rawMerges = sheet['!merges'];
        if (Array.isArray(rawMerges) && rawMerges.length > 0) {
          merges[name] = rawMerges
            .filter((m) => m && m.s && m.e)
            .map((m) => ({
              s: { r: m.s.r, c: m.s.c },
              e: { r: m.e.r, c: m.e.c },
            }));
        } else {
          merges[name] = [];
        }
      } else {
        sheets[name] = [];
        merges[name] = [];
      }
    });

    const result: ParsedWorkbook = {
      sheetNames: sheetNames.length > 0 ? sheetNames : ['Sheet1'],
      sheets: Object.keys(sheets).length > 0 ? sheets : { Sheet1: [] },
      merges,
    };

    if (dataSignature) {
      excelCache.set(dataSignature, result);
    }

    return result;
  } catch (err) {
    console.error('엑셀/CSV 파싱 오류:', err);
    return {
      sheetNames: ['Sheet1'],
      sheets: { Sheet1: [['(파일을 읽을 수 없습니다)']] },
    };
  }
}

/**
 * 2차원 배열 데이터들을 다시 XLSX Data URL로 생성합니다.
 */
export function generateExcelDataUrl(
  sheets: Record<string, (string | number)[][]>,
  format: 'xlsx' | 'csv' = 'xlsx'
): string {
  const workbook = XLSX.utils.book_new();

  const sheetNames = Object.keys(sheets);
  if (sheetNames.length === 0) {
    const ws = XLSX.utils.aoa_to_sheet([['']]);
    XLSX.utils.book_append_sheet(workbook, ws, 'Sheet1');
  } else {
    sheetNames.forEach((name) => {
      const rows = sheets[name] || [];
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(workbook, ws, name.slice(0, 31));
    });
  }

  if (format === 'csv') {
    const firstSheetName = workbook.SheetNames[0] || 'Sheet1';
    const csvStr = XLSX.utils.sheet_to_csv(workbook.Sheets[firstSheetName]);
    const b64 = btoa(unescape(encodeURIComponent(csvStr)));
    return `data:text/csv;charset=utf-8;base64,${b64}`;
  }

  const base64Out = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
  return `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64Out}`;
}

/**
 * 2차원 배열 표 데이터를 깔끔한 마크다운 표 문자열로 변환합니다.
 */
export function sheetRowsToMarkdownTable(rows: (string | number)[][]): string {
  if (!rows || rows.length === 0) return '';

  let maxCols = 0;
  rows.forEach((r) => {
    maxCols = Math.max(maxCols, r.length);
  });
  if (maxCols === 0) return '';

  const headerRow = rows[0] || [];
  const headerCols: string[] = [];
  for (let c = 0; c < maxCols; c++) {
    const val = String(headerRow[c] ?? '').trim();
    headerCols.push(val || `열 ${c + 1}`);
  }

  const lines: string[] = [];
  lines.push(`| ${headerCols.join(' | ')} |`);
  lines.push(`| ${headerCols.map(() => '---').join(' | ')} |`);

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const cols: string[] = [];
    for (let c = 0; c < maxCols; c++) {
      const val = String(row[c] ?? '').replace(/\|/g, '\\|').replace(/\n/g, ' ').trim();
      cols.push(val);
    }
    lines.push(`| ${cols.join(' | ')} |`);
  }

  return lines.join('\n');
}

/**
 * 열 인덱스(0, 1, 2...)를 엑셀 열 문자('A', 'B', 'Z', 'AA'...)로 변환합니다.
 */
export function getColumnLabel(index: number): string {
  let label = '';
  let num = index;
  while (num >= 0) {
    label = String.fromCharCode((num % 26) + 65) + label;
    num = Math.floor(num / 26) - 1;
  }
  return label;
}

function parseArgbColor(colorObj: any): string | undefined {
  if (!colorObj) return undefined;
  if (typeof colorObj === 'string') {
    if (colorObj.startsWith('#')) return colorObj;
    if (colorObj.length === 8) return `#${colorObj.slice(2)}`;
    return `#${colorObj}`;
  }
  if (colorObj.argb && typeof colorObj.argb === 'string') {
    const argb = colorObj.argb;
    if (argb.length === 8) {
      return `#${argb.slice(2)}`;
    }
    return `#${argb}`;
  }
  return undefined;
}

/**
 * ExcelJS를 사용하여 원본 셀 서식(글자색, 채우기색, 굵게, 정렬, 테두리, numFmt, richText, 수식)을 포함해 워크북을 정밀 파싱합니다.
 */
export async function parseExcelWorkbookWithStyles(dataOrUrl: string, cacheKey?: string): Promise<ParsedWorkbook> {
  const dataSignature = cacheKey
    ? `${cacheKey}_${(dataOrUrl || '').slice(0, 100)}_${(dataOrUrl || '').length}_${(dataOrUrl || '').slice(-100)}`
    : undefined;

  if (dataSignature && excelCache.has(dataSignature)) {
    const cached = excelCache.get(dataSignature)!;
    if (cached.cellStyles) {
      return cached;
    }
  }

  try {
    const bytes = dataUrlToUint8Array(dataOrUrl);
    const isCsv = (dataOrUrl || '').startsWith('data:text/csv') || (dataOrUrl || '').includes('csv;base64');
    if (isCsv) {
      return parseExcelWorkbook(dataOrUrl, cacheKey);
    }

    const ExcelJS = await import('exceljs');
    const WorkbookClass = (ExcelJS as any).Workbook || (ExcelJS as any).default?.Workbook || ExcelJS;
    const wb = new WorkbookClass();
    await wb.xlsx.load(bytes);

    const sheetNames: string[] = [];
    const sheets: Record<string, (string | number)[][]> = {};
    const merges: Record<string, CellMergeRange[]> = {};
    const cellStyles: Record<string, CellVisualInfo[][]> = {};

    wb.worksheets.forEach((ws: any) => {
      const sName = ws.name || 'Sheet1';
      sheetNames.push(sName);

      const rows: (string | number)[][] = [];
      const sheetStyles: CellVisualInfo[][] = [];
      const rowCount = ws.rowCount || 0;
      const colCount = ws.columnCount || 0;

      // 셀 참조 맵 (수식 평가용)
      const cellValMap = new Map<string, any>();
      for (let r = 1; r <= rowCount; r++) {
        const row = ws.getRow(r);
        for (let c = 1; c <= colCount; c++) {
          const cell = row.getCell(c);
          const ref = `${getColumnLabel(c - 1)}${r}`;
          let v = cell.value;
          if (v && typeof v === 'object') {
            if ('result' in v && v.result !== undefined && v.result !== null) {
              v = v.result;
            } else if (v instanceof Date) {
              // Date
            }
          }
          cellValMap.set(ref, v);
        }
      }

      for (let r = 1; r <= rowCount; r++) {
        const row = ws.getRow(r);
        const rowVals: (string | number)[] = [];
        const rowStyles: CellVisualInfo[] = [];

        for (let c = 1; c <= colCount; c++) {
          const cell = row.getCell(c);
          const numFmt = cell.numFmt;
          let val: any = cell.value;
          let formattedText = '';
          let richTextRuns: RichTextRun[] | undefined = undefined;
          let isFormulaFallback = false;

          if (val && typeof val === 'object') {
            // 1. 수식(formula) 객체 처리
            if ('formula' in val) {
              const formulaStr = String(val.formula || '');
              const hasVolatile = formulaStr.includes('TODAY') || formulaStr.includes('NOW');
              const hasCachedResult = val.result !== undefined && val.result !== null;

              if (!hasCachedResult || hasVolatile) {
                // 뷰어에서 동적 평가 시도
                const evalResult = evaluateFormula(formulaStr, (ref) => cellValMap.get(ref));
                if (evalResult.success) {
                  val = evalResult.result;
                  formattedText = formatCellValue(val, numFmt);
                } else if (hasCachedResult) {
                  val = val.result;
                  formattedText = formatCellValue(val, numFmt);
                } else {
                  // 계산 실패 시 수식 문자열을 회색으로 표시
                  val = evalResult.formulaText;
                  formattedText = evalResult.formulaText;
                  isFormulaFallback = true;
                }
              } else {
                val = val.result;
                formattedText = formatCellValue(val, numFmt);
              }
            } else if ('richText' in val && Array.isArray(val.richText)) {
              // 2. 서식 있는 텍스트(RichText) run 단위 처리
              const runs: RichTextRun[] = val.richText.map((item: any) => ({
                text: item.text || '',
                color: parseArgbColor(item.font?.color),
                bold: !!item.font?.bold,
                italic: !!item.font?.italic,
                size: item.font?.size,
              }));
              richTextRuns = runs;
              val = runs.map((t) => t.text).join('');
              formattedText = val;
            } else if ('text' in val) {
              val = val.text;
              formattedText = formatCellValue(val, numFmt);
            } else if (val instanceof Date) {
              formattedText = formatCellValue(val, numFmt);
            } else {
              val = String(val);
              formattedText = formatCellValue(val, numFmt);
            }
          } else {
            formattedText = formatCellValue(val, numFmt);
          }

          if (val === undefined || val === null) val = '';
          rowVals.push(formattedText || String(val));

          const font = cell.font || {};
          const fill = cell.fill || {};
          const alignment = cell.alignment || {};
          const border = cell.border || {};

          let textColor = parseArgbColor(font.color);
          let bgColor: string | undefined = undefined;
          if (fill.type === 'pattern' && fill.fgColor) {
            bgColor = parseArgbColor(fill.fgColor);
          }

          let align: 'left' | 'center' | 'right' | undefined = undefined;
          if (alignment.horizontal === 'center' || alignment.horizontal === 'right' || alignment.horizontal === 'left') {
            align = alignment.horizontal;
          }

          const hasBorder = !!(border.top || border.bottom || border.left || border.right);
          const borderColor = parseArgbColor(border.top?.color || border.bottom?.color || border.left?.color || border.right?.color);

          rowStyles.push({
            value: val,
            formattedText,
            richTextRuns,
            isFormulaFallback,
            numFmt,
            textColor,
            bgColor,
            bold: !!font.bold,
            italic: !!font.italic,
            align,
            border: hasBorder ? {
              top: !!border.top,
              bottom: !!border.bottom,
              left: !!border.left,
              right: !!border.right,
              borderColor,
            } : undefined,
          });
        }
        rows.push(rowVals);
        sheetStyles.push(rowStyles);
      }

      sheets[sName] = rows;
      cellStyles[sName] = sheetStyles;

      const sheetMerges: CellMergeRange[] = [];
      if (ws.model && ws.model.merges && Array.isArray(ws.model.merges)) {
        ws.model.merges.forEach((mStr: string) => {
          const parts = mStr.split(':');
          if (parts.length === 2) {
            const startCell = ws.getCell(parts[0]);
            const endCell = ws.getCell(parts[1]);
            sheetMerges.push({
              s: { r: Number(startCell.row) - 1, c: Number(startCell.col) - 1 },
              e: { r: Number(endCell.row) - 1, c: Number(endCell.col) - 1 },
            });
          }
        });
      }
      merges[sName] = sheetMerges;
    });

    const parsedWb: ParsedWorkbook = {
      sheetNames: sheetNames.length > 0 ? sheetNames : ['Sheet1'],
      sheets: Object.keys(sheets).length > 0 ? sheets : { Sheet1: [] },
      cellStyles,
      merges,
    };

    if (dataSignature) {
      excelCache.set(dataSignature, parsedWb);
    }

    return parsedWb;
  } catch (err) {
    console.warn('ExcelJS 스타일 파싱 실패, XLSX 기본 파서로 폴백:', err);
    return parseExcelWorkbook(dataOrUrl, cacheKey);
  }
}

/**
 * 특정 시트의 테이블 헤더 및 바디 HTML을 앱 테마와 분리된 완전한 엑셀 룩앤필(기본 흰 바탕/검정 글자, 원본 서식 보존)로 생성합니다.
 */
export function renderExcelTableHtml(
  parsed: ParsedWorkbook,
  sheetName: string,
  maxRows = 12
): {
  tableHeaderHtml: string;
  tableBodyHtml: string;
  colCount: number;
  totalRows: number;
} {
  const rows = parsed.sheets[sheetName] || [];
  const sheetStyles = parsed.cellStyles?.[sheetName];
  const sheetMerges = parsed.merges?.[sheetName] || [];
  const totalRows = rows.length;

  let maxCols = 0;
  rows.forEach((r) => {
    maxCols = Math.max(maxCols, r.length);
  });
  sheetMerges.forEach((m) => {
    maxCols = Math.max(maxCols, m.e.c + 1);
  });
  const colCount = Math.max(maxCols, 1);
  const previewRows = rows.slice(0, maxRows);
  const mergeMap = buildMergeMap(sheetMerges);

  const colHeaders = Array.from({ length: colCount })
    .map((_, cIdx) => {
      return `<th class="excel-th" style="padding: 6px 12px; font-family: monospace; font-size: 11px; font-weight: 600; color: #334155 !important; background-color: #f1f5f9 !important; border-right: 1px solid #cbd5e1 !important; border-bottom: 1px solid #cbd5e1 !important; min-width: 85px; text-align: center !important; user-select: none;">${getColumnLabel(cIdx)}</th>`;
    })
    .join('');

  const tableHeaderHtml = `
    <thead style="background-color: #f1f5f9 !important; position: sticky; top: 0; z-index: 10;">
      <tr>
        <th class="excel-th-num" style="width: 40px; min-width: 40px; padding: 6px 8px; text-align: center !important; font-size: 10px; font-weight: 700; color: #64748b !important; background-color: #e2e8f0 !important; border-right: 1px solid #cbd5e1 !important; border-bottom: 1px solid #cbd5e1 !important; position: sticky; left: 0; z-index: 20; user-select: none;">#</th>
        ${colHeaders}
      </tr>
    </thead>
  `;

  const tableBodyHtml = previewRows
    .map((row, rIdx) => {
      const cellsHtml: string[] = [];
      const rowStyle = sheetStyles ? sheetStyles[rIdx] : undefined;

      for (let cIdx = 0; cIdx < colCount; cIdx++) {
        const mergeInfo = mergeMap.get(`${rIdx},${cIdx}`);
        if (mergeInfo?.isCovered) {
          continue;
        }

        const colSpanAttr = mergeInfo && mergeInfo.colSpan > 1 ? ` colspan="${mergeInfo.colSpan}"` : '';
        const rowSpanAttr = mergeInfo && mergeInfo.rowSpan > 1 ? ` rowspan="${mergeInfo.rowSpan}"` : '';
        const isMerged = mergeInfo && (mergeInfo.colSpan > 1 || mergeInfo.rowSpan > 1);

        const cellVisual = rowStyle ? rowStyle[cIdx] : undefined;
        // 기본 바탕 흰색, 기본 글자 검정(rgb(0,0,0)), 원본 지정 색상이 있으면 원본 색상 적용
        const effectiveBg = cellVisual?.bgColor || (isMerged ? '#f8fafc' : '#ffffff');
        const effectiveColor = cellVisual?.isFormulaFallback
          ? '#94a3b8'
          : (cellVisual?.textColor || '#000000');
        const isBold = cellVisual?.bold !== undefined ? cellVisual.bold : isMerged;
        const textAlign = cellVisual?.align || (typeof row[cIdx] === 'number' ? 'right' : 'left');
        const borderColor = cellVisual?.border?.borderColor || '#e2e8f0';

        // 셀 내부 내용: 서식 있는 텍스트(RichText) run 렌더링 지원
        let cellInnerHtml = '';
        if (cellVisual?.richTextRuns && cellVisual.richTextRuns.length > 0) {
          cellInnerHtml = cellVisual.richTextRuns
            .map((run) => {
              const runColor = run.color ? `color: ${run.color} !important;` : 'color: inherit !important;';
              const runWeight = run.bold ? 'font-weight: 700 !important;' : 'font-weight: inherit !important;';
              const runStyle = run.italic ? 'font-style: italic !important;' : '';
              return `<span style="${runColor} ${runWeight} ${runStyle}">${run.text}</span>`;
            })
            .join('');
        } else {
          const textVal = cellVisual?.formattedText !== undefined
            ? cellVisual.formattedText
            : (row[cIdx] !== undefined && row[cIdx] !== null ? String(row[cIdx]) : '');
          cellInnerHtml = textVal;
        }

        const styleAttr = `style="background-color: ${effectiveBg} !important; color: ${effectiveColor} !important; font-weight: ${isBold ? '700' : '400'} !important; text-align: ${textAlign} !important; border-right: 1px solid ${borderColor} !important; border-bottom: 1px solid ${borderColor} !important; padding: 6px 12px; font-size: 12px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.5; word-break: break-word; white-space: pre-wrap; max-width: 320px;"`;

        cellsHtml.push(
          `<td${colSpanAttr}${rowSpanAttr} class="excel-td" ${styleAttr}>${cellInnerHtml}</td>`
        );
      }

      return `
        <tr class="excel-tr" style="background-color: #ffffff;">
          <td class="excel-td-num" style="padding: 6px 8px; text-align: center !important; font-size: 10px; font-family: monospace; font-weight: 600; color: #64748b !important; background-color: #f8fafc !important; border-right: 1px solid #cbd5e1 !important; border-bottom: 1px solid #e2e8f0 !important; position: sticky; left: 0; z-index: 10; user-select: none;">${rIdx + 1}</td>
          ${cellsHtml.join('')}
        </tr>
      `;
    })
    .join('');

  return { tableHeaderHtml, tableBodyHtml, colCount, totalRows };
}

/**
 * 인라인 카드에서 시트 탭을 클릭했을 때 카드 내부의 테이블과 상태를 즉시 해당 시트로 전환합니다.
 */
export async function switchExcelInlineSheet(
  cardElement: HTMLElement,
  sheetName: string,
  dataOrUrl: string
): Promise<void> {
  if (!cardElement || !sheetName || !dataOrUrl) return;

  const attId = cardElement.getAttribute('data-att-id') || '';
  let parsed: ParsedWorkbook;
  try {
    parsed = await parseExcelWorkbookWithStyles(dataOrUrl, attId);
  } catch {
    parsed = parseExcelWorkbook(dataOrUrl, attId);
  }

  const { tableHeaderHtml, tableBodyHtml, colCount, totalRows } = renderExcelTableHtml(parsed, sheetName, 12);

  const tableEl = cardElement.querySelector<HTMLTableElement>('table');
  if (tableEl) {
    const thead = tableEl.querySelector('thead');
    if (thead) {
      thead.innerHTML = tableHeaderHtml.replace(/<thead[^>]*>/, '').replace(/<\/thead>/, '');
    } else {
      tableEl.insertAdjacentHTML('afterbegin', tableHeaderHtml);
    }

    const tbody = tableEl.querySelector('tbody');
    if (tbody) {
      tbody.innerHTML = tableBodyHtml;
    }
  }

  // 상단 시트 탭 활성/비활성 클래스 업데이트
  cardElement.querySelectorAll<HTMLElement>('.excel-sheet-pill').forEach((pill) => {
    const pName = pill.getAttribute('data-sheet-name');
    if (pName === sheetName) {
      pill.className = 'excel-sheet-pill px-2.5 py-1 text-[11px] font-semibold rounded-md shrink-0 whitespace-nowrap transition-all cursor-pointer bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 shadow-xs';
    } else {
      pill.className = 'excel-sheet-pill px-2.5 py-1 text-[11px] font-semibold rounded-md shrink-0 whitespace-nowrap transition-all cursor-pointer bg-slate-800 text-slate-400 hover:text-slate-200 border border-slate-700/60';
    }
  });

  // 하단 행/열 카운트 텍스트 갱신
  const footerSpan = cardElement.querySelector<HTMLElement>('.excel-footer-info');
  if (footerSpan) {
    footerSpan.textContent = `전체 ${totalRows}개 행 • ${colCount}개 열 ${totalRows > 12 ? '(상위 12개 행 미리보기)' : ''}`;
  }
}

/**
 * 인라인 카드들에 원본 서식(글자색, 채우기색, 테두리 등)을 비동기로 로드하여 부드럽게 주입합니다.
 */
export async function hydrateExcelInlineStyles(
  container: HTMLElement,
  getAttachmentData: (attId: string, fileName?: string) => string | undefined
): Promise<void> {
  if (!container) return;

  const cards = container.querySelectorAll<HTMLElement>('.excel-inline-viewer');
  for (const card of Array.from(cards)) {
    const attId = card.getAttribute('data-att-id') || '';
    const fileName = card.getAttribute('data-file-name') || '';
    const data = getAttachmentData(attId, fileName);
    if (!data) continue;

    // 현재 활성화된 시트 이름
    const activePill = card.querySelector<HTMLElement>('.excel-sheet-pill.bg-emerald-500\\/20');
    const sheetName = activePill?.getAttribute('data-sheet-name') || undefined;

    try {
      const styledWb = await parseExcelWorkbookWithStyles(data, attId);
      const targetSheet = sheetName || styledWb.sheetNames[0] || 'Sheet1';
      if (styledWb.cellStyles?.[targetSheet]) {
        const { tableHeaderHtml, tableBodyHtml } = renderExcelTableHtml(styledWb, targetSheet, 12);
        const tableEl = card.querySelector<HTMLTableElement>('table');
        if (tableEl) {
          const thead = tableEl.querySelector('thead');
          if (thead) {
            thead.innerHTML = tableHeaderHtml.replace(/<thead[^>]*>/, '').replace(/<\/thead>/, '');
          }
          const tbody = tableEl.querySelector('tbody');
          if (tbody) {
            tbody.innerHTML = tableBodyHtml;
          }
        }
      }
    } catch {
      // 무시
    }
  }
}
