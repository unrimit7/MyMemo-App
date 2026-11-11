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

export interface ParsedWorkbook {
  sheetNames: string[];
  sheets: Record<string, (string | number)[][]>;
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
    // attId로 시작하는 모든 동적 시그니처 캐시 엔트리 삭제
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
 * XLSX, XLS, CSV Data URL 또는 Base64 문자열을 워크북 객체로 파싱합니다.
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
      if (sheet) {
        // 2차원 배열 형태로 추출 (빈 셀은 빈 문자열로 보존)
        const rows = XLSX.utils.sheet_to_json<(string | number)[]>(sheet, {
          header: 1,
          defval: '',
          blankrows: false,
        });
        sheets[name] = rows;

        // 셀 병합(!merges) 정보 안전하게 추출
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
      XLSX.utils.book_append_sheet(workbook, ws, name.slice(0, 31)); // 엑셀 시트 이름 최대 31자
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

  // 최대 컬럼 수 계산
  let maxCols = 0;
  rows.forEach((r) => {
    maxCols = Math.max(maxCols, r.length);
  });
  if (maxCols === 0) return '';

  // 1행을 헤더로 처리 (1행만 있는 경우 기본 헤더 생성)
  const headerRow = rows[0] || [];
  const headerCols: string[] = [];
  for (let c = 0; c < maxCols; c++) {
    const val = String(headerRow[c] ?? '').trim();
    headerCols.push(val || `열 ${c + 1}`);
  }

  const lines: string[] = [];
  // 헤더
  lines.push(`| ${headerCols.join(' | ')} |`);
  // 구분선
  lines.push(`| ${headerCols.map(() => '---').join(' | ')} |`);

  // 데이터 행들 (2행부터)
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
