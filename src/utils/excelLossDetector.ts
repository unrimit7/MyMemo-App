import JSZip from 'jszip';

export interface UnsupportedFeatureReport {
  hasCharts: boolean;
  hasPivotTables: boolean;
  hasImagesOrDrawings: boolean;
  hasVbaMacro: boolean;
  hasSparklines: boolean;
  unsupportedList: string[];
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

/**
 * .xlsx 파일의 ZIP 내부 엔트리를 초고속으로 검사하여
 * ExcelJS 및 웹 스프레드시트에서 완벽히 보존되지 않는 고급 요소(차트, 피벗, 이미지, 매크로 등)를 감지합니다.
 */
export async function detectExcelUnsupportedFeatures(
  dataOrUrl: string
): Promise<UnsupportedFeatureReport> {
  const result: UnsupportedFeatureReport = {
    hasCharts: false,
    hasPivotTables: false,
    hasImagesOrDrawings: false,
    hasVbaMacro: false,
    hasSparklines: false,
    unsupportedList: [],
  };

  try {
    const bytes = dataUrlToUint8Array(dataOrUrl);
    const zip = await JSZip.loadAsync(bytes);

    zip.forEach((relativePath) => {
      const lower = relativePath.toLowerCase();

      // 1. 차트 감지
      if (lower.startsWith('xl/charts/') || lower.includes('chart')) {
        result.hasCharts = true;
      }
      // 2. 피벗 테이블 감지
      if (lower.startsWith('xl/pivottables/') || lower.startsWith('xl/pivotrecords/')) {
        result.hasPivotTables = true;
      }
      // 3. 이미지 및 미디어/그리기 객체 감지
      if (
        lower.startsWith('xl/media/') ||
        lower.startsWith('xl/drawings/') ||
        lower.includes('drawing')
      ) {
        result.hasImagesOrDrawings = true;
      }
      // 4. VBA 매크로 감지
      if (lower.includes('vbaproject.bin') || lower.endsWith('.bin')) {
        result.hasVbaMacro = true;
      }
      // 5. 스파크라인 감지
      if (lower.includes('sparkline')) {
        result.hasSparklines = true;
      }
    });

    if (result.hasCharts) {
      result.unsupportedList.push('차트 (그래프)');
    }
    if (result.hasPivotTables) {
      result.unsupportedList.push('피벗 테이블');
    }
    if (result.hasImagesOrDrawings) {
      result.unsupportedList.push('삽입된 이미지 / 도형');
    }
    if (result.hasVbaMacro) {
      result.unsupportedList.push('VBA 매크로 코드');
    }
    if (result.hasSparklines) {
      result.unsupportedList.push('스파크라인 (미니 차트)');
    }
  } catch (err) {
    console.warn('엑셀 구조 사전 검사 중 오류 (무시 가능):', err);
  }

  return result;
}
