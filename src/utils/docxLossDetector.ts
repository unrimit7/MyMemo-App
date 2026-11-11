import JSZip from 'jszip';

export interface DocxUnsupportedFeatureReport {
  hasVbaMacro: boolean;
  hasCharts: boolean;
  hasOleEmbeddings: boolean;
  hasSmartArt: boolean;
  unsupportedList: string[];
}

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
 * .docx 파일의 ZIP 구조를 사전 검사하여
 * 웹 워드 편집기에서 주의가 필요한 매크로, OLE 임베딩, 차트 등을 감지합니다.
 */
export async function detectDocxUnsupportedFeatures(
  dataOrUrl: string
): Promise<DocxUnsupportedFeatureReport> {
  const result: DocxUnsupportedFeatureReport = {
    hasVbaMacro: false,
    hasCharts: false,
    hasOleEmbeddings: false,
    hasSmartArt: false,
    unsupportedList: [],
  };

  try {
    const bytes = dataUrlToUint8Array(dataOrUrl);
    const zip = await JSZip.loadAsync(bytes);

    zip.forEach((relativePath) => {
      const lower = relativePath.toLowerCase();

      // 1. 매크로 파일 (.docm / vbaProject.bin)
      if (lower.includes('vbaproject.bin') || lower.includes('vba/')) {
        result.hasVbaMacro = true;
      }
      // 2. 차트 객체
      if (lower.startsWith('word/charts/') || lower.includes('chart')) {
        result.hasCharts = true;
      }
      // 3. OLE 임베디드 개체 (임베디드 엑셀, 파워포인트 등)
      if (lower.startsWith('word/embeddings/') || lower.includes('oleobject')) {
        result.hasOleEmbeddings = true;
      }
      // 4. SmartArt 다이어그램
      if (lower.startsWith('word/diagrams/') || lower.includes('smartart')) {
        result.hasSmartArt = true;
      }
    });

    if (result.hasVbaMacro) {
      result.unsupportedList.push('VBA 매크로 코드');
    }
    if (result.hasOleEmbeddings) {
      result.unsupportedList.push('OLE 삽입 개체 (임베디드 오피스 파일)');
    }
  } catch (err) {
    console.warn('DOCX 고급 요소 사전 검사 중 예외 발생 (무시 가능):', err);
  }

  return result;
}
