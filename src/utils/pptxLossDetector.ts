import JSZip from 'jszip';

export interface PptxUnsupportedFeatureReport {
  hasVbaMacro: boolean;
  hasOleEmbeddings: boolean;
  hasSmartArt: boolean;
  hasMediaAudioVideo: boolean;
  hasInk: boolean;
  has3DModel: boolean;
  hasEmbeddedFonts: boolean;
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

const MEDIA_EXTENSIONS = [
  '.mp4', '.avi', '.mov', '.wmv', '.m4v', '.webm', '.mkv',
  '.mp3', '.wav', '.m4a', '.wma', '.aac', '.ogg'
];

/**
 * .pptx 파일의 ZIP 구조를 사전 검사하여
 * 웹 프론트엔드 환경에서 주의가 필요한 매크로, OLE, SmartArt, 비디오/오디오 등을 감지합니다.
 */
export async function detectPptxUnsupportedFeatures(
  dataOrUrl: string
): Promise<PptxUnsupportedFeatureReport> {
  const result: PptxUnsupportedFeatureReport = {
    hasVbaMacro: false,
    hasOleEmbeddings: false,
    hasSmartArt: false,
    hasMediaAudioVideo: false,
    hasInk: false,
    has3DModel: false,
    hasEmbeddedFonts: false,
    unsupportedList: [],
  };

  try {
    const bytes = dataUrlToUint8Array(dataOrUrl);
    const zip = await JSZip.loadAsync(bytes);

    zip.forEach((relativePath) => {
      const lower = relativePath.toLowerCase();

      // 1. 매크로 파일 (.pptm / vbaProject.bin)
      if (lower.includes('vbaproject.bin') || lower.includes('vba/')) {
        result.hasVbaMacro = true;
      }
      // 2. OLE 개체 (임베디드 엑셀, 파워포인트 등)
      if (lower.startsWith('ppt/embeddings/') || lower.includes('oleobject')) {
        result.hasOleEmbeddings = true;
      }
      // 3. SmartArt 다이어그램 (dgm)
      if (lower.startsWith('ppt/diagrams/') || lower.includes('smartart') || lower.includes('/dgm')) {
        result.hasSmartArt = true;
      }
      // 4. 동영상 및 오디오 미디어
      if (lower.startsWith('ppt/media/')) {
        for (const ext of MEDIA_EXTENSIONS) {
          if (lower.endsWith(ext)) {
            result.hasMediaAudioVideo = true;
            break;
          }
        }
      }
      // 5. 잉크 주석 (Ink)
      if (lower.startsWith('ppt/ink/') || lower.includes('/ink/')) {
        result.hasInk = true;
      }
      // 6. 3D 모델
      if (lower.startsWith('ppt/models/') || lower.includes('model3d') || lower.includes('3dmodel')) {
        result.has3DModel = true;
      }
      // 7. 포함 글꼴 (Embedded Fonts)
      if (lower.startsWith('ppt/fonts/') || lower.includes('/fonts/')) {
        result.hasEmbeddedFonts = true;
      }
    });

    if (result.hasVbaMacro) {
      result.unsupportedList.push('VBA 매크로 코드 (.pptm)');
    }
    if (result.hasOleEmbeddings) {
      result.unsupportedList.push('OLE 삽입 개체 (임베디드 문서)');
    }
    if (result.hasSmartArt) {
      result.unsupportedList.push('SmartArt 그래픽');
    }
    if (result.hasMediaAudioVideo) {
      result.unsupportedList.push('동영상 및 오디오 미디어');
    }
    if (result.hasInk) {
      result.unsupportedList.push('잉크 드로잉 주석');
    }
    if (result.has3DModel) {
      result.unsupportedList.push('3D 모델 개체');
    }
    if (result.hasEmbeddedFonts) {
      result.unsupportedList.push('문서 포함 글꼴 (Embedded Fonts)');
    }
  } catch (err) {
    console.warn('PPTX 고급 요소 사전 검사 중 예외 발생:', err);
  }

  return result;
}
