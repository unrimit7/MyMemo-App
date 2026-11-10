import JSZip from 'jszip';
import { loadPdfDocument } from './pdfViewer';

export interface PptxTextRun {
  text: string;
  isBold?: boolean;
  isItalic?: boolean;
  isUnderline?: boolean;
  fontSize?: number; // pt 단위
  color?: string; // hex 색상 (예: #FFFFFF, #333333)
}

export interface PptxParagraph {
  text: string;
  runs: PptxTextRun[];
  level: number;
  align?: 'left' | 'center' | 'right' | 'justify';
  isTitle?: boolean;
  bulletChar?: string;
  marL?: number; // 왼쪽 여백 (EMU)
  indent?: number; // 첫 줄 들여쓰기 (EMU)
  isBlank?: boolean;
}

export interface PptxTableCell {
  text: string;
  paragraphs: PptxParagraph[];
  colSpan?: number;
  rowSpan?: number;
  fillColor?: string;
  textColor?: string;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  align?: 'left' | 'center' | 'right';
  borders?: {
    top?: boolean;
    bottom?: boolean;
    left?: boolean;
    right?: boolean;
    color?: string;
  };
}

export interface PptxTableRow {
  cells: PptxTableCell[];
  height?: number; // percentage
}

export interface PptxTableElement {
  id: string;
  type: 'table';
  left: number; // percentage (0-100)
  top: number; // percentage (0-100)
  width: number; // percentage (0-100)
  height: number; // percentage (0-100)
  colWidths: number[]; // percentage of table width
  rows: PptxTableRow[];
  zIndex: number;
}

export interface PptxShapeElement {
  id: string;
  type: 'shape';
  shapeType?: string; // 'rect', 'roundRect', 'ellipse', etc.
  isRounded?: boolean;
  left: number; // percentage (0-100)
  top: number; // percentage (0-100)
  width: number; // percentage (0-100)
  height: number; // percentage (0-100)
  fillColor?: string;
  borderColor?: string;
  borderWidth?: number;
  paragraphs: PptxParagraph[];
  isTitle?: boolean;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  zIndex: number;
}

export interface PptxImageElement {
  id: string;
  type: 'image';
  left: number; // percentage (0-100)
  top: number; // percentage (0-100)
  width: number; // percentage (0-100)
  height: number; // percentage (0-100)
  dataUrl: string;
  contentType: string;
  zIndex: number;
}

export type PptxSlideElement = PptxShapeElement | PptxTableElement | PptxImageElement;

export interface PptxSlide {
  slideNumber: number;
  title: string;
  slidePath: string;
  width: number;
  height: number;
  aspectRatio: number;
  backgroundColor?: string;
  elements: PptxSlideElement[]; // 2차원 절대 좌표 렌더링용 엘리먼트 목록
  paragraphs: PptxParagraph[]; // 레거시/텍스트 추출 호환용
  tables: { headers: string[]; rows: string[][] }[]; // 레거시 호환용
  images: { id: string; dataUrl: string; contentType: string }[];
  notes?: string;
}

export interface PptxPresentation {
  title: string;
  slideCount: number;
  slideWidth: number;
  slideHeight: number;
  aspectRatio: number; // 16/9 = 1.7777...
  isLandscape?: boolean;
  orientation?: 'landscape' | 'portrait';
  slides: PptxSlide[];
  rawZip?: JSZip;
}

// Data URL에서 Base64 바이너리 추출
function dataUrlToBinary(dataUrl: string): string {
  const commaIdx = dataUrl.indexOf(',');
  const base64 = commaIdx !== -1 ? dataUrl.slice(commaIdx + 1) : dataUrl;
  return atob(base64);
}

// 바이너리 문자열을 Uint8Array로 변환
function binaryToUint8Array(binary: string): Uint8Array {
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

// 캐시 맵
const presentationCache = new Map<string, { presentation: PptxPresentation; timestamp: number }>();

export function invalidatePptxCache(attId?: string) {
  if (attId) {
    presentationCache.delete(attId);
  } else {
    presentationCache.clear();
  }
}

// XML 문자열을 DOM으로 파싱
function parseXml(xmlText: string): Document {
  const parser = new DOMParser();
  return parser.parseFromString(xmlText, 'application/xml');
}

// 노드에서 로컬 태그 이름으로 요소들 찾기 (네임스페이스 무관)
function getElementsByLocalName(parent: Element | Document, localName: string): Element[] {
  const results: Element[] = [];
  const elements = parent.getElementsByTagName('*');
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    if (el.localName === localName) {
      results.push(el);
    }
  }
  return results;
}

// 첫 번째 일치하는 로컬 태그 이름 요소 찾기
function findFirstByLocalName(parent: Element | Document, localName: string): Element | null {
  const elements = parent.getElementsByTagName('*');
  for (let i = 0; i < elements.length; i++) {
    if (elements[i].localName === localName) {
      return elements[i];
    }
  }
  return null;
}

// 특정 노드의 직속 자식 중 로컬 태그 이름이 일치하는 요소들 찾기
function getDirectChildrenByLocalName(parent: Element, localName: string): Element[] {
  const results: Element[] = [];
  for (let i = 0; i < parent.children.length; i++) {
    const child = parent.children[i];
    if (child.localName === localName) {
      results.push(child);
    }
  }
  return results;
}

// DrawingML 색상 파서 (srgbClr, schemeClr, prstClr)
function parseDrawingColor(parent: Element, themeMap?: Record<string, string>): string | undefined {
  // 1. srgbClr (직접 지정 16진수 색상)
  const srgb = getElementsByLocalName(parent, 'srgbClr');
  if (srgb.length > 0) {
    const val = srgb[0].getAttribute('val');
    if (val) return `#${val}`;
  }

  // 2. schemeClr (테마 색상 맵핑)
  const scheme = getElementsByLocalName(parent, 'schemeClr');
  if (scheme.length > 0) {
    const val = scheme[0].getAttribute('val');
    if (val) {
      const lumMod = getElementsByLocalName(scheme[0], 'lumMod');
      if (val === 'bg1' && lumMod.length > 0) {
        return '#F1F5F9';
      }
      if (themeMap && themeMap[val]) {
        return themeMap[val];
      }
      const defaultSchemeMap: Record<string, string> = {
        accent1: '#4F81BD',
        accent2: '#C0504D',
        accent3: '#5CB85C',
        accent4: '#8064A2',
        accent5: '#4BACC6',
        accent6: '#F79646',
        dk1: '#111827',
        lt1: '#FFFFFF',
        dk2: '#1F497D',
        lt2: '#F3F4F6',
        bg1: '#FFFFFF',
        tx1: '#111827',
        bg2: '#F3F4F6',
        tx2: '#1F497D',
        hlink: '#2563EB',
        folHlink: '#7C3AED',
      };
      if (defaultSchemeMap[val]) {
        return defaultSchemeMap[val];
      }
    }
  }

  // 3. prstClr (사전 정의 색상 이름)
  const prst = getElementsByLocalName(parent, 'prstClr');
  if (prst.length > 0) {
    const val = prst[0].getAttribute('val');
    if (val) return val;
  }

  return undefined;
}

// 테마 파일(theme1.xml)에서 프레젠테이션 고유 색상 팔레트 파싱
async function parseThemeColors(zip: JSZip): Promise<Record<string, string>> {
  const themeMap: Record<string, string> = {};
  const themeFiles = Object.keys(zip.files).filter(
    (f) => f.startsWith('ppt/theme/theme') && f.endsWith('.xml')
  );
  if (themeFiles.length === 0) return themeMap;

  const themeFile = zip.file(themeFiles[0]);
  if (!themeFile) return themeMap;

  try {
    const themeText = await themeFile.async('text');
    const themeDoc = parseXml(themeText);
    const clrScheme = findFirstByLocalName(themeDoc, 'clrScheme');
    if (clrScheme) {
      for (let i = 0; i < clrScheme.children.length; i++) {
        const child = clrScheme.children[i];
        const tag = child.localName;
        const srgbList = getElementsByLocalName(child, 'srgbClr');
        if (srgbList.length > 0) {
          const val = srgbList[0].getAttribute('val');
          if (val) themeMap[tag] = `#${val}`;
        } else {
          const sysList = getElementsByLocalName(child, 'sysClr');
          if (sysList.length > 0) {
            const lastClr = sysList[0].getAttribute('lastClr');
            if (lastClr) themeMap[tag] = `#${lastClr}`;
          }
        }
      }
    }
  } catch (e) {
    console.warn('테마 색상 파싱 오류:', e);
  }

  return themeMap;
}

// EMU 좌표를 슬라이드 기준 백분율(0-100%)로 변환
function parseElementCoordinates(
  parent: Element,
  slideWidth: number,
  slideHeight: number
): { left: number; top: number; width: number; height: number } | null {
  const xfrmList = getElementsByLocalName(parent, 'xfrm');
  if (xfrmList.length === 0) return null;
  const xfrm = xfrmList[0];

  const offList = getElementsByLocalName(xfrm, 'off');
  const extList = getElementsByLocalName(xfrm, 'ext');

  if (offList.length === 0 || extList.length === 0) return null;

  const rawX = parseInt(offList[0].getAttribute('x') || '0', 10);
  const rawY = parseInt(offList[0].getAttribute('y') || '0', 10);
  const rawW = parseInt(extList[0].getAttribute('cx') || '0', 10);
  const rawH = parseInt(extList[0].getAttribute('cy') || '0', 10);

  if (slideWidth <= 0 || slideHeight <= 0) return null;

  const left = Math.max(0, Math.min(100, (rawX / slideWidth) * 100));
  const top = Math.max(0, Math.min(100, (rawY / slideHeight) * 100));
  const width = Math.max(0, Math.min(100, (rawW / slideWidth) * 100));
  const height = Math.max(0, Math.min(100, (rawH / slideHeight) * 100));

  return { left, top, width, height };
}

// 단락(Paragraph) 파서
function parseParagraph(p: Element, defaultAlign?: 'left' | 'center' | 'right' | 'justify'): PptxParagraph {
  let level = 0;
  let align: 'left' | 'center' | 'right' | 'justify' = defaultAlign || 'left';
  let marL: number | undefined;
  let indent: number | undefined;

  const pPrList = getDirectChildrenByLocalName(p, 'pPr');
  let bulletPrefix = '';
  if (pPrList.length > 0) {
    const lvlAttr = pPrList[0].getAttribute('lvl');
    if (lvlAttr) level = parseInt(lvlAttr, 10);

    const marLAttr = pPrList[0].getAttribute('marL');
    if (marLAttr) marL = parseInt(marLAttr, 10);

    const indentAttr = pPrList[0].getAttribute('indent');
    if (indentAttr) indent = parseInt(indentAttr, 10);

    const algnAttr = pPrList[0].getAttribute('algn');
    if (algnAttr === 'ctr') align = 'center';
    else if (algnAttr === 'r') align = 'right';
    else if (algnAttr === 'just') align = 'justify';
    else if (algnAttr === 'l') align = 'left';

    const buAutoNum = getElementsByLocalName(pPrList[0], 'buAutoNum');
    if (buAutoNum.length > 0) {
      bulletPrefix = '1. ';
    }
  }

  const runs: PptxTextRun[] = [];
  let paragraphFullText = '';

  for (let i = 0; i < p.children.length; i++) {
    const child = p.children[i];
    const localName = child.localName;

    if (localName === 'br') {
      runs.push({ text: '\n' });
      paragraphFullText += '\n';
      continue;
    }

    if (localName === 'r') {
      const tElements = getDirectChildrenByLocalName(child, 't');
      for (const t of tElements) {
        const textContent = t.textContent || '';
        if (!textContent) continue;

        let isBold = false;
        let isItalic = false;
        let isUnderline = false;
        let fontSize: number | undefined;
        let color: string | undefined;

        const rPrList = getDirectChildrenByLocalName(child, 'rPr');
        if (rPrList.length > 0) {
          const rPr = rPrList[0];
          isBold = rPr.getAttribute('b') === '1' || rPr.getAttribute('b') === 'true';
          isItalic = rPr.getAttribute('i') === '1' || rPr.getAttribute('i') === 'true';
          isUnderline = rPr.getAttribute('u') === 'sng';

          const szAttr = rPr.getAttribute('sz');
          if (szAttr) {
            fontSize = Math.round(parseInt(szAttr, 10) / 100);
          }

          color = parseDrawingColor(rPr);
        }

        runs.push({
          text: textContent,
          isBold,
          isItalic,
          isUnderline,
          fontSize,
          color,
        });
        paragraphFullText += textContent;
      }
    }
  }

  // 직속 t 요소 (필드 내 텍스트 등) fallback
  if (runs.length === 0) {
    const allT = getElementsByLocalName(p, 't');
    for (const t of allT) {
      const textContent = t.textContent || '';
      if (textContent) {
        runs.push({ text: textContent });
        paragraphFullText += textContent;
      }
    }
  }

  if (bulletPrefix && runs.length > 0) {
    const trimmed = runs[0].text.trim();
    if (!trimmed.startsWith('1.') && !trimmed.startsWith('2.') && !trimmed.startsWith('3.')) {
      runs[0].text = bulletPrefix + runs[0].text;
      paragraphFullText = bulletPrefix + paragraphFullText;
    }
  }

  return {
    text: paragraphFullText,
    runs: runs.length > 0 ? runs : [{ text: paragraphFullText }],
    level,
    align,
    marL,
    indent,
    isBlank: !paragraphFullText.trim(),
  };
}

/**
 * EMF 파일 내부에 포함된 벡터 PDF 데이터를 추출하여 PNG dataUrl로 렌더링
 */
async function convertEmfWithEmbeddedPdf(emfBytes: Uint8Array): Promise<string | null> {
  const pdfSig = [0x25, 0x50, 0x44, 0x46, 0x2d]; // %PDF-
  let start = -1;
  for (let i = 0; i < emfBytes.length - 5; i++) {
    if (
      emfBytes[i] === pdfSig[0] &&
      emfBytes[i + 1] === pdfSig[1] &&
      emfBytes[i + 2] === pdfSig[2] &&
      emfBytes[i + 3] === pdfSig[3] &&
      emfBytes[i + 4] === pdfSig[4]
    ) {
      start = i;
      break;
    }
  }
  if (start === -1) return null;

  const eofSig = [0x25, 0x25, 0x45, 0x4f, 0x46]; // %%EOF
  let end = -1;
  for (let i = emfBytes.length - 5; i >= start; i--) {
    if (
      emfBytes[i] === eofSig[0] &&
      emfBytes[i + 1] === eofSig[1] &&
      emfBytes[i + 2] === eofSig[2] &&
      emfBytes[i + 3] === eofSig[3] &&
      emfBytes[i + 4] === eofSig[4]
    ) {
      end = i + 5;
      break;
    }
  }
  const pdfBytes = emfBytes.subarray(start, end !== -1 ? end : undefined);
  if (pdfBytes.length < 100) return null;

  try {
    let binary = '';
    const chunk = 8192;
    for (let i = 0; i < pdfBytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, Array.from(pdfBytes.subarray(i, i + chunk)));
    }
    const pdfDataUrl = `data:application/pdf;base64,${btoa(binary)}`;
    const pdfDoc = await loadPdfDocument(pdfDataUrl, `emf_${start}_${pdfBytes.length}`);
    const page = await pdfDoc.getPage(1);
    const viewport = page.getViewport({ scale: 2.0 });

    if (typeof document !== 'undefined') {
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(viewport.width);
      canvas.height = Math.round(viewport.height);
      const ctx = canvas.getContext('2d');
      if (ctx) {
        await (page as any).render({ canvasContext: ctx, viewport }).promise;
        return canvas.toDataURL('image/png');
      }
    }
  } catch (e) {
    console.warn('EMF 내부 PDF 변환 실패:', e);
  }
  return null;
}

/**
 * PPTX 미디어 파일(PNG, JPG, SVG, EMF 등)을 안전하게 dataUrl로 로드
 */
async function loadMediaAsDataUrl(
  zip: JSZip,
  mediaPath: string,
  hintName?: string
): Promise<{ dataUrl: string; mime: string } | null> {
  const mediaFile = zip.file(mediaPath);
  if (!mediaFile) return null;

  const ext = mediaPath.split('.').pop()?.toLowerCase() || 'png';
  if (ext === 'emf' || ext === 'wmf') {
    // 1. EMF 내부의 벡터 PDF 추출 및 렌더링 시도
    const bytes = await mediaFile.async('uint8array');
    const pdfPngUrl = await convertEmfWithEmbeddedPdf(bytes);
    if (pdfPngUrl) {
      return { dataUrl: pdfPngUrl, mime: 'image/png' };
    }

    // 2. Carrier 로고 등 번들 대체 이미지가 있는 경우 (image4.png 등)
    const lowerHint = (hintName || '').toLowerCase();
    const isCarrierLogo = lowerHint.includes('carrier') || mediaPath.includes('image2');
    if (isCarrierLogo) {
      const fallbackFile = zip.file('ppt/media/image4.png');
      if (fallbackFile) {
        const fbBase64 = await fallbackFile.async('base64');
        return { dataUrl: `data:image/png;base64,${fbBase64}`, mime: 'image/png' };
      }
    }
    return null;
  }

  const mime =
    ext === 'jpg' || ext === 'jpeg'
      ? 'image/jpeg'
      : ext === 'gif'
      ? 'image/gif'
      : ext === 'svg'
      ? 'image/svg+xml'
      : 'image/png';
  const base64 = await mediaFile.async('base64');
  return { dataUrl: `data:${mime};base64,${base64}`, mime };
}

interface SlideLayoutData {
  elements: PptxSlideElement[];
  phStyles: Map<string, { align?: 'left' | 'center' | 'right' }>;
}

/**
 * 슬라이드 레이아웃 XML 및 관계 분석 (배경 이미지, 마스터 장식 선, 브랜딩 로고 추출)
 */
async function parseSlideLayout(
  zip: JSZip,
  layoutPath: string,
  slideWidth: number,
  slideHeight: number,
  themeMap?: Record<string, string>
): Promise<SlideLayoutData> {
  const lFile = zip.file(layoutPath);
  if (!lFile) {
    return { elements: [], phStyles: new Map() };
  }

  const lXmlText = await lFile.async('text');
  const lDoc = parseXml(lXmlText);

  // Layout rels
  const pathParts = layoutPath.split('/');
  const lFileName = pathParts.pop() || '';
  const relsPath = `${pathParts.join('/')}/_rels/${lFileName}.rels`;
  const relsFile = zip.file(relsPath);

  const layoutImageMap = new Map<string, string>();
  if (relsFile) {
    try {
      const sRelsText = await relsFile.async('text');
      const sRelsDoc = parseXml(sRelsText);
      const rels = getElementsByLocalName(sRelsDoc, 'Relationship');
      for (const r of rels) {
        const type = r.getAttribute('Type') || '';
        if (type.includes('image')) {
          const id = r.getAttribute('Id');
          let target = r.getAttribute('Target') || '';
          if (id && target) {
            if (target.startsWith('../')) {
              target = `ppt/${target.substring(3)}`;
            } else if (!target.startsWith('ppt/')) {
              target = `ppt/slideLayouts/${target}`;
            }
            layoutImageMap.set(id, target);
          }
        }
      }
    } catch (e) {
      console.warn(`레이아웃 rels 파싱 실패 (${relsPath}):`, e);
    }
  }

  const elements: PptxSlideElement[] = [];
  const phStyles = new Map<string, { align?: 'left' | 'center' | 'right' }>();
  let lZOrder = 1;

  const spTree = findFirstByLocalName(lDoc, 'spTree');
  const shapeContainers: Element[] = spTree ? Array.from(spTree.children) : [];

  for (const item of shapeContainers) {
    const localName = item.localName;

    // 플레이스홀더 여부 확인
    const phList = getElementsByLocalName(item, 'ph');
    const hasPh = phList.length > 0;
    const phType = hasPh ? phList[0].getAttribute('type') || 'body' : '';

    if (hasPh) {
      let align: 'left' | 'center' | 'right' | undefined;
      const txBodyList = getElementsByLocalName(item, 'txBody');
      if (txBodyList.length > 0) {
        const lvl1List = getElementsByLocalName(txBodyList[0], 'lvl1pPr');
        if (lvl1List.length > 0) {
          const algnAttr = lvl1List[0].getAttribute('algn');
          if (algnAttr === 'ctr') align = 'center';
          else if (algnAttr === 'r') align = 'right';
          else if (algnAttr === 'l') align = 'left';
        }
      }
      if (align) {
        phStyles.set(phType, { align });
      }
      continue;
    }

    // 정적 레이아웃 그래픽 요소들
    if (localName === 'pic') {
      const coords = parseElementCoordinates(item, slideWidth, slideHeight);
      const blipList = getElementsByLocalName(item, 'blip');
      const cNvPrList = getElementsByLocalName(item, 'cNvPr');
      const picName = cNvPrList.length > 0 ? cNvPrList[0].getAttribute('name') || '' : '';

      if (coords && blipList.length > 0) {
        const embedId = blipList[0].getAttribute('r:embed');
        if (embedId && layoutImageMap.has(embedId)) {
          const mediaPath = layoutImageMap.get(embedId)!;
          const mediaResult = await loadMediaAsDataUrl(zip, mediaPath, picName);
          if (mediaResult) {
            elements.push({
              id: `layout-img-${lZOrder}`,
              type: 'image',
              left: coords.left,
              top: coords.top,
              width: coords.width,
              height: coords.height,
              dataUrl: mediaResult.dataUrl,
              contentType: mediaResult.mime,
              zIndex: lZOrder++,
            });
          }
        }
      }
    } else if (localName === 'cxnSp') {
      const coords = parseElementCoordinates(item, slideWidth, slideHeight);
      if (coords) {
        let lineColor: string | undefined;
        let lineWidth: number | undefined;

        const spPrList = getElementsByLocalName(item, 'spPr');
        if (spPrList.length > 0) {
          const lnList = getElementsByLocalName(spPrList[0], 'ln');
          if (lnList.length > 0) {
            lineColor = parseDrawingColor(lnList[0], themeMap) || '#CBD5E1';
            const wAttr = lnList[0].getAttribute('w');
            lineWidth = wAttr ? Math.max(1, Math.round(parseInt(wAttr, 10) / 12700)) : 1;
          }
        }

        elements.push({
          id: `layout-cxn-${lZOrder}`,
          type: 'shape',
          shapeType: 'line',
          left: coords.left,
          top: coords.top,
          width: coords.width,
          height: coords.height,
          borderColor: lineColor || '#CBD5E1',
          borderWidth: lineWidth || 1,
          paragraphs: [],
          zIndex: lZOrder++,
        });
      }
    } else if (localName === 'sp') {
      const coords = parseElementCoordinates(item, slideWidth, slideHeight);
      if (coords) {
        let shapeType = 'rect';
        const prstGeomList = getElementsByLocalName(item, 'prstGeom');
        if (prstGeomList.length > 0) {
          shapeType = prstGeomList[0].getAttribute('prst') || 'rect';
        }
        if (coords.height <= 0.8 || shapeType.toLowerCase().includes('line')) {
          shapeType = 'line';
        }

        let fillColor: string | undefined;
        let borderColor: string | undefined;
        let borderWidth: number | undefined;

        const spPrList = getElementsByLocalName(item, 'spPr');
        if (spPrList.length > 0) {
          fillColor = parseDrawingColor(spPrList[0], themeMap);
          const lnList = getElementsByLocalName(spPrList[0], 'ln');
          if (lnList.length > 0) {
            borderColor = parseDrawingColor(lnList[0], themeMap) || '#CBD5E1';
            const wAttr = lnList[0].getAttribute('w');
            borderWidth = wAttr ? Math.max(1, Math.round(parseInt(wAttr, 10) / 12700)) : 1;
          }
        }

        elements.push({
          id: `layout-sp-${lZOrder}`,
          type: 'shape',
          shapeType,
          left: coords.left,
          top: coords.top,
          width: coords.width,
          height: coords.height,
          fillColor,
          borderColor,
          borderWidth,
          paragraphs: [],
          zIndex: lZOrder++,
        });
      }
    }
  }

  return { elements, phStyles };
}

// PPTX 파일 파싱 (2차원 절대 좌표 및 벡터 그래픽 복원)
export async function parsePptxPresentation(dataUrl: string): Promise<PptxPresentation> {
  const binary = dataUrlToBinary(dataUrl);
  const bytes = binaryToUint8Array(binary);
  const zip = await JSZip.loadAsync(bytes);

  // 1. 슬라이드 크기 및 해상도 확인 (presentation.xml)
  let slideWidth = 12192000; // 16:9 기본 너비 (EMU)
  let slideHeight = 6858000; // 16:9 기본 높이 (EMU)
  let slidePaths: string[] = [];

  const presFile = zip.file('ppt/presentation.xml');
  const presRelsFile = zip.file('ppt/_rels/presentation.xml.rels');

  if (presFile) {
    try {
      const presText = await presFile.async('text');
      const presDoc = parseXml(presText);

      // sldSz (슬라이드 크기 및 방향 속성)
      const sldSzList = getElementsByLocalName(presDoc, 'sldSz');
      if (sldSzList.length > 0) {
        let cx = parseInt(sldSzList[0].getAttribute('cx') || '0', 10);
        let cy = parseInt(sldSzList[0].getAttribute('cy') || '0', 10);
        const orient = (sldSzList[0].getAttribute('orient') || '').toLowerCase();

        // orient 속성이 명시된 경우 가로/세로 방향 보정
        if (orient === 'portrait' && cx > cy) {
          const tmp = cx;
          cx = cy;
          cy = tmp;
        } else if (orient === 'landscape' && cy > cx) {
          const tmp = cx;
          cx = cy;
          cy = tmp;
        }

        if (cx > 0 && cy > 0) {
          slideWidth = cx;
          slideHeight = cy;
        }
      }

      // 슬라이드 순서 및 관계 매핑
      if (presRelsFile) {
        const relsText = await presRelsFile.async('text');
        const relsDoc = parseXml(relsText);

        const relMap = new Map<string, string>();
        const relElements = getElementsByLocalName(relsDoc, 'Relationship');
        for (const rel of relElements) {
          const id = rel.getAttribute('Id');
          const target = rel.getAttribute('Target');
          if (id && target) {
            const normalized = target.startsWith('ppt/') ? target : `ppt/${target.replace(/^\//, '')}`;
            relMap.set(id, normalized);
          }
        }

        const sldIdElements = getElementsByLocalName(presDoc, 'sldId');
        for (const sld of sldIdElements) {
          const rId = sld.getAttribute('r:id') || sld.getAttribute('id');
          if (rId && relMap.has(rId)) {
            slidePaths.push(relMap.get(rId)!);
          }
        }
      }
    } catch (e) {
      console.warn('presentation.xml 파싱 오류:', e);
    }
  }

  // 슬라이드 경로 fallback
  if (slidePaths.length === 0) {
    const allFiles = Object.keys(zip.files);
    const regex = /^ppt\/slides\/slide(\d+)\.xml$/;
    const slideEntries: { path: string; num: number }[] = [];

    for (const f of allFiles) {
      const match = f.match(regex);
      if (match) {
        slideEntries.push({ path: f, num: parseInt(match[1], 10) });
      }
    }

    slideEntries.sort((a, b) => a.num - b.num);
    slidePaths = slideEntries.map((e) => e.path);
  }

  const aspectRatio = slideWidth / slideHeight;
  const slides: PptxSlide[] = [];
  let presentationTitle = '';
  const layoutCache = new Map<string, SlideLayoutData>();
  const themeMap = await parseThemeColors(zip);

  // 2. 각 슬라이드 정밀 파싱
  for (let idx = 0; idx < slidePaths.length; idx++) {
    const sPath = slidePaths[idx];
    const sFile = zip.file(sPath);
    if (!sFile) continue;

    const sXmlText = await sFile.async('text');
    const sDoc = parseXml(sXmlText);

    // 슬라이드 관계(이미지 및 레이아웃) 매핑
    const pathParts = sPath.split('/');
    const sFileName = pathParts.pop() || '';
    const relsPath = `${pathParts.join('/')}/_rels/${sFileName}.rels`;
    const relsFile = zip.file(relsPath);

    const slideImageMap = new Map<string, string>();
    let layoutPath: string | undefined;

    if (relsFile) {
      try {
        const sRelsText = await relsFile.async('text');
        const sRelsDoc = parseXml(sRelsText);
        const rels = getElementsByLocalName(sRelsDoc, 'Relationship');
        for (const r of rels) {
          const type = r.getAttribute('Type') || '';
          const id = r.getAttribute('Id');
          let target = r.getAttribute('Target') || '';

          if (type.includes('image') && id && target) {
            if (target.startsWith('../')) {
              target = `ppt/${target.substring(3)}`;
            } else if (!target.startsWith('ppt/')) {
              target = `ppt/slides/${target}`;
            }
            slideImageMap.set(id, target);
          } else if (type.includes('slideLayout') && target) {
            layoutPath = target.startsWith('../')
              ? `ppt/${target.substring(3)}`
              : target.startsWith('ppt/')
              ? target
              : `ppt/slides/${target}`;
          }
        }
      } catch (e) {
        console.warn(`슬라이드 rels 파싱 오류 (${relsPath}):`, e);
      }
    }

    // 슬라이드 레이아웃 데이터 가져오기
    let layoutData: SlideLayoutData = { elements: [], phStyles: new Map() };
    if (layoutPath) {
      if (!layoutCache.has(layoutPath)) {
        const parsedLayout = await parseSlideLayout(zip, layoutPath, slideWidth, slideHeight, themeMap);
        layoutCache.set(layoutPath, parsedLayout);
      }
      layoutData = layoutCache.get(layoutPath)!;
    }

    // 슬라이드 배경색
    let backgroundColor: string | undefined;
    const bgList = getElementsByLocalName(sDoc, 'bg');
    if (bgList.length > 0) {
      backgroundColor = parseDrawingColor(bgList[0], themeMap);
    }

    const elements: PptxSlideElement[] = [];
    let zOrder = 1;

    // 1. 레이아웃 고정 요소 (배경 이미지, 브랜딩 로고, 헤더 장식 선 등)
    for (const lEl of layoutData.elements) {
      elements.push({
        ...lEl,
        id: `${lEl.id}-s${idx + 1}`,
        zIndex: zOrder++,
      });
    }

    // spTree 내의 모든 엘리먼트 순차 탐색
    const spTree = findFirstByLocalName(sDoc, 'spTree');
    const shapeContainers: Element[] = spTree ? Array.from(spTree.children) : [];

    // 편의용 레거시 목록
    let slideTitle = '';
    const legacyParagraphs: PptxParagraph[] = [];
    const legacyTables: { headers: string[]; rows: string[][] }[] = [];
    const legacyImages: { id: string; dataUrl: string; contentType: string }[] = [];

    for (const item of shapeContainers) {
      const localName = item.localName;

      // 1. 도형 및 텍스트 상자 (<p:sp>)
      if (localName === 'sp') {
        const coords = parseElementCoordinates(item, slideWidth, slideHeight);
        if (!coords) continue;

        let shapeType = 'rect';
        let isRounded = false;
        const prstGeomList = getElementsByLocalName(item, 'prstGeom');
        if (prstGeomList.length > 0) {
          shapeType = prstGeomList[0].getAttribute('prst') || 'rect';
          if (shapeType.toLowerCase().includes('round')) {
            isRounded = true;
          }
        }
        if (coords.height <= 0.8 || shapeType.toLowerCase().includes('line')) {
          shapeType = 'line';
        }

        let fillColor: string | undefined;
        let borderColor: string | undefined;
        let borderWidth: number | undefined;

        const spPrList = getElementsByLocalName(item, 'spPr');
        if (spPrList.length > 0) {
          const spPr = spPrList[0];
          fillColor = parseDrawingColor(spPr, themeMap);

          const lnList = getElementsByLocalName(spPr, 'ln');
          if (lnList.length > 0) {
            borderColor = parseDrawingColor(lnList[0], themeMap) || '#CBD5E1';
            const wAttr = lnList[0].getAttribute('w');
            borderWidth = wAttr ? Math.max(1, Math.round(parseInt(wAttr, 10) / 12700)) : 1;
          }
        }

        // spPr에 배경색이 없다면 <p:style> -> <a:fillRef> 확인
        if (!fillColor) {
          const styleList = getElementsByLocalName(item, 'style');
          if (styleList.length > 0) {
            const fillRefList = getElementsByLocalName(styleList[0], 'fillRef');
            if (fillRefList.length > 0) {
              fillColor = parseDrawingColor(fillRefList[0], themeMap);
            }
          }
        }

        // spPr에 테두리가 없다면 <p:style> -> <a:lnRef> 확인
        if (!borderColor) {
          const styleList = getElementsByLocalName(item, 'style');
          if (styleList.length > 0) {
            const lnRefList = getElementsByLocalName(styleList[0], 'lnRef');
            if (lnRefList.length > 0) {
              borderColor = parseDrawingColor(lnRefList[0], themeMap);
              if (borderColor) borderWidth = 1.5;
            }
          }
        }

        let verticalAlign: 'top' | 'middle' | 'bottom' = 'top';
        const txBodyList = getElementsByLocalName(item, 'txBody');
        if (txBodyList.length > 0) {
          const bodyPrList = getElementsByLocalName(txBodyList[0], 'bodyPr');
          if (bodyPrList.length > 0) {
            const anchor = bodyPrList[0].getAttribute('anchor');
            if (anchor === 'ctr' || anchor === 'mid') verticalAlign = 'middle';
            else if (anchor === 'b') verticalAlign = 'bottom';
          }
        }

        // 제목 여부 판별 및 기본 정렬 상속
        let isTitle = false;
        let defaultAlign: 'left' | 'center' | 'right' | 'justify' | undefined;
        const phList = getElementsByLocalName(item, 'ph');
        if (phList.length > 0) {
          const phType = phList[0].getAttribute('type') || '';
          if (phType === 'title' || phType === 'ctrTitle') {
            isTitle = true;
          }
          const inherited = layoutData.phStyles.get(phType);
          if (inherited?.align) {
            defaultAlign = inherited.align;
          }
        }
        // 명시적 ph 태그가 없더라도, 슬라이드 상단에 위치한 첫 대형 텍스트 상자면 제목으로 인식
        if (!isTitle && !slideTitle && coords.top <= 16 && coords.width >= 35) {
          isTitle = true;
        }

        // 표지 슬라이드(1페이지)의 텍스트는 기본적으로 중앙 정렬
        if (idx === 0) {
          defaultAlign = defaultAlign || 'center';
        }

        // 텍스트 단락들 파싱
        const paragraphs: PptxParagraph[] = [];
        if (txBodyList.length > 0) {
          const pElements = getDirectChildrenByLocalName(txBodyList[0], 'p');
          const titleParts: string[] = [];

          for (const p of pElements) {
            const parsedP = parseParagraph(p, defaultAlign);
            if (parsedP.text.trim()) {
              if (isTitle) {
                parsedP.isTitle = true;
                titleParts.push(parsedP.text.trim());
              } else {
                legacyParagraphs.push(parsedP);
              }
              paragraphs.push(parsedP);
            } else {
              // 빈 단락도 줄바꿈/간격 유지를 위해 보존
              paragraphs.push(parsedP);
            }
          }

          if (isTitle && titleParts.length > 0 && !slideTitle) {
            slideTitle = titleParts.join(' ');
            if (idx === 0 && !presentationTitle) {
              presentationTitle = slideTitle;
            }
          }

          // 양 끝(앞뒤)의 불필요한 빈 단락만 정리 (중간 빈 줄은 보존)
          while (paragraphs.length > 0 && !paragraphs[0].text.trim()) {
            paragraphs.shift();
          }
          while (paragraphs.length > 0 && !paragraphs[paragraphs.length - 1].text.trim()) {
            paragraphs.pop();
          }
        }

        elements.push({
          id: `sp-${zOrder}`,
          type: 'shape',
          shapeType,
          isRounded,
          left: coords.left,
          top: coords.top,
          width: coords.width,
          height: coords.height,
          fillColor,
          borderColor,
          borderWidth,
          paragraphs,
          isTitle,
          verticalAlign,
          zIndex: zOrder++,
        });
      }

      // 2. 표 (<p:graphicFrame> -> <a:tbl>)
      else if (localName === 'graphicFrame') {
        const coords = parseElementCoordinates(item, slideWidth, slideHeight);
        const tblList = getElementsByLocalName(item, 'tbl');

        if (coords && tblList.length > 0) {
          const tbl = tblList[0];

          // 열 너비 백분율
          const gridCols = getElementsByLocalName(tbl, 'gridCol');
          const colWidthsEmu = gridCols.map((gc) => parseInt(gc.getAttribute('w') || '1000000', 10));
          const totalColsEmu = colWidthsEmu.reduce((a, b) => a + b, 0) || 1;
          const colWidthsPct = colWidthsEmu.map((w) => (w / totalColsEmu) * 100);

          // 행 데이터
          const trElements = getElementsByLocalName(tbl, 'tr');
          const rows: PptxTableRow[] = [];
          const legacyRows: string[][] = [];
          const tableHeightEmu = slideHeight > 0 ? (coords.height / 100) * slideHeight : 1;

          for (const tr of trElements) {
            const rawH = parseInt(tr.getAttribute('h') || '0', 10);
            const rowHeightPct = tableHeightEmu > 0 ? (rawH / tableHeightEmu) * 100 : undefined;

            const cells: PptxTableCell[] = [];
            const legacyCellTexts: string[] = [];
            const tcElements = getElementsByLocalName(tr, 'tc');

            for (const tc of tcElements) {
              // 1. 병합 검사: hMerge 및 vMerge는 이전/상단 셀에 포함되므로 렌더링에서 제외
              const hMerge = tc.getAttribute('hMerge');
              const vMerge = tc.getAttribute('vMerge');
              if (hMerge === '1' || hMerge === 'true' || vMerge === '1' || vMerge === 'true') {
                continue;
              }

              const colSpan = parseInt(tc.getAttribute('gridSpan') || '1', 10);
              const rowSpan = parseInt(tc.getAttribute('rowSpan') || '1', 10);

              let fillColor: string | undefined;
              let verticalAlign: 'top' | 'middle' | 'bottom' = 'top';

              const tcPrList = getElementsByLocalName(tc, 'tcPr');
              if (tcPrList.length > 0) {
                const tcPr = tcPrList[0];
                const anchor = tcPr.getAttribute('anchor');
                if (anchor === 'ctr' || anchor === 'mid') verticalAlign = 'middle';
                else if (anchor === 'b') verticalAlign = 'bottom';
                else verticalAlign = 'top';

                // 중요: tcPr 직속의 solidFill만 추출하여 테두리 선(lnL, lnR 등)의 tx1(검정색) 오인 방지
                const solidFillList = getDirectChildrenByLocalName(tcPr, 'solidFill');
                if (solidFillList.length > 0) {
                  fillColor = parseDrawingColor(solidFillList[0], themeMap);
                }
              }

              const paragraphs: PptxParagraph[] = [];
              const pElements = getElementsByLocalName(tc, 'p');

              for (const p of pElements) {
                const parsedP = parseParagraph(p);
                paragraphs.push(parsedP);
              }

              // 셀 앞뒤의 불필요한 빈 단락 정리 (중간 빈 줄은 보존)
              while (paragraphs.length > 0 && !paragraphs[0].text.trim()) {
                paragraphs.shift();
              }
              while (paragraphs.length > 0 && !paragraphs[paragraphs.length - 1].text.trim()) {
                paragraphs.pop();
              }

              const cellText = paragraphs.map((p) => p.text).join('\n');

              cells.push({
                text: cellText,
                paragraphs,
                colSpan,
                rowSpan,
                fillColor,
                verticalAlign,
              });
              legacyCellTexts.push(cellText);
            }

            if (cells.length > 0) {
              rows.push({ cells, height: rowHeightPct });
              legacyRows.push(legacyCellTexts);
            }
          }

          if (legacyRows.length > 0) {
            legacyTables.push({
              headers: legacyRows[0] || [],
              rows: legacyRows.slice(1),
            });
          }

          elements.push({
            id: `tbl-${zOrder}`,
            type: 'table',
            left: coords.left,
            top: coords.top,
            width: coords.width,
            height: coords.height,
            colWidths: colWidthsPct,
            rows,
            zIndex: zOrder++,
          });
        }
      }

      // 3. 그림 / 이미지 (<p:pic>)
      else if (localName === 'pic') {
        const coords = parseElementCoordinates(item, slideWidth, slideHeight);
        const blipList = getElementsByLocalName(item, 'blip');
        const cNvPrList = getElementsByLocalName(item, 'cNvPr');
        const picName = cNvPrList.length > 0 ? cNvPrList[0].getAttribute('name') || '' : '';

        if (coords && blipList.length > 0) {
          const embedId = blipList[0].getAttribute('r:embed');
          if (embedId && slideImageMap.has(embedId)) {
            const mediaPath = slideImageMap.get(embedId)!;
            try {
              const mediaResult = await loadMediaAsDataUrl(zip, mediaPath, picName);
              if (mediaResult) {
                elements.push({
                  id: `img-${zOrder}`,
                  type: 'image',
                  left: coords.left,
                  top: coords.top,
                  width: coords.width,
                  height: coords.height,
                  dataUrl: mediaResult.dataUrl,
                  contentType: mediaResult.mime,
                  zIndex: zOrder++,
                });

                legacyImages.push({
                  id: embedId,
                  dataUrl: mediaResult.dataUrl,
                  contentType: mediaResult.mime,
                });
              }
            } catch (e) {
              console.warn(`이미지 로드 실패: ${mediaPath}`, e);
            }
          }
        }
      }

      // 4. 연결선 / 구분선 (<p:cxnSp>)
      else if (localName === 'cxnSp') {
        const coords = parseElementCoordinates(item, slideWidth, slideHeight);
        if (coords) {
          let lineColor: string | undefined;
          let lineWidth: number | undefined;

          const spPrList = getElementsByLocalName(item, 'spPr');
          if (spPrList.length > 0) {
            const lnList = getElementsByLocalName(spPrList[0], 'ln');
            if (lnList.length > 0) {
              lineColor = parseDrawingColor(lnList[0], themeMap) || '#334155';
              const wAttr = lnList[0].getAttribute('w');
              lineWidth = wAttr ? Math.max(1, Math.round(parseInt(wAttr, 10) / 12700)) : 1;
            }
          }

          elements.push({
            id: `cxn-${zOrder}`,
            type: 'shape',
            shapeType: 'line',
            left: coords.left,
            top: coords.top,
            width: coords.width,
            height: coords.height,
            borderColor: lineColor || '#334155',
            borderWidth: lineWidth || 1,
            paragraphs: [],
            zIndex: zOrder++,
          });
        }
      }
    }

    if (!slideTitle) {
      if (legacyParagraphs.length > 0) {
        slideTitle = legacyParagraphs[0].text;
      } else {
        slideTitle = `슬라이드 ${idx + 1}`;
      }
    }

    slides.push({
      slideNumber: idx + 1,
      title: slideTitle,
      slidePath: sPath,
      width: slideWidth,
      height: slideHeight,
      aspectRatio,
      backgroundColor,
      elements,
      paragraphs: legacyParagraphs,
      tables: legacyTables,
      images: legacyImages,
    });
  }

  return {
    title: presentationTitle || (slides[0]?.title ?? '파워포인트 프레젠테이션'),
    slideCount: slides.length,
    slideWidth,
    slideHeight,
    aspectRatio,
    isLandscape: slideWidth >= slideHeight,
    orientation: slideWidth >= slideHeight ? 'landscape' : 'portrait',
    slides,
    rawZip: zip,
  };
}

// 캐시 기반 파싱
export async function getCachedPptxPresentation(attId: string, dataUrl: string): Promise<PptxPresentation> {
  const cached = presentationCache.get(attId);
  if (cached && Date.now() - cached.timestamp < 1000 * 60 * 30) {
    return cached.presentation;
  }
  const presentation = await parsePptxPresentation(dataUrl);
  presentationCache.set(attId, { presentation, timestamp: Date.now() });
  return presentation;
}

// 마크다운 및 텍스트 추출 (메모로 가져오기용)
export async function extractPptxTextAndMarkdown(dataUrl: string): Promise<{ title: string; markdown: string; text: string }> {
  const pres = await parsePptxPresentation(dataUrl);

  const mdLines: string[] = [];
  const textLines: string[] = [];

  mdLines.push(`# 📊 ${pres.title || '파워포인트 발표 자료'}\n`);
  textLines.push(`${pres.title || '파워포인트 발표 자료'}\n`);

  for (const slide of pres.slides) {
    mdLines.push(`\n## 📄 슬라이드 ${slide.slideNumber}: ${slide.title}\n`);
    textLines.push(`\n[슬라이드 ${slide.slideNumber}] ${slide.title}\n`);

    for (const p of slide.paragraphs) {
      if (p.isTitle && p.text === slide.title) continue;
      const indent = '  '.repeat(Math.max(0, p.level));
      mdLines.push(`${indent}- ${p.text}`);
      textLines.push(`${indent}• ${p.text}`);
    }

    for (const tbl of slide.tables) {
      if (tbl.headers.length > 0) {
        mdLines.push('\n| ' + tbl.headers.join(' | ') + ' |');
        mdLines.push('| ' + tbl.headers.map(() => '---').join(' | ') + ' |');
        for (const row of tbl.rows) {
          mdLines.push('| ' + row.join(' | ') + ' |');
        }
        mdLines.push('');
      }
    }
  }

  return {
    title: pres.title,
    markdown: mdLines.join('\n'),
    text: textLines.join('\n'),
  };
}

export interface UpdatedSlideContent {
  slideNumber: number;
  title: string;
  bodyText: string;
  tables?: { rows: string[][] }[];
}

// 슬라이드 텍스트 및 표 내용 수정 후 새로운 PPTX Data URL 생성
export async function updatePptxSlideTexts(
  originalDataUrl: string,
  updatedSlides: UpdatedSlideContent[]
): Promise<string> {
  const binary = dataUrlToBinary(originalDataUrl);
  const bytes = binaryToUint8Array(binary);
  const zip = await JSZip.loadAsync(bytes);

  for (const upd of updatedSlides) {
    const slidePath = `ppt/slides/slide${upd.slideNumber}.xml`;
    const sFile = zip.file(slidePath);
    if (!sFile) continue;

    const xmlText = await sFile.async('text');
    const doc = parseXml(xmlText);

    // 1. 슬라이드 내 제목 도형 찾기
    const spList = getElementsByLocalName(doc, 'sp');
    let titleSp: Element | null = null;
    for (const sp of spList) {
      const phList = getElementsByLocalName(sp, 'ph');
      if (phList.length > 0) {
        const phType = phList[0].getAttribute('type');
        if (phType === 'title' || phType === 'ctrTitle') {
          titleSp = sp;
          break;
        }
      }
    }
    if (!titleSp && spList.length > 0) {
      titleSp = spList[0];
    }

    // 제목 업데이트
    if (titleSp) {
      const tElements = getElementsByLocalName(titleSp, 't');
      if (tElements.length > 0) {
        tElements[0].textContent = upd.title;
        for (let i = 1; i < tElements.length; i++) {
          tElements[i].textContent = '';
        }
      }
    }

    // 2. 본문 도형 텍스트 업데이트
    const otherSpList = spList.filter((sp) => sp !== titleSp);
    const bodyLines = upd.bodyText
      .split('\n')
      .map((l) => l.trim().replace(/^[-*•]\s*/, ''))
      .filter(Boolean);

    if (otherSpList.length > 0 && bodyLines.length > 0) {
      const bodySp = otherSpList[0];
      const pElements = getElementsByLocalName(bodySp, 'p');
      for (let i = 0; i < bodyLines.length; i++) {
        if (i < pElements.length) {
          const p = pElements[i];
          const pTs = getElementsByLocalName(p, 't');
          if (pTs.length > 0) {
            pTs[0].textContent = bodyLines[i];
            for (let k = 1; k < pTs.length; k++) pTs[k].textContent = '';
          }
        }
      }
    }

    // 3. 표(Table) 셀 데이터 업데이트
    if (upd.tables && upd.tables.length > 0) {
      const tblList = getElementsByLocalName(doc, 'tbl');
      for (let tblIdx = 0; tblIdx < upd.tables.length; tblIdx++) {
        if (tblIdx >= tblList.length) break;
        const tbl = tblList[tblIdx];
        const trList = getElementsByLocalName(tbl, 'tr');
        const updatedRows = upd.tables[tblIdx].rows;

        for (let rIdx = 0; rIdx < updatedRows.length; rIdx++) {
          if (rIdx >= trList.length) break;
          const tr = trList[rIdx];
          const tcList = getElementsByLocalName(tr, 'tc');
          const updatedCells = updatedRows[rIdx];

          for (let cIdx = 0; cIdx < updatedCells.length; cIdx++) {
            if (cIdx >= tcList.length) break;
            const tc = tcList[cIdx];
            const newText = updatedCells[cIdx];
            const tcTs = getElementsByLocalName(tc, 't');
            if (tcTs.length > 0) {
              tcTs[0].textContent = newText;
              for (let k = 1; k < tcTs.length; k++) {
                tcTs[k].textContent = '';
              }
            }
          }
        }
      }
    }

    const serializer = new XMLSerializer();
    const updatedXml = serializer.serializeToString(doc);
    zip.file(slidePath, updatedXml);
  }

  const updatedBase64 = await zip.generateAsync({
    type: 'base64',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });

  return `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${updatedBase64}`;
}
