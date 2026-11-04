import init, { HwpDocument } from '@rhwp/core';
import wasmUrl from '@rhwp/core/rhwp_bg.wasm?url';

let isInitialized = false;
let initPromise: Promise<any> | null = null;

// rhwp WASM이 텍스트 너비를 계산할 때 사용하는 필수 콜백 등록
let measureCtx: CanvasRenderingContext2D | null = null;
let lastMeasureFont = '';

function ensureMeasureTextRegistered() {
  if (typeof window === 'undefined') return;
  if (!(globalThis as any).measureTextWidth) {
    (globalThis as any).measureTextWidth = (font: string, text: string) => {
      if (!measureCtx) {
        const c = document.createElement('canvas');
        measureCtx = c.getContext('2d');
      }
      if (measureCtx) {
        if (font !== lastMeasureFont) {
          measureCtx.font = font;
          lastMeasureFont = font;
        }
        return measureCtx.measureText(text || '').width;
      }
      return (text || '').length * 8;
    };
  }
}

export async function initHwpWasm(): Promise<void> {
  if (isInitialized) return;
  if (initPromise) return initPromise;

  ensureMeasureTextRegistered();

  initPromise = (async () => {
    try {
      await init({ module_or_path: wasmUrl });
      isInitialized = true;
    } catch (err) {
      console.error('rhwp WASM 초기화 실패:', err);
      initPromise = null;
      throw err;
    }
  })();

  return initPromise;
}

const hwpDocCache = new Map<string, HwpDocument>();

export async function loadHwpDocument(dataOrUrl: string, cacheKey?: string): Promise<HwpDocument> {
  const key = cacheKey || dataOrUrl.slice(0, 80);
  if (hwpDocCache.has(key)) {
    return hwpDocCache.get(key)!;
  }

  await initHwpWasm();

  let bytes: Uint8Array;
  if (dataOrUrl.startsWith('data:')) {
    const commaIdx = dataOrUrl.indexOf(',');
    const base64 = commaIdx !== -1 ? dataOrUrl.slice(commaIdx + 1) : dataOrUrl;
    const binary = atob(base64);
    bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
  } else if (dataOrUrl.startsWith('blob:') || dataOrUrl.startsWith('http')) {
    const resp = await fetch(dataOrUrl);
    const buf = await resp.arrayBuffer();
    bytes = new Uint8Array(buf);
  } else {
    const binary = atob(dataOrUrl);
    bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
  }

  const doc = new HwpDocument(bytes);
  hwpDocCache.set(key, doc);
  return doc;
}

/**
 * HWP 문서의 특정 페이지를 SVG 문자열로 렌더링합니다.
 */
export function renderHwpPageToSvg(doc: HwpDocument, pageIndex: number): string {
  try {
    return doc.renderPageSvg(pageIndex);
  } catch (err) {
    console.error('HWP 페이지 SVG 렌더링 실패:', err);
    throw err;
  }
}

/**
 * HWP 문서의 본문 텍스트를 추출하여 마크다운 메모에 바로 삽입할 수 있는 텍스트로 변환합니다.
 */
export function extractHwpText(doc: HwpDocument): string {
  try {
    let rawText = '';
    if (typeof doc.getTextFileUnicode === 'function') {
      rawText = doc.getTextFileUnicode();
    } else if (typeof doc.getTextFileText === 'function') {
      rawText = doc.getTextFileText();
    }
    return rawText.trim();
  } catch (err) {
    console.error('HWP 텍스트 추출 실패:', err);
    return '';
  }
}

/**
 * SVG 문자열을 90도 단위(0, 90, 180, 270도)로 시계방향 회전하여
 * viewBox 및 내부 좌표계를 완벽히 변환합니다.
 */
export function rotateSvg(svgStr: string, degrees: number): string {
  const normalizedDeg = ((degrees % 360) + 360) % 360;
  if (normalizedDeg === 0) return svgStr;

  const vbMatch = svgStr.match(/viewBox=["']\s*([-\d.]+)\s+([-\d.]+)\s+([\d.]+)\s+([\d.]+)\s*["']/i);
  if (!vbMatch) return svgStr;

  const w = parseFloat(vbMatch[3]);
  const h = parseFloat(vbMatch[4]);

  const openTagEnd = svgStr.indexOf('>');
  if (openTagEnd === -1) return svgStr;

  const closeTagStart = svgStr.lastIndexOf('</svg>');
  if (closeTagStart === -1) return svgStr;

  let openTag = svgStr.slice(0, openTagEnd + 1);
  const innerContent = svgStr.slice(openTagEnd + 1, closeTagStart);

  let newVb = '';
  let transform = '';

  if (normalizedDeg === 90) {
    newVb = `0 0 ${h} ${w}`;
    transform = `rotate(90) translate(0, -${h})`;
  } else if (normalizedDeg === 180) {
    newVb = `0 0 ${w} ${h}`;
    transform = `rotate(180) translate(-${w}, -${h})`;
  } else if (normalizedDeg === 270) {
    newVb = `0 0 ${h} ${w}`;
    transform = `rotate(270) translate(-${w}, 0)`;
  }

  openTag = openTag.replace(/viewBox=["'][^"']*["']/i, `viewBox="${newVb}"`);
  openTag = openTag.replace(/\s+(?:width|height)=["'][^"']*["']/gi, '');

  return `${openTag}<g transform="${transform}">${innerContent}</g></svg>`;
}

