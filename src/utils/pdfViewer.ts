// TC39 Map/WeakMap getOrInsertComputed & getOrInsert 폴리필 (pdf.js v6 모바일 렌더링 필수)
if (typeof Map !== 'undefined') {
  if (!(Map.prototype as any).getOrInsertComputed) {
    (Map.prototype as any).getOrInsertComputed = function (key: any, callback: (k: any) => any) {
      if (this.has(key)) {
        return this.get(key);
      }
      const value = callback(key);
      this.set(key, value);
      return value;
    };
  }
  if (!(Map.prototype as any).getOrInsert) {
    (Map.prototype as any).getOrInsert = function (key: any, defaultValue: any) {
      if (this.has(key)) {
        return this.get(key);
      }
      this.set(key, defaultValue);
      return defaultValue;
    };
  }
}
if (typeof WeakMap !== 'undefined') {
  if (!(WeakMap.prototype as any).getOrInsertComputed) {
    (WeakMap.prototype as any).getOrInsertComputed = function (key: any, callback: (k: any) => any) {
      if (this.has(key)) {
        return this.get(key);
      }
      const value = callback(key);
      this.set(key, value);
      return value;
    };
  }
  if (!(WeakMap.prototype as any).getOrInsert) {
    (WeakMap.prototype as any).getOrInsert = function (key: any, defaultValue: any) {
      if (this.has(key)) {
        return this.get(key);
      }
      this.set(key, defaultValue);
      return defaultValue;
    };
  }
}

import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
// Vite URL import for pdf.worker
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';

if (typeof window !== 'undefined' && 'GlobalWorkerOptions' in pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;
}

// PDF 문서 캐시 (중복 파싱 방지)
const pdfDocCache = new Map<string, PDFDocumentProxy>();

/**
 * 특정 첨부파일 ID 또는 전체 PDF 캐시를 무효화(삭제)합니다.
 */
export function invalidatePdfCache(cacheKey?: string): void {
  if (cacheKey) {
    for (const key of Array.from(pdfDocCache.keys())) {
      if (key === cacheKey || key.startsWith(`${cacheKey}_`)) {
        try {
          const doc = pdfDocCache.get(key) as any;
          doc?.destroy?.();
          doc?.cleanup?.();
        } catch {
          // ignore
        }
        pdfDocCache.delete(key);
      }
    }
  } else {
    pdfDocCache.forEach((doc: any) => {
      try {
        doc?.destroy?.();
        doc?.cleanup?.();
      } catch {
        // ignore
      }
    });
    pdfDocCache.clear();
  }
}

// 진행 중인 캔버스 렌더링 태스크 추적 (다중 렌더링 충돌 취소 방지)
const activeRenderTasks = new WeakMap<HTMLCanvasElement, any>();

function getAssetBaseUrl(): string {
  if (typeof window === 'undefined') return './';
  let href = window.location.href.split('#')[0].split('?')[0];
  if (!href.endsWith('/')) {
    const lastPart = href.substring(href.lastIndexOf('/') + 1);
    if (!lastPart.includes('.')) {
      href += '/';
    } else {
      href = href.substring(0, href.lastIndexOf('/') + 1);
    }
  }
  return href;
}

export async function loadPdfDocument(dataOrUrl: string, cacheKey?: string): Promise<PDFDocumentProxy> {
  // 데이터 내용의 길이와 서명을 포함시켜 내용이 수정되면 자동으로 새 캐시 키가 생성되도록 보장
  const dataSignature = `${dataOrUrl.length}_${dataOrUrl.slice(-32)}`;
  const key = cacheKey ? `${cacheKey}_${dataSignature}` : `${dataSignature}_${dataOrUrl.slice(0, 40)}`;
  if (pdfDocCache.has(key)) {
    return pdfDocCache.get(key)!;
  }

  const baseUrl = getAssetBaseUrl();
  const cMapUrl = `${baseUrl}cmaps/`;
  const standardFontDataUrl = `${baseUrl}standard_fonts/`;
  const wasmUrl = `${baseUrl}wasm/`;
  const iccUrl = `${baseUrl}iccs/`;

  // 한글(CJK) 및 스캔 문서(JBIG2, OpenJPEG) 유실 방지를 위한 필수 디코더 파라미터
  const commonParams = {
    cMapUrl,
    cMapPacked: true,
    standardFontDataUrl,
    wasmUrl,
    iccUrl,
    enableXfa: true,
  };

  let loadingTask;
  if (dataOrUrl.startsWith('data:')) {
    const commaIdx = dataOrUrl.indexOf(',');
    const base64 = commaIdx !== -1 ? dataOrUrl.slice(commaIdx + 1) : dataOrUrl;
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    loadingTask = pdfjsLib.getDocument({
      data: bytes,
      ...commonParams,
    });
  } else {
    loadingTask = pdfjsLib.getDocument({
      url: dataOrUrl,
      ...commonParams,
    });
  }

  const doc = await loadingTask.promise;
  pdfDocCache.set(key, doc);
  return doc;
}

export async function renderPdfPageToCanvas(
  pdfDoc: PDFDocumentProxy,
  pageNumber: number,
  canvas: HTMLCanvasElement,
  targetWidth?: number,
  zoomScale: number = 1.0,
  rotationOffset: number = 0
): Promise<{ width: number; height: number }> {
  // 이전 렌더링 태스크가 진행 중이면 취소하여 캔버스 충돌 방지
  if (activeRenderTasks.has(canvas)) {
    try {
      activeRenderTasks.get(canvas).cancel();
    } catch {
      // ignore
    }
    activeRenderTasks.delete(canvas);
  }

  const page = await pdfDoc.getPage(pageNumber);
  const effectiveRotation = (((page.rotate || 0) + rotationOffset) % 360 + 360) % 360;
  const unscaledViewport = page.getViewport({ scale: 1.0, rotation: effectiveRotation });

  // 화면 가로폭에 맞춘 반응형 스케일 계산
  let baseScale = 1.0;
  if (targetWidth && targetWidth > 50) {
    baseScale = (targetWidth - 10) / unscaledViewport.width;
  } else {
    baseScale = 1.2;
  }

  const finalScale = baseScale * zoomScale;
  const viewport = page.getViewport({ scale: finalScale, rotation: effectiveRotation });

  // 모바일 레티나 및 고해상도 모니터 선명도 보정 (안전 범위 2.0 제한)
  let dpr = Math.min(window.devicePixelRatio || 1, 2.0);
  
  // 모바일 브라우저(iOS Safari / Chrome 등)의 캔버스 최대 해상도 초과로 인한 캔버스 크래시 방지
  const maxCanvasDimension = 3840;
  let canvasW = Math.floor(viewport.width * dpr);
  let canvasH = Math.floor(viewport.height * dpr);
  if (canvasW > maxCanvasDimension || canvasH > maxCanvasDimension) {
    const scaleFactor = maxCanvasDimension / Math.max(canvasW, canvasH);
    canvasW = Math.floor(canvasW * scaleFactor);
    canvasH = Math.floor(canvasH * scaleFactor);
    dpr = dpr * scaleFactor;
  }

  canvas.width = canvasW;
  canvas.height = canvasH;
  canvas.style.width = `${Math.floor(viewport.width)}px`;
  canvas.style.height = `${Math.floor(viewport.height)}px`;

  const ctx = canvas.getContext('2d', { alpha: false });
  if (!ctx) return { width: viewport.width, height: viewport.height };

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  ctx.save();
  ctx.scale(dpr, dpr);

  const renderContext = {
    canvasContext: ctx,
    viewport: viewport,
  };

  const renderTask = page.render(renderContext as any);
  activeRenderTasks.set(canvas, renderTask);

  try {
    await (renderTask as any).promise;
  } catch (err: any) {
    if (err?.name === 'RenderingCancelledException') {
      return { width: viewport.width, height: viewport.height };
    }
    throw err;
  } finally {
    if (activeRenderTasks.get(canvas) === renderTask) {
      activeRenderTasks.delete(canvas);
    }
    ctx.restore();
  }

  return { width: viewport.width, height: viewport.height };
}

/**
 * PDF 문서 전체에서 텍스트를 페이지별로 추출하여 마크다운 포맷으로 반환
 */
export async function extractPdfText(pdfDoc: PDFDocumentProxy): Promise<string> {
  const pageTexts: string[] = [];
  const numPages = pdfDoc.numPages;

  for (let i = 1; i <= numPages; i++) {
    const page = await pdfDoc.getPage(i);
    const textContent = await page.getTextContent();
    let lastY: number | null = null;
    let pageStr = '';

    for (const item of textContent.items as any[]) {
      if ('str' in item) {
        if (lastY !== null && Math.abs(item.transform[5] - lastY) > 5) {
          pageStr += '\n';
        } else if (pageStr.length > 0 && !pageStr.endsWith(' ') && !pageStr.endsWith('\n')) {
          pageStr += ' ';
        }
        pageStr += item.str;
        lastY = item.transform[5];
      }
    }

    const trimmed = pageStr.trim();
    if (trimmed) {
      pageTexts.push(`#### [${i} 페이지]\n\n${trimmed}`);
    }
  }

  return pageTexts.join('\n\n---\n\n');
}
