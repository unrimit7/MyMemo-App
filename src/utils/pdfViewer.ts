import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
// Vite URL import for pdf.worker
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';

if (typeof window !== 'undefined' && 'GlobalWorkerOptions' in pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;
}

// PDF 문서 캐시 (중복 파싱 방지)
const pdfDocCache = new Map<string, PDFDocumentProxy>();

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
  const key = cacheKey || dataOrUrl.slice(0, 80);
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
  zoomScale: number = 1.0
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
  const unscaledViewport = page.getViewport({ scale: 1.0 });

  // 화면 가로폭에 맞춘 반응형 스케일 계산
  let baseScale = 1.0;
  if (targetWidth && targetWidth > 50) {
    baseScale = (targetWidth - 10) / unscaledViewport.width;
  } else {
    baseScale = 1.2;
  }

  const finalScale = baseScale * zoomScale;
  const viewport = page.getViewport({ scale: finalScale });

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
