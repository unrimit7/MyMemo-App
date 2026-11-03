import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
// Vite URL import for pdf.worker
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';

if (typeof window !== 'undefined' && 'GlobalWorkerOptions' in pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;
}

// PDF 문서 캐시 (중복 파싱 방지)
const pdfDocCache = new Map<string, PDFDocumentProxy>();

export async function loadPdfDocument(dataOrUrl: string, cacheKey?: string): Promise<PDFDocumentProxy> {
  const key = cacheKey || dataOrUrl.slice(0, 80);
  if (pdfDocCache.has(key)) {
    return pdfDocCache.get(key)!;
  }

  let loadingTask;
  if (dataOrUrl.startsWith('data:')) {
    const commaIdx = dataOrUrl.indexOf(',');
    const base64 = commaIdx !== -1 ? dataOrUrl.slice(commaIdx + 1) : dataOrUrl;
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    loadingTask = pdfjsLib.getDocument({ data: bytes });
  } else {
    loadingTask = pdfjsLib.getDocument({ url: dataOrUrl });
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

  // 모바일 레티나 및 고해상도 모니터 선명도 보정
  const dpr = Math.min(window.devicePixelRatio || 1, 2.5);
  canvas.width = Math.floor(viewport.width * dpr);
  canvas.height = Math.floor(viewport.height * dpr);
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

  await (page.render(renderContext as any) as any).promise;
  ctx.restore();

  return { width: viewport.width, height: viewport.height };
}
