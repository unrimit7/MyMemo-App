import React, { useState, useEffect, useRef, useCallback } from 'react';
import { X, ChevronLeft, ChevronRight, ZoomIn, ZoomOut, Download } from 'lucide-react';
import { loadPdfDocument, renderPdfPageToCanvas } from '../utils/pdfViewer';
import type { PDFDocumentProxy } from 'pdfjs-dist';

interface PdfViewerModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  initialPage?: number;
}

export const PdfViewerModal: React.FC<PdfViewerModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  initialPage = 1,
}) => {
  const [pdfDoc, setPdfDoc] = useState<PDFDocumentProxy | null>(null);
  const [currentPage, setCurrentPage] = useState<number>(initialPage);
  const [totalPages, setTotalPages] = useState<number>(1);
  const [zoom, setZoom] = useState<number>(1.0);
  const [renderedZoom, setRenderedZoom] = useState<number>(1.0);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [renderError, setRenderError] = useState<string | null>(null);

  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const debounceTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const lastTapRef = useRef<number>(0);

  // PDF 문서 로딩
  useEffect(() => {
    if (!isOpen || !data) return;

    let isMounted = true;
    setIsLoading(true);
    setRenderError(null);
    setCurrentPage(initialPage);
    setZoom(1.0);
    setRenderedZoom(1.0);

    loadPdfDocument(data)
      .then((doc) => {
        if (!isMounted) return;
        setPdfDoc(doc);
        setTotalPages(doc.numPages);
        setIsLoading(false);
      })
      .catch((err) => {
        if (!isMounted) return;
        console.error('PDF 로드 실패:', err);
        setRenderError('PDF 문서를 불러오는 데 실패했습니다.');
        setIsLoading(false);
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, data, initialPage]);

  // 줌 변경 시: 부드러운 GPU CSS 확대/축소 후 100ms 정지 시 고해상도 캔버스 재렌더링 (디바운스)
  useEffect(() => {
    if (debounceTimerRef.current) {
      clearTimeout(debounceTimerRef.current);
    }
    debounceTimerRef.current = setTimeout(() => {
      setRenderedZoom(zoom);
    }, 100);

    return () => {
      if (debounceTimerRef.current) clearTimeout(debounceTimerRef.current);
    };
  }, [zoom]);

  // 페이지 및 줌 배율 변경 시 캔버스 렌더링
  useEffect(() => {
    if (!isOpen || !pdfDoc || !canvasRef.current) return;

    let isMounted = true;
    const canvas = canvasRef.current;
    const containerWidth = containerRef.current?.clientWidth || window.innerWidth;
    const baseWidth = Math.min(Math.max(260, containerWidth - 24), 850);

    renderPdfPageToCanvas(pdfDoc, currentPage, canvas, baseWidth, renderedZoom)
      .catch((err) => {
        if (isMounted) {
          console.error('PDF 페이지 렌더링 실패:', err);
        }
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, pdfDoc, currentPage, renderedZoom]);

  // 터치 제스처 (스마트폰 두 손가락 핀치 줌: 20% 아주 작게 ~ 400% 아주 크게)
  const touchStartDistRef = useRef<number | null>(null);
  const touchStartZoomRef = useRef<number>(1.0);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const getDistance = (t1: Touch, t2: Touch) => {
      const dx = t1.clientX - t2.clientX;
      const dy = t1.clientY - t2.clientY;
      return Math.hypot(dx, dy);
    };

    const handleTouchStart = (e: TouchEvent) => {
      if (e.touches.length === 2) {
        touchStartDistRef.current = getDistance(e.touches[0], e.touches[1]);
        touchStartZoomRef.current = zoom;
      }
    };

    const handleTouchMove = (e: TouchEvent) => {
      if (e.touches.length === 2 && touchStartDistRef.current !== null) {
        e.preventDefault(); // 브라우저 자체 제스처 간섭 방지
        const currentDist = getDistance(e.touches[0], e.touches[1]);
        const scaleRatio = currentDist / touchStartDistRef.current;
        const newZoom = Math.min(4.0, Math.max(0.2, Math.round(touchStartZoomRef.current * scaleRatio * 100) / 100));
        setZoom(newZoom);
      }
    };

    const handleTouchEnd = (e: TouchEvent) => {
      if (e.touches.length < 2) {
        touchStartDistRef.current = null;
      }
    };

    container.addEventListener('touchstart', handleTouchStart, { passive: true });
    container.addEventListener('touchmove', handleTouchMove, { passive: false });
    container.addEventListener('touchend', handleTouchEnd, { passive: true });
    container.addEventListener('touchcancel', handleTouchEnd, { passive: true });

    return () => {
      container.removeEventListener('touchstart', handleTouchStart);
      container.removeEventListener('touchmove', handleTouchMove);
      container.removeEventListener('touchend', handleTouchEnd);
      container.removeEventListener('touchcancel', handleTouchEnd);
    };
  }, [zoom]);

  // PC: Ctrl + 마우스 휠 및 노트북 트랙패드 핀치 줌 지원 (20% ~ 400%)
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const handleWheel = (e: WheelEvent) => {
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        const factor = e.deltaY < 0 ? 1.15 : 0.85;
        setZoom((prev) => Math.min(4.0, Math.max(0.2, Math.round(prev * factor * 100) / 100)));
      }
    };

    container.addEventListener('wheel', handleWheel, { passive: false });
    return () => container.removeEventListener('wheel', handleWheel);
  }, []);

  // 더블 탭 시 100% <-> 200% 배율 토글 (모바일 편의 기능)
  const handleDoubleTap = useCallback(() => {
    const now = Date.now();
    if (now - lastTapRef.current < 300) {
      setZoom((prev) => (prev >= 1.5 ? 1.0 : 2.0));
    }
    lastTapRef.current = now;
  }, []);

  // Esc 키 닫기 및 키보드 좌/우 페이지 이동
  useEffect(() => {
    if (!isOpen) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose();
      } else if (e.key === 'ArrowLeft') {
        setCurrentPage((prev) => Math.max(1, prev - 1));
      } else if (e.key === 'ArrowRight') {
        setCurrentPage((prev) => Math.min(totalPages, prev + 1));
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose, totalPages]);

  if (!isOpen) return null;

  const downloadFilename = title.toLowerCase().endsWith('.pdf') ? title : `${title}.pdf`;

  const handleDownload = () => {
    const a = document.createElement('a');
    a.href = data;
    a.download = downloadFilename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const handleZoomOut = () => {
    setZoom((z) => {
      const step = z <= 0.6 ? 0.1 : 0.2;
      return Math.max(0.2, Math.round((z - step) * 10) / 10);
    });
  };

  const handleZoomIn = () => {
    setZoom((z) => {
      const step = z < 0.6 ? 0.1 : 0.2;
      return Math.min(4.0, Math.round((z + step) * 10) / 10);
    });
  };

  // 실시간 GPU CSS 스케일 배율 (디바운스 렌더링 중 즉각적인 반응성 제공)
  const visualScale = renderedZoom > 0 ? zoom / renderedZoom : 1.0;

  return (
    <div 
      className="fixed inset-0 z-50 flex flex-col bg-slate-950/95 backdrop-blur-md animate-in fade-in duration-200 select-none"
      role="dialog"
      aria-modal="true"
    >
      {/* 상단 툴바 (모바일/PC 모두 터치 친화적 컴팩트 배치) */}
      <div className="flex items-center justify-between gap-1 sm:gap-2 px-2 sm:px-4 py-2 bg-slate-900 border-b border-slate-800 text-slate-200 shrink-0 z-10 shadow-lg">
        {/* 제목 */}
        <div className="flex items-center gap-1.5 min-w-0 max-w-[110px] xs:max-w-[150px] sm:max-w-xs md:max-w-md">
          <span className="text-base sm:text-lg shrink-0">📄</span>
          <span className="font-bold text-xs sm:text-sm truncate text-slate-100" title={title}>
            {title}
          </span>
        </div>

        {/* 페이지 네비게이션 */}
        <div className="flex items-center gap-0.5 sm:gap-1 bg-slate-950/80 border border-slate-800 rounded-lg px-1.5 py-0.5 sm:px-2 sm:py-1 text-xs shrink-0">
          <button
            type="button"
            onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
            disabled={currentPage <= 1 || isLoading}
            className="p-1 hover:bg-slate-800 rounded text-slate-300 disabled:opacity-30 disabled:hover:bg-transparent"
            title="이전 페이지 (←)"
          >
            <ChevronLeft size={15} />
          </button>

          <span className="font-semibold text-slate-200 px-0.5 sm:px-1 text-[11px] sm:text-xs">
            {currentPage} <span className="text-slate-500 font-normal">/</span> {totalPages}
          </span>

          <button
            type="button"
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            disabled={currentPage >= totalPages || isLoading}
            className="p-1 hover:bg-slate-800 rounded text-slate-300 disabled:opacity-30 disabled:hover:bg-transparent"
            title="다음 페이지 (→)"
          >
            <ChevronRight size={15} />
          </button>
        </div>

        {/* 줌 배율 컨트롤 (모바일 & PC 공통 표시: 20% ~ 400% 자유 조절) */}
        <div className="flex items-center gap-0.5 sm:gap-1 bg-slate-950/80 border border-slate-800 rounded-lg px-1 py-0.5 shrink-0">
          <button
            type="button"
            onClick={handleZoomOut}
            className="p-1 sm:p-1.5 hover:bg-slate-800 text-slate-300 rounded transition-colors active:scale-95"
            title="축소 (최소 20%까지 아주 작게 보기)"
          >
            <ZoomOut size={14} className="sm:w-4 sm:h-4" />
          </button>

          <button
            type="button"
            onClick={() => setZoom(1.0)}
            className="px-1 sm:px-1.5 py-0.5 text-[10px] sm:text-xs font-bold text-indigo-300 hover:text-white rounded transition-colors"
            title="배율 초기화 (100% 클릭 시 기본 화면 맞춤)"
          >
            {Math.round(zoom * 100)}%
          </button>

          <button
            type="button"
            onClick={handleZoomIn}
            className="p-1 sm:p-1.5 hover:bg-slate-800 text-slate-300 rounded transition-colors active:scale-95"
            title="확대 (최대 400%까지 아주 크게 보기)"
          >
            <ZoomIn size={14} className="sm:w-4 sm:h-4" />
          </button>
        </div>

        {/* 액션 컨트롤 (다운로드 & 닫기) */}
        <div className="flex items-center gap-1 shrink-0">
          <button
            type="button"
            onClick={handleDownload}
            className="flex items-center gap-1 p-1.5 sm:px-2.5 sm:py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all active:scale-95"
            title="PDF 파일 다운로드"
          >
            <Download size={14} className="text-indigo-400" />
            <span className="hidden md:inline">다운로드</span>
          </button>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 rounded-lg transition-colors"
            title="닫기 (Esc)"
          >
            <X size={18} />
          </button>
        </div>
      </div>

      {/* 캔버스 뷰어 본체 (가로/세로 자유 스크롤 및 핀치 줌 지원) */}
      <div 
        ref={containerRef}
        onClick={handleDoubleTap}
        className="flex-1 overflow-auto p-2 sm:p-4 md:p-6 flex items-start justify-center bg-slate-950 cursor-grab active:cursor-grabbing touch-pan-x touch-pan-y"
      >
        {isLoading && (
          <div className="flex flex-col items-center justify-center h-64 gap-3 text-slate-400 text-sm m-auto">
            <div className="w-8 h-8 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin" />
            <span>PDF 문서를 렌더링하는 중입니다...</span>
          </div>
        )}

        {renderError && (
          <div className="p-6 bg-rose-950/40 border border-rose-800 rounded-2xl text-center text-rose-300 space-y-2 max-w-md m-auto">
            <p className="font-bold text-sm">⚠️ {renderError}</p>
            <button
              type="button"
              onClick={handleDownload}
              className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-semibold"
            >
              기기 다운로드로 열기
            </button>
          </div>
        )}

        <div 
          className={`m-auto transition-opacity duration-150 ${isLoading || renderError ? 'hidden' : 'block'}`}
          style={{
            transform: visualScale !== 1.0 ? `scale(${visualScale})` : undefined,
            transformOrigin: 'top center',
            transition: 'transform 75ms ease-out',
          }}
        >
          <canvas
            ref={canvasRef}
            style={{ maxWidth: 'none' }}
            className="rounded-xl shadow-2xl bg-white block m-auto"
          />
        </div>
      </div>
    </div>
  );
};
