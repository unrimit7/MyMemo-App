import React, { useState, useEffect, useRef, useCallback } from 'react';
import { X, ChevronLeft, ChevronRight, ZoomIn, ZoomOut, Download, RotateCw, Edit3 } from 'lucide-react';
import { loadPdfDocument, renderPdfPageToCanvas } from '../utils/pdfViewer';
import type { PDFDocumentProxy } from 'pdfjs-dist';
import { useFloatingWindow } from './useFloatingWindow';
import { useWheelPageNavigation } from '../hooks/useWheelPageNavigation';

interface PdfViewerModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  initialPage?: number;
  onOpenEditor?: () => void;
}

export const PdfViewerModal: React.FC<PdfViewerModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  initialPage = 1,
  onOpenEditor,
}) => {
  const [pdfDoc, setPdfDoc] = useState<PDFDocumentProxy | null>(null);
  const [currentPage, setCurrentPage] = useState<number>(initialPage);
  const [totalPages, setTotalPages] = useState<number>(1);
  const [zoom, setZoom] = useState<number>(1.0);
  const [renderedZoom, setRenderedZoom] = useState<number>(1.0);
  const [rotation, setRotation] = useState<number>(0);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [renderError, setRenderError] = useState<string | null>(null);

  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const debounceTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const lastTapRef = useRef<number>(0);
  const floating = useFloatingWindow({ id: 'pdf-viewer', defaultWidth: 880, defaultHeight: 800 });

  // 상/하단 마우스 휠 페이지 이동 (상단에서 휠 올리면 이전 페이지, 하단에서 휠 내리면 다음 페이지)
  const { pendingScrollRef, pageToast } = useWheelPageNavigation({
    containerRef,
    currentPage,
    totalPages,
    onPageChange: setCurrentPage,
    enabled: isOpen && !isLoading,
  });

  // PDF 문서 로딩
  useEffect(() => {
    if (!isOpen || !data) return;

    let isMounted = true;
    setIsLoading(true);
    setRenderError(null);
    setCurrentPage(initialPage);
    setZoom(1.0);
    setRenderedZoom(1.0);
    setRotation(0);

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

  // 페이지, 줌 배율 및 회전각 변경 시 캔버스 렌더링
  useEffect(() => {
    if (!isOpen || !pdfDoc || !canvasRef.current) return;

    let isMounted = true;
    const canvas = canvasRef.current;
    const containerWidth = containerRef.current?.clientWidth || window.innerWidth;
    const baseWidth = Math.min(Math.max(260, containerWidth - 24), 850);

    renderPdfPageToCanvas(pdfDoc, currentPage, canvas, baseWidth, renderedZoom, rotation)
      .then(() => {
        if (!isMounted) return;
        const container = containerRef.current;
        if (container && pendingScrollRef.current) {
          if (pendingScrollRef.current === 'bottom') {
            container.scrollTop = container.scrollHeight - container.clientHeight;
          } else if (pendingScrollRef.current === 'top') {
            container.scrollTop = 0;
          }
          pendingScrollRef.current = null;
        }
      })
      .catch((err) => {
        if (isMounted) {
          console.error('PDF 페이지 렌더링 실패:', err);
        }
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, pdfDoc, currentPage, renderedZoom, rotation]);

  // 터치 제스처 (스마트폰 두 손가락 핀치 줌: 20% 아주 작게 ~ 500% 아주 크게)
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
        const newZoom = Math.min(5.0, Math.max(0.2, Math.round(touchStartZoomRef.current * scaleRatio * 100) / 100));
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

  // PC: Ctrl + 마우스 휠 및 노트북 트랙패드 핀치 줌 지원 (20% ~ 500%)
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const handleWheel = (e: WheelEvent) => {
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        const factor = e.deltaY < 0 ? 1.15 : 0.85;
        setZoom((prev) => Math.min(5.0, Math.max(0.2, Math.round(prev * factor * 100) / 100)));
      }
    };

    container.addEventListener('wheel', handleWheel, { passive: false });
    return () => container.removeEventListener('wheel', handleWheel);
  }, []);

  // PC: 마우스 클릭 드래그로 화면 이동 (Pan Hand Tool)
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const dragStartRef = useRef<{ x: number; y: number; scrollLeft: number; scrollTop: number } | null>(null);
  const hasMovedRef = useRef<boolean>(false);

  const handleMouseDown = (e: React.MouseEvent<HTMLDivElement>) => {
    if (e.button !== 0) return; // 좌클릭만 처리
    const container = containerRef.current;
    if (!container) return;

    // 네이티브 드래그 방지
    e.preventDefault();

    dragStartRef.current = {
      x: e.clientX,
      y: e.clientY,
      scrollLeft: container.scrollLeft,
      scrollTop: container.scrollTop,
    };
    hasMovedRef.current = false;
    setIsDragging(true);
  };

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!dragStartRef.current || !containerRef.current) return;
      const dx = e.clientX - dragStartRef.current.x;
      const dy = e.clientY - dragStartRef.current.y;

      if (Math.abs(dx) > 3 || Math.abs(dy) > 3) {
        hasMovedRef.current = true;
      }

      containerRef.current.scrollLeft = dragStartRef.current.scrollLeft - dx;
      containerRef.current.scrollTop = dragStartRef.current.scrollTop - dy;
    };

    const handleMouseUp = () => {
      if (dragStartRef.current) {
        dragStartRef.current = null;
        setIsDragging(false);
      }
    };

    if (isDragging) {
      window.addEventListener('mousemove', handleMouseMove, { passive: false });
      window.addEventListener('mouseup', handleMouseUp);
    }

    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isDragging]);

  // 더블 탭 시 100% <-> 200% 배율 토글 (드래그 이동 중에는 무시)
  const handleDoubleTap = useCallback(() => {
    if (hasMovedRef.current) return;
    const now = Date.now();
    if (now - lastTapRef.current < 300) {
      setZoom((prev) => (prev >= 1.5 ? 1.0 : 2.0));
    }
    lastTapRef.current = now;
  }, []);

  // Esc 키 닫기 및 키보드 좌/우 페이지 이동, R 회전
  useEffect(() => {
    if (!isOpen) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose();
      } else if (e.key === 'ArrowLeft') {
        setCurrentPage((prev) => Math.max(1, prev - 1));
      } else if (e.key === 'ArrowRight') {
        setCurrentPage((prev) => Math.min(totalPages, prev + 1));
      } else if (e.key.toLowerCase() === 'r') {
        setRotation((r) => (r + 90) % 360);
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
      return Math.min(5.0, Math.round((z + step) * 10) / 10);
    });
  };

  // 실시간 GPU CSS 스케일 배율 (디바운스 렌더링 중 즉각적인 반응성 제공)
  const visualScale = renderedZoom > 0 ? zoom / renderedZoom : 1.0;

  return (
    <div 
      className={floating.modalWrapperClass}
      role="dialog"
      aria-modal="true"
    >
      <div {...floating.windowContainerProps}>
        {/* 상단 툴바 (모바일/PC 모두 터치 친화적 컴팩트 배치) */}
        <div 
          {...floating.headerProps}
          className="flex items-center justify-between gap-1 sm:gap-2 px-2 sm:px-4 py-2 bg-slate-900 border-b border-slate-800 text-slate-200 shrink-0 z-10 shadow-lg"
        >
        {/* 제목 */}
        <div className="flex items-center gap-1.5 min-w-0 max-w-[100px] xs:max-w-[140px] sm:max-w-xs md:max-w-md">
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

        {/* 줌 배율 컨트롤 (20% 아주 작게 ~ 500% 아주 크게) & 회전 버튼 */}
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
            title="확대 (최대 500%까지 자유롭게 확대하여 글씨 읽기)"
          >
            <ZoomIn size={14} className="sm:w-4 sm:h-4" />
          </button>

          {/* 90도 회전 버튼 */}
          <button
            type="button"
            onClick={() => setRotation((r) => (r + 90) % 360)}
            className="p-1 sm:p-1.5 hover:bg-slate-800 text-amber-300 hover:text-amber-200 rounded transition-colors active:scale-95 flex items-center gap-0.5 border-l border-slate-800 pl-1.5 ml-0.5"
            title="90도 회전 (가로 문서를 세로로 눕혀서 보기)"
          >
            <RotateCw size={13} className="sm:w-3.5 sm:h-3.5" />
            <span className="text-[10px] font-bold">{rotation ? `${rotation}°` : '회전'}</span>
          </button>
        </div>

        {/* 액션 컨트롤 (편집, 다운로드 & 닫기) */}
        <div className="flex items-center gap-1 shrink-0">
          {onOpenEditor && (
            <button
              type="button"
              onClick={onOpenEditor}
              className="flex items-center gap-1 p-1.5 sm:px-2.5 sm:py-1.5 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 rounded-lg text-xs font-semibold border border-emerald-500/30 transition-all active:scale-95"
              title="PDF 문서 위에 텍스트 입력, 형광펜, 펜 드로잉, 서명/도장 편집"
            >
              <Edit3 size={14} />
              <span className="hidden sm:inline">주석/서명</span>
            </button>
          )}

          <button
            type="button"
            onClick={handleDownload}
            className="flex items-center gap-1 p-1.5 sm:px-2.5 sm:py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all active:scale-95"
            title="PDF 파일 다운로드"
          >
            <Download size={14} className="text-indigo-400" />
            <span className="hidden md:inline">다운로드</span>
          </button>

          {floating.toggleFloatingButton}

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

      {/* 캔버스 뷰어 본체 (화면 너비에 갇히지 않고 줌 배율대로 자유 확대 & 마우스 드래그 및 터치 스크롤 지원) */}
      <div 
        ref={containerRef}
        onMouseDown={handleMouseDown}
        onClick={handleDoubleTap}
        className={`flex-1 overflow-auto bg-slate-950 p-2 sm:p-4 touch-pan-x touch-pan-y relative ${
          isDragging ? 'cursor-grabbing select-none' : 'cursor-grab'
        }`}
      >
        <div className="w-max min-w-full min-h-full flex flex-col items-center justify-start py-4 sm:py-6 px-2">
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
            className={`shrink-0 m-auto transition-opacity duration-150 ${isLoading || renderError ? 'hidden' : 'block'}`}
            style={{
              transform: visualScale !== 1.0 ? `scale(${visualScale})` : undefined,
              transformOrigin: 'center center',
              transition: 'transform 75ms ease-out',
            }}
          >
            <canvas
              ref={canvasRef}
              style={{ maxWidth: 'none', display: 'block' }}
              className="rounded-xl shadow-2xl bg-white"
            />
          </div>
        </div>

        {/* 마우스 휠 페이지 이동 안내 토스트 */}
        {pageToast && (
          <div className="absolute bottom-6 left-1/2 -translate-x-1/2 z-40 bg-slate-900/95 text-indigo-300 border border-indigo-500/40 px-3.5 py-1.5 rounded-full text-xs font-mono font-semibold shadow-2xl backdrop-blur-sm pointer-events-none transition-all animate-in fade-in duration-150">
            {pageToast}
          </div>
        )}
      </div>

      {floating.resizeHandle}
    </div>
  </div>
);
};
