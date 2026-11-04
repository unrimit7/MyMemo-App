import React, { useState, useEffect, useRef, useCallback } from 'react';
import { X, ChevronLeft, ChevronRight, ZoomIn, ZoomOut, Download, FileText, Check, RotateCw } from 'lucide-react';
import { loadHwpDocument, renderHwpPageToSvg, extractHwpText, rotateSvg } from '../utils/hwpViewer';
import type { HwpDocument } from '@rhwp/core';

interface HwpViewerModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  initialPage?: number;
  onImportText?: (text: string) => void;
}

export const HwpViewerModal: React.FC<HwpViewerModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  initialPage = 1,
  onImportText,
}) => {
  const [hwpDoc, setHwpDoc] = useState<HwpDocument | null>(null);
  const [currentPage, setCurrentPage] = useState<number>(initialPage);
  const [totalPages, setTotalPages] = useState<number>(1);
  const [zoom, setZoom] = useState<number>(1.0);
  const [rotation, setRotation] = useState<number>(0);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [renderError, setRenderError] = useState<string | null>(null);
  const [svgHtml, setSvgHtml] = useState<string>('');
  const [isCopied, setIsCopied] = useState<boolean>(false);

  const containerRef = useRef<HTMLDivElement>(null);
  const lastTapRef = useRef<number>(0);

  // HWP 문서 로드
  useEffect(() => {
    if (!isOpen || !data) return;

    let isMounted = true;
    setIsLoading(true);
    setRenderError(null);
    setCurrentPage(initialPage);
    setZoom(1.0);
    setRotation(0);

    loadHwpDocument(data)
      .then((doc) => {
        if (!isMounted) return;
        setHwpDoc(doc);
        const count = doc.pageCount();
        setTotalPages(Math.max(1, count));
        setIsLoading(false);
      })
      .catch((err) => {
        if (!isMounted) return;
        console.error('HWP 로드 실패:', err);
        setRenderError('한글(HWP/HWPX) 문서를 불러오는 데 실패했습니다.');
        setIsLoading(false);
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, data, initialPage]);

  // 페이지 및 회전각 변경 시 SVG 렌더링
  useEffect(() => {
    if (!isOpen || !hwpDoc) return;

    try {
      const pageIdx = Math.max(0, Math.min(totalPages - 1, currentPage - 1));
      const rawSvg = renderHwpPageToSvg(hwpDoc, pageIdx);
      const transformedSvg = rotation !== 0 ? rotateSvg(rawSvg, rotation) : rawSvg;
      setSvgHtml(transformedSvg);
    } catch (err) {
      console.error('HWP 페이지 SVG 렌더링 오류:', err);
    }
  }, [isOpen, hwpDoc, currentPage, totalPages, rotation]);

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

  // 더블 탭 시 100% <-> 200% 배율 토글 (모바일 편의 기능)
  const handleDoubleTap = useCallback(() => {
    const now = Date.now();
    if (now - lastTapRef.current < 300) {
      setZoom((prev) => (prev >= 1.5 ? 1.0 : 2.0));
    }
    lastTapRef.current = now;
  }, []);

  // 키보드 단축키 (Esc: 닫기, 좌/우: 페이지 이동, R: 회전)
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

  const downloadFilename = title.toLowerCase().endsWith('.hwp') || title.toLowerCase().endsWith('.hwpx')
    ? title
    : `${title}.hwp`;

  const handleDownload = () => {
    const a = document.createElement('a');
    a.href = data;
    a.download = downloadFilename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // 한글 문서 텍스트를 메모 본문으로 가져오기
  const handleImportToMemo = () => {
    if (!hwpDoc) return;
    const text = extractHwpText(hwpDoc);
    if (!text.trim()) {
      alert('추출할 수 있는 본문 텍스트가 없습니다.');
      return;
    }

    if (onImportText) {
      onImportText(text);
      setIsCopied(true);
      setTimeout(() => setIsCopied(false), 2500);
    } else {
      navigator.clipboard.writeText(text);
      setIsCopied(true);
      setTimeout(() => setIsCopied(false), 2500);
    }
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

  // 화면 폭에 비례한 기준 너비 및 줌 배율 적용 너비 (화면 폭 제한 없이 줌만큼 실제로 무한 확장)
  const containerWidth = containerRef.current?.clientWidth || window.innerWidth;
  const baseWidth = Math.min(Math.max(260, containerWidth - 32), 850);
  const targetWidth = Math.round(baseWidth * zoom);

  return (
    <div
      className="fixed inset-0 z-50 flex flex-col bg-slate-950/95 backdrop-blur-md animate-in fade-in duration-200 select-none"
      role="dialog"
      aria-modal="true"
    >
      {/* 상단 툴바 (모바일/PC 컴팩트 구성) */}
      <div className="flex items-center justify-between gap-1 sm:gap-2 px-2 sm:px-4 py-2 bg-slate-900 border-b border-slate-800 text-slate-200 shrink-0 z-10 shadow-lg">
        {/* 제목 */}
        <div className="flex items-center gap-1.5 min-w-0 max-w-[90px] xs:max-w-[130px] sm:max-w-xs md:max-w-md">
          <span className="text-base sm:text-lg shrink-0">📝</span>
          <span className="font-bold text-xs sm:text-sm truncate text-slate-100" title={title}>
            {title}
          </span>
          <span className="text-[9px] sm:text-[10px] px-1 py-0.5 rounded bg-blue-500/10 text-blue-300 border border-blue-500/20 font-bold shrink-0 hidden xs:inline-block">
            HWP
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

        {/* 액션 컨트롤 (본문 복사, 다운로드 & 닫기) */}
        <div className="flex items-center gap-1 shrink-0">
          <button
            type="button"
            onClick={handleImportToMemo}
            className="flex items-center gap-1 p-1.5 sm:px-2 sm:py-1.5 bg-blue-600/20 hover:bg-blue-600/30 text-blue-300 rounded-lg text-xs font-semibold border border-blue-500/30 transition-all active:scale-95"
            title="HWP 본문 텍스트 추출"
          >
            {isCopied ? <Check size={14} className="text-emerald-400" /> : <FileText size={14} className="text-blue-400" />}
            <span className="hidden md:inline">{isCopied ? '복사 완료!' : '본문 추출'}</span>
          </button>

          <button
            type="button"
            onClick={handleDownload}
            className="flex items-center gap-1 p-1.5 sm:px-2.5 sm:py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all active:scale-95"
            title="원본 파일 다운로드"
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

      {/* 본문 뷰어 영역 (화면 너비에 갇히지 않고 확대 시 실제로 커지며 가로/세로 자유 스크롤) */}
      <div
        ref={containerRef}
        onClick={handleDoubleTap}
        className="flex-1 overflow-auto bg-slate-950 p-2 sm:p-4 touch-pan-x touch-pan-y relative"
      >
        <div className="min-w-full min-h-full flex items-center justify-center p-1 sm:p-3">
          {isLoading && (
            <div className="flex flex-col items-center justify-center h-64 gap-3 text-slate-400 text-sm m-auto">
              <div className="w-8 h-8 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin" />
              <span>한글 문서를 렌더링하는 중입니다...</span>
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

          {svgHtml && !isLoading && !renderError && (
            <div
              style={{
                width: `${targetWidth}px`,
                minWidth: `${targetWidth}px`,
                maxWidth: 'none',
              }}
              className="shrink-0 shadow-2xl rounded-xl bg-white p-2 sm:p-4 my-auto text-slate-900 transition-all duration-75 [&>svg]:w-full [&>svg]:h-auto [&>svg]:block [&>svg]:max-w-none"
              dangerouslySetInnerHTML={{ __html: svgHtml }}
            />
          )}
        </div>
      </div>
    </div>
  );
};
