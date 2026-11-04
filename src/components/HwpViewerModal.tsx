import React, { useState, useEffect, useRef } from 'react';
import { X, ChevronLeft, ChevronRight, ZoomIn, ZoomOut, Download, FileText, Check } from 'lucide-react';
import { loadHwpDocument, renderHwpPageToSvg, extractHwpText } from '../utils/hwpViewer';
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
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [renderError, setRenderError] = useState<string | null>(null);
  const [svgHtml, setSvgHtml] = useState<string>('');
  const [isCopied, setIsCopied] = useState<boolean>(false);

  const containerRef = useRef<HTMLDivElement>(null);

  // HWP 문서 로드
  useEffect(() => {
    if (!isOpen || !data) return;

    let isMounted = true;
    setIsLoading(true);
    setRenderError(null);
    setCurrentPage(initialPage);
    setZoom(1.0);

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

  // 페이지 및 줌 배율 변경 시 SVG 렌더링
  useEffect(() => {
    if (!isOpen || !hwpDoc) return;

    try {
      const pageIdx = Math.max(0, Math.min(totalPages - 1, currentPage - 1));
      const rawSvg = renderHwpPageToSvg(hwpDoc, pageIdx);
      setSvgHtml(rawSvg);
    } catch (err) {
      console.error('HWP 페이지 SVG 렌더링 오류:', err);
    }
  }, [isOpen, hwpDoc, currentPage, totalPages]);

  // 키보드 단축키 (Esc: 닫기, 좌/우: 페이지 이동)
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

  return (
    <div
      className="fixed inset-0 z-50 flex flex-col bg-slate-950/95 backdrop-blur-md animate-in fade-in duration-200 select-none"
      role="dialog"
      aria-modal="true"
    >
      {/* 상단 툴바 */}
      <div className="flex flex-wrap items-center justify-between gap-2 px-3 md:px-5 py-2.5 bg-slate-900 border-b border-slate-800 text-slate-200 shrink-0 z-10 shadow-lg">
        {/* 제목 */}
        <div className="flex items-center gap-2 min-w-0 max-w-[200px] md:max-w-md">
          <span className="text-lg">📝</span>
          <span className="font-bold text-xs md:text-sm truncate text-slate-100" title={title}>
            {title}
          </span>
          <span className="text-[10px] px-1.5 py-0.5 rounded bg-blue-500/10 text-blue-300 border border-blue-500/20 font-bold shrink-0">
            HWP 뷰어
          </span>
        </div>

        {/* 페이지 네비게이션 */}
        <div className="flex items-center gap-1.5 bg-slate-950/70 border border-slate-800 rounded-lg px-2 py-1 text-xs">
          <button
            type="button"
            onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
            disabled={currentPage <= 1 || isLoading}
            className="p-1 hover:bg-slate-800 rounded text-slate-300 disabled:opacity-30 disabled:hover:bg-transparent"
            title="이전 페이지 (←)"
          >
            <ChevronLeft size={16} />
          </button>

          <span className="font-semibold text-slate-200 px-1 text-xs">
            {currentPage} <span className="text-slate-500 font-normal">/</span> {totalPages}
          </span>

          <button
            type="button"
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            disabled={currentPage >= totalPages || isLoading}
            className="p-1 hover:bg-slate-800 rounded text-slate-300 disabled:opacity-30 disabled:hover:bg-transparent"
            title="다음 페이지 (→)"
          >
            <ChevronRight size={16} />
          </button>
        </div>

        {/* 줌 배율 & 액션 컨트롤 */}
        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={() => setZoom((z) => Math.max(0.6, Math.round((z - 0.2) * 10) / 10))}
            className="p-1.5 hover:bg-slate-800 text-slate-300 rounded-lg transition-colors hidden sm:block"
            title="축소"
          >
            <ZoomOut size={16} />
          </button>

          <button
            type="button"
            onClick={() => setZoom(1.0)}
            className="px-2 py-1 text-[11px] font-semibold text-slate-400 hover:text-slate-200 hover:bg-slate-800 rounded hidden sm:block"
            title="배율 초기화 (100%)"
          >
            {Math.round(zoom * 100)}%
          </button>

          <button
            type="button"
            onClick={() => setZoom((z) => Math.min(2.5, Math.round((z + 0.2) * 10) / 10))}
            className="p-1.5 hover:bg-slate-800 text-slate-300 rounded-lg transition-colors hidden sm:block"
            title="확대"
          >
            <ZoomIn size={16} />
          </button>

          <div className="h-4 w-[1px] bg-slate-800 mx-1 hidden sm:block" />

          {/* 본문 메모로 가져오기 버튼 */}
          <button
            type="button"
            onClick={handleImportToMemo}
            disabled={isLoading}
            className="flex items-center gap-1 px-2.5 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95 disabled:opacity-50"
            title="한글 문서 내용을 마크다운 메모 본문으로 바로 가져옵니다."
          >
            {isCopied ? <Check size={14} className="text-emerald-300" /> : <FileText size={14} />}
            <span>{isCopied ? '메모에 반영됨!' : '메모로 가져오기'}</span>
          </button>

          {/* 다운로드 버튼 */}
          <button
            type="button"
            onClick={handleDownload}
            className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all active:scale-95"
            title="원본 파일 다운로드"
          >
            <Download size={14} className="text-indigo-400" />
            <span className="hidden md:inline">다운로드</span>
          </button>

          {/* 닫기 버튼 */}
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 rounded-lg transition-colors ml-1"
            title="닫기 (Esc)"
          >
            <X size={18} />
          </button>
        </div>
      </div>

      {/* 본문 뷰어 영역 (SVG 렌더링) */}
      <div
        ref={containerRef}
        className="flex-1 overflow-auto p-3 md:p-6 flex items-start justify-center bg-slate-950"
      >
        {isLoading && (
          <div className="flex flex-col items-center justify-center h-64 gap-3 text-slate-400 text-sm">
            <div className="w-8 h-8 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin" />
            <span>한글 문서를 렌더링하는 중입니다...</span>
          </div>
        )}

        {renderError && (
          <div className="p-6 bg-rose-950/40 border border-rose-800 rounded-2xl text-center text-rose-300 space-y-2 max-w-md my-auto">
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
            style={{ transform: `scale(${zoom})`, transformOrigin: 'top center' }}
            className="transition-transform duration-100 shadow-2xl rounded-xl bg-white p-2 md:p-4 my-2 text-slate-900 max-w-full"
            dangerouslySetInnerHTML={{ __html: svgHtml }}
          />
        )}
      </div>
    </div>
  );
};
