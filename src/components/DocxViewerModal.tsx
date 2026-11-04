import React, { useState, useEffect, useRef } from 'react';
import {
  X,
  Download,
  Edit3,
  FileText,
  Copy,
  Check,
  ZoomIn,
  ZoomOut,
  RotateCcw,
  Loader2
} from 'lucide-react';
import { renderDocxToContainer, extractDocxTextAndHtml } from '../utils/docxViewer';

interface DocxViewerModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onOpenEditor?: () => void;
  onImportText?: (text: string) => void;
}

export const DocxViewerModal: React.FC<DocxViewerModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  onOpenEditor,
  onImportText,
}) => {
  const [zoom, setZoom] = useState<number>(1.0);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [copied, setCopied] = useState<boolean>(false);

  const containerRef = useRef<HTMLDivElement>(null);

  // 워드 문서 렌더링
  useEffect(() => {
    if (!isOpen || !data) return;

    let isMounted = true;
    setIsLoading(true);
    setErrorMessage(null);
    setZoom(1.0);

    const timer = setTimeout(() => {
      if (!containerRef.current) return;

      renderDocxToContainer(data, containerRef.current)
        .then(() => {
          if (isMounted) setIsLoading(false);
        })
        .catch((err) => {
          console.error('워드 문서 렌더링 실패:', err);
          if (isMounted) {
            setErrorMessage('워드(DOCX) 문서를 렌더링하는 중 오류가 발생했습니다.');
            setIsLoading(false);
          }
        });
    }, 50);

    return () => {
      isMounted = false;
      clearTimeout(timer);
    };
  }, [isOpen, data]);

  // 텍스트 복사 핸들러
  const handleCopyText = async () => {
    try {
      const { text } = await extractDocxTextAndHtml(data);
      if (!text.trim()) {
        alert('복사할 텍스트가 없습니다.');
        return;
      }
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error('텍스트 복사 실패:', err);
    }
  };

  // 메모 본문으로 가져오기 핸들러
  const handleImport = async () => {
    try {
      const { markdown, text } = await extractDocxTextAndHtml(data);
      const contentToImport = markdown || text;
      if (!contentToImport.trim()) {
        alert('가져올 수 있는 문서 내용이 없습니다.');
        return;
      }
      if (onImportText) {
        onImportText(contentToImport);
        onClose();
      }
    } catch (err) {
      console.error('메모 가져오기 실패:', err);
      alert('워드 문서 내용을 추출하지 못했습니다.');
    }
  };

  // 다운로드 핸들러
  const handleDownload = () => {
    const link = document.createElement('a');
    link.href = data;
    link.download = title.endsWith('.docx') || title.endsWith('.doc') ? title : `${title}.docx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-slate-950/90 backdrop-blur-md flex flex-col items-center justify-center p-2 md:p-4 select-none animate-fadeIn">
      {/* 팝업 컨테이너 */}
      <div className="w-full max-w-6xl h-[95vh] bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
        {/* 상단 툴바 헤더 */}
        <div className="px-4 py-3 bg-slate-900/95 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-2.5 min-w-0">
            <span className="p-1.5 rounded-lg bg-blue-500/10 text-blue-400 border border-blue-500/20">
              <FileText size={18} />
            </span>
            <div className="truncate">
              <div className="flex items-center gap-2">
                <h2 className="text-sm md:text-base font-bold text-slate-100 truncate" title={title}>
                  {title}
                </h2>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-blue-500/10 text-blue-300 border border-blue-500/20 font-bold uppercase shrink-0">
                  Word DOCX
                </span>
              </div>
            </div>
          </div>

          {/* 중앙 확대/축소 컨트롤 */}
          <div className="flex items-center gap-1 bg-slate-950 border border-slate-800 rounded-lg p-0.5 text-xs text-slate-300">
            <button
              type="button"
              onClick={() => setZoom((prev) => Math.max(0.6, prev - 0.1))}
              className="p-1 hover:text-white hover:bg-slate-800 rounded"
              title="축소"
            >
              <ZoomOut size={13} />
            </button>
            <span className="px-1 text-[11px] font-mono min-w-[42px] text-center">
              {Math.round(zoom * 100)}%
            </span>
            <button
              type="button"
              onClick={() => setZoom((prev) => Math.min(1.8, prev + 0.1))}
              className="p-1 hover:text-white hover:bg-slate-800 rounded"
              title="확대"
            >
              <ZoomIn size={13} />
            </button>
            <button
              type="button"
              onClick={() => setZoom(1.0)}
              className="p-1 hover:text-white hover:bg-slate-800 rounded"
              title="100% 기본 크기"
            >
              <RotateCcw size={12} />
            </button>
          </div>

          {/* 우측 액션 버튼들 */}
          <div className="flex items-center gap-1.5 shrink-0">
            {onOpenEditor && (
              <button
                type="button"
                onClick={() => {
                  onClose();
                  onOpenEditor();
                }}
                className="px-2.5 py-1.5 bg-blue-600/20 hover:bg-blue-600/30 text-blue-300 border border-blue-500/30 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors active:scale-95"
                title="웹에서 워드 문서 직접 편집"
              >
                <Edit3 size={13} />
                <span>문서 직접 편집</span>
              </button>
            )}

            <button
              type="button"
              onClick={handleCopyText}
              className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-medium border border-slate-700 flex items-center gap-1.5 transition-colors active:scale-95"
              title="문서 본문 텍스트 복사"
            >
              {copied ? <Check size={13} className="text-emerald-400" /> : <Copy size={13} />}
              <span>{copied ? '복사됨!' : '텍스트 복사'}</span>
            </button>

            {onImportText && (
              <button
                type="button"
                onClick={handleImport}
                className="px-2.5 py-1.5 bg-indigo-600/30 hover:bg-indigo-600 text-indigo-200 hover:text-white rounded-lg text-xs font-semibold border border-indigo-500/40 flex items-center gap-1.5 transition-colors active:scale-95"
                title="워드 문서 내용을 마크다운으로 메모 본문에 가져오기"
              >
                <span>📝 메모로 가져오기</span>
              </button>
            )}

            <button
              type="button"
              onClick={handleDownload}
              className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors"
              title="파일 다운로드"
            >
              <Download size={15} />
            </button>

            <button
              type="button"
              onClick={onClose}
              className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700 transition-colors"
              title="닫기 (Esc)"
            >
              <X size={16} />
            </button>
          </div>
        </div>

        {/* 메인 뷰어 영역 (워드 용지 스타일) */}
        <div className="flex-1 overflow-auto bg-slate-950 p-4 flex flex-col items-center">
          {isLoading && (
            <div className="flex items-center justify-center gap-2 py-20 text-slate-400 text-xs">
              <Loader2 size={16} className="animate-spin text-blue-500" />
              <span>워드(DOCX) 문서를 렌더링하는 중입니다...</span>
            </div>
          )}

          {errorMessage && (
            <div className="p-4 bg-rose-950/40 border border-rose-500/30 rounded-xl text-center text-xs text-rose-300 max-w-md my-10">
              {errorMessage}
            </div>
          )}

          {/* docx-preview가 렌더링할 컨테이너 */}
          <div
            style={{
              transform: `scale(${zoom})`,
              transformOrigin: 'top center',
              transition: 'transform 0.1s ease-out',
            }}
            className="w-full flex justify-center py-2"
          >
            <div
              ref={containerRef}
              className="docx-render-wrapper bg-white text-slate-900 shadow-2xl rounded-sm min-h-[600px] w-full max-w-[850px] p-6 md:p-12 overflow-x-auto select-text font-serif leading-relaxed"
            ></div>
          </div>
        </div>

        {/* 하단 안내 바 */}
        <div className="px-4 py-2 bg-slate-900 border-t border-slate-800 flex items-center justify-between text-[11px] text-slate-400">
          <span>워드 문서 레이아웃, 표, 이미지 및 서식이 원본 규격으로 렌더링됩니다.</span>
          <span className="hidden sm:inline">문서 내용 전체를 메모에 넣으려면 [📝 메모로 가져오기]를 누르세요.</span>
        </div>
      </div>
    </div>
  );
};
