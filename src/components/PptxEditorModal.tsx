import React, { useState, useEffect } from 'react';
import {
  X,
  Save,
  Download,
  Presentation,
  CheckCircle,
  AlertCircle,
  Loader2,
  Plus,
  Trash2,
  ArrowUp,
  ArrowDown,
  Eye,
  Type
} from 'lucide-react';
import {
  parsePptxPresentation,
  updatePptxSlideTexts,
  type PptxPresentation,
} from '../utils/pptxViewer';

interface PptxEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
}

interface EditableSlide {
  slideNumber: number;
  title: string;
  bodyText: string;
  images: { id: string; dataUrl: string }[];
}

export const PptxEditorModal: React.FC<PptxEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  onSave,
}) => {
  const [slides, setSlides] = useState<EditableSlide[]>([]);
  const [activeSlideIndex, setActiveSlideIndex] = useState<number>(0);
  const [loading, setLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [hasChanges, setHasChanges] = useState<boolean>(false);

  const [saveToast, setSaveToast] = useState<string | null>(null);
  const [errorToast, setErrorToast] = useState<string | null>(null);
  const [viewMode, setViewMode] = useState<'split' | 'editOnly' | 'previewOnly'>('split');

  // 원본 PPTX 파싱 및 편집용 슬라이드 상태 초기화
  useEffect(() => {
    if (!isOpen || !data) return;

    let isMounted = true;
    setLoading(true);
    setErrorToast(null);

    parsePptxPresentation(data)
      .then((pres: PptxPresentation) => {
        if (!isMounted) return;

        const editableList: EditableSlide[] = pres.slides.map((s) => {
          const bodyLines = s.paragraphs
            .filter((p) => !p.isTitle || p.text !== s.title)
            .map((p) => {
              const indent = '  '.repeat(Math.max(0, p.level));
              return `${indent}• ${p.text}`;
            });

          return {
            slideNumber: s.slideNumber,
            title: s.title,
            bodyText: bodyLines.join('\n'),
            images: s.images.map((img) => ({ id: img.id, dataUrl: img.dataUrl })),
          };
        });

        setSlides(editableList);
        setActiveSlideIndex(0);
        setHasChanges(false);
        setLoading(false);
      })
      .catch((err) => {
        if (!isMounted) return;
        console.error('PPTX 파싱 실패:', err);
        setErrorToast('파워포인트 문서를 편집 모드로 불러오지 못했습니다.');
        setLoading(false);
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, data]);

  // 활성 슬라이드 제목 수정
  const handleUpdateTitle = (newTitle: string) => {
    setSlides((prev) => {
      const copy = [...prev];
      copy[activeSlideIndex] = {
        ...copy[activeSlideIndex],
        title: newTitle,
      };
      return copy;
    });
    setHasChanges(true);
  };

  // 활성 슬라이드 본문 수정
  const handleUpdateBody = (newBody: string) => {
    setSlides((prev) => {
      const copy = [...prev];
      copy[activeSlideIndex] = {
        ...copy[activeSlideIndex],
        bodyText: newBody,
      };
      return copy;
    });
    setHasChanges(true);
  };

  // 새 슬라이드 추가
  const handleAddSlide = () => {
    const nextNum = slides.length + 1;
    const newSlide: EditableSlide = {
      slideNumber: nextNum,
      title: `새 슬라이드 ${nextNum}`,
      bodyText: '• 새로운 내용을 입력하세요\n• 부가 설명을 입력하세요',
      images: [],
    };
    setSlides((prev) => [...prev, newSlide]);
    setActiveSlideIndex(slides.length);
    setHasChanges(true);
  };

  // 슬라이드 삭제
  const handleDeleteSlide = (index: number) => {
    if (slides.length <= 1) {
      alert('최소 1개의 슬라이드가 유지되어야 합니다.');
      return;
    }
    if (!confirm(`'슬라이드 ${slides[index].slideNumber}'을(를) 삭제하시겠습니까?`)) {
      return;
    }
    setSlides((prev) => {
      const filtered = prev.filter((_, i) => i !== index);
      return filtered.map((s, idx) => ({ ...s, slideNumber: idx + 1 }));
    });
    setActiveSlideIndex((prev) => Math.max(0, Math.min(prev, slides.length - 2)));
    setHasChanges(true);
  };

  // 슬라이드 위로 이동
  const handleMoveUp = (index: number) => {
    if (index <= 0) return;
    setSlides((prev) => {
      const copy = [...prev];
      const temp = copy[index - 1];
      copy[index - 1] = copy[index];
      copy[index] = temp;
      return copy.map((s, idx) => ({ ...s, slideNumber: idx + 1 }));
    });
    setActiveSlideIndex(index - 1);
    setHasChanges(true);
  };

  // 슬라이드 아래로 이동
  const handleMoveDown = (index: number) => {
    if (index >= slides.length - 1) return;
    setSlides((prev) => {
      const copy = [...prev];
      const temp = copy[index + 1];
      copy[index + 1] = copy[index];
      copy[index] = temp;
      return copy.map((s, idx) => ({ ...s, slideNumber: idx + 1 }));
    });
    setActiveSlideIndex(index + 1);
    setHasChanges(true);
  };

  // 저장 및 메모 반영
  const handleSave = async () => {
    setIsSaving(true);
    setSaveToast(null);
    setErrorToast(null);

    try {
      const updatedDataUrl = await updatePptxSlideTexts(
        data,
        slides.map((s) => ({
          slideNumber: s.slideNumber,
          title: s.title,
          bodyText: s.bodyText,
        }))
      );

      if (onSave) {
        await onSave(updatedDataUrl);
      }

      setHasChanges(false);
      setSaveToast('파워포인트 내용이 성공적으로 저장 및 메모에 반영되었습니다!');
      setTimeout(() => setSaveToast(null), 3000);
    } catch (err) {
      console.error('PPTX 저장 실패:', err);
      setErrorToast('파워포인트 파일을 저장하는 중 오류가 발생했습니다.');
    } finally {
      setIsSaving(false);
    }
  };

  // 다운로드
  const handleDownload = async () => {
    try {
      const updatedDataUrl = await updatePptxSlideTexts(
        data,
        slides.map((s) => ({
          slideNumber: s.slideNumber,
          title: s.title,
          bodyText: s.bodyText,
        }))
      );

      const a = document.createElement('a');
      a.href = updatedDataUrl;
      a.download = title.endsWith('.pptx') ? title : `${title}.pptx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch (err) {
      console.error('다운로드 실패:', err);
      alert('파일을 다운로드하지 못했습니다.');
    }
  };

  // 닫기
  const handleClose = () => {
    if (hasChanges) {
      if (!confirm('저장되지 않은 파워포인트 수정 사항이 있습니다. 닫으시겠습니까?')) {
        return;
      }
    }
    onClose();
  };

  if (!isOpen) return null;

  const currentSlide: EditableSlide | undefined = slides[activeSlideIndex];

  return (
    <div className="fixed inset-0 z-50 flex flex-col bg-black/90 backdrop-blur-sm p-2 md:p-4 select-none animate-fadeIn">
      <div className="w-full h-full max-w-7xl mx-auto flex flex-col bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl overflow-hidden text-slate-200">
        {/* 상단 헤더 바 */}
        <div className="px-4 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-2.5 min-w-0">
            <span className="p-1.5 rounded-lg bg-orange-500/15 text-orange-400 border border-orange-500/30">
              <Presentation size={18} />
            </span>
            <div className="truncate">
              <div className="flex items-center gap-2">
                <h2 className="text-sm md:text-base font-bold text-slate-100 truncate" title={title}>
                  {title}
                </h2>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-orange-500/20 text-orange-300 border border-orange-500/40 font-bold uppercase shrink-0">
                  PPTX 에디터
                </span>
                {hasChanges && (
                  <span className="text-[11px] text-amber-400 flex items-center gap-1 font-medium">
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse"></span>
                    수정됨
                  </span>
                )}
              </div>
            </div>
          </div>

          {/* 중앙 뷰 분할 모드 토글 */}
          <div className="hidden sm:flex items-center bg-slate-950 p-1 rounded-xl border border-slate-800 text-xs gap-1">
            <button
              type="button"
              onClick={() => setViewMode('split')}
              className={`px-2.5 py-1 rounded-lg font-medium transition-all cursor-pointer ${
                viewMode === 'split' ? 'bg-orange-600 text-white shadow-sm' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              편집 + 미리보기
            </button>
            <button
              type="button"
              onClick={() => setViewMode('editOnly')}
              className={`px-2.5 py-1 rounded-lg font-medium transition-all cursor-pointer ${
                viewMode === 'editOnly' ? 'bg-orange-600 text-white shadow-sm' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              편집창만
            </button>
            <button
              type="button"
              onClick={() => setViewMode('previewOnly')}
              className={`px-2.5 py-1 rounded-lg font-medium transition-all cursor-pointer ${
                viewMode === 'previewOnly' ? 'bg-orange-600 text-white shadow-sm' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              미리보기만
            </button>
          </div>

          {/* 우측 상단 액션 */}
          <div className="flex items-center gap-2 shrink-0">
            {saveToast && (
              <div className="hidden sm:flex items-center gap-1.5 text-xs text-emerald-300 bg-emerald-950/40 border border-emerald-500/30 px-2.5 py-1 rounded-lg">
                <CheckCircle size={13} />
                <span>{saveToast}</span>
              </div>
            )}
            {errorToast && (
              <div className="hidden sm:flex items-center gap-1.5 text-xs text-rose-300 bg-rose-950/40 border border-rose-500/30 px-2.5 py-1 rounded-lg">
                <AlertCircle size={13} />
                <span>{errorToast}</span>
              </div>
            )}

            {/* 저장 및 반영 */}
            <button
              type="button"
              onClick={handleSave}
              disabled={isSaving}
              className="px-3.5 py-1.5 bg-orange-600 hover:bg-orange-500 text-white rounded-lg text-xs font-semibold shadow-md flex items-center gap-1.5 transition-all active:scale-95 disabled:opacity-50 cursor-pointer"
              title="수정한 슬라이드를 메모 첨부파일에 저장 (Ctrl + S)"
            >
              {isSaving ? <Loader2 size={13} className="animate-spin" /> : <Save size={13} />}
              <span>{isSaving ? '저장 중...' : '💾 저장 및 반영'}</span>
            </button>

            {/* 다운로드 */}
            <button
              type="button"
              onClick={handleDownload}
              className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors cursor-pointer"
              title="수정된 파워포인트 파일 다운로드"
            >
              <Download size={15} />
            </button>

            {/* 닫기 */}
            <button
              type="button"
              onClick={handleClose}
              className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700 transition-colors cursor-pointer"
              title="닫기 (Esc)"
            >
              <X size={16} />
            </button>
          </div>
        </div>

        {/* 메인 뷰: 좌측 슬라이드 목록 + 중앙 편집 영역 / 실시간 미리보기 */}
        <div className="flex-1 flex overflow-hidden min-h-0 relative">
          {loading && (
            <div className="absolute inset-0 z-30 flex flex-col items-center justify-center bg-slate-950/80 gap-3">
              <Loader2 size={32} className="animate-spin text-orange-400" />
              <span className="text-xs text-slate-400 font-medium">슬라이드 에디터 준비 중...</span>
            </div>
          )}

          {/* 좌측 슬라이드 관리 패널 */}
          <div className="w-56 sm:w-64 bg-slate-950 border-r border-slate-800 flex flex-col shrink-0 overflow-y-auto p-3 gap-2 select-none">
            <div className="flex items-center justify-between pb-2 border-b border-slate-800">
              <span className="text-xs font-bold text-slate-300">슬라이드 ({slides.length})</span>
              <button
                type="button"
                onClick={handleAddSlide}
                className="px-2 py-1 bg-orange-600/30 hover:bg-orange-600 text-orange-200 hover:text-white rounded-lg text-xs font-medium border border-orange-500/40 flex items-center gap-1 cursor-pointer"
                title="새 슬라이드 추가"
              >
                <Plus size={12} />
                <span>추가</span>
              </button>
            </div>

            <div className="space-y-2 flex-1 overflow-y-auto pr-0.5">
              {slides.map((s, idx) => {
                const isActive = idx === activeSlideIndex;
                return (
                  <div
                    key={idx}
                    onClick={() => setActiveSlideIndex(idx)}
                    className={`p-2.5 rounded-xl border transition-all cursor-pointer flex flex-col gap-1.5 ${
                      isActive
                        ? 'bg-orange-950/30 border-orange-500/60 ring-2 ring-orange-500/40'
                        : 'bg-slate-900/70 border-slate-800 hover:bg-slate-850 hover:border-slate-700'
                    }`}
                  >
                    <div className="flex items-center justify-between text-xs">
                      <span className={`px-1.5 py-0.5 rounded font-mono font-bold text-[10px] ${
                        isActive ? 'bg-orange-500 text-white' : 'bg-slate-800 text-slate-400'
                      }`}>
                        #{s.slideNumber}
                      </span>

                      {/* 순서 변경 및 삭제 */}
                      <div className="flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
                        <button
                          type="button"
                          onClick={() => handleMoveUp(idx)}
                          disabled={idx === 0}
                          className="p-1 hover:bg-slate-800 disabled:opacity-20 text-slate-400 hover:text-white rounded"
                          title="위로 이동"
                        >
                          <ArrowUp size={11} />
                        </button>
                        <button
                          type="button"
                          onClick={() => handleMoveDown(idx)}
                          disabled={idx === slides.length - 1}
                          className="p-1 hover:bg-slate-800 disabled:opacity-20 text-slate-400 hover:text-white rounded"
                          title="아래로 이동"
                        >
                          <ArrowDown size={11} />
                        </button>
                        <button
                          type="button"
                          onClick={() => handleDeleteSlide(idx)}
                          disabled={slides.length <= 1}
                          className="p-1 hover:bg-rose-950/60 disabled:opacity-20 text-slate-400 hover:text-rose-300 rounded"
                          title="슬라이드 삭제"
                        >
                          <Trash2 size={11} />
                        </button>
                      </div>
                    </div>

                    <div className="font-semibold text-xs text-slate-200 truncate">
                      {s.title || `(제목 없음)`}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* 중앙/우측 편집 및 실시간 미리보기 */}
          {currentSlide && (
            <div className="flex-1 flex flex-col md:flex-row overflow-hidden bg-slate-925">
              {/* 편집기 영역 */}
              {viewMode !== 'previewOnly' && (
                <div className={`p-4 md:p-6 flex flex-col overflow-y-auto border-r border-slate-800/80 ${
                  viewMode === 'split' ? 'w-full md:w-1/2' : 'w-full'
                }`}>
                  <div className="flex items-center gap-2 mb-3 text-xs font-bold text-orange-400 uppercase tracking-wider">
                    <Type size={14} />
                    <span>슬라이드 {currentSlide.slideNumber} 내용 직접 수정</span>
                  </div>

                  {/* 제목 입력 */}
                  <div className="mb-4">
                    <label className="block text-xs font-semibold text-slate-300 mb-1.5">
                      슬라이드 제목
                    </label>
                    <input
                      type="text"
                      value={currentSlide.title}
                      onChange={(e) => handleUpdateTitle(e.target.value)}
                      placeholder="슬라이드 제목을 입력하세요..."
                      className="w-full px-3.5 py-2.5 bg-slate-900 border border-slate-700/80 rounded-xl text-slate-100 font-bold text-base focus:outline-none focus:border-orange-500 transition-colors shadow-inner"
                    />
                  </div>

                  {/* 본문 단락 / 불릿 입력 */}
                  <div className="flex-1 flex flex-col">
                    <div className="flex items-center justify-between mb-1.5">
                      <label className="block text-xs font-semibold text-slate-300">
                        본문 내용 (줄 단위 글머리 기호)
                      </label>
                      <span className="text-[11px] text-slate-500">
                        줄 바꿈(Enter)으로 항목을 나눕니다
                      </span>
                    </div>
                    <textarea
                      value={currentSlide.bodyText}
                      onChange={(e) => handleUpdateBody(e.target.value)}
                      placeholder="• 첫 번째 항목&#10;• 두 번째 항목..."
                      className="w-full flex-1 min-h-[220px] p-3.5 bg-slate-900 border border-slate-700/80 rounded-xl text-slate-100 font-sans text-sm leading-6 resize-none focus:outline-none focus:border-orange-500 transition-colors shadow-inner"
                    />
                  </div>
                </div>
              )}

              {/* 실시간 16:9 슬라이드 프리뷰 영역 */}
              {viewMode !== 'editOnly' && (
                <div className={`p-4 md:p-6 flex flex-col items-center justify-center bg-slate-950 overflow-y-auto ${
                  viewMode === 'split' ? 'w-full md:w-1/2' : 'w-full'
                }`}>
                  <div className="w-full max-w-xl flex items-center justify-between text-xs text-slate-400 mb-2">
                    <span className="flex items-center gap-1 font-semibold text-orange-400">
                      <Eye size={13} />
                      <span>실시간 슬라이드 미리보기</span>
                    </span>
                    <span className="font-mono">16 : 9</span>
                  </div>

                  {/* 16:9 슬라이드 카드 */}
                  <div className="w-full max-w-xl aspect-video bg-white text-slate-900 rounded-2xl shadow-2xl overflow-hidden flex flex-col border border-slate-200">
                    <div className="h-2 bg-gradient-to-r from-orange-500 via-amber-500 to-indigo-600 shrink-0" />
                    <div className="flex-1 p-5 sm:p-6 flex flex-col overflow-y-auto">
                      <h1 className="text-lg sm:text-xl font-extrabold text-slate-900 pb-2 mb-3 border-b border-orange-500/30">
                        {currentSlide.title || '(제목 없음)'}
                      </h1>

                      <div className="space-y-2 flex-1">
                        {currentSlide.bodyText
                          .split('\n')
                          .filter((l) => l.trim().length > 0)
                          .map((line, lIdx) => (
                            <div key={lIdx} className="flex items-start gap-2 text-xs sm:text-sm leading-relaxed text-slate-800">
                              <span className="text-orange-500 font-bold select-none">•</span>
                              <span>{line.replace(/^[-*•]\s*/, '')}</span>
                            </div>
                          ))}

                        {currentSlide.images.length > 0 && (
                          <div className="mt-3 flex gap-2">
                            {currentSlide.images.map((img) => (
                              <img
                                key={img.id}
                                src={img.dataUrl}
                                alt="Slide image"
                                className="max-h-24 rounded border object-contain shadow-xs"
                              />
                            ))}
                          </div>
                        )}
                      </div>

                      <div className="mt-auto pt-2 flex items-center justify-between text-[10px] text-slate-400 select-none">
                        <span>{title}</span>
                        <span>Slide {currentSlide.slideNumber}</span>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
