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
  Type,
  ImageIcon,
  Maximize2,
  Table as TableIcon,
  LayoutTemplate,
} from 'lucide-react';
import {
  parsePptxPresentation,
  type PptxPresentation,
} from '../utils/pptxViewer';
import {
  savePptxPresentation,
} from '../utils/pptxAdapter';
import {
  detectPptxUnsupportedFeatures,
  type PptxUnsupportedFeatureReport,
} from '../utils/pptxLossDetector';
import { useFloatingWindow } from './useFloatingWindow';

interface PptxEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
}

export interface EditableTableData {
  id: string;
  rows: string[][];
}

interface EditableSlide {
  slideNumber: number;
  originalSlideIndex?: number;
  title: string;
  bodyText: string;
  images: { id: string; dataUrl: string }[];
  tables: EditableTableData[];
}

export const PptxEditorModal: React.FC<PptxEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  onSave,
}) => {
  const [slides, setSlides] = useState<EditableSlide[]>([]);
  const [deletedSlideIndices, setDeletedSlideIndices] = useState<number[]>([]);
  const [activeSlideIndex, setActiveSlideIndex] = useState<number>(0);
  const [loading, setLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [showExitConfirm, setShowExitConfirm] = useState<boolean>(false);
  const [unsupportedReport, setUnsupportedReport] = useState<PptxUnsupportedFeatureReport | null>(null);

  // 브라우저 탭/창 종료 시 미저장 경고
  useEffect(() => {
    if (!isOpen || !hasChanges) return;

    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      e.preventDefault();
      e.returnValue = '';
      return '';
    };

    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
    };
  }, [isOpen, hasChanges]);

  const [saveToast, setSaveToast] = useState<string | null>(null);
  const [errorToast, setErrorToast] = useState<string | null>(null);
  const [viewMode, setViewMode] = useState<'split' | 'editOnly' | 'previewOnly'>('split');
  const [previewZoom, setPreviewZoom] = useState<number>(100);
  const [lockAspectRatio, setLockAspectRatio] = useState<boolean>(true);
  const [lightboxImage, setLightboxImage] = useState<string | null>(null);
  const floating = useFloatingWindow({ id: 'pptx-editor', defaultWidth: 940, defaultHeight: 820 });

  // 슬라이드 화면 비율 및 방향(가로/세로) 정보
  const [presentationAspect, setPresentationAspect] = useState<{
    slideWidth: number;
    slideHeight: number;
    aspectRatio: number;
    isLandscape: boolean;
    label: string;
  }>({
    slideWidth: 12192000,
    slideHeight: 6858000,
    aspectRatio: 16 / 9,
    isLandscape: true,
    label: '16:9 와이드 가로',
  });

  // 원본 PPTX 파싱 및 편집용 슬라이드 상태 초기화
  useEffect(() => {
    if (!isOpen || !data) return;

    let isMounted = true;
    setLoading(true);
    setErrorToast(null);

    detectPptxUnsupportedFeatures(data).then((rep) => {
      if (isMounted) setUnsupportedReport(rep);
    });

    parsePptxPresentation(data)
      .then((pres: PptxPresentation) => {
        if (!isMounted) return;

        // 프레젠테이션 페이지 속성(가로/세로 및 비율) 분석
        const w = pres.slideWidth || 12192000;
        const h = pres.slideHeight || 6858000;
        const ratio = pres.aspectRatio || (w / h);
        const isLandscape = pres.isLandscape ?? (ratio >= 1.0);

        let label = isLandscape ? '가로형' : '세로형';
        if (Math.abs(ratio - 16 / 9) < 0.05) {
          label = '16:9 와이드 가로';
        } else if (Math.abs(ratio - 16 / 10) < 0.05) {
          label = '16:10 와이드 가로';
        } else if (Math.abs(ratio - 4 / 3) < 0.05) {
          label = '4:3 일반 가로';
        } else if (Math.abs(ratio - Math.SQRT2) < 0.05) {
          label = 'A4 가로';
        } else if (Math.abs(ratio - (1 / Math.SQRT2)) < 0.05) {
          label = 'A4 세로';
        } else if (Math.abs(ratio - 9 / 16) < 0.05) {
          label = '9:16 와이드 세로';
        } else {
          label = ratio >= 1 ? `${ratio.toFixed(2)}:1 가로` : `1:${(1 / ratio).toFixed(2)} 세로`;
        }

        setPresentationAspect({
          slideWidth: w,
          slideHeight: h,
          aspectRatio: ratio,
          isLandscape,
          label,
        });

        const editableList: EditableSlide[] = pres.slides.map((s, idx) => {
          const bodyLines = s.paragraphs
            .filter((p) => !p.isTitle && p.text !== s.title)
            .map((p) => {
              const indent = '  '.repeat(Math.max(0, p.level));
              return `${indent}• ${p.text}`;
            });

          // 슬라이드 내 표(Table) 목록 추출
          const slideTables: EditableTableData[] = [];
          for (const el of s.elements) {
            if (el.type === 'table') {
              const tableRows = el.rows.map((row) =>
                row.cells.map((cell) => cell.text)
              );
              slideTables.push({
                id: el.id,
                rows: tableRows,
              });
            }
          }

          return {
            slideNumber: s.slideNumber,
            originalSlideIndex: idx,
            title: s.title,
            bodyText: bodyLines.join('\n'),
            images: s.images.map((img) => ({ id: img.id, dataUrl: img.dataUrl })),
            tables: slideTables,
          };
        });

        setSlides(editableList);
        setDeletedSlideIndices([]);
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

  // 활성 슬라이드 표 셀 내용 수정
  const handleUpdateTableCell = (
    tableIdx: number,
    rowIndex: number,
    colIndex: number,
    newText: string
  ) => {
    setSlides((prev) => {
      const copy = [...prev];
      const current = { ...copy[activeSlideIndex] };
      current.tables = current.tables.map((tbl, tIdx) => {
        if (tIdx !== tableIdx) return tbl;
        const newRows = tbl.rows.map((r, rIdx) => {
          if (rIdx !== rowIndex) return [...r];
          const newCells = [...r];
          newCells[colIndex] = newText;
          return newCells;
        });
        return { ...tbl, rows: newRows };
      });
      copy[activeSlideIndex] = current;
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
      tables: [],
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
    const targetSlide = slides[index];
    if (!confirm(`'슬라이드 ${targetSlide.slideNumber}'을(를) 삭제하시겠습니까?`)) {
      return;
    }
    if (targetSlide.originalSlideIndex !== undefined && targetSlide.originalSlideIndex >= 0) {
      setDeletedSlideIndices((prev) => [...prev, targetSlide.originalSlideIndex!]);
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

  // 저장 및 메모 반영 (새 OpenXML In-Place 어댑터로 무손실 diff -> run 저장)
  const handleSave = async () => {
    setIsSaving(true);
    setSaveToast(null);
    setErrorToast(null);

    try {
      const updatedDataUrl = await savePptxPresentation(
        data,
        slides.map((s) => ({
          slideNumber: s.slideNumber,
          originalSlideIndex: s.originalSlideIndex,
          title: s.title,
          bodyText: s.bodyText,
          tables: s.tables.map((t) => ({ rows: t.rows })),
        })),
        deletedSlideIndices
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
      const updatedDataUrl = await savePptxPresentation(
        data,
        slides.map((s) => ({
          slideNumber: s.slideNumber,
          originalSlideIndex: s.originalSlideIndex,
          title: s.title,
          bodyText: s.bodyText,
          tables: s.tables.map((t) => ({ rows: t.rows })),
        })),
        deletedSlideIndices
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

  // 원본으로 되돌리기
  const handleRevertToOriginal = () => {
    if (window.confirm('모든 수정을 취소하고 파일의 원래 상태로 되돌리시겠습니까?')) {
      setLoading(true);
      parsePptxPresentation(data)
        .then((pres) => {
          const editableList: EditableSlide[] = pres.slides.map((s, idx) => {
            const bodyLines = s.paragraphs
              .filter((p) => !p.isTitle && p.text !== s.title)
              .map((p) => {
                const indent = '  '.repeat(Math.max(0, p.level));
                return `${indent}• ${p.text}`;
              });
            const slideTables: EditableTableData[] = [];
            for (const el of s.elements) {
              if (el.type === 'table') {
                const tableRows = el.rows.map((row) => row.cells.map((cell) => cell.text));
                slideTables.push({ id: el.id, rows: tableRows });
              }
            }
            return {
              slideNumber: s.slideNumber,
              originalSlideIndex: idx,
              title: s.title,
              bodyText: bodyLines.join('\n'),
              images: s.images.map((img) => ({ id: img.id, dataUrl: img.dataUrl })),
              tables: slideTables,
            };
          });
          setSlides(editableList);
          setDeletedSlideIndices([]);
          setHasChanges(false);
          setLoading(false);
          setSaveToast('원본 상태로 되돌렸습니다.');
          setTimeout(() => setSaveToast(null), 3000);
        })
        .catch(() => {
          setLoading(false);
        });
    }
  };

  // 닫기
  const handleClose = () => {
    if (hasChanges) {
      setShowExitConfirm(true);
      return;
    }
    onClose();
  };

  const handleSaveAndExit = async () => {
    await handleSave();
    setShowExitConfirm(false);
    onClose();
  };

  const handleDiscardAndExit = () => {
    setHasChanges(false);
    setShowExitConfirm(false);
    onClose();
  };

  if (!isOpen) return null;

  const currentSlide: EditableSlide | undefined = slides[activeSlideIndex];

  return (
    <div className={floating.modalWrapperClass}>
      <div {...floating.windowContainerProps} className={floating.windowContainerProps.className + ' text-slate-200'}>
        {/* 상단 헤더 바 */}
        <div 
          {...floating.headerProps}
          className="px-3 sm:px-4 py-2 bg-slate-900 border-b border-slate-800 flex items-center justify-between gap-2 sm:gap-3 shrink-0 select-none"
        >
          {/* 좌측: 파일 제목 & 상태 배지 */}
          <div className="flex items-center gap-2 min-w-0 max-w-[32%] sm:max-w-[42%] md:max-w-md">
            <span className="p-1.5 rounded-lg bg-orange-500/15 text-orange-400 border border-orange-500/30 shrink-0">
              <Presentation size={17} />
            </span>
            <div className="truncate min-w-0">
              <div className="flex items-center gap-1.5 sm:gap-2">
                <h2 className="text-xs sm:text-sm font-bold text-slate-100 truncate" title={title}>
                  {title}
                </h2>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-orange-500/20 text-orange-300 border border-orange-500/40 font-bold uppercase shrink-0 hidden xs:inline-block">
                  PPTX 에디터
                </span>
                {hasChanges && (
                  <span className="text-[11px] text-amber-400 flex items-center gap-1 font-medium shrink-0">
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse"></span>
                    <span className="hidden sm:inline">수정됨</span>
                  </span>
                )}
              </div>
            </div>
          </div>

          {/* 중앙 뷰 분할 모드 토글 */}
          <div className="flex items-center bg-slate-950 p-0.5 sm:p-1 rounded-xl border border-slate-800 text-xs gap-0.5 sm:gap-1 shrink-0">
            <button
              type="button"
              onClick={() => setViewMode('split')}
              className={`px-2 sm:px-2.5 py-1 rounded-lg font-medium transition-all cursor-pointer ${
                viewMode === 'split' ? 'bg-orange-600 text-white shadow-sm' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <span className="hidden sm:inline">편집 + 미리보기</span>
              <span className="sm:hidden">분할</span>
            </button>
            <button
              type="button"
              onClick={() => setViewMode('editOnly')}
              className={`px-2 sm:px-2.5 py-1 rounded-lg font-medium transition-all cursor-pointer ${
                viewMode === 'editOnly' ? 'bg-orange-600 text-white shadow-sm' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <span className="hidden sm:inline">편집창만</span>
              <span className="sm:hidden">편집</span>
            </button>
            <button
              type="button"
              onClick={() => setViewMode('previewOnly')}
              className={`px-2 sm:px-2.5 py-1 rounded-lg font-medium transition-all cursor-pointer ${
                viewMode === 'previewOnly' ? 'bg-orange-600 text-white shadow-sm' : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <span className="hidden sm:inline">미리보기만</span>
              <span className="sm:hidden">미리보기</span>
            </button>
          </div>

          {/* 우측 상단 액션 (항상 우측 최상단 끝에 고정) */}
          <div className="flex items-center gap-1.5 sm:gap-2 shrink-0 ml-auto">
            {saveToast && (
              <div className="hidden lg:flex items-center gap-1.5 text-xs text-emerald-300 bg-emerald-950/40 border border-emerald-500/30 px-2.5 py-1 rounded-lg">
                <CheckCircle size={13} />
                <span>{saveToast}</span>
              </div>
            )}
            {errorToast && (
              <div className="hidden lg:flex items-center gap-1.5 text-xs text-rose-300 bg-rose-950/40 border border-rose-500/30 px-2.5 py-1 rounded-lg">
                <AlertCircle size={13} />
                <span>{errorToast}</span>
              </div>
            )}

            {/* 저장 및 반영 */}
            <button
              type="button"
              onClick={handleSave}
              disabled={isSaving}
              className="px-2.5 sm:px-3.5 py-1.5 bg-orange-600 hover:bg-orange-500 text-white rounded-lg text-xs font-semibold shadow-md flex items-center gap-1.5 transition-all active:scale-95 disabled:opacity-50 cursor-pointer"
              title="수정한 슬라이드를 메모 첨부파일에 저장 (Ctrl + S)"
            >
              {isSaving ? <Loader2 size={13} className="animate-spin" /> : <Save size={13} />}
              <span className="hidden sm:inline">{isSaving ? '저장 중...' : '저장 및 반영'}</span>
              <span className="sm:hidden">저장</span>
            </button>

            {/* 원본으로 되돌리기 */}
            {hasChanges && (
              <button
                type="button"
                onClick={handleRevertToOriginal}
                className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg text-xs font-medium border border-slate-700 flex items-center gap-1 transition-colors cursor-pointer"
                title="모든 수정을 취소하고 원본 상태로 되돌리기"
              >
                <span>되돌리기</span>
              </button>
            )}

            {/* 다운로드 */}
            <button
              type="button"
              onClick={handleDownload}
              className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors cursor-pointer"
              title="수정된 파워포인트 파일 다운로드"
            >
              <Download size={15} />
            </button>

            {floating.toggleFloatingButton}

            {/* 닫기 (우측 최상단 모서리 배치) */}
            <button
              type="button"
              onClick={handleClose}
              className="p-1.5 bg-slate-800 hover:bg-rose-900/50 text-slate-400 hover:text-rose-200 rounded-lg border border-slate-700 hover:border-rose-500/50 transition-colors cursor-pointer"
              title="닫기 (Esc)"
            >
              <X size={16} />
            </button>
          </div>
        </div>

        {/* 미지원/고급 요소 감지 경고 배너 */}
        {unsupportedReport && unsupportedReport.unsupportedList.length > 0 && (
          <div className="px-4 py-2 bg-amber-950/60 border-b border-amber-800/80 text-amber-200 text-xs flex items-center justify-between gap-2 shrink-0">
            <div className="flex items-center gap-2">
              <AlertCircle size={15} className="text-amber-400 shrink-0" />
              <span>
                <strong>고급 개체 감지:</strong> 이 프레젠테이션에는 {unsupportedReport.unsupportedList.join(', ')}가 포함되어 있습니다. 본문 텍스트나 표 수정 후 저장 시 해당 고급 개체는 100% 무손실 보존됩니다.
              </span>
            </div>
          </div>
        )}

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

                    {s.images.length > 0 && (
                      <span className="text-[10px] text-slate-400 flex items-center gap-1">
                        <ImageIcon size={10} className="text-orange-400" /> 이미지 {s.images.length}개 보존됨
                      </span>
                    )}
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
                    <textarea
                      rows={2}
                      value={currentSlide.title}
                      onChange={(e) => handleUpdateTitle(e.target.value)}
                      placeholder="슬라이드 제목을 입력하세요..."
                      className="w-full px-3.5 py-2.5 bg-slate-900 border border-slate-700/80 rounded-xl text-slate-100 font-bold text-sm sm:text-base leading-snug resize-y focus:outline-none focus:border-orange-500 transition-colors shadow-inner"
                    />
                  </div>

                  {/* 본문 단락 / 불릿 입력 */}
                  <div className="flex-1 flex flex-col mb-4">
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
                      className="w-full min-h-[120px] p-3.5 bg-slate-900 border border-slate-700/80 rounded-xl text-slate-100 font-sans text-sm leading-6 resize-y focus:outline-none focus:border-orange-500 transition-colors shadow-inner"
                    />
                  </div>

                  {/* 슬라이드 표(Table) 편집 섹션 */}
                  {currentSlide.tables.length > 0 && (
                    <div className="mt-2 pt-4 border-t border-slate-800">
                      <div className="flex items-center justify-between mb-2.5">
                        <label className="flex items-center gap-1.5 text-xs font-bold text-orange-400 uppercase tracking-wider">
                          <TableIcon size={14} />
                          <span>슬라이드 표 내용 직접 수정 ({currentSlide.tables.length}개 표)</span>
                        </label>
                        <span className="text-[11px] text-slate-400">
                          셀 내용을 입력하면 실시간으로 미리보기 및 파일에 반영됩니다
                        </span>
                      </div>

                      <div className="space-y-4">
                        {currentSlide.tables.map((tbl, tblIdx) => (
                          <div
                            key={tbl.id || tblIdx}
                            className="rounded-xl border border-slate-700/80 bg-slate-900/90 overflow-hidden shadow-inner"
                          >
                            <div className="px-3 py-1.5 bg-slate-800/80 border-b border-slate-700/60 flex items-center justify-between text-xs text-slate-300">
                              <span className="font-semibold text-orange-300">
                                표 #{tblIdx + 1} ({tbl.rows.length}행 × {tbl.rows[0]?.length || 0}열)
                              </span>
                              <span className="text-[10px] text-emerald-400 font-medium">
                                편집 가능
                              </span>
                            </div>

                            <div className="max-h-72 overflow-auto p-2">
                              <table className="w-full text-xs border-collapse border border-slate-700">
                                <tbody>
                                  {tbl.rows.map((row, rIdx) => (
                                    <tr
                                      key={rIdx}
                                      className={
                                        rIdx === 0
                                          ? 'bg-slate-800/90 font-bold text-slate-100'
                                          : 'hover:bg-slate-850/60'
                                      }
                                    >
                                      {row.map((cellText, cIdx) => (
                                        <td
                                          key={cIdx}
                                          className="border border-slate-700/80 p-1 min-w-[100px]"
                                        >
                                          <input
                                            type="text"
                                            value={cellText}
                                            onChange={(e) =>
                                              handleUpdateTableCell(tblIdx, rIdx, cIdx, e.target.value)
                                            }
                                            className="w-full px-2 py-1 bg-slate-950/60 border border-slate-700/50 rounded text-slate-100 text-xs focus:outline-none focus:border-orange-500 focus:bg-slate-900 transition-colors"
                                            placeholder={`(${rIdx + 1}, ${cIdx + 1})`}
                                          />
                                        </td>
                                      ))}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* 포함된 이미지 확인 및 관리 섹션 */}
                  {currentSlide.images.length > 0 && (
                    <div className="mt-4 pt-3.5 border-t border-slate-800">
                      <div className="flex items-center justify-between mb-2">
                        <label className="flex items-center gap-1.5 text-xs font-semibold text-slate-200">
                          <ImageIcon size={14} className="text-orange-400" />
                          <span>슬라이드 포함 이미지 ({currentSlide.images.length}개)</span>
                        </label>
                        <span className="text-[11px] text-emerald-400 font-medium flex items-center gap-1 bg-emerald-950/40 px-2 py-0.5 rounded border border-emerald-500/30">
                          <CheckCircle size={11} />
                          <span>저장 시 원본 무손실 안전 보존</span>
                        </span>
                      </div>

                      <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">
                        {currentSlide.images.map((img, imgIdx) => (
                          <div
                            key={img.id || imgIdx}
                            className="relative group rounded-xl border border-slate-800 bg-slate-900 p-2 flex flex-col items-center overflow-hidden hover:border-orange-500/50 transition-colors"
                          >
                            <img
                              src={img.dataUrl}
                              alt={`Slide element ${imgIdx + 1}`}
                              onClick={() => setLightboxImage(img.dataUrl)}
                              className="w-full h-24 object-contain rounded-lg cursor-zoom-in group-hover:scale-105 transition-transform bg-white/5"
                              title="클릭하여 원본 고해상도로 확대 보기"
                            />
                            <div className="w-full mt-2 flex items-center justify-between text-[10px] text-slate-400">
                              <span className="font-mono">이미지 #{imgIdx + 1}</span>
                              <div className="flex items-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => setLightboxImage(img.dataUrl)}
                                  className="p-1 hover:bg-slate-800 text-slate-400 hover:text-orange-300 rounded cursor-pointer"
                                  title="원본 크게 보기"
                                >
                                  <Maximize2 size={12} />
                                </button>
                                <a
                                  href={img.dataUrl}
                                  download={`slide_${currentSlide.slideNumber}_img_${imgIdx + 1}.png`}
                                  className="p-1 hover:bg-slate-800 text-slate-400 hover:text-white rounded cursor-pointer no-underline"
                                  title="이미지 파일 다운로드"
                                >
                                  <Download size={12} />
                                </a>
                              </div>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* 실시간 슬라이드 프리뷰 영역 (PPTX 페이지 속성: 가로/세로 비율 연동) */}
              {viewMode !== 'editOnly' && (() => {
                const slideAspect = presentationAspect.aspectRatio || (16 / 9);
                const isLandscape = presentationAspect.isLandscape;
                const previewMaxWidth = viewMode === 'previewOnly' && lockAspectRatio
                  ? isLandscape
                    ? `min(94vw, calc(80vh * ${slideAspect}))`
                    : `min(88vw, calc(82vh * ${slideAspect}))`
                  : '100%';

                return (
                  <div className={`p-3 sm:p-5 md:p-6 flex flex-col items-center justify-start bg-slate-950 overflow-y-auto ${
                    viewMode === 'split' ? 'w-full md:w-1/2' : 'w-full'
                  }`}>
                    {/* 상단 컨트롤 바 (슬라이드 번호, 페이지 속성 표시, 비율 고정 토글, 줌) */}
                    <div
                      style={{ maxWidth: previewMaxWidth }}
                      className="w-full flex flex-wrap items-center justify-between text-xs text-slate-400 mb-2.5 px-1 gap-2"
                    >
                      <div className="flex items-center gap-2">
                        <span className="flex items-center gap-1.5 font-semibold text-orange-400">
                          <Eye size={13} />
                          <span>슬라이드 #{activeSlideIndex + 1} 미리보기</span>
                        </span>

                        {/* PPTX 페이지 속성 배지 (가로형/세로형 및 실제 비율) */}
                        <span
                          className={`px-2 py-0.5 rounded-md text-[11px] font-mono font-medium flex items-center gap-1 border ${
                            isLandscape
                              ? 'bg-emerald-950/40 border-emerald-500/40 text-emerald-300'
                              : 'bg-indigo-950/40 border-indigo-500/40 text-indigo-300'
                          }`}
                          title={`파워포인트 원본 페이지 속성: ${presentationAspect.label} (${presentationAspect.slideWidth} × ${presentationAspect.slideHeight} EMU)`}
                        >
                          <LayoutTemplate size={11} />
                          <span>{presentationAspect.label}</span>
                        </span>
                      </div>

                      <div className="flex items-center gap-2">
                        {/* 비율 고정 vs 전체 펼치기 토글 */}
                        <button
                          type="button"
                          onClick={() => setLockAspectRatio((prev) => !prev)}
                          className={`px-2 py-1 rounded-lg text-[11px] font-medium border transition-colors cursor-pointer flex items-center gap-1 ${
                            lockAspectRatio
                              ? 'bg-orange-500/15 border-orange-500/40 text-orange-300 hover:bg-orange-500/25'
                              : 'bg-slate-900 border-slate-800 text-slate-400 hover:text-slate-200'
                          }`}
                          title={lockAspectRatio ? '클릭 시 비율 제한을 풀고 내용 길이에 맞춰 전체를 펼쳐서 봅니다' : '클릭 시 실제 파워포인트 가로 슬라이드 비율(16:9)로 고정합니다'}
                        >
                          <span>{lockAspectRatio ? '슬라이드 비율 고정' : '내용 펼쳐보기'}</span>
                        </button>

                        {/* 줌 조절 버튼 */}
                        <div className="flex items-center gap-1 bg-slate-900 border border-slate-800 rounded-lg p-0.5 text-[11px]">
                          <button
                            type="button"
                            onClick={() => setPreviewZoom((z) => Math.max(50, z - 15))}
                            className="px-2 py-0.5 hover:bg-slate-800 rounded text-slate-300 cursor-pointer font-bold"
                            title="축소"
                          >
                            -
                          </button>
                          <span className="font-mono text-slate-300 px-1 min-w-[36px] text-center">{previewZoom}%</span>
                          <button
                            type="button"
                            onClick={() => setPreviewZoom((z) => Math.min(160, z + 15))}
                            className="px-2 py-0.5 hover:bg-slate-800 rounded text-slate-300 cursor-pointer font-bold"
                            title="확대"
                          >
                            +
                          </button>
                          <button
                            type="button"
                            onClick={() => setPreviewZoom(100)}
                            className="px-2 py-0.5 hover:bg-slate-800 rounded text-slate-400 hover:text-slate-200 cursor-pointer text-[10px]"
                            title="100% 기본 크기 복원"
                          >
                            100%
                          </button>
                        </div>
                      </div>
                    </div>

                    {/* 슬라이드 카드 (가로형/세로형 페이지 속성 엄격 준수) */}
                    <div
                      style={{
                        aspectRatio: lockAspectRatio ? `${slideAspect}` : undefined,
                        maxWidth: previewMaxWidth,
                        width: lockAspectRatio && !isLandscape ? `calc(82vh * ${slideAspect})` : '100%',
                        maxHeight: lockAspectRatio
                          ? isLandscape
                            ? viewMode === 'previewOnly' ? '80vh' : undefined
                            : '82vh'
                          : undefined,
                        transform: previewZoom !== 100 ? `scale(${previewZoom / 100})` : undefined,
                        transformOrigin: 'top center',
                        transition: 'transform 0.12s ease-out',
                      }}
                      className="bg-white text-slate-900 rounded-2xl shadow-2xl border border-slate-200 flex flex-col shrink-0 mb-8 overflow-hidden"
                    >
                      <div className="h-2.5 bg-gradient-to-r from-orange-500 via-amber-500 to-indigo-600 rounded-t-2xl shrink-0" />
                      <div className="p-6 sm:p-8 flex flex-col justify-between flex-1 overflow-y-auto">
                      <div>
                        {/* 슬라이드 대제목 */}
                        <h1 className="text-base sm:text-xl md:text-2xl font-extrabold text-slate-900 pb-2.5 mb-3.5 border-b-2 border-orange-500/30 whitespace-pre-wrap leading-snug">
                          {currentSlide.title || '(제목 없음)'}
                        </h1>

                        {/* 본문 텍스트 (글머리 기호 및 순번 자동 정리) */}
                        {currentSlide.bodyText.trim().length > 0 && (
                          <div className="space-y-2 mb-5">
                            {currentSlide.bodyText
                              .split('\n')
                              .map((l) => l.trim())
                              .filter((l) => l.length > 0)
                              .map((line, lIdx) => {
                                const isNumbered = /^\d+[\.\)]\s*/.test(line);
                                const cleanLine = line.replace(/^([-*•\u2022\u25E6\u25AA\u25CF]+\s*)+/, '');
                                return (
                                  <div key={lIdx} className="flex items-start gap-2.5 text-xs sm:text-sm leading-relaxed text-slate-800">
                                    {!isNumbered && (
                                      <span className="text-orange-500 font-bold select-none shrink-0">•</span>
                                    )}
                                    <span className="whitespace-pre-wrap">{cleanLine}</span>
                                  </div>
                                );
                              })}
                          </div>
                        )}

                        {/* 표(Table) 실시간 렌더링 (넓은 폭 & 깔끔한 셀 패딩 & 가로 잘림 방지) */}
                        {currentSlide.tables.length > 0 && (
                          <div className="space-y-4 my-4">
                            {currentSlide.tables.map((tbl, tblIdx) => (
                              <div
                                key={tbl.id || tblIdx}
                                className="overflow-x-auto rounded-xl border border-slate-300 shadow-xs bg-white"
                              >
                                <table className="w-full text-xs sm:text-sm border-collapse min-w-[500px]">
                                  <tbody>
                                    {tbl.rows.map((row, rIdx) => (
                                      <tr
                                        key={rIdx}
                                        className={
                                          rIdx === 0
                                            ? 'bg-slate-100 font-bold text-slate-900 border-b-2 border-slate-300'
                                            : rIdx % 2 === 1
                                            ? 'bg-white text-slate-800 border-b border-slate-200'
                                            : 'bg-slate-50/70 text-slate-800 border-b border-slate-200'
                                        }
                                      >
                                        {row.map((cellText, cIdx) => (
                                          <td
                                            key={cIdx}
                                            className="border-r border-slate-200 last:border-r-0 px-3 py-2 sm:px-4 sm:py-2.5 whitespace-pre-wrap align-top leading-relaxed min-w-[90px]"
                                          >
                                            {cellText || <span className="text-slate-300 italic">-</span>}
                                          </td>
                                        ))}
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            ))}
                          </div>
                        )}

                        {/* 슬라이드 이미지 목록 */}
                        {currentSlide.images.length > 0 && (
                          <div className="mt-4 flex flex-wrap gap-3 items-center justify-center">
                            {currentSlide.images.map((img, imgIdx) => (
                              <img
                                key={img.id || imgIdx}
                                src={img.dataUrl}
                                alt="Slide image"
                                onClick={() => setLightboxImage(img.dataUrl)}
                                className="max-h-36 sm:max-h-48 rounded-lg border border-slate-300 shadow-sm object-contain cursor-zoom-in hover:scale-105 transition-transform"
                                title="클릭하여 원본 고해상도로 확대 보기"
                              />
                            ))}
                          </div>
                        )}
                      </div>

                      {/* 하단 슬라이드 정보 바 */}
                      <div className="pt-3 border-t border-slate-200/80 flex items-center justify-between text-[11px] text-slate-400 font-mono mt-6 select-none">
                        <span className="truncate max-w-[280px]">{title}</span>
                        <span className="font-semibold text-slate-500">#{activeSlideIndex + 1} / {slides.length}</span>
                      </div>
                    </div>
                  </div>
                </div>
              );
            })()}
            </div>
          )}
        </div>

        {/* 저장되지 않은 변경사항 확인 모달 */}
        {showExitConfirm && (
          <div className="fixed inset-0 z-100 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
            <div className="bg-slate-900 border border-slate-700 rounded-2xl p-6 max-w-md w-full shadow-2xl space-y-4">
              <div className="flex items-center gap-3 text-amber-400">
                <div className="p-2.5 rounded-xl bg-amber-500/10 border border-amber-500/20">
                  <AlertCircle size={24} />
                </div>
                <div>
                  <h3 className="font-bold text-base text-slate-100">저장되지 않은 변경사항</h3>
                  <p className="text-xs text-slate-400">파워포인트 슬라이드 수정 내용이 아직 메모에 저장되지 않았습니다.</p>
                </div>
              </div>

              <p className="text-xs text-slate-300 leading-relaxed bg-slate-800/60 p-3 rounded-xl border border-slate-700/50">
                문서를 저장하지 않고 닫으시면 수정한 내용이 취소되고 원래 내용으로 복원됩니다.
              </p>

              <div className="flex flex-col sm:flex-row gap-2 pt-1">
                <button
                  type="button"
                  onClick={handleSaveAndExit}
                  disabled={isSaving}
                  className="flex-1 py-2.5 px-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold transition-all shadow-md active:scale-95 flex items-center justify-center gap-1.5 disabled:opacity-50"
                >
                  {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
                  <span>저장하고 닫기</span>
                </button>
                <button
                  type="button"
                  onClick={handleDiscardAndExit}
                  className="py-2.5 px-3 bg-rose-600/20 hover:bg-rose-600/30 text-rose-300 border border-rose-500/30 rounded-xl text-xs font-bold transition-all active:scale-95"
                >
                  저장하지 않고 닫기
                </button>
                <button
                  type="button"
                  onClick={() => setShowExitConfirm(false)}
                  className="py-2.5 px-3 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-medium transition-all"
                >
                  취소
                </button>
              </div>
            </div>
          </div>
        )}

        {floating.resizeHandle}
      </div>

      {/* 고해상도 이미지 라이트박스 (클릭 시 원본 확대 모달) */}
      {lightboxImage && (
        <div
          className="fixed inset-0 z-[100] bg-black/95 backdrop-blur-md flex flex-col items-center justify-center p-4 select-none animate-fadeIn"
          onClick={() => setLightboxImage(null)}
        >
          <div
            className="relative max-w-5xl w-full max-h-[92vh] flex flex-col items-center"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="w-full flex items-center justify-between pb-2 text-xs text-slate-300">
              <span className="font-semibold text-orange-400 flex items-center gap-1.5">
                <ImageIcon size={14} />
                <span>슬라이드 포함 이미지 원본 고해상도 확대 (글씨 선명하게 확인 가능)</span>
              </span>
              <div className="flex items-center gap-2">
                <a
                  href={lightboxImage}
                  download={`slide_image_${Date.now()}.png`}
                  className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-white rounded-lg border border-slate-700 flex items-center gap-1 text-xs no-underline cursor-pointer"
                  title="이미지 파일로 다운로드"
                >
                  <Download size={12} />
                  <span>다운로드</span>
                </a>
                <button
                  type="button"
                  onClick={() => setLightboxImage(null)}
                  className="p-1 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700 cursor-pointer"
                  title="닫기 (Esc)"
                >
                  <X size={16} />
                </button>
              </div>
            </div>
            <img
              src={lightboxImage}
              alt="Expanded slide element"
              className="max-w-full max-h-[84vh] rounded-xl shadow-2xl object-contain border border-slate-700 bg-slate-900"
            />
          </div>
        </div>
      )}
    </div>
  );
};
