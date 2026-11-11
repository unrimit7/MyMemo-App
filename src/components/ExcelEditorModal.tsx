import React, { useState, useEffect, useRef } from 'react';
import {
  X,
  Save,
  Download,
  FileSpreadsheet,
  CheckCircle,
  AlertCircle,
  Loader2,
} from 'lucide-react';
import { createUniver, defaultTheme, LocaleType } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreKoKR from '@univerjs/preset-sheets-core/locales/ko-KR';
import '@univerjs/preset-sheets-core/lib/index.css';

import { dataUrlToUniverWorkbook, univerSnapshotToDataUrl } from '../utils/univerAdapter';
import { invalidateExcelCache } from '../utils/excelViewer';
import { useFloatingWindow } from './useFloatingWindow';

interface ExcelEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
}

export const ExcelEditorModal: React.FC<ExcelEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  attId,
  onSave,
}) => {
  const isCsv = title.toLowerCase().endsWith('.csv');
  const format: 'xlsx' | 'csv' = isCsv ? 'csv' : 'xlsx';

  const floating = useFloatingWindow({ id: 'excel-editor', defaultWidth: 1040, defaultHeight: 840 });

  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [saveToast, setSaveToast] = useState<string | null>(null);
  const [errorToast, setErrorToast] = useState<string | null>(null);
  const [showExitConfirm, setShowExitConfirm] = useState<boolean>(false);

  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<any>(null);
  const univerAPIRef = useRef<any>(null);

  // 저장되지 않은 변경사항이 있을 때 브라우저 닫기/새로고침 방지
  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (hasChanges) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, [hasChanges]);

  // Univer 인스턴스 초기화 및 라이프사이클 관리
  useEffect(() => {
    if (!isOpen || !data || !containerRef.current) return;

    let isDisposed = false;
    let univerInstance: any = null;

    // 약간의 지연 후 렌더링 컨테이너가 안정화되었을 때 Univer 초기화
    const timer = setTimeout(() => {
      if (isDisposed || !containerRef.current) return;

      try {
        // 기존 인스턴스 정리
        if (univerRef.current) {
          try { univerRef.current.dispose(); } catch {}
          univerRef.current = null;
          univerAPIRef.current = null;
        }

        const workbookData = dataUrlToUniverWorkbook(data, title);

        const { univer, univerAPI } = createUniver({
          locale: LocaleType.KO_KR,
          locales: {
            [LocaleType.KO_KR]: UniverPresetSheetsCoreKoKR,
          },
          theme: defaultTheme,
          presets: [
            UniverSheetsCorePreset({
              container: containerRef.current,
              header: true,
              toolbar: true,
              formulaBar: true,
              footer: {
                sheetBar: true,
                statisticBar: true,
              },
            }),
          ],
        });

        univerInstance = univer;
        univerRef.current = univer;
        univerAPIRef.current = univerAPI;

        univerAPI.createWorkbook(workbookData);

        // 창 크기 변경에 따른 Univer 리사이즈 동기화
        const resizeObserver = new ResizeObserver(() => {
          window.dispatchEvent(new Event('resize'));
        });
        if (containerRef.current) {
          resizeObserver.observe(containerRef.current);
        }

        // 변경 감지 (커맨드 실행 시)
        let isInitial = true;
        setTimeout(() => {
          isInitial = false;
        }, 800);

        const commandService = (univerAPI as any)._commandService;
        let disposable: any = null;
        if (commandService && typeof commandService.onCommandExecuted === 'function') {
          disposable = commandService.onCommandExecuted(() => {
            if (!isInitial && !isDisposed) {
              setHasChanges(true);
            }
          });
        }

        return () => {
          resizeObserver.disconnect();
          try { disposable?.dispose?.(); } catch {}
        };
      } catch (err) {
        console.error('Univer 엑셀 엔진 초기화 실패:', err);
        setErrorToast('엑셀 뷰어를 불러오는 중 오류가 발생했습니다.');
      }
    }, 60);

    return () => {
      isDisposed = true;
      clearTimeout(timer);
      if (univerInstance) {
        try { univerInstance.dispose(); } catch {}
      }
      univerRef.current = null;
      univerAPIRef.current = null;
    };
  }, [isOpen, data, title]);

  // 저장 처리
  const handleSave = async () => {
    if (!univerAPIRef.current || isSaving) return;

    try {
      setIsSaving(true);
      const activeWb = univerAPIRef.current.getActiveWorkbook();
      if (!activeWb) throw new Error('활성 워크북을 찾을 수 없습니다.');

      const snapshot = activeWb.save();
      const updatedDataUrl = univerSnapshotToDataUrl(snapshot, format);

      invalidateExcelCache(attId);

      if (onSave) {
        await onSave(updatedDataUrl);
      }

      setHasChanges(false);
      setSaveToast('메모에 엑셀이 저장되었습니다!');
      setTimeout(() => setSaveToast(null), 2500);
    } catch (err) {
      console.error('엑셀 저장 실패:', err);
      setErrorToast('저장에 실패했습니다.');
      setTimeout(() => setErrorToast(null), 3000);
    } finally {
      setIsSaving(false);
    }
  };

  // 다운로드 처리
  const handleDownload = () => {
    if (!univerAPIRef.current) return;
    try {
      const activeWb = univerAPIRef.current.getActiveWorkbook();
      if (!activeWb) return;

      const snapshot = activeWb.save();
      const dataUrl = univerSnapshotToDataUrl(snapshot, format);

      const a = document.createElement('a');
      a.href = dataUrl;
      const downloadName = title.toLowerCase().endsWith(`.${format}`) ? title : `${title}.${format}`;
      a.download = downloadName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch (err) {
      console.error('다운로드 중 오류:', err);
    }
  };

  // 닫기 요청
  const handleRequestClose = () => {
    if (hasChanges) {
      setShowExitConfirm(true);
    } else {
      onClose();
    }
  };

  // 변경사항 저장 후 닫기
  const handleSaveAndExit = async () => {
    await handleSave();
    setShowExitConfirm(false);
    onClose();
  };

  // 저장하지 않고 닫기
  const handleDiscardAndExit = () => {
    setShowExitConfirm(false);
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div
      className={floating.modalWrapperClass}
      tabIndex={0}
      onKeyDown={(e) => {
        if (e.key === 'Escape' && !showExitConfirm) {
          handleRequestClose();
        }
      }}
    >
      {/* 팝업 컨테이너 */}
      <div {...floating.windowContainerProps}>
        {/* 상단 1단: 제목 & 저장 / 다운로드 / 닫기 헤더 */}
        <div
          {...floating.headerProps}
          className="px-4 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0 select-none"
        >
          <div className="flex items-center gap-2.5 min-w-0">
            <span className="p-1.5 rounded-lg bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
              <FileSpreadsheet size={18} />
            </span>
            <div className="truncate">
              <div className="flex items-center gap-2">
                <h2 className="text-sm md:text-base font-bold text-slate-100 truncate" title={title}>
                  {title}
                </h2>
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-emerald-500/10 text-emerald-300 border border-emerald-500/20 font-bold uppercase shrink-0">
                  {format} 엑셀 편집기
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

          {/* 우측 저장 및 창 제어 버튼 */}
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

            {/* 저장 및 반영 버튼 */}
            <button
              type="button"
              onClick={handleSave}
              disabled={isSaving}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold shadow-xs transition-all cursor-pointer ${
                hasChanges
                  ? 'bg-emerald-600 hover:bg-emerald-500 text-white shadow-emerald-900/30 animate-pulse'
                  : 'bg-emerald-700/80 hover:bg-emerald-600 text-white'
              } disabled:opacity-50`}
              title="편집한 내용을 메모에 즉시 저장합니다 (Ctrl+S)"
            >
              {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
              <span>{isSaving ? '저장 중...' : '저장 및 반영'}</span>
            </button>

            {/* 다운로드 버튼 */}
            <button
              type="button"
              onClick={handleDownload}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-medium border border-slate-700 transition-colors cursor-pointer"
              title="내 컴퓨터에 엑셀 파일로 다운로드"
            >
              <Download size={14} />
              <span className="hidden sm:inline">다운로드</span>
            </button>

            {/* 창 분리 (플로팅 모드) 토글 버튼 */}
            {floating.toggleFloatingButton}

            {/* 닫기 버튼 */}
            <button
              type="button"
              onClick={handleRequestClose}
              className="p-1.5 text-slate-400 hover:text-rose-400 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer"
              title="닫기 (Esc)"
            >
              <X size={16} />
            </button>
          </div>
        </div>

        {/* 바디: Univer 진짜 엑셀 스프레드시트 뷰어/편집기 */}
        <div className="flex-1 min-h-0 relative w-full h-full bg-white overflow-hidden">
          <div
            ref={containerRef}
            className="w-full h-full univer-container"
            style={{ width: '100%', height: '100%', position: 'relative' }}
          />
        </div>

        {/* 창 크기 조절 핸들 (플로팅 모드일 때) */}
        {floating.resizeHandle}
      </div>

      {/* 저장되지 않은 변경사항 닫기 확인 모달 */}
      {showExitConfirm && (
        <div className="fixed inset-0 z-60 bg-black/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700/80 rounded-2xl shadow-2xl max-w-sm w-full p-5 space-y-4 animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-start gap-3">
              <div className="p-2 rounded-xl bg-amber-500/10 text-amber-400 border border-amber-500/20 shrink-0">
                <AlertCircle size={22} />
              </div>
              <div>
                <h3 className="text-sm font-bold text-slate-100">저장되지 않은 변경사항이 있습니다</h3>
                <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                  수정하신 엑셀 내용이 아직 저장되지 않았습니다. 저장하지 않고 닫으면 작업 내용이 사라집니다.
                </p>
              </div>
            </div>
            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-800">
              <button
                type="button"
                onClick={() => setShowExitConfirm(false)}
                className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium transition-colors"
              >
                취소
              </button>
              <button
                type="button"
                onClick={handleDiscardAndExit}
                className="px-3 py-1.5 rounded-lg bg-rose-900/40 hover:bg-rose-900/60 text-rose-300 border border-rose-700/40 text-xs font-medium transition-colors"
              >
                저장하지 않고 닫기
              </button>
              <button
                type="button"
                onClick={handleSaveAndExit}
                className="px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold shadow-xs transition-colors"
              >
                저장하고 닫기
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

