import React, { useState, useEffect, useRef } from 'react';
import {
  X,
  Save,
  Download,
  FileSpreadsheet,
  CheckCircle,
  AlertCircle,
  AlertTriangle,
  RotateCcw,
  Loader2,
  FolderDown,
} from 'lucide-react';
import { createUniver, defaultTheme, LocaleType } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreKoKR from '@univerjs/preset-sheets-core/locales/ko-KR';
import '@univerjs/preset-sheets-core/lib/index.css';

import { dataUrlToUniverWorkbook, univerSnapshotToDataUrl, type StructuralOperation } from '../utils/univerAdapter';
import { invalidateExcelCache } from '../utils/excelViewer';
import { detectExcelUnsupportedFeatures, type UnsupportedFeatureReport } from '../utils/excelLossDetector';
import { useFloatingWindow } from './useFloatingWindow';
import {
  saveOfficeDraft,
  getOfficeDraft,
  deleteOfficeDraft,
  computeDataHash,
  type OfficeDraft,
} from '../utils/draftStorage';
import { useOfficeShortcuts } from '../utils/useOfficeShortcuts';
import { saveFileAs } from '../utils/fileSavePicker';
import {
  type FileHandleContext,
  checkFileHandleConflict,
  ensureOriginalBackup,
  saveToFileHandle,
} from '../utils/fileHandleSaver';

interface ExcelEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  originalData?: string;
  originalDate?: string;
  initialEngine?: 'exceljs' | 'sheetjs';
  fileHandle?: any;
  onSave?: (updatedDataUrl: string, engine: 'exceljs' | 'sheetjs') => Promise<void> | void;
  onRestoreOriginal?: () => Promise<void> | void;
}

export const ExcelEditorModal: React.FC<ExcelEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  attId,
  originalData,
  originalDate,
  initialEngine = 'exceljs',
  fileHandle,
  onSave,
  onRestoreOriginal,
}) => {
  const isCsv = title.toLowerCase().endsWith('.csv');
  const format: 'xlsx' | 'csv' = isCsv ? 'csv' : 'xlsx';

  const floating = useFloatingWindow({ id: 'excel-editor', defaultWidth: 1040, defaultHeight: 840 });

  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [saveToast, setSaveToast] = useState<string | null>(null);
  const [errorToast, setErrorToast] = useState<string | null>(null);
  const [isFallbackMode, setIsFallbackMode] = useState(false);
  const [showExitConfirm, setShowExitConfirm] = useState<boolean>(false);
  const [showRestoreConfirm, setShowRestoreConfirm] = useState<boolean>(false);
  const [showLossWarningModal, setShowLossWarningModal] = useState<boolean>(false);
  const [lossReport, setLossReport] = useState<UnsupportedFeatureReport | null>(null);

  // 5단계 임시본 복구 모달 및 파일 핸들러 충돌 모달 상태
  const draftKey = `excel_${attId || ''}_${title}`;
  const [draftToRecover, setDraftToRecover] = useState<OfficeDraft | null>(null);
  const [showDraftRecoveryModal, setShowDraftRecoveryModal] = useState<boolean>(false);
  const [showFileConflictModal, setShowFileConflictModal] = useState<boolean>(false);
  const fileContextRef = useRef<FileHandleContext | null>(null);

  // FileHandle 초기화 및 최초 덮어쓰기 전 백업 준비
  useEffect(() => {
    if (!fileHandle) return;
    (async () => {
      try {
        const file = await fileHandle.getFile();
        fileContextRef.current = {
          handle: fileHandle,
          initialSize: file.size,
          initialLastModified: file.lastModified,
          fileName: title,
        };
      } catch (err) {
        console.warn('FileHandle 초기화 실패:', err);
      }
    })();
  }, [fileHandle, title]);

  // 임시본 검사 및 원본 해시 일치 판별 (불일치 시 자동 폐기)
  useEffect(() => {
    if (!isOpen || !data) return;
    let isCancelled = false;

    (async () => {
      const draft = await getOfficeDraft(draftKey);
      if (!draft || isCancelled) return;

      const currentHash = await computeDataHash(data);
      if (draft.baseOriginalHash !== currentHash) {
        // 동기화 등으로 원본이 바뀐 경우 임시본 폐기 (복구 제안 원천 차단)
        await deleteOfficeDraft(draftKey);
      } else {
        setDraftToRecover(draft);
        setShowDraftRecoveryModal(true);
      }
    })();

    return () => {
      isCancelled = true;
    };
  }, [isOpen, data, draftKey]);

  // 30초 주기 + visibilitychange / blur 시점에만 동작하는 안전 자동 저장 (IndexedDB 전용)
  useEffect(() => {
    if (!isOpen || !hasChanges) return;

    let autoSaveIntervalMs = 30000;

    const performDraftSave = async () => {
      if (!univerRef.current || !hasChanges) return;
      try {
        const t0 = performance.now();
        const activeWb = univerRef.current.getActiveWorkbook();
        if (!activeWb) return;
        const snapshot = activeWb.save();
        const duration = performance.now() - t0;
        const baseHash = await computeDataHash(data);
        const snapshotStr = JSON.stringify(snapshot);

        await saveOfficeDraft({
          draftKey,
          attId,
          title,
          fileType: 'excel',
          baseOriginalHash: baseHash,
          data: snapshot,
          dataType: 'univer_snapshot',
          updatedAt: Date.now(),
          byteLength: snapshotStr.length,
          saveDurationMs: Math.round(duration),
        });

        // 1초 이상 걸리면 주기 60초로 자동 연장
        if (duration >= 1000) {
          autoSaveIntervalMs = 60000;
        }
      } catch (err) {
        console.warn('[Excel] 임시본 자동 저장 오류:', err);
      }
    };

    const interval = setInterval(performDraftSave, autoSaveIntervalMs);

    const handleVisibility = () => {
      if (document.visibilityState === 'hidden') {
        performDraftSave();
      }
    };
    const handleBlur = () => {
      performDraftSave();
    };

    document.addEventListener('visibilitychange', handleVisibility);
    window.addEventListener('blur', handleBlur);

    return () => {
      clearInterval(interval);
      document.removeEventListener('visibilitychange', handleVisibility);
      window.removeEventListener('blur', handleBlur);
    };
  }, [isOpen, hasChanges, draftKey, attId, title, data]);

  // 저장 엔진 선택: 'exceljs' (서식 보존 권장) vs 'sheetjs' (기존 경량 모드)
  const [engine, setEngine] = useState<'exceljs' | 'sheetjs'>(initialEngine);

  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<any>(null);
  const univerAPIRef = useRef<any>(null);
  const currentDataRef = useRef<string>(data);
  const structuralOpsRef = useRef<StructuralOperation[]>([]);

  useEffect(() => {
    currentDataRef.current = data;
  }, [data]);

  // 원본 문서의 미보존 고급 요소(차트, 피벗 등) 사전 검사
  useEffect(() => {
    if (!isOpen || !data) return;
    let isCancelled = false;
    detectExcelUnsupportedFeatures(data).then((report) => {
      if (!isCancelled) {
        setLossReport(report);
      }
    });
    return () => {
      isCancelled = true;
    };
  }, [isOpen, data]);

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
    const timer = setTimeout(async () => {
      if (isDisposed || !containerRef.current) return;

      try {
        // 기존 인스턴스 정리
        if (univerRef.current) {
          try { univerRef.current.dispose(); } catch {}
          univerRef.current = null;
          univerAPIRef.current = null;
        }

        const workbookData = await dataUrlToUniverWorkbook(data, title);
        if (isDisposed || !containerRef.current) return;
        if ((workbookData as any).isFallbackMode) {
          setIsFallbackMode(true);
        }

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

        if (typeof window !== 'undefined') {
          (window as any).__excelUniverAPI = univerAPI;
          (window as any).__excelModalExecuteSave = () => executeSave();
          (window as any).__excelIsReady = () => Boolean(univerAPIRef.current);
        }

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
        }, 100);

        const commandService = (univerAPI as any)._commandService;
        let disposable: any = null;
        if (commandService && typeof commandService.onCommandExecuted === 'function') {
          disposable = commandService.onCommandExecuted((cmd: any) => {
            if (isDisposed) return;
            if (!isInitial) {
              setHasChanges(true);
            }

            // 행/열 삽입 및 삭제 mutation 감지하여 In-Place 패치용 큐에 기록 (mutation ID 정확 매칭)
            const id = cmd?.id || '';
            const range = cmd?.params?.range;
            const subUnitId = cmd.params?.subUnitId;

            if (range) {
              let op: StructuralOperation | null = null;
              if (id === 'sheet.mutation.insert-row') {
                op = {
                  type: 'insertRow',
                  subUnitId,
                  start: range.startRow + 1,
                  count: range.endRow - range.startRow + 1,
                };
              } else if (id === 'sheet.mutation.remove-rows') {
                op = {
                  type: 'removeRow',
                  subUnitId,
                  start: range.startRow + 1,
                  count: range.endRow - range.startRow + 1,
                };
              } else if (id === 'sheet.mutation.insert-col') {
                op = {
                  type: 'insertCol',
                  subUnitId,
                  start: range.startColumn + 1,
                  count: range.endColumn - range.startColumn + 1,
                };
              } else if (id === 'sheet.mutation.remove-col') {
                op = {
                  type: 'removeCol',
                  subUnitId,
                  start: range.startColumn + 1,
                  count: range.endColumn - range.startColumn + 1,
                };
              }

              if (op) {
                setHasChanges(true);
                // 실행 취소(Undo) / 다시 실행(Redo) 시 반대 연산 상쇄 처리
                  const ops = structuralOpsRef.current;
                  if (ops.length > 0) {
                    const lastOp = ops[ops.length - 1];
                    const isSameSheet = !lastOp.subUnitId || !op.subUnitId || lastOp.subUnitId === op.subUnitId;
                    const isOppositeRow =
                      isSameSheet &&
                      ((lastOp.type === 'insertRow' && op.type === 'removeRow') ||
                       (lastOp.type === 'removeRow' && op.type === 'insertRow')) &&
                      lastOp.start === op.start &&
                      lastOp.count === op.count;
                    const isOppositeCol =
                      isSameSheet &&
                      ((lastOp.type === 'insertCol' && op.type === 'removeCol') ||
                       (lastOp.type === 'removeCol' && op.type === 'insertCol')) &&
                      lastOp.start === op.start &&
                      lastOp.count === op.count;

                    if (isOppositeRow || isOppositeCol) {
                      ops.pop();
                      return;
                    }
                  }
                  ops.push(op);
                }
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

  // 실제 저장 실행 함수
  const executeSave = async () => {
    if (!univerAPIRef.current || isSaving) return;

    try {
      setIsSaving(true);
      const activeWb = univerAPIRef.current.getActiveWorkbook();
      if (!activeWb) throw new Error('활성 워크북을 찾을 수 없습니다.');

      const snapshot = activeWb.save();
      // 원본 데이터(currentDataRef.current) 및 구조 변경 연산(structuralOpsRef.current)을 전달하여 in-place 패치
      const ops = [...structuralOpsRef.current];
      let updatedDataUrl = '';
      try {
        updatedDataUrl = await univerSnapshotToDataUrl(snapshot, format, engine, currentDataRef.current, ops);
      } catch (err: any) {
        if (err?.message === 'ORIGINAL_WORKBOOK_LOAD_FAILED') {
          const proceed = window.confirm(
            '⚠️ 원본 엑셀 문서의 내부 구조를 로드하는 중 오류가 발생했습니다.\n' +
            '이대로 저장하면 원본의 고급 서식(드롭다운 목록, 조건부 서식, 메모, 틀 고정, 인쇄 설정 등)이 누락될 수 있습니다.\n\n' +
            '[확인]을 누르면 기본 서식으로 계속 저장하고, [취소]를 누르면 저장을 중단합니다.'
          );
          if (!proceed) {
            setIsSaving(false);
            return;
          }
          // 그래도 저장을 선택한 경우 originalDataOrUrl 없이 새 워크북으로 저장
          updatedDataUrl = await univerSnapshotToDataUrl(snapshot, format, engine, undefined, ops);
        } else {
          throw err;
        }
      }

      currentDataRef.current = updatedDataUrl;
      structuralOpsRef.current = [];

      invalidateExcelCache(attId);

      // 5단계 임시본 정리
      await deleteOfficeDraft(draftKey);

      // FileHandle 연동 (PWA 탐색기 열기 파일 직접 저장)
      if (fileContextRef.current) {
        const base64 = data.split(',')[1] || '';
        const binary = atob(base64);
        const origBytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) origBytes[i] = binary.charCodeAt(i);
        await ensureOriginalBackup(fileContextRef.current, origBytes);

        if (await checkFileHandleConflict(fileContextRef.current)) {
          setShowFileConflictModal(true);
          setIsSaving(false);
          return;
        }

        const updatedB64 = updatedDataUrl.split(',')[1] || '';
        const updatedBin = atob(updatedB64);
        const saveBytes = new Uint8Array(updatedBin.length);
        for (let i = 0; i < updatedBin.length; i++) saveBytes[i] = updatedBin.charCodeAt(i);

        await saveToFileHandle(fileContextRef.current, saveBytes);
      }

      if (onSave) {
        await onSave(updatedDataUrl, engine);
      }

      setHasChanges(false);
      setSaveToast(`메모에 엑셀이 저장되었습니다 (${engine === 'exceljs' ? 'ExcelJS 서식보존' : 'SheetJS 경량'})`);
      setTimeout(() => setSaveToast(null), 3000);
    } catch (err) {
      console.error('엑셀 저장 실패:', err);
      setErrorToast('저장에 실패했습니다.');
      setTimeout(() => setErrorToast(null), 3000);
    } finally {
      setIsSaving(false);
    }
  };

  // 저장 버튼 클릭 (미보존 요소 감지 시 경고 다이얼로그 띄움)
  const handleSave = async () => {
    if (lossReport && lossReport.unsupportedList.length > 0) {
      setShowLossWarningModal(true);
      return;
    }
    await executeSave();
  };

  // 다운로드 처리
  const handleDownload = async () => {
    if (!univerAPIRef.current) return;
    try {
      const activeWb = univerAPIRef.current.getActiveWorkbook();
      if (!activeWb) return;

      const snapshot = activeWb.save();
      const ops = [...structuralOpsRef.current];
      const dataUrl = await univerSnapshotToDataUrl(snapshot, format, engine, currentDataRef.current, ops);

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

  // 다른 이름으로 저장 (showSaveFilePicker 연동)
  const handleSaveAs = async () => {
    if (!univerAPIRef.current) return;
    try {
      const activeWb = univerAPIRef.current.getActiveWorkbook();
      if (!activeWb) return;

      const snapshot = activeWb.save();
      const ops = [...structuralOpsRef.current];
      const dataUrl = await univerSnapshotToDataUrl(snapshot, format, engine, currentDataRef.current, ops);
      const base64 = dataUrl.split(',')[1] || '';
      const binary = atob(base64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

      const mime = format === 'csv' ? 'text/csv' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      await saveFileAs(bytes, title, mime, format);
    } catch (err) {
      console.error('다른 이름으로 저장 실패:', err);
    }
  };

  // 5단계 임시본 복구 적용 및 폐기
  const handleApplyDraftRecovery = () => {
    if (!draftToRecover || !univerAPIRef.current) return;
    try {
      univerAPIRef.current.createWorkbook({ id: 'workbook', ...draftToRecover.data });
      setHasChanges(true);
      setShowDraftRecoveryModal(false);
      setSaveToast('임시 복구본이 성공적으로 복구되었습니다.');
      setTimeout(() => setSaveToast(null), 3000);
    } catch (err) {
      console.error('임시본 복구 실패:', err);
    }
  };

  const handleDiscardDraftRecovery = async () => {
    setShowDraftRecoveryModal(false);
    setDraftToRecover(null);
    await deleteOfficeDraft(draftKey);
  };

  // 닫기 요청
  const handleRequestClose = () => {
    if (hasChanges) {
      setShowExitConfirm(true);
    } else {
      onClose();
    }
  };

  // 5단계 공통 단축키 훅 (Ctrl+S 통일, Alt+Enter 전체화면, Esc 셀편집 보호)
  useOfficeShortcuts({
    enabled: isOpen,
    onSave: handleSave,
    onClose: handleRequestClose,
    hasChanges,
    onConfirmCloseNeeded: () => setShowExitConfirm(true),
    isEditingCellOrDialog: () => {
      const active = document.activeElement;
      if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.getAttribute('contenteditable') === 'true')) {
        if (containerRef.current?.contains(active)) return true;
      }
      if (document.querySelector('.univer-cell-editor, .univer-dialog, .univer-modal, [class*="cell-editor"], [class*="find-replace"]')) {
        return true;
      }
      return false;
    },
    toggleFullscreen: floating.toggleFullscreen,
  });

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

  // 원본 복원 실행
  const handleConfirmRestore = async () => {
    setShowRestoreConfirm(false);
    if (onRestoreOriginal) {
      await onRestoreOriginal();
    }
  };

  if (!isOpen) return null;

  return (
    <div className={floating.modalWrapperClass}>
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
                  <span className="text-[11px] text-amber-400 flex items-center gap-1 font-medium shrink-0">
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse"></span>
                    수정됨
                  </span>
                )}
                {isFallbackMode && (
                  <span className="text-[10px] px-2 py-0.5 rounded bg-amber-500/20 text-amber-300 border border-amber-500/40 font-semibold shrink-0" title="ExcelJS 대신 기본 데이터 모드로 로드되었습니다.">
                    ⚠️ 호환 모드: 서식 일부가 표시되지 않습니다
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

            {/* 원본으로 되돌리기 버튼 (원본 백업이 존재하는 경우) */}
            {originalData && onRestoreOriginal && (
              <button
                type="button"
                onClick={() => setShowRestoreConfirm(true)}
                className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-amber-500/10 hover:bg-amber-500/20 text-amber-300 border border-amber-500/30 text-xs font-semibold transition-colors cursor-pointer"
                title={`웹 수정 이전 최초 원본 버전으로 복원합니다 (${originalDate ? new Date(originalDate).toLocaleString() : '최초 버전'})`}
              >
                <RotateCcw size={13} />
                <span className="hidden md:inline">원본으로 되돌리기</span>
              </button>
            )}

            {/* 저장 엔진 선택 (ExcelJS vs SheetJS) */}
            <div className="hidden lg:flex items-center gap-1 bg-slate-950 border border-slate-800 rounded-lg p-0.5 text-[11px]">
              <button
                type="button"
                onClick={() => setEngine('exceljs')}
                className={`px-2 py-0.5 rounded font-semibold transition-all cursor-pointer ${
                  engine === 'exceljs'
                    ? 'bg-emerald-600 text-white shadow-xs'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
                title="ExcelJS: 글꼴, 배경색, 테두리, 열너비, 행높이, 수식, 병합을 보존하는 고품질 엔진"
              >
                ExcelJS (서식보존)
              </button>
              <button
                type="button"
                onClick={() => setEngine('sheetjs')}
                className={`px-2 py-0.5 rounded font-semibold transition-all cursor-pointer ${
                  engine === 'sheetjs'
                    ? 'bg-slate-700 text-white shadow-xs'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
                title="SheetJS: 순수 표 데이터 중심의 기존 경량 엔진"
              >
                SheetJS (경량)
              </button>
            </div>

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
              title="편집한 내용을 메모에 즉시 저장합니다 (단축키: Ctrl+S)"
            >
              {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
              <span>{isSaving ? '저장 중...' : '저장 및 반영'}</span>
            </button>

            {/* 다른 이름으로 저장 버튼 */}
            <button
              type="button"
              onClick={handleSaveAs}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-medium border border-slate-700 transition-colors cursor-pointer"
              title="다른 이름으로 저장"
            >
              <FolderDown size={14} />
              <span className="hidden sm:inline">다른 이름으로 저장</span>
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

            {/* 전체 화면 (Alt+Enter) 토글 버튼 */}
            {floating.toggleFullscreenButton}

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

      {/* 미보존 요소(차트, 피벗 등) 손실 경고 다이얼로그 */}
      {showLossWarningModal && lossReport && (
        <div className="fixed inset-0 z-70 bg-black/75 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-amber-500/40 rounded-2xl shadow-2xl max-w-md w-full p-5 space-y-4 animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-start gap-3">
              <div className="p-2.5 rounded-xl bg-amber-500/10 text-amber-400 border border-amber-500/30 shrink-0">
                <AlertTriangle size={24} />
              </div>
              <div className="space-y-1.5">
                <h3 className="text-sm font-bold text-slate-100">고급 서식 보존 주의 안내</h3>
                <p className="text-xs text-slate-300 leading-relaxed">
                  현재 열려 있는 엑셀 문서에 웹 편집기(ExcelJS)에서 완벽히 보존되지 않는 다음 요소가 포함되어 있습니다:
                </p>
                <div className="bg-slate-950 border border-slate-800 rounded-lg p-2.5 mt-2 space-y-1">
                  {lossReport.unsupportedList.map((item, idx) => (
                    <div key={idx} className="flex items-center gap-1.5 text-xs text-amber-300 font-medium">
                      <span>•</span>
                      <span>{item}</span>
                    </div>
                  ))}
                </div>
                <p className="text-[11px] text-slate-400 mt-2 leading-relaxed">
                  * 셀 데이터, 수식, 병합, 폰트, 배경색, 열 너비는 정상 보존되나, 위 고급 요소는 저장 시 누락될 수 있습니다. (첫 저장 시 원본 파일은 자동으로 백업 보관됩니다)
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-800">
              <button
                type="button"
                onClick={() => setShowLossWarningModal(false)}
                className="px-3.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium transition-colors cursor-pointer"
              >
                취소 (저장 안 함)
              </button>
              <button
                type="button"
                onClick={async () => {
                  setShowLossWarningModal(false);
                  await executeSave();
                }}
                className="px-3.5 py-1.5 rounded-lg bg-amber-600 hover:bg-amber-500 text-white text-xs font-semibold shadow-xs transition-colors cursor-pointer"
              >
                그래도 저장
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 원본 복원 확인 모달 */}
      {showRestoreConfirm && (
        <div className="fixed inset-0 z-70 bg-black/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700/80 rounded-2xl shadow-2xl max-w-sm w-full p-5 space-y-4 animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-start gap-3">
              <div className="p-2 rounded-xl bg-amber-500/10 text-amber-400 border border-amber-500/20 shrink-0">
                <RotateCcw size={22} />
              </div>
              <div>
                <h3 className="text-sm font-bold text-slate-100">최초 원본 버전으로 되돌리시겠습니까?</h3>
                <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                  웹에서 수정한 모든 내용이 취소되고, 처음 업로드했던 원본 엑셀 문서 상태로 완전히 복원됩니다.
                </p>
                {originalDate && (
                  <p className="text-[11px] text-slate-500 mt-1">
                    (백업 시점: {new Date(originalDate).toLocaleString()})
                  </p>
                )}
              </div>
            </div>
            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-800">
              <button
                type="button"
                onClick={() => setShowRestoreConfirm(false)}
                className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium transition-colors cursor-pointer"
              >
                취소
              </button>
              <button
                type="button"
                onClick={handleConfirmRestore}
                className="px-3.5 py-1.5 rounded-lg bg-amber-600 hover:bg-amber-500 text-white text-xs font-semibold shadow-xs transition-colors cursor-pointer"
              >
                원본으로 되돌리기
              </button>
            </div>
          </div>
        </div>
      )}

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
                className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium transition-colors cursor-pointer"
              >
                취소
              </button>
              <button
                type="button"
                onClick={handleDiscardAndExit}
                className="px-3 py-1.5 rounded-lg bg-rose-900/40 hover:bg-rose-900/60 text-rose-300 border border-rose-700/40 text-xs font-medium transition-colors cursor-pointer"
              >
                저장하지 않고 닫기
              </button>
              <button
                type="button"
                onClick={handleSaveAndExit}
                className="px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold shadow-xs transition-colors cursor-pointer"
              >
                저장하고 닫기
              </button>
            </div>
          </div>
        </div>
      )}
      {/* 5단계 임시 복구본 복구 모달 */}
      {showDraftRecoveryModal && draftToRecover && (
        <div className="fixed inset-0 z-100 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
          <div className="bg-slate-900 border border-emerald-500/30 rounded-2xl p-6 max-w-md w-full shadow-2xl space-y-4">
            <div className="flex items-center gap-3 text-emerald-400">
              <div className="p-2.5 rounded-xl bg-emerald-500/10 border border-emerald-500/20">
                <RotateCcw size={24} />
              </div>
              <div>
                <h3 className="font-bold text-base text-slate-100">임시 복구본 복구</h3>
                <p className="text-xs text-emerald-300/80">
                  {new Date(draftToRecover.updatedAt).toLocaleString()}에 자동 저장된 작업 내용이 있습니다.
                </p>
              </div>
            </div>
            <p className="text-xs text-slate-300 leading-relaxed bg-slate-800/60 p-3 rounded-xl border border-slate-700/50">
              이전에 저장되지 않은 엑셀 작업 내용(임시본)이 기기 내부에 보관되어 있습니다. 복구하여 작업을 계속하시겠습니까?
            </p>
            <div className="flex flex-col sm:flex-row gap-2 pt-1">
              <button
                type="button"
                id="office-confirm-recover-btn"
                onClick={handleApplyDraftRecovery}
                className="flex-1 py-2.5 px-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold transition-all shadow-md active:scale-95 flex items-center justify-center gap-1.5 cursor-pointer"
              >
                복구 적용
              </button>
              <button
                type="button"
                id="office-discard-recover-btn"
                onClick={handleDiscardDraftRecovery}
                className="py-2.5 px-3 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-medium transition-all cursor-pointer"
              >
                복구 안 함 (삭제)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 5단계 파일 충돌 모달 (디스크 파일 변경 감지) */}
      {showFileConflictModal && (
        <div className="fixed inset-0 z-100 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
          <div className="bg-slate-900 border border-rose-500/30 rounded-2xl p-6 max-w-md w-full shadow-2xl space-y-4">
            <div className="flex items-center gap-3 text-rose-400">
              <div className="p-2.5 rounded-xl bg-rose-500/10 border border-rose-500/20">
                <AlertCircle size={24} />
              </div>
              <div>
                <h3 className="font-bold text-base text-slate-100">파일 충돌 경고</h3>
                <p className="text-xs text-rose-300/80">다른 프로그램에서 파일이 변경되었습니다.</p>
              </div>
            </div>
            <p className="text-xs text-slate-300 leading-relaxed bg-slate-800/60 p-3 rounded-xl border border-slate-700/50">
              문서를 연 이후 외부 프로그램에서 파일 내용이나 수정 일시가 바뀌었습니다. 원래 파일에 그대로 덮어쓰시겠습니까, 아니면 다른 이름으로 저장하시겠습니까?
            </p>
            <div className="flex flex-col gap-2 pt-1">
              <button
                type="button"
                onClick={async () => {
                  setShowFileConflictModal(false);
                  if (fileContextRef.current && univerAPIRef.current) {
                    const activeWb = univerAPIRef.current.getActiveWorkbook();
                    const snapshot = activeWb.save();
                    const ops = [...structuralOpsRef.current];
                    const updatedDataUrl = await univerSnapshotToDataUrl(snapshot, format, engine, currentDataRef.current, ops);
                    const updatedB64 = updatedDataUrl.split(',')[1] || '';
                    const updatedBin = atob(updatedB64);
                    const saveBytes = new Uint8Array(updatedBin.length);
                    for (let i = 0; i < updatedBin.length; i++) saveBytes[i] = updatedBin.charCodeAt(i);
                    await saveToFileHandle(fileContextRef.current, saveBytes);
                    setHasChanges(false);
                    setSaveToast('원래 파일에 성공적으로 덮어썼습니다.');
                    setTimeout(() => setSaveToast(null), 3000);
                  }
                }}
                className="py-2.5 px-3 bg-rose-600 hover:bg-rose-500 text-white rounded-xl text-xs font-bold transition-all cursor-pointer"
              >
                원래 파일에 그대로 덮어쓰기
              </button>
              <button
                type="button"
                onClick={() => {
                  setShowFileConflictModal(false);
                  handleSaveAs();
                }}
                className="py-2.5 px-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold transition-all cursor-pointer"
              >
                다른 이름으로 저장
              </button>
              <button
                type="button"
                onClick={() => setShowFileConflictModal(false)}
                className="py-2 px-3 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-medium transition-all cursor-pointer"
              >
                취소
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
