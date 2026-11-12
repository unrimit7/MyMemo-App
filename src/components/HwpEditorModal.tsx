import React, { useState, useEffect, useRef } from 'react';
import {
  X,
  Save,
  Download,
  AlertCircle,
  CheckCircle,
  Loader2,
  FolderDown,
  RotateCcw,
  FileText,
} from 'lucide-react';
import { createEditor, type RhwpEditor } from '@rhwp/editor';
import { useFloatingWindow } from './useFloatingWindow';
import { postProcessHwpx, postProcessHwp, hasHwpTrackChanges } from '../utils/hwpxAdapter';
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

interface HwpEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  originalData?: string;
  originalDate?: string;
  fileHandle?: any;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
  onRestoreOriginal?: () => Promise<void> | void;
}

// Data URL 또는 Base64 문자열을 안전하게 Uint8Array로 변환
function dataUrlToUint8Array(dataOrUrl: string): Uint8Array {
  const commaIdx = dataOrUrl.indexOf(',');
  const base64 = commaIdx !== -1 ? dataOrUrl.slice(commaIdx + 1) : dataOrUrl;
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

// Uint8Array 바이트를 Base64 Data URL로 안전하게 비동기 변환 (메모리 낭비 및 콜스택 초과 방지)
function uint8ArrayToDataUrl(bytes: Uint8Array, mimeType: string = 'application/x-hwp'): Promise<string> {
  return new Promise((resolve, reject) => {
    const blob = new Blob([bytes as any], { type: mimeType });
    const reader = new FileReader();
    reader.onloadend = () => {
      resolve(reader.result as string);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

export const HwpEditorModal: React.FC<HwpEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  attId,
  originalData,
  originalDate,
  fileHandle,
  onSave,
  onRestoreOriginal,
}) => {
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [saveSuccess, setSaveSuccess] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [showExitConfirm, setShowExitConfirm] = useState<boolean>(false);
  const [showRestoreConfirm, setShowRestoreConfirm] = useState<boolean>(false);
  const [showTrackChangesWarning, setShowTrackChangesWarning] = useState<boolean>(false);
  const [trackChangesWarningAcknowledged, setTrackChangesWarningAcknowledged] = useState<boolean>(false);
  const [pendingSaveAction, setPendingSaveAction] = useState<'save' | 'download' | null>(null);

  // 5단계 임시본 복구 모달 및 파일 핸들러 충돌 모달 상태
  const draftKey = `hwp_${attId || ''}_${title}`;
  const [draftToRecover, setDraftToRecover] = useState<OfficeDraft | null>(null);
  const [showDraftRecoveryModal, setShowDraftRecoveryModal] = useState<boolean>(false);
  const [showFileConflictModal, setShowFileConflictModal] = useState<boolean>(false);
  const fileContextRef = useRef<FileHandleContext | null>(null);

  const containerRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<RhwpEditor | null>(null);
  const floating = useFloatingWindow({ id: 'hwp-editor', defaultWidth: 920, defaultHeight: 820 });

  const isHwpx = title.toLowerCase().endsWith('.hwpx');
  const mimeType = isHwpx ? 'application/vnd.hancom.hwpx' : 'application/x-hwp';

  // FileHandle 초기화
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

  // 30초 주기 + visibilitychange / blur 자동 저장
  useEffect(() => {
    if (!isOpen || !hasChanges || !editorRef.current) return;

    let autoSaveIntervalMs = 30000;

    const performDraftSave = async () => {
      if (!hasChanges || !editorRef.current) return;
      try {
        const t0 = performance.now();
        const bytes = isHwpx ? await editorRef.current.exportHwpx() : await editorRef.current.exportHwp();
        const duration = performance.now() - t0;
        const draftDataUrl = await uint8ArrayToDataUrl(bytes, mimeType);
        const baseHash = await computeDataHash(data);

        await saveOfficeDraft({
          draftKey,
          attId,
          title,
          fileType: isHwpx ? 'hwpx' : 'hwp',
          baseOriginalHash: baseHash,
          data: draftDataUrl,
          dataType: 'data_url',
          updatedAt: Date.now(),
          byteLength: bytes.length,
          saveDurationMs: Math.round(duration),
        });

        if (duration >= 1000) {
          autoSaveIntervalMs = 60000;
        }
      } catch (err) {
        console.warn('[Hwp] 임시본 자동 저장 오류:', err);
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
  }, [isOpen, hasChanges, draftKey, attId, title, data, isHwpx, mimeType]);

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

  // rhwp 에디터 내부의 dirty(미저장 수정 상태)를 실시간 감지하여 상단 배지 및 닫기 방지 연동
  useEffect(() => {
    if (!isOpen || isLoading) return;

    let isMounted = true;
    const checkDirtyInterval = setInterval(async () => {
      if (!editorRef.current || !isMounted) return;
      try {
        const state = await editorRef.current.getDocumentState();
        if (state && typeof state.dirty === 'boolean') {
          setHasChanges(state.dirty);
        }
      } catch {
        // rhwp 환경에 따라 document-state-v1 미지원 시 무시
      }
    }, 1000);

    return () => {
      isMounted = false;
      clearInterval(checkDirtyInterval);
    };
  }, [isOpen, isLoading]);

  useEffect(() => {
    if (!isOpen || !data || !containerRef.current) return;

    let isMounted = true;
    setIsLoading(true);
    setErrorMessage(null);
    setSaveSuccess(false);
    setHasChanges(false);

    const initEditor = async () => {
      try {
        if (!containerRef.current) return;

        // 기존 에디터 정리
        if (editorRef.current) {
          try {
            editorRef.current.destroy();
          } catch {
            // ignore
          }
          editorRef.current = null;
        }

        // rhwp 웹 에디터 임베드 인스턴스 생성 (자체 호스팅 오프라인 지원)
        const studioUrl = new URL('rhwp/index.html', document.baseURI).href;
        const editor = await createEditor(containerRef.current, {
          studioUrl,
          width: '100%',
          height: '100%',
          requestTimeoutMs: 30000,
        });

        if (!isMounted) {
          editor.destroy();
          return;
        }

        editorRef.current = editor;

        // E2E 및 테스트용 전역 참조 등록
        (window as any).__currentHwpEditor = editor;

        // 문서 변경 감지
        if (typeof editor.onDocumentChanged === 'function') {
          editor.onDocumentChanged(() => {
            if (isMounted) setHasChanges(true);
          });
        }

        // 문서 바이트 로드
        const bytes = dataUrlToUint8Array(data);
        await editor.loadFile(bytes, title, { suppressDialogs: true, skipUnsavedGuard: true });

        if (isMounted) {
          setIsLoading(false);
          (window as any).__hwpIsReady = true;
        }
      } catch (err: any) {
        console.error('HWP 에디터 로드 실패:', err);
        if (isMounted) {
          setErrorMessage(
            err.message?.includes('timeout')
              ? '에디터 초기화 시간이 초과되었습니다. 네트워크 상태를 확인해 주세요.'
              : `한글 에디터를 불러오지 못했습니다: ${err.message || '네트워크 오류'}`
          );
          setIsLoading(false);
          (window as any).__hwpIsReady = false;
        }
      }
    };

    initEditor();

    return () => {
      isMounted = false;
      (window as any).__currentHwpEditor = null;
      (window as any).__hwpIsReady = false;
      if (editorRef.current) {
        try {
          editorRef.current.destroy();
        } catch {
          // ignore
        }
        editorRef.current = null;
      }
    };
  }, [isOpen, data, title]);

  if (!isOpen) return null;

  const checkIsDirty = async (): Promise<boolean> => {
    if (hasChanges) return true;
    if (editorRef.current) {
      try {
        const state = await editorRef.current.getDocumentState();
        if (state && state.dirty) {
          setHasChanges(true);
          return true;
        }
      } catch {
        // ignore
      }
    }
    return false;
  };

  // 수정된 문서 저장 및 메모 본문에 반영
  const handleSave = async (options?: { bypassTrackWarning?: boolean }) => {
    if (!editorRef.current || isSaving) return;

    setIsSaving(true);
    setErrorMessage(null);
    try {
      const editor = editorRef.current;
      const originalBytes = dataUrlToUint8Array(data);
      const isDirty = await checkIsDirty();

      let bytes: Uint8Array;
      if (!isDirty) {
        // 1-a: 사용자가 편집하지 않았으면 재직렬화 없이 원본 바이트 그대로 보존
        bytes = originalBytes;
      } else {
        // 1-b: 문서에 변경 추적이 있는 경우 경고 모달 표시
        if (!options?.bypassTrackWarning && !trackChangesWarningAcknowledged) {
          const trackInfo = await hasHwpTrackChanges(originalBytes, isHwpx);
          if (trackInfo.hasTrackChanges) {
            setPendingSaveAction('save');
            setShowTrackChangesWarning(true);
            setIsSaving(false);
            return;
          }
        }

        bytes = isHwpx ? await editor.exportHwpx() : await editor.exportHwp();
        if (isHwpx) {
          bytes = await postProcessHwpx(bytes, originalBytes);
        } else {
          bytes = postProcessHwp(bytes);
        }
      }

      const updatedDataUrl = await uint8ArrayToDataUrl(bytes, mimeType);

      // 5단계 임시본 정리
      await deleteOfficeDraft(draftKey);

      // FileHandle 연동 (PWA 파일 연결)
      if (fileContextRef.current) {
        const origBytes = dataUrlToUint8Array(data);
        await ensureOriginalBackup(fileContextRef.current, origBytes);

        if (await checkFileHandleConflict(fileContextRef.current)) {
          setShowFileConflictModal(true);
          setIsSaving(false);
          return;
        }

        await saveToFileHandle(fileContextRef.current, bytes);
      }

      if (onSave) {
        await onSave(updatedDataUrl);
      }

      try {
        await editor.notifySaved(title);
      } catch {
        // notifySaved가 지원되지 않는 구버전 스튜디오의 경우 무시
      }

      setHasChanges(false);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (err: any) {
      console.error('HWP 저장 실패:', err);
      setErrorMessage(`저장 실패: ${err.message || '문서를 내보낼 수 없습니다.'}`);
    } finally {
      setIsSaving(false);
    }
  };

  // E2E 테스트용 저장 핸들러 노출
  useEffect(() => {
    (window as any).__hwpEditorExecuteSave = (options?: { bypassTrackWarning?: boolean }) => handleSave(options);
    return () => {
      delete (window as any).__hwpEditorExecuteSave;
    };
  }, [handleSave]);

  // 수정된 원본 파일 내 PC에 바로 다운로드
  const handleDownload = async (options?: { bypassTrackWarning?: boolean }) => {
    if (!editorRef.current) return;

    try {
      const editor = editorRef.current;
      const originalBytes = dataUrlToUint8Array(data);
      const isDirty = await checkIsDirty();

      let bytes: Uint8Array;
      if (!isDirty) {
        bytes = originalBytes;
      } else {
        if (!options?.bypassTrackWarning && !trackChangesWarningAcknowledged) {
          const trackInfo = await hasHwpTrackChanges(originalBytes, isHwpx);
          if (trackInfo.hasTrackChanges) {
            setPendingSaveAction('download');
            setShowTrackChangesWarning(true);
            return;
          }
        }

        bytes = isHwpx ? await editor.exportHwpx() : await editor.exportHwp();
        if (isHwpx) {
          bytes = await postProcessHwpx(bytes, originalBytes);
        } else {
          bytes = postProcessHwp(bytes);
        }
      }

      const blob = new Blob([bytes as any], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = title;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: any) {
      console.error('HWP 다운로드 실패:', err);
      alert('파일 다운로드 실패: ' + err.message);
    }
  };

  // 다른 이름으로 저장 (showSaveFilePicker 연동)
  const handleSaveAs = async (options?: { bypassTrackWarning?: boolean }) => {
    if (!editorRef.current) return;
    try {
      const editor = editorRef.current;
      const originalBytes = dataUrlToUint8Array(data);
      const isDirty = await checkIsDirty();

      let bytes: Uint8Array;
      if (!isDirty) {
        bytes = originalBytes;
      } else {
        if (!options?.bypassTrackWarning && !trackChangesWarningAcknowledged) {
          const trackInfo = await hasHwpTrackChanges(originalBytes, isHwpx);
          if (trackInfo.hasTrackChanges) {
            setPendingSaveAction('saveAs' as any);
            setShowTrackChangesWarning(true);
            return;
          }
        }

        bytes = isHwpx ? await editor.exportHwpx() : await editor.exportHwp();
        if (isHwpx) {
          bytes = await postProcessHwpx(bytes, originalBytes);
        } else {
          bytes = postProcessHwp(bytes);
        }
      }

      await saveFileAs(bytes, title, mimeType, isHwpx ? 'hwpx' : 'hwp');
    } catch (err: any) {
      console.error('HWP 다른 이름으로 저장 실패:', err);
    }
  };

  // 5단계 임시본 복구 적용 및 폐기
  const handleApplyDraftRecovery = async () => {
    if (!draftToRecover || !editorRef.current) return;
    try {
      const bytes = dataUrlToUint8Array(draftToRecover.data);
      await editorRef.current.loadFile(bytes, title, { suppressDialogs: true, skipUnsavedGuard: true });
      setHasChanges(true);
      setShowDraftRecoveryModal(false);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (err) {
      console.error('HWP 임시본 복구 실패:', err);
    }
  };

  const handleDiscardDraftRecovery = async () => {
    setShowDraftRecoveryModal(false);
    setDraftToRecover(null);
    await deleteOfficeDraft(draftKey);
  };

  const handleConfirmRestore = async () => {
    setShowRestoreConfirm(false);
    if (onRestoreOriginal) {
      await onRestoreOriginal();
    }
  };

  // 조건 4: rhwp iframe 내부 Ctrl+S 감지 및 location.origin + iframe.contentWindow 검증 postMessage 연동
  useEffect(() => {
    if (!isOpen || isLoading || !containerRef.current) return;

    const iframe = containerRef.current.querySelector('iframe');
    if (!iframe || !iframe.contentWindow) return;

    const targetOrigin = window.location.origin;

    const handleIframeKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
        e.preventDefault();
        e.stopPropagation();
        iframe.contentWindow?.parent.postMessage({ type: 'MYMEMO_OFFICE_SAVE' }, targetOrigin);
      }
    };

    try {
      iframe.contentWindow.addEventListener('keydown', handleIframeKeyDown);
    } catch (err) {
      console.warn('rhwp iframe keydown 리스너 등록 실패:', err);
    }

    const handleMessage = (e: MessageEvent) => {
      // 조건 4: postMessage 대상 및 발신지 origin & source 엄격 검증
      if (e.origin !== window.location.origin) return;
      if (e.source !== iframe.contentWindow) return;
      if (e.data?.type === 'MYMEMO_OFFICE_SAVE') {
        handleSave();
      }
    };

    window.addEventListener('message', handleMessage);

    return () => {
      try {
        iframe.contentWindow?.removeEventListener('keydown', handleIframeKeyDown);
      } catch {}
      window.removeEventListener('message', handleMessage);
    };
  }, [isOpen, isLoading]);

  const handleConfirmSaveAnyway = async () => {
    setShowTrackChangesWarning(false);
    setTrackChangesWarningAcknowledged(true);
    if (pendingSaveAction === 'download') {
      setPendingSaveAction(null);
      await handleDownload({ bypassTrackWarning: true });
    } else {
      setPendingSaveAction(null);
      await handleSave({ bypassTrackWarning: true });
    }
  };

  const handleCancelTrackChangesWarning = () => {
    setShowTrackChangesWarning(false);
    setPendingSaveAction(null);
  };

  const handleClose = async () => {
    const isDirty = await checkIsDirty();
    if (isDirty) {
      setShowExitConfirm(true);
      return;
    }
    onClose();
  };

  // 5단계 공통 단축키 훅 (Ctrl+S 단일 실행, Alt+Enter 전체화면, Esc 대화상자 보호)
  useOfficeShortcuts({
    enabled: isOpen,
    onSave: () => handleSave(),
    onClose: handleClose,
    hasChanges,
    onConfirmCloseNeeded: () => setShowExitConfirm(true),
    isEditingCellOrDialog: () => {
      if (document.querySelector('[class*="dialog"], [class*="modal"]')) {
        return true;
      }
      return false;
    },
    toggleFullscreen: floating.toggleFullscreen,
  });

  const handleSaveAndExit = async () => {
    await handleSave();
    setShowExitConfirm(false);
    onClose();
  };

  const handleDiscardAndExit = async () => {
    try {
      if (editorRef.current) {
        // notifySaved()를 호출하면 rhwp 스튜디오 내부의 자동 복구 임시본(draft)이 완전히 삭제되어 다음 번 실행 시 '문서 복구' 안내창이 뜨지 않습니다.
        await editorRef.current.notifySaved(title);
      }
    } catch (err) {
      console.warn('rhwp notifySaved 호출 중 오류(무시 가능):', err);
    }
    setHasChanges(false);
    setShowExitConfirm(false);
    onClose();
  };

  return (
    <div
      className={floating.modalWrapperClass}
      role="dialog"
      aria-modal="true"
    >
      <div {...floating.windowContainerProps}>
        {/* 상단 헤더 바 */}
        <div 
          {...floating.headerProps}
          className="flex items-center justify-between gap-2 px-3 sm:px-5 py-2.5 bg-slate-900 border-b border-slate-800 text-slate-200 shrink-0 shadow-md"
        >
        <div className="flex items-center gap-2 min-w-0">
          <span className="text-lg">✏️</span>
          <div className="min-w-0">
            <div className="flex items-center gap-2">
              <span className="font-bold text-xs sm:text-sm text-slate-100 truncate" title={title}>
                {title}
              </span>
              <span className="text-[10px] px-1.5 py-0.5 rounded bg-blue-500/10 text-blue-300 border border-blue-500/20 font-bold shrink-0">
                한글 직접 편집기
              </span>
              {hasChanges && (
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-amber-500/20 text-amber-300 font-semibold shrink-0">
                  수정됨
                </span>
              )}
            </div>
            <p className="text-[10px] text-slate-400 hidden sm:block">
              표, 글자, 서식을 클릭하여 직접 수정하고 [저장]을 누르면 메모에 즉시 반영됩니다.
            </p>
          </div>
        </div>

        {/* 액션 버튼 그룹 */}
        <div className="flex items-center gap-1.5 shrink-0">
          {saveSuccess && (
            <span className="text-xs text-emerald-400 font-semibold flex items-center gap-1 animate-in fade-in">
              <CheckCircle size={14} />
              <span className="hidden sm:inline">메모에 저장 완료!</span>
            </span>
          )}

          {/* 최초 원본으로 되돌리기 */}
          {originalData && onRestoreOriginal && (
            <button
              type="button"
              onClick={() => setShowRestoreConfirm(true)}
              className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-amber-300 border border-slate-700 hover:border-amber-500/40 rounded-lg text-xs font-medium flex items-center gap-1 transition-all cursor-pointer"
              title={`최초 편집 전 원본 버전(${originalDate || '이전'})으로 되돌립니다.`}
            >
              <RotateCcw size={13} />
              <span className="hidden sm:inline">최초 원본 되돌리기</span>
            </button>
          )}

          {/* 저장 버튼 */}
          <button
            type="button"
            onClick={() => handleSave()}
            disabled={isLoading || isSaving}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95 disabled:opacity-50"
            title="수정한 내용을 메모에 즉시 저장"
          >
            {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
            <span>저장 및 반영</span>
          </button>

          {/* 다운로드 버튼 */}
          <button
            type="button"
            onClick={() => handleDownload()}
            disabled={isLoading}
            className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all active:scale-95 disabled:opacity-50"
            title="수정된 한글 파일 내 PC로 다운로드"
          >
            <Download size={14} className="text-indigo-400" />
            <span className="hidden md:inline">다운로드</span>
          </button>

          {/* 다른 이름으로 저장 버튼 */}
          <button
            type="button"
            onClick={() => handleSaveAs()}
            disabled={isLoading}
            className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all active:scale-95 disabled:opacity-50"
            title="원하는 위치에 다른 이름으로 파일 저장"
          >
            <FolderDown size={14} className="text-emerald-400" />
            <span className="hidden md:inline">다른 이름으로 저장</span>
          </button>

          {/* 전체화면 전환 버튼 */}
          {floating.toggleFullscreenButton}

          {floating.toggleFloatingButton}

          {/* 닫기 버튼 */}
          <button
            type="button"
            onClick={handleClose}
            className="p-1.5 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 rounded-lg transition-colors ml-1"
            title="에디터 닫기"
          >
            <X size={18} />
          </button>
        </div>
      </div>

      {/* 에러 메시지 알림 바 */}
      {errorMessage && (
        <div className="px-4 py-2 bg-rose-950/80 border-b border-rose-800 text-rose-300 text-xs flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <AlertCircle size={14} className="shrink-0" />
            <span>{errorMessage}</span>
          </div>
          <button
            type="button"
            onClick={() => setErrorMessage(null)}
            className="text-rose-400 hover:text-white text-xs underline"
          >
            닫기
          </button>
        </div>
      )}

      {/* 에디터 임베드 영역 */}
      <div className="flex-1 relative w-full h-full bg-slate-900 overflow-hidden">
        {isLoading && (
          <div className="absolute inset-0 flex flex-col items-center justify-center gap-3 bg-slate-950/80 backdrop-blur-xs text-slate-300 text-sm z-20">
            <div className="w-9 h-9 border-3 border-indigo-500 border-t-transparent rounded-full animate-spin" />
            <div className="text-center space-y-1">
              <p className="font-semibold text-slate-100">한글(HWP) 편집기를 준비하는 중입니다...</p>
              <p className="text-xs text-slate-400">문서 서식과 표 구조를 불러오고 있습니다.</p>
            </div>
          </div>
        )}

        {/* rhwp 에디터가 마운트될 DOM 컨테이너 */}
        <div ref={containerRef} className="w-full h-full" />
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
                <p className="text-xs text-slate-400">한글 문서 수정 내용이 아직 메모에 저장되지 않았습니다.</p>
              </div>
            </div>

            <p className="text-xs text-slate-300 leading-relaxed bg-slate-800/60 p-3 rounded-xl border border-slate-700/50">
              <strong>[저장하지 않고 닫기]</strong>를 누르면 수정한 내용이 취소되며, 브라우저에 남은 임시 복구본도 깨끗하게 삭제되어 다음에 열 때 '문서 복구' 안내창이 뜨지 않습니다.
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

      {/* 변경 추적 손실 주의 경고 모달 */}
      {showTrackChangesWarning && (
        <div className="fixed inset-0 z-100 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
          <div className="bg-slate-900 border border-amber-500/30 rounded-2xl p-6 max-w-md w-full shadow-2xl space-y-4">
            <div className="flex items-center gap-3 text-amber-400">
              <div className="p-2.5 rounded-xl bg-amber-500/10 border border-amber-500/20">
                <AlertCircle size={24} />
              </div>
              <div>
                <h3 className="font-bold text-base text-slate-100">변경 추적 손실 주의</h3>
                <p className="text-xs text-amber-300/80">이 문서에는 변경 추적(교정 이력)이 포함되어 있습니다.</p>
              </div>
            </div>

            <p className="text-xs text-slate-300 leading-relaxed bg-slate-800/60 p-3 rounded-xl border border-slate-700/50">
              이 문서에는 변경 추적(삽입/삭제 이력)이 포함되어 있습니다. <strong>변경 추적이 걸린 문단을 직접 수정하면 그 문단의 변경 추적이 일반 텍스트로 합쳐질 수 있습니다</strong>. (.hwp로 저장하면 변경 내용이 수락된 상태로 저장됩니다) 그래도 저장하시겠습니까?
            </p>

            <div className="flex flex-col sm:flex-row gap-2 pt-1">
              <button
                type="button"
                id="hwp-confirm-save-anyway-btn"
                onClick={handleConfirmSaveAnyway}
                className="flex-1 py-2.5 px-3 bg-amber-600 hover:bg-amber-500 text-white rounded-xl text-xs font-bold transition-all shadow-md active:scale-95 flex items-center justify-center gap-1.5"
              >
                그래도 저장
              </button>
              <button
                type="button"
                id="hwp-cancel-save-warning-btn"
                onClick={handleCancelTrackChangesWarning}
                className="py-2.5 px-3 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-medium transition-all"
              >
                취소
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 원본 되돌리기 확인 모달 */}
      {showRestoreConfirm && (
        <div className="fixed inset-0 z-100 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
          <div className="bg-slate-900 border border-slate-800 rounded-xl max-w-sm w-full p-5 shadow-2xl space-y-4">
            <div className="flex items-center gap-3 text-amber-400">
              <RotateCcw size={22} />
              <h3 className="font-bold text-sm text-slate-100">최초 원본 버전으로 되돌리기</h3>
            </div>
            <p className="text-xs text-slate-300 leading-relaxed">
              편집을 시작하기 전의 최초 한글 문서 원본 상태로 되돌립니다.<br />
              <span className="text-slate-400 text-[11px] mt-1 block">
                (백업 시점: {originalDate || '편집 시작 전'})
              </span>
              되돌리기를 실행하면 현재 수정한 모든 내용이 취소됩니다. 진행하시겠습니까?
            </p>
            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setShowRestoreConfirm(false)}
                className="px-3 py-1.5 text-xs text-slate-400 hover:text-slate-200"
              >
                취소
              </button>
              <button
                type="button"
                onClick={handleConfirmRestore}
                className="px-3.5 py-1.5 text-xs bg-amber-600 hover:bg-amber-500 text-white rounded-lg font-semibold shadow-md"
              >
                원본으로 되돌리기
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 5단계 임시 복구본 발견 모달 */}
      {showDraftRecoveryModal && draftToRecover && (
        <div className="fixed inset-0 z-100 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
          <div className="bg-slate-900 border border-emerald-500/30 rounded-2xl p-6 max-w-md w-full shadow-2xl space-y-4">
            <div className="flex items-center gap-3 text-emerald-400">
              <div className="p-2.5 rounded-xl bg-emerald-500/10 border border-emerald-500/20">
                <FileText size={24} />
              </div>
              <div>
                <h3 className="font-bold text-base text-slate-100">이전 편집 임시본 발견</h3>
                <p className="text-xs text-slate-400">
                  저장되지 않은 자동 임시본이 보관되어 있습니다.
                </p>
              </div>
            </div>
            <div className="text-xs text-slate-300 leading-relaxed bg-slate-800/60 p-3 rounded-xl border border-slate-700/50 space-y-1.5">
              <div className="flex justify-between text-slate-400">
                <span>임시본 저장 시점:</span>
                <span className="text-slate-200 font-mono">
                  {new Date(draftToRecover.updatedAt).toLocaleTimeString()}
                </span>
              </div>
              <div className="flex justify-between text-slate-400">
                <span>임시 데이터 크기:</span>
                <span className="text-slate-200 font-mono">
                  {(draftToRecover.byteLength / 1024).toFixed(1)} KB
                </span>
              </div>
              <p className="text-[11px] text-emerald-300/80 pt-1 border-t border-slate-700/50">
                비정상 종료 전 작업하던 한글 문서 내용을 불러오시겠습니까?
              </p>
            </div>
            <div className="flex items-center gap-2 pt-1">
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
                  if (fileContextRef.current && editorRef.current) {
                    const originalBytes = dataUrlToUint8Array(data);
                    let bytes = isHwpx ? await editorRef.current.exportHwpx() : await editorRef.current.exportHwp();
                    if (isHwpx) {
                      bytes = await postProcessHwpx(bytes, originalBytes);
                    } else {
                      bytes = postProcessHwp(bytes);
                    }
                    await saveToFileHandle(fileContextRef.current, bytes);
                    setHasChanges(false);
                    setSaveSuccess(true);
                    setTimeout(() => setSaveSuccess(false), 3000);
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

      {floating.resizeHandle}
    </div>
  </div>
);
};
