import React, { useState, useEffect, useRef } from 'react';
import { X, Save, Download, AlertCircle, CheckCircle, Loader2 } from 'lucide-react';
import { createEditor, type RhwpEditor } from '@rhwp/editor';
import { useFloatingWindow } from './useFloatingWindow';

interface HwpEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
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
  onSave,
}) => {
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [saveSuccess, setSaveSuccess] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [showExitConfirm, setShowExitConfirm] = useState<boolean>(false);

  const containerRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<RhwpEditor | null>(null);
  const floating = useFloatingWindow({ id: 'hwp-editor', defaultWidth: 920, defaultHeight: 820 });

  const isHwpx = title.toLowerCase().endsWith('.hwpx');
  const mimeType = isHwpx ? 'application/vnd.hancom.hwpx' : 'application/x-hwp';

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

        // rhwp 웹 에디터 임베드 인스턴스 생성
        const editor = await createEditor(containerRef.current, {
          studioUrl: 'https://edwardkim.github.io/rhwp/',
          width: '100%',
          height: '100%',
          requestTimeoutMs: 30000,
        });

        if (!isMounted) {
          editor.destroy();
          return;
        }

        editorRef.current = editor;

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
        }
      }
    };

    initEditor();

    return () => {
      isMounted = false;
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

  // 수정된 문서 저장 및 메모 본문에 반영
  const handleSave = async () => {
    if (!editorRef.current || isSaving) return;

    setIsSaving(true);
    setErrorMessage(null);
    try {
      const editor = editorRef.current;
      const bytes = isHwpx ? await editor.exportHwpx() : await editor.exportHwp();
      const updatedDataUrl = await uint8ArrayToDataUrl(bytes, mimeType);

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

  // 수정된 원본 파일 내 PC에 바로 다운로드
  const handleDownload = async () => {
    if (!editorRef.current) return;

    try {
      const editor = editorRef.current;
      const bytes = isHwpx ? await editor.exportHwpx() : await editor.exportHwp();
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

  const handleClose = async () => {
    let isDirty = hasChanges;

    // 닫기 버튼 클릭 시 rhwp 에디터의 최신 dirty 상태를 즉시 최종 확인
    if (!isDirty && editorRef.current) {
      try {
        const state = await editorRef.current.getDocumentState();
        if (state && state.dirty) {
          isDirty = true;
          setHasChanges(true);
        }
      } catch {
        // ignore
      }
    }

    if (isDirty) {
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

          {/* 저장 버튼 */}
          <button
            type="button"
            onClick={handleSave}
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
            onClick={handleDownload}
            disabled={isLoading}
            className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all active:scale-95 disabled:opacity-50"
            title="수정된 한글 파일 내 PC로 다운로드"
          >
            <Download size={14} className="text-indigo-400" />
            <span className="hidden md:inline">다운로드</span>
          </button>

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

      {floating.resizeHandle}
    </div>
  </div>
);
};
