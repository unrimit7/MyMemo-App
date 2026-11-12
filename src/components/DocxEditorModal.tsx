import React, { useState, useEffect, useRef } from 'react';
import {
  X,
  Save,
  Download,
  Bold,
  Italic,
  Heading1,
  Heading2,
  Heading3,
  List,
  ListOrdered,
  Quote,
  Table,
  CheckCircle,
  AlertCircle,
  AlertTriangle,
  RotateCcw,
  Loader2,
  FileText,
  Eye,
  Edit3,
  FolderDown,
} from 'lucide-react';
import { SuperDoc } from 'superdoc';
import 'superdoc/style.css';

import { extractDocxTextAndHtml, generateDocxDataUrl } from '../utils/docxViewer';
import { detectDocxUnsupportedFeatures, type DocxUnsupportedFeatureReport } from '../utils/docxLossDetector';
import { marked } from 'marked';
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

interface DocxEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  originalData?: string;
  originalDate?: string;
  initialEngine?: 'superdoc' | 'legacy';
  fileHandle?: any;
  onSave?: (updatedDataUrl: string, engine?: 'superdoc' | 'legacy') => Promise<void> | void;
  onRestoreOriginal?: () => Promise<void> | void;
}

function dataUrlToBlob(dataOrUrl: string, mime: string = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'): Blob {
  const commaIdx = dataOrUrl.indexOf(',');
  const b64 = commaIdx !== -1 ? dataOrUrl.slice(commaIdx + 1) : dataOrUrl;
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new Blob([bytes], { type: mime });
}

export const DocxEditorModal: React.FC<DocxEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  attId,
  originalData,
  originalDate,
  initialEngine = 'superdoc',
  fileHandle,
  onSave,
  onRestoreOriginal,
}) => {
  const [engine, setEngine] = useState<'superdoc' | 'legacy'>(initialEngine);
  const [docTitle, setDocTitle] = useState<string>(title.replace(/\.docx?$/i, ''));
  const [content, setContent] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [saveToast, setSaveToast] = useState<string | null>(null);
  const [errorToast, setErrorToast] = useState<string | null>(null);
  const [viewMode, setViewMode] = useState<'edit' | 'preview' | 'split'>('split');

  const [showExitConfirm, setShowExitConfirm] = useState<boolean>(false);
  const [showRestoreConfirm, setShowRestoreConfirm] = useState<boolean>(false);
  const [showLossWarningModal, setShowLossWarningModal] = useState<boolean>(false);
  const [lossReport, setLossReport] = useState<DocxUnsupportedFeatureReport | null>(null);

  // 5단계 임시본 복구 모달 및 파일 핸들러 충돌 모달 상태
  const draftKey = `docx_${attId || ''}_${title}`;
  const [draftToRecover, setDraftToRecover] = useState<OfficeDraft | null>(null);
  const [showDraftRecoveryModal, setShowDraftRecoveryModal] = useState<boolean>(false);
  const [showFileConflictModal, setShowFileConflictModal] = useState<boolean>(false);
  const fileContextRef = useRef<FileHandleContext | null>(null);

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
    if (!isOpen || !hasChanges) return;

    let autoSaveIntervalMs = 30000;

    const performDraftSave = async () => {
      if (!hasChanges) return;
      try {
        const t0 = performance.now();
        let draftData: any = content;
        if (engine === 'superdoc' && superdocRef.current) {
          draftData = (superdocRef.current as any).editor?.getJSON?.() || (superdocRef.current as any).getContent?.() || content;
        }
        const duration = performance.now() - t0;
        const baseHash = await computeDataHash(data);
        const dataStr = typeof draftData === 'string' ? draftData : JSON.stringify(draftData);

        await saveOfficeDraft({
          draftKey,
          attId,
          title,
          fileType: 'docx',
          baseOriginalHash: baseHash,
          data: draftData,
          dataType: engine === 'superdoc' ? 'superdoc_json' : 'data_url',
          updatedAt: Date.now(),
          byteLength: dataStr.length,
          saveDurationMs: Math.round(duration),
        });

        if (duration >= 1000) {
          autoSaveIntervalMs = 60000;
        }
      } catch (err) {
        console.warn('[Docx] 임시본 자동 저장 오류:', err);
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
  }, [isOpen, hasChanges, draftKey, attId, title, data, engine, content]);

  const containerRef = useRef<HTMLDivElement>(null);
  const superdocRef = useRef<SuperDoc | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const floating = useFloatingWindow({ id: 'docx-editor', defaultWidth: 1040, defaultHeight: 840 });

  // 문서 사전 검사 (매크로, OLE 등)
  useEffect(() => {
    if (!isOpen || !data) return;
    let isCancelled = false;
    detectDocxUnsupportedFeatures(data).then((report) => {
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



  // 레거시 마크다운 편집기 내용 추출
  useEffect(() => {
    if (!isOpen || !data || engine !== 'legacy') return;
    let isMounted = true;
    setIsLoading(true);

    extractDocxTextAndHtml(data)
      .then(({ markdown, text }) => {
        if (!isMounted) return;
        setContent(markdown || text || '');
        setIsLoading(false);
      })
      .catch((err) => {
        console.error('워드 텍스트 추출 실패:', err);
        if (isMounted) {
          setErrorToast('워드 문서 텍스트를 불러오지 못했습니다.');
          setIsLoading(false);
        }
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, data, engine]);

  // SuperDoc 인스턴스 초기화 및 라이프사이클 관리
  useEffect(() => {
    if (!isOpen || !data || engine !== 'superdoc') return;

    let isDisposed = false;
    setIsLoading(true);

    const timer = setTimeout(() => {
      if (isDisposed || !containerRef.current) return;

      try {
        if (superdocRef.current) {
          try {
            superdocRef.current.destroy();
          } catch {}
          superdocRef.current = null;
        }

        // 이전 렌더링 엘리먼트 비우기
        if (containerRef.current) {
          containerRef.current.innerHTML = '';
        }

        const blob = dataUrlToBlob(data);
        const superdoc = new SuperDoc({
          selector: containerRef.current,
          document: blob,
          documentMode: 'editing',
        });

        superdocRef.current = superdoc;

        superdoc.on('ready', () => {
          if (!isDisposed) {
            setIsLoading(false);
          }
        });

        superdoc.on('editor-update', () => {
          if (!isDisposed) {
            setHasChanges(true);
          }
        });

        superdoc.on('history-change', () => {
          if (!isDisposed) {
            setHasChanges(true);
          }
        });
      } catch (err) {
        console.error('SuperDoc 초기화 실패:', err);
        if (!isDisposed) {
          setErrorToast('SuperDoc 워드 엔진 로드 실패. 레거시 모드로 전환할 수 있습니다.');
          setIsLoading(false);
        }
      }
    }, 100);

    return () => {
      isDisposed = true;
      clearTimeout(timer);
      if (superdocRef.current) {
        try {
          superdocRef.current.destroy();
        } catch {}
        superdocRef.current = null;
      }
    };
  }, [isOpen, data, engine]);

  // 레거시 서식 도구 적용 헬퍼
  const applyFormat = (prefix: string, suffix: string = '') => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const selected = content.substring(start, end);

    const replacement = `${prefix}${selected || '텍스트'}${suffix}`;
    const newContent = content.substring(0, start) + replacement + content.substring(end);

    setContent(newContent);
    setHasChanges(true);

    setTimeout(() => {
      textarea.focus();
      textarea.setSelectionRange(start + prefix.length, start + replacement.length - suffix.length);
    }, 10);
  };

  const insertTable = () => {
    const sampleTable = `\n| 항목 | 내용 | 비고 |\n| --- | --- | --- |\n| 항목 1 | 내용 입력 |  |\n| 항목 2 | 내용 입력 |  |\n`;
    applyFormat(sampleTable);
  };

  // 저장 트리거 (손실 경고 검사)
  const triggerSave = () => {
    if (lossReport && lossReport.unsupportedList.length > 0) {
      setShowLossWarningModal(true);
      return;
    }
    executeSave();
  };

  // 실제 저장 실행
  const executeSave = async () => {
    setShowLossWarningModal(false);
    setIsSaving(true);
    setSaveToast(null);
    setErrorToast(null);

    try {
      let updatedDataUrl = '';

      if (engine === 'superdoc' && superdocRef.current) {
        const exportedBlob = await superdocRef.current.export({ exportType: ['docx'], triggerDownload: false });
        const reader = new FileReader();
        await new Promise<void>((resolve, reject) => {
          reader.onloadend = () => {
            updatedDataUrl = reader.result as string;
            resolve();
          };
          reader.onerror = reject;
          reader.readAsDataURL(exportedBlob);
        });
      } else {
        updatedDataUrl = await generateDocxDataUrl(docTitle, content);
      }

      if (onSave) {
        await onSave(updatedDataUrl, engine);
      }

      // 5단계 임시본 정리
      await deleteOfficeDraft(draftKey);

      // FileHandle 연동 (PWA 파일 연결)
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

      setHasChanges(false);
      setSaveToast('워드 문서가 성공적으로 저장 및 메모에 반영되었습니다!');
      setTimeout(() => setSaveToast(null), 3000);
    } catch (err) {
      console.error('워드 문서 저장 실패:', err);
      setErrorToast('워드 문서를 저장하는 중 오류가 발생했습니다.');
    } finally {
      setIsSaving(false);
    }
  };

  // 다운로드 핸들러
  const handleDownload = async () => {
    try {
      if (engine === 'superdoc' && superdocRef.current) {
        const exportedBlob = await superdocRef.current.export({ exportType: ['docx'] });
        const url = URL.createObjectURL(exportedBlob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${docTitle}.docx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      } else {
        const exportUrl = await generateDocxDataUrl(docTitle, content);
        const link = document.createElement('a');
        link.href = exportUrl;
        link.download = `${docTitle}.docx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      }
    } catch (err) {
      console.error('다운로드 생성 실패:', err);
      alert('파일을 다운로드하지 못했습니다.');
    }
  };

  // 다른 이름으로 저장 (showSaveFilePicker 연동)
  const handleSaveAs = async () => {
    try {
      let bytes: Uint8Array;
      if (engine === 'superdoc' && superdocRef.current) {
        const exportedBlob = await superdocRef.current.export({ exportType: ['docx'], triggerDownload: false });
        const ab = await exportedBlob.arrayBuffer();
        bytes = new Uint8Array(ab);
      } else {
        const exportUrl = await generateDocxDataUrl(docTitle, content);
        const b64 = exportUrl.split(',')[1] || '';
        const bin = atob(b64);
        bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      }
      await saveFileAs(bytes, `${docTitle}.docx`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'docx');
    } catch (err) {
      console.error('다른 이름으로 저장 실패:', err);
    }
  };

  // 5단계 임시본 복구 적용 및 폐기
  const handleApplyDraftRecovery = () => {
    if (!draftToRecover) return;
    try {
      if (engine === 'superdoc' && superdocRef.current && typeof draftToRecover.data === 'object') {
        (superdocRef.current as any).editor?.commands?.setContent?.(draftToRecover.data);
      } else if (typeof draftToRecover.data === 'string') {
        setContent(draftToRecover.data);
      }
      setHasChanges(true);
      setShowDraftRecoveryModal(false);
      setSaveToast('임시 복구본이 성공적으로 복구되었습니다.');
      setTimeout(() => setSaveToast(null), 3000);
    } catch (err) {
      console.error('워드 임시본 복구 실패:', err);
    }
  };

  const handleDiscardDraftRecovery = async () => {
    setShowDraftRecoveryModal(false);
    setDraftToRecover(null);
    await deleteOfficeDraft(draftKey);
  };

  // 닫기 요청 핸들러
  const handleRequestClose = () => {
    if (hasChanges) {
      setShowExitConfirm(true);
      return;
    }
    onClose();
  };

  // 5단계 공통 단축키 훅 (Ctrl+S 통일, Alt+Enter 전체화면, Esc 대화상자 보호)
  useOfficeShortcuts({
    enabled: isOpen,
    onSave: triggerSave,
    onClose: handleRequestClose,
    hasChanges,
    onConfirmCloseNeeded: () => setShowExitConfirm(true),
    isEditingCellOrDialog: () => {
      if (document.querySelector('.superdoc-dialog, .superdoc-modal, [class*="dialog"], [class*="modal"]')) {
        return true;
      }
      return false;
    },
    toggleFullscreen: floating.toggleFullscreen,
  });

  // 원본으로 되돌리기 실행
  const handleExecuteRestore = async () => {
    setShowRestoreConfirm(false);
    if (onRestoreOriginal) {
      try {
        await onRestoreOriginal();
        setHasChanges(false);
        setSaveToast('원본 버전으로 되돌렸습니다.');
        setTimeout(() => setSaveToast(null), 3000);
      } catch (err) {
        console.error('원본 되돌리기 실패:', err);
        setErrorToast('원본으로 되돌리지 못했습니다.');
      }
    }
  };

  if (!isOpen) return null;

  return (
    <div className={floating.modalWrapperClass}>
      <div {...floating.windowContainerProps}>
        {/* 상단 1단: 헤더 */}
        <div 
          {...floating.headerProps}
          className="px-4 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0"
        >
          <div className="flex items-center gap-2.5 min-w-0">
            <span className="p-1.5 rounded-lg bg-blue-500/10 text-blue-400 border border-blue-500/20">
              <FileText size={18} />
            </span>
            <div className="truncate">
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  value={docTitle}
                  onChange={(e) => {
                    setDocTitle(e.target.value);
                    setHasChanges(true);
                  }}
                  className="bg-transparent font-bold text-sm md:text-base text-slate-100 border-b border-transparent hover:border-slate-700 focus:border-blue-500 focus:outline-none px-1"
                  title="클릭하여 워드 문서 제목 수정"
                />
                <span className="text-[10px] px-1.5 py-0.5 rounded bg-blue-500/10 text-blue-300 border border-blue-500/20 font-bold uppercase shrink-0">
                  DOCX 편집기
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

          {/* 우측 도구: 엔진 선택, 저장, 다운로드, 원본 되돌리기, 닫기 */}
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

            {/* 편집 엔진 선택 셀렉터 */}
            <div className="flex items-center gap-1 bg-slate-950/70 border border-slate-800 rounded-lg px-2 py-1 text-xs">
              <span className="text-[10px] text-slate-400 font-medium">엔진:</span>
              <select
                value={engine}
                onChange={(e) => setEngine(e.target.value as 'superdoc' | 'legacy')}
                className="bg-transparent text-slate-200 text-xs font-semibold focus:outline-none cursor-pointer"
                title="워드 편집 엔진을 전환합니다."
              >
                <option value="superdoc" className="bg-slate-900 text-slate-100">
                  SuperDoc (워드 정품 서식 보존)
                </option>
                <option value="legacy" className="bg-slate-900 text-slate-100">
                  레거시 (텍스트/마크다운)
                </option>
              </select>
            </div>

            {/* 원본으로 되돌리기 버튼 */}
            {originalData && onRestoreOriginal && (
              <button
                type="button"
                onClick={() => setShowRestoreConfirm(true)}
                className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-amber-300 border border-slate-700 hover:border-amber-500/40 rounded-lg text-xs font-medium flex items-center gap-1 transition-all"
                title={`최초 편집 전 원본 버전(${originalDate || '이전'})으로 되돌립니다.`}
              >
                <RotateCcw size={13} />
                <span className="hidden sm:inline">원본 되돌리기</span>
              </button>
            )}

            {/* 레거시 보기 모드 전환 (편집, 분할, 미리보기) */}
            {engine === 'legacy' && (
              <div className="hidden sm:flex items-center gap-0.5 bg-slate-950 border border-slate-800 rounded-lg p-0.5 text-xs">
                <button
                  type="button"
                  onClick={() => setViewMode('edit')}
                  className={`px-2 py-1 rounded transition-colors ${
                    viewMode === 'edit' ? 'bg-slate-800 text-blue-400 font-bold' : 'text-slate-400 hover:text-slate-200'
                  }`}
                  title="편집 전용"
                >
                  <Edit3 size={13} />
                </button>
                <button
                  type="button"
                  onClick={() => setViewMode('split')}
                  className={`px-2 py-1 rounded transition-colors ${
                    viewMode === 'split' ? 'bg-slate-800 text-blue-400 font-bold' : 'text-slate-400 hover:text-slate-200'
                  }`}
                  title="분할 보기"
                >
                  반반
                </button>
                <button
                  type="button"
                  onClick={() => setViewMode('preview')}
                  className={`px-2 py-1 rounded transition-colors ${
                    viewMode === 'preview' ? 'bg-slate-800 text-blue-400 font-bold' : 'text-slate-400 hover:text-slate-200'
                  }`}
                  title="인쇄 용지 미리보기"
                >
                  <Eye size={13} />
                </button>
              </div>
            )}

            {/* 다운로드 버튼 */}
            <button
              type="button"
              onClick={handleDownload}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded-lg transition-colors border border-slate-800"
              title="워드 문서(.docx) 다운로드"
            >
              <Download size={15} />
            </button>

            {/* 다른 이름으로 저장 버튼 */}
            <button
              type="button"
              onClick={handleSaveAs}
              className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 rounded-lg text-xs font-medium flex items-center gap-1 transition-all"
              title="원하는 위치에 다른 이름으로 파일 저장"
            >
              <FolderDown size={14} />
              <span className="hidden sm:inline">다른 이름으로 저장</span>
            </button>

            {/* 전체화면 전환 버튼 */}
            {floating.toggleFullscreenButton}

            {/* 저장 및 반영 버튼 */}
            <button
              type="button"
              onClick={triggerSave}
              disabled={isSaving}
              className="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-semibold shadow-md flex items-center gap-1.5 transition-all active:scale-95 disabled:opacity-50"
              title="Ctrl+S: 저장 후 메모에 즉시 반영"
            >
              {isSaving ? <Loader2 size={13} className="animate-spin" /> : <Save size={13} />}
              <span>{isSaving ? '저장 중...' : '저장 및 반영'}</span>
            </button>

            {/* 닫기 버튼 */}
            <button
              type="button"
              onClick={handleRequestClose}
              className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors"
              title="닫기"
            >
              <X size={18} />
            </button>
          </div>
        </div>

        {/* 상단 2단: 레거시 모드 마크다운 서식 툴바 */}
        {engine === 'legacy' && (
          <div className="px-4 py-1.5 bg-slate-900/90 border-b border-slate-800/80 flex items-center gap-1 overflow-x-auto text-xs shrink-0">
            <button
              type="button"
              onClick={() => applyFormat('**', '**')}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
              title="굵게 (**텍스트**)"
            >
              <Bold size={14} />
            </button>
            <button
              type="button"
              onClick={() => applyFormat('*', '*')}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
              title="기울임 (*텍스트*)"
            >
              <Italic size={14} />
            </button>
            <div className="h-4 w-px bg-slate-800 mx-1" />
            <button
              type="button"
              onClick={() => applyFormat('# ')}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
              title="제목 1 (# )"
            >
              <Heading1 size={14} />
            </button>
            <button
              type="button"
              onClick={() => applyFormat('## ')}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
              title="제목 2 (## )"
            >
              <Heading2 size={14} />
            </button>
            <button
              type="button"
              onClick={() => applyFormat('### ')}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
              title="제목 3 (### )"
            >
              <Heading3 size={14} />
            </button>
            <div className="h-4 w-px bg-slate-800 mx-1" />
            <button
              type="button"
              onClick={() => applyFormat('- ')}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
              title="글머리 기호 목록 (- )"
            >
              <List size={14} />
            </button>
            <button
              type="button"
              onClick={() => applyFormat('1. ')}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
              title="번호 매기기 목록 (1. )"
            >
              <ListOrdered size={14} />
            </button>
            <button
              type="button"
              onClick={() => applyFormat('> ')}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
              title="인용구 (> )"
            >
              <Quote size={14} />
            </button>
            <div className="h-4 w-px bg-slate-800 mx-1" />
            <button
              type="button"
              onClick={insertTable}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
              title="표 삽입"
            >
              <Table size={14} />
            </button>
          </div>
        )}

        {/* 바디 컨테이너 */}
        <div className="flex-1 relative overflow-hidden bg-slate-950 flex flex-col">
          {isLoading && (
            <div className="absolute inset-0 z-30 bg-slate-900/80 backdrop-blur-xs flex flex-col items-center justify-center gap-3 text-slate-300">
              <Loader2 size={32} className="animate-spin text-blue-400" />
              <p className="text-sm font-medium">
                {engine === 'superdoc' ? 'SuperDoc 워드 편집기를 불러오는 중입니다...' : '문서 내용을 파싱하는 중입니다...'}
              </p>
            </div>
          )}

          {/* SuperDoc WYSIWYG 편집기 */}
          {engine === 'superdoc' ? (
            <div
              ref={containerRef}
              className="w-full h-full superdoc-editor-container bg-white text-slate-900 overflow-hidden"
              style={{ minHeight: '100%', width: '100%' }}
            />
          ) : (
            /* 레거시 마크다운 편집기 */
            <div className="flex-1 flex overflow-hidden">
              {(viewMode === 'edit' || viewMode === 'split') && (
                <div className={`h-full flex flex-col ${viewMode === 'split' ? 'w-1/2 border-r border-slate-800' : 'w-full'}`}>
                  <textarea
                    ref={textareaRef}
                    value={content}
                    onChange={(e) => {
                      setContent(e.target.value);
                      setHasChanges(true);
                    }}
                    placeholder="워드 문서 내용을 마크다운으로 편집하세요..."
                    className="flex-1 w-full p-4 bg-slate-950 text-slate-100 font-mono text-xs md:text-sm resize-none focus:outline-none focus:ring-0 leading-relaxed overflow-y-auto"
                    spellCheck={false}
                  />
                </div>
              )}

              {(viewMode === 'preview' || viewMode === 'split') && (
                <div className={`h-full overflow-y-auto bg-slate-900/50 p-6 flex justify-center ${viewMode === 'split' ? 'w-1/2' : 'w-full'}`}>
                  <div className="w-full max-w-3xl bg-white text-slate-900 shadow-xl rounded-sm p-8 min-h-full">
                    <h1 className="text-2xl font-bold border-b pb-2 mb-6 text-slate-900">{docTitle}</h1>
                    <div
                      className="prose prose-slate max-w-none text-sm leading-relaxed"
                      dangerouslySetInnerHTML={{
                        __html: marked.parse(content || '*문서 내용이 비어 있습니다.*') as string,
                      }}
                    />
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* 1. 닫기 전 변경사항 경고 모달 */}
        {showExitConfirm && (
          <div className="absolute inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-xl max-w-sm w-full p-5 shadow-2xl space-y-4">
              <div className="flex items-center gap-3 text-amber-400">
                <AlertTriangle size={24} />
                <h3 className="font-bold text-sm text-slate-100">저장되지 않은 변경사항</h3>
              </div>
              <p className="text-xs text-slate-300 leading-relaxed">
                수정된 워드 문서 내용이 아직 저장되지 않았습니다.<br />
                저장하지 않고 닫으시면 변경된 서식과 텍스트가 모두 손실됩니다.
              </p>
              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setShowExitConfirm(false)}
                  className="px-3 py-1.5 text-xs text-slate-400 hover:text-slate-200"
                >
                  계속 편집
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setShowExitConfirm(false);
                    onClose();
                  }}
                  className="px-3 py-1.5 text-xs bg-rose-600/80 hover:bg-rose-600 text-white rounded-lg font-medium"
                >
                  저장 안 함
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setShowExitConfirm(false);
                    triggerSave();
                  }}
                  className="px-3.5 py-1.5 text-xs bg-blue-600 hover:bg-blue-500 text-white rounded-lg font-semibold shadow-md"
                >
                  저장하고 닫기
                </button>
              </div>
            </div>
          </div>
        )}

        {/* 2. 원본 되돌리기 확인 모달 */}
        {showRestoreConfirm && (
          <div className="absolute inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-xl max-w-sm w-full p-5 shadow-2xl space-y-4">
              <div className="flex items-center gap-3 text-amber-400">
                <RotateCcw size={22} />
                <h3 className="font-bold text-sm text-slate-100">원본 버전으로 되돌리기</h3>
              </div>
              <p className="text-xs text-slate-300 leading-relaxed">
                편집을 시작하기 전의 최초 워드 문서 원본 상태로 되돌립니다.<br />
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
                  onClick={handleExecuteRestore}
                  className="px-3.5 py-1.5 text-xs bg-amber-600 hover:bg-amber-500 text-white rounded-lg font-semibold shadow-md"
                >
                  원본으로 되돌리기
                </button>
              </div>
            </div>
          </div>
        )}

        {/* 3. 손실 경고 다이얼로그 */}
        {showLossWarningModal && lossReport && (
          <div className="absolute inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-amber-500/30 rounded-xl max-w-md w-full p-5 shadow-2xl space-y-4">
              <div className="flex items-center gap-3 text-amber-400">
                <AlertTriangle size={24} className="shrink-0" />
                <h3 className="font-bold text-sm text-slate-100">워드 서식 보존 주의 안내</h3>
              </div>
              <div className="space-y-2 text-xs text-slate-300 leading-relaxed">
                <p>
                  이 워드 파일에는 웹 오피스 엔진에서 완벽히 보존되지 않을 수 있는 특수 요소가 포함되어 있습니다:
                </p>
                <ul className="list-disc list-inside bg-slate-950/60 p-2.5 rounded-lg border border-slate-800 text-amber-300/90 font-medium space-y-1">
                  {lossReport.unsupportedList.map((item, idx) => (
                    <li key={idx}>{item}</li>
                  ))}
                </ul>
                <p className="text-[11px] text-slate-400">
                  위 요소를 제외한 본문 문단 서식, 글꼴, 표, 이미지 등은 원본대로 안전하게 보존됩니다.
                </p>
              </div>
              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setShowLossWarningModal(false)}
                  className="px-3 py-1.5 text-xs text-slate-400 hover:text-slate-200"
                >
                  취소
                </button>
                <button
                  type="button"
                  onClick={executeSave}
                  className="px-3.5 py-1.5 text-xs bg-amber-600 hover:bg-amber-500 text-white rounded-lg font-semibold shadow-md"
                >
                  그래도 저장
                </button>
              </div>
            </div>
          </div>
        )}

        {/* 4. 5단계 임시 복구본 발견 모달 */}
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
                  비정상 종료 전 작업하던 워드 문서 내용을 불러오시겠습니까?
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

        {/* 5. 5단계 파일 충돌 모달 (디스크 파일 변경 감지) */}
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
                    if (fileContextRef.current) {
                      let updatedDataUrl = '';
                      if (engine === 'superdoc' && superdocRef.current) {
                        const exportedBlob = await superdocRef.current.export({ exportType: ['docx'], triggerDownload: false });
                        const reader = new FileReader();
                        await new Promise<void>((resolve, reject) => {
                          reader.onloadend = () => {
                            updatedDataUrl = reader.result as string;
                            resolve();
                          };
                          reader.onerror = reject;
                          reader.readAsDataURL(exportedBlob);
                        });
                      } else {
                        updatedDataUrl = await generateDocxDataUrl(docTitle, content);
                      }
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
    </div>
  );
};
