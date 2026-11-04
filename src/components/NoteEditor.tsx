import React, { useState, useRef, useEffect, useLayoutEffect, useMemo } from 'react';
import type { Note, Project } from '../types';
import { marked } from 'marked';
import { compressImageToDataUrl } from '../utils/image';
import { processFileForNote } from '../utils/file';
import { cleanLegacyAttachments, renderAttachmentsInHtml, invalidateAttachmentBlobUrl } from '../utils/attachment';
import { exportNoteToPdf } from '../utils/exportPdf';
import { getFolderHierarchyPath, updateTitleWithFolderPath } from '../utils/folderPath';
import { PdfViewerModal } from './PdfViewerModal';
import { PdfEditorModal } from './PdfEditorModal';
import { ImageEditorModal } from './ImageEditorModal';
import { loadPdfDocument, renderPdfPageToCanvas, extractPdfText, invalidatePdfCache } from '../utils/pdfViewer';
import { HwpViewerModal } from './HwpViewerModal';
import { HwpEditorModal } from './HwpEditorModal';
import { loadHwpDocument, renderHwpPageToSvg, extractHwpText, invalidateHwpCache } from '../utils/hwpViewer';
import type { AttachmentItem } from '../types';
import { 
  Lock, 
  Unlock, 
  Eye, 
  Edit3, 
  Columns, 
  ArrowLeft, 
  Bold, 
  Italic, 
  Heading1, 
  Heading2, 
  Heading3,
  Strikethrough,
  CheckSquare, 
  CheckCheck,
  List, 
  ListOrdered,
  Indent,
  Outdent,
  Code, 
  FileCode,
  Quote, 
  Image as ImageIcon,
  AlertTriangle,
  Folder,
  Loader2,
  Trash2,
  PanelLeftOpen,
  FileDown,
  Keyboard,
  X,
  Table,
  Link2,
  Minus,
  CloudUpload,
  Undo2,
  Redo2,
  FileText,
  Palette,
  Download
} from 'lucide-react';

// 마크다운 번호 매기기 목록 렌더링 시 사용자가 직접 입력한 번호(예: 1., 2., 3.)와 보기 모드의 번호가 일치하도록 커스텀 렌더러 설정
marked.use({
  renderer: {
    listitem(item: any) {
      const match = item.raw.match(/^\s*(\d+)\.\s+/);
      const val = match ? ` value="${match[1]}"` : '';
      return `<li${val}>${(this as any).parser.parse(item.tokens)}</li>\n`;
    },
    link({ href, title, tokens }: any) {
      const text = (this as any).parser.parseInline(tokens);
      const titleAttr = title ? ` title="${title}"` : '';
      return `<a href="${href}" target="_blank" rel="noopener noreferrer"${titleAttr} class="text-indigo-400 hover:text-indigo-300 underline underline-offset-2">${text}</a>`;
    }
  }
});

// 에디터 지원 폰트 색상 및 형광펜 배경색 팔레트 상수
const TEXT_COLORS = [
  { name: '빨강', value: '#ef4444' },
  { name: '주황', value: '#f97316' },
  { name: '노랑', value: '#eab308' },
  { name: '초록', value: '#10b981' },
  { name: '파랑', value: '#3b82f6' },
  { name: '보라', value: '#a855f7' },
  { name: '회색', value: '#94a3b8' },
];

const HIGHLIGHT_COLORS = [
  { name: '노랑', value: '#fef08a' },
  { name: '연두', value: '#bbf7d0' },
  { name: '하늘', value: '#bae6fd' },
  { name: '분홍', value: '#fbcfe8' },
];

interface NoteEditorProps {
  note: Note | null;
  projects: Project[];
  onUpdateNote: (updated: Partial<Note>) => void;
  onBackToList: () => void;
  onDeleteNote?: (id: string) => void;
  isSidebarCollapsed?: boolean;
  onToggleSidebar?: () => void;
  isNoteListCollapsed?: boolean;
  onToggleNoteList?: () => void;
  onSync?: () => void;
  isSyncing?: boolean;
}

export const NoteEditor: React.FC<NoteEditorProps> = ({
  note,
  projects,
  onUpdateNote,
  onBackToList,
  onDeleteNote,
  isSidebarCollapsed,
  onToggleSidebar,
  isNoteListCollapsed,
  onToggleNoteList,
  onSync,
  isSyncing,
}) => {
  // 모바일/휴대폰 환경 기본 감지 (휴대폰에서는 기본 모드를 '보기 모드'로 제공)
  const isMobileEnvironment = () => {
    if (typeof window === 'undefined') return false;
    return window.innerWidth < 768 || /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent);
  };

  const [viewMode, setViewMode] = useState<'edit' | 'preview' | 'split'>(() => {
    if (typeof window === 'undefined') return 'edit';
    if (isMobileEnvironment()) return 'preview';
    if (window.innerWidth >= 1280) return 'split';
    return 'edit';
  });
  const [isProcessingImage, setIsProcessingImage] = useState(false);
  const [isProcessingPdf, setIsProcessingPdf] = useState(false);
  const [isProcessingHwp, setIsProcessingHwp] = useState(false);
  const [isExportingPdf, setIsExportingPdf] = useState(false);
  const [isDragOverEditor, setIsDragOverEditor] = useState(false);
  const [showShortcutsModal, setShowShortcutsModal] = useState(false);
  const [lightboxImage, setLightboxImage] = useState<{ attId: string; src: string; alt: string } | null>(null);
  const [activePdfModal, setActivePdfModal] = useState<{
    data: string;
    title: string;
    initialPage: number;
  } | null>(null);
  const [activePdfEditorModal, setActivePdfEditorModal] = useState<{
    attId: string;
    data: string;
    title: string;
    initialPage: number;
  } | null>(null);
  const [activeHwpModal, setActiveHwpModal] = useState<{
    data: string;
    title: string;
    initialPage: number;
  } | null>(null);
  const [activeHwpEditorModal, setActiveHwpEditorModal] = useState<{
    attId: string;
    data: string;
    title: string;
  } | null>(null);
  const [activeImageEditorModal, setActiveImageEditorModal] = useState<{
    attId: string;
    data: string;
    title: string;
  } | null>(null);
  const pdfPagesRef = useRef<Map<string, number>>(new Map());
  const hwpPagesRef = useRef<Map<string, number>>(new Map());

  // 편집창 & 미리보기 분할(Split) 비율 (기본 50%, 15%~85% 조절 가능, 로컬 스토리지 보존)
  const [splitRatio, setSplitRatio] = useState<number>(() => {
    if (typeof window === 'undefined') return 50;
    const saved = localStorage.getItem('mymemo_editor_split_ratio');
    if (saved) {
      const parsed = parseFloat(saved);
      if (!isNaN(parsed) && parsed >= 15 && parsed <= 85) return parsed;
    }
    return 50;
  });
  const [isResizingSplit, setIsResizingSplit] = useState(false);
  const splitContainerRef = useRef<HTMLDivElement>(null);

  // 한글 IME 조합 깨짐 방지를 위한 로컬 버퍼 상태
  const [localTitle, setLocalTitle] = useState(note?.title || '');
  const [localContent, setLocalContent] = useState(note?.content || '');
  const [localAttachments, setLocalAttachments] = useState<Record<string, AttachmentItem>>(note?.attachments || {});
  const localAttachmentsRef = useRef(localAttachments);
  localAttachmentsRef.current = localAttachments;

  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);
  const hwpInputRef = useRef<HTMLInputElement>(null);
  const updateTimerRef = useRef<any>(null);
  const currentNoteIdRef = useRef(note?.id);
  const lastSentContentRef = useRef<string | null>(null);
  const lastSentTitleRef = useRef<string | null>(null);
  const sentContentsHistoryRef = useRef<Set<string>>(new Set());

  // 실행 취소(Undo) / 다시 실행(Redo) 전용 히스토리 스택
  const historyRef = useRef<{ content: string; cursorPos?: number }[]>([]);
  const historyIndexRef = useRef<number>(-1);
  const isUndoRedoActionRef = useRef<boolean>(false);
  const pushHistoryTimerRef = useRef<any>(null);

  // 히스토리 스택에 스냅샷 기록
  const pushHistory = (content: string, cursorPos?: number) => {
    if (isUndoRedoActionRef.current) return;
    const history = historyRef.current;
    const currentIndex = historyIndexRef.current;
    const newHistory = history.slice(0, currentIndex + 1);
    if (newHistory.length > 0 && newHistory[newHistory.length - 1].content === content) {
      return;
    }
    newHistory.push({ content, cursorPos });
    if (newHistory.length > 60) {
      newHistory.shift();
    }
    historyRef.current = newHistory;
    historyIndexRef.current = newHistory.length - 1;
  };

  // 타이핑 중 디바운스 히스토리 기록
  const debouncedPushHistory = (content: string, cursorPos?: number) => {
    if (pushHistoryTimerRef.current) {
      clearTimeout(pushHistoryTimerRef.current);
    }
    pushHistoryTimerRef.current = setTimeout(() => {
      pushHistory(content, cursorPos);
    }, 350);
  };

  // 메모가 변경(ID 변경)되거나 외부(동기화 등)에서 최신 버전으로 갱신될 때 로컬 버퍼 동기화
  useEffect(() => {
    if (note) {
      const isDifferentNote = note.id !== currentNoteIdRef.current;
      currentNoteIdRef.current = note.id;

      if (isDifferentNote) {
        // 완전히 다른 메모로 이동한 경우 무조건 새 메모 내용으로 로드
        lastSentContentRef.current = null;
        lastSentTitleRef.current = null;
        setLocalTitle(note.title);

        const { content: cleanedContent, attachments: cleanedAttachments, hasChanges } = cleanLegacyAttachments(
          note.content,
          note.attachments
        );
        setLocalContent(cleanedContent);
        setLocalAttachments(cleanedAttachments);

        // 히스토리 초기화
        historyRef.current = [{ content: cleanedContent, cursorPos: 0 }];
        historyIndexRef.current = 0;

        if (hasChanges) {
          onUpdateNote({ content: cleanedContent, attachments: cleanedAttachments });
        }
      } else {
        // 같은 메모를 보고 있는 경우:
        if (note.title !== localTitle && note.title !== lastSentTitleRef.current) {
          setLocalTitle(note.title);
        }

        // 사용자가 현재 textarea에 포커스하여 타이핑 중이거나,
        // 내가 방금 상위로 보낸 내용(sentContentsHistoryRef)의 에코인 경우 덮어쓰지 않음 (커서 튐 방지!)
        const isEditingNow = typeof document !== 'undefined' && textareaRef.current === document.activeElement;
        const isSentByMe = sentContentsHistoryRef.current.has(note.content);

        if (note.content !== localContent && !isSentByMe && !isEditingNow) {
          const textarea = textareaRef.current;
          const prevStart = textarea ? textarea.selectionStart : 0;
          const prevEnd = textarea ? textarea.selectionEnd : 0;

          const { content: cleanedContent, attachments: cleanedAttachments } = cleanLegacyAttachments(
            note.content,
            note.attachments
          );
          setLocalContent(cleanedContent);
          setLocalAttachments(cleanedAttachments);
          pushHistory(cleanedContent, prevStart);

          // 외부 동기화로 갱신되더라도 커서 위치 보존
          if (textarea) {
            setTimeout(() => {
              textarea.setSelectionRange(prevStart, prevEnd);
            }, 0);
          }
        }
      }
    }
  }, [note?.id, note?.updatedAt]);

  // 상위 prop에서 attachments가 갱신될 때 로컬 첨부파일 상태 동기화
  useEffect(() => {
    if (note?.attachments) {
      setLocalAttachments(note.attachments);
    }
  }, [note?.attachments]);

  // 화면 너비 및 기기 환경에 따라 기본 뷰 모드 자동 조정
  useEffect(() => {
    if (isMobileEnvironment()) {
      // 휴대폰 환경: 기본 모드를 '보기(미리보기)' 모드로 열기
      // 단, 방금 새로 추가된 빈 메모(내용이 비어있음)인 경우에만 바로 작성할 수 있도록 편집 모드로 시작
      const isNewlyCreatedEmpty = note && (
        !note.content.trim() || 
        (Date.now() - new Date(note.createdAt).getTime() < 3000 && !note.content.includes('#'))
      );

      if (isNewlyCreatedEmpty) {
        setViewMode('edit');
      } else {
        setViewMode('preview');
      }
    } else if (window.innerWidth >= 1280) {
      setViewMode('split');
    } else {
      setViewMode('edit');
    }
  }, [note?.id]);

  // 보기 모드 순환 토글 함수 (단일 단축키로 번갈아 가며 전환)
  const cycleViewMode = () => {
    const isLargeScreen = window.innerWidth >= 1024;
    setViewMode((prev) => {
      if (isLargeScreen) {
        // PC/대화면: 편집 -> 미리보기 -> 분할 -> 편집 순환
        if (prev === 'edit') return 'preview';
        if (prev === 'preview') return 'split';
        return 'edit';
      } else {
        // 모바일/소화면: 편집 <-> 미리보기 순환
        return prev === 'edit' ? 'preview' : 'edit';
      }
    });
  };

  // 모드 변경 시 편집 가능 상태면 textarea에 포커스 자동 복원
  useEffect(() => {
    if (viewMode === 'edit' || viewMode === 'split') {
      setTimeout(() => {
        textareaRef.current?.focus({ preventScroll: true });
      }, 50);
    }
  }, [viewMode]);

  // 전역 단축키 이벤트 리스너 (미리보기 화면에서도 모드 전환 가능하도록 window 레벨 등록)
  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      // 일반 input(제목 등)에 포커스되어 있을 때는 텍스트 입력을 방해하지 않음
      if (target && target.tagName === 'INPUT') return;

      const isCtrlOrCmd = e.ctrlKey || e.metaKey;

      // 1. 단일 단축키 순환 토글: Ctrl + M 또는 Alt + V
      if ((isCtrlOrCmd && e.key.toLowerCase() === 'm') || (e.altKey && e.key.toLowerCase() === 'v')) {
        e.preventDefault();
        cycleViewMode();
        return;
      }

      // 2. 각각의 모드 직행 단축키: Alt + 1(편집), Alt + 2(미리보기), Alt + 3(분할)
      if (e.altKey && !e.ctrlKey && !e.metaKey) {
        if (e.key === '1') {
          e.preventDefault();
          setViewMode('edit');
          return;
        }
        if (e.key === '2') {
          e.preventDefault();
          setViewMode('preview');
          return;
        }
        if (e.key === '3') {
          e.preventDefault();
          setViewMode('split');
          return;
        }
      }
    };

    window.addEventListener('keydown', handleGlobalKeyDown);
    return () => {
      window.removeEventListener('keydown', handleGlobalKeyDown);
    };
  }, []);

  // 부모 컴포넌트에 안전하게 디바운스 업데이트 전송
  const debouncedSyncToParent = (updatedFields: Partial<Note>) => {
    if (updatedFields.content !== undefined) {
      lastSentContentRef.current = updatedFields.content;
      sentContentsHistoryRef.current.add(updatedFields.content);
      if (sentContentsHistoryRef.current.size > 30) {
        const first = sentContentsHistoryRef.current.values().next().value;
        if (first) sentContentsHistoryRef.current.delete(first);
      }
    }
    if (updatedFields.title !== undefined) {
      lastSentTitleRef.current = updatedFields.title;
    }
    if (updateTimerRef.current) {
      clearTimeout(updateTimerRef.current);
    }
    updateTimerRef.current = setTimeout(() => {
      onUpdateNote(updatedFields);
    }, 250);
  };

  // 제목 변경 핸들러 (한글 완성형 완벽 지원)
  const handleTitleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newTitle = e.target.value;
    setLocalTitle(newTitle);
    debouncedSyncToParent({ title: newTitle });
  };

  // 본문 변경 핸들러 (한글 조합 중 리렌더링 차단)
  const handleContentChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const newContent = e.target.value;
    setLocalContent(newContent);
    debouncedSyncToParent({ content: newContent });
    debouncedPushHistory(newContent, e.target.selectionStart);
  };

  // 폴더 계층 빵가루(Breadcrumb) 계산
  const folderBreadcrumbs = useMemo(() => {
    if (!note) return [];
    const chain: Project[] = [];
    let curr = projects.find((p) => p.id === note.projectId);
    while (curr) {
      chain.unshift(curr);
      if (!curr.parentId) break;
      curr = projects.find((p) => p.id === curr?.parentId);
    }
    return chain;
  }, [note?.projectId, projects]);

  // 사이드바와 100% 동일한 계층 트리 순서로 정렬된 폴더 목록 생성
  const hierarchicalProjects = useMemo(() => {
    const rootProjects: Project[] = [];
    const childrenMap: Record<string, Project[]> = {};

    projects.forEach((p) => {
      if (!p.parentId) {
        rootProjects.push(p);
      } else {
        if (!childrenMap[p.parentId]) childrenMap[p.parentId] = [];
        childrenMap[p.parentId].push(p);
      }
    });

    const result: { project: Project; depth: number }[] = [];

    const traverse = (proj: Project, depth: number) => {
      result.push({ project: proj, depth });
      const children = childrenMap[proj.id] || [];
      children.forEach((child) => traverse(child, depth + 1));
    };

    rootProjects.forEach((root) => traverse(root, 0));

    // 혹시 parentId가 유효하지 않아 누락된 고아 폴더가 있다면 맨 뒤에 안전하게 추가
    projects.forEach((p) => {
      if (!result.some((r) => r.project.id === p.id)) {
        result.push({ project: p, depth: 0 });
      }
    });

    return result;
  }, [projects]);

  const renderMarkdown = (content: string) => {
    try {
      // 마크다운 체크박스(- [ ] 또는 - [x]) 뒤에 띄어쓰기를 빠뜨린 경우에도 정상 체크박스로 렌더링되도록 자동 보정
      let normalizedContent = (content || '').replace(/^(\s*-\s*\[[ xX]\])([^\s\n])/gm, '$1 $2');
      // ==형광펜== 문법 지원
      normalizedContent = normalizedContent.replace(/==([^=\n]+)==/g, '<mark style="background-color: #fef08a; color: #0f172a; padding: 2px 5px; border-radius: 4px;">$1</mark>');
      let rawHtml = marked.parse(normalizedContent, { breaks: true, gfm: true }) as string;

      // 표(table)를 모바일 가로 스크롤 컨테이너로 감싸 좁은 화면에서 컬럼이 찌그러지는 현상 방지
      rawHtml = rawHtml.replace(/<table>([\s\S]*?)<\/table>/g, '<div class="table-scroll-container"><table>$1</table></div>');

      // 마크다운 체크박스를 보기 모드에서도 직접 클릭해 체크/해제할 수 있도록 속성 변환
      let taskCounter = 0;
      rawHtml = rawHtml.replace(/<input\s+([^>]*?)type="checkbox"([^>]*?)>/g, (_match, before, after) => {
        const currentIndex = taskCounter++;
        const cleanBefore = before.replace(/disabled(?:="[^"]*")?/g, '');
        const cleanAfter = after.replace(/disabled(?:="[^"]*")?/g, '');
        return `<input ${cleanBefore}type="checkbox"${cleanAfter} data-task-index="${currentIndex}" class="task-checkbox cursor-pointer">`;
      });

      return { __html: renderAttachmentsInHtml(rawHtml, localAttachments) };
    } catch {
      return { __html: '<p>마크다운을 렌더링할 수 없습니다.</p>' };
    }
  };

  // 분할 모드 타이핑 중 미리보기(PDF 및 마크다운) 잦은 갱신 및 깜빡임 방지 디바운스
  const [debouncedPreviewContent, setDebouncedPreviewContent] = useState(localContent);
  const previewTimerRef = useRef<any>(null);

  useEffect(() => {
    if (viewMode === 'preview') {
      setDebouncedPreviewContent(localContent);
      return;
    }
    if (previewTimerRef.current) {
      clearTimeout(previewTimerRef.current);
    }
    previewTimerRef.current = setTimeout(() => {
      setDebouncedPreviewContent(localContent);
    }, 280);

    return () => {
      if (previewTimerRef.current) {
        clearTimeout(previewTimerRef.current);
      }
    };
  }, [localContent, viewMode]);

  const renderedPreviewHtml = useMemo(() => {
    return renderMarkdown(debouncedPreviewContent);
  }, [debouncedPreviewContent, localAttachments]);

  const previewContainerRef = useRef<HTMLDivElement>(null);
  const previewScrollTopRef = useRef<number>(0);
  const isSyncingScrollRef = useRef<'editor' | 'preview' | null>(null);

  // 분할 뷰(Split View) 양방향 스크롤 실시간 동기화 핸들러
  const handleEditorScroll = (e: React.UIEvent<HTMLTextAreaElement>) => {
    if (viewMode !== 'split') return;
    if (isSyncingScrollRef.current === 'preview') return;

    const editor = e.currentTarget;
    const preview = previewContainerRef.current;
    if (!preview) return;

    isSyncingScrollRef.current = 'editor';
    const editorScrollable = editor.scrollHeight - editor.clientHeight;
    if (editorScrollable > 0) {
      const percentage = editor.scrollTop / editorScrollable;
      const previewScrollable = preview.scrollHeight - preview.clientHeight;
      preview.scrollTop = percentage * previewScrollable;
      previewScrollTopRef.current = preview.scrollTop;
    }

    requestAnimationFrame(() => {
      isSyncingScrollRef.current = null;
    });
  };

  const handlePreviewScroll = (e: React.UIEvent<HTMLDivElement>) => {
    previewScrollTopRef.current = e.currentTarget.scrollTop;
    if (viewMode !== 'split') return;
    if (isSyncingScrollRef.current === 'editor') return;

    const preview = e.currentTarget;
    const editor = textareaRef.current;
    if (!editor) return;

    isSyncingScrollRef.current = 'preview';
    const previewScrollable = preview.scrollHeight - preview.clientHeight;
    if (previewScrollable > 0) {
      const percentage = preview.scrollTop / previewScrollable;
      const editorScrollable = editor.scrollHeight - editor.clientHeight;
      editor.scrollTop = percentage * editorScrollable;
    }

    requestAnimationFrame(() => {
      isSyncingScrollRef.current = null;
    });
  };

  // 메모 변경 시 미리보기 스크롤 상단 리셋
  useEffect(() => {
    previewScrollTopRef.current = 0;
    if (previewContainerRef.current) {
      previewContainerRef.current.scrollTop = 0;
    }
  }, [note?.id]);

  // 미리보기 갱신 시 스크롤 위치 유지 (PDF 임베드 등 리렌더링으로 인한 하단 스크롤 튐 원천 방지)
  useLayoutEffect(() => {
    if (previewContainerRef.current) {
      previewContainerRef.current.scrollTop = previewScrollTopRef.current;
    }
  }, [renderedPreviewHtml]);

  useEffect(() => {
    const timer1 = setTimeout(() => {
      if (previewContainerRef.current) {
        previewContainerRef.current.scrollTop = previewScrollTopRef.current;
      }
    }, 50);
    const timer2 = setTimeout(() => {
      if (previewContainerRef.current) {
        previewContainerRef.current.scrollTop = previewScrollTopRef.current;
      }
    }, 200);
    return () => {
      clearTimeout(timer1);
      clearTimeout(timer2);
    };
  }, [renderedPreviewHtml]);

  // 인라인 PDF 페이지 렌더링 핸들러
  const renderInlinePdf = useRef<(attId: string, customData?: string) => void>(() => {});
  renderInlinePdf.current = (attId: string, customData?: string) => {
    const preview = previewContainerRef.current;
    if (!preview) return;

    const viewer = preview.querySelector<HTMLElement>(`.pdf-inline-viewer[data-att-id="${attId}"]`);
    if (!viewer) return;

    const att = localAttachments[attId];
    const dataToUse = customData || att?.data;
    if (!dataToUse) return;

    const canvas = viewer.querySelector<HTMLCanvasElement>(`#pdf-canvas-${attId}`);
    const loadingBox = viewer.querySelector<HTMLElement>(`#pdf-loading-${attId}`);
    const pageIndicator = viewer.querySelector<HTMLElement>(`#pdf-page-indicator-${attId}`);
    const prevBtn = viewer.querySelector<HTMLButtonElement>('.pdf-btn-prev');
    const nextBtn = viewer.querySelector<HTMLButtonElement>('.pdf-btn-next');

    const currentPage = pdfPagesRef.current.get(attId) || 1;

    loadPdfDocument(dataToUse, attId)
      .then((doc) => {
        const total = doc.numPages;
        if (pageIndicator) pageIndicator.textContent = `${currentPage} / ${total}`;
        if (prevBtn) prevBtn.disabled = currentPage <= 1;
        if (nextBtn) nextBtn.disabled = currentPage >= total;

        if (canvas) {
          const containerWidth = viewer.clientWidth || preview.clientWidth || 600;
          renderPdfPageToCanvas(doc, currentPage, canvas, Math.min(containerWidth, 800))
            .then(() => {
              if (loadingBox) loadingBox.classList.add('hidden');
              canvas.classList.remove('hidden');
            })
            .catch((err) => {
              console.error('인라인 PDF 페이지 렌더링 실패:', err);
              if (loadingBox) {
                const errMsg = err?.message || String(err);
                loadingBox.innerHTML = `<div class="p-3 bg-rose-950/40 border border-rose-500/30 rounded-xl text-center space-y-1.5 max-w-md mx-auto">
                  <div class="text-rose-400 font-semibold text-xs">⚠️ PDF 페이지를 표시하지 못했습니다.</div>
                  <div class="text-[11px] text-slate-400 font-mono break-all">${errMsg}</div>
                  <div class="flex items-center justify-center gap-2 pt-1">
                    <button type="button" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs rounded border border-slate-700" onclick="window.location.reload()">새로고침</button>
                  </div>
                </div>`;
              }
            });
        }
      })
      .catch((err) => {
        console.error('인라인 PDF 렌더링 실패:', err);
        if (loadingBox) {
          const errMsg = err?.message || String(err);
          loadingBox.innerHTML = `<div class="p-3 bg-rose-950/40 border border-rose-500/30 rounded-xl text-center space-y-1.5 max-w-md mx-auto">
            <div class="text-rose-400 font-semibold text-xs">⚠️ PDF 문서를 로드하지 못했습니다.</div>
            <div class="text-[11px] text-slate-400 font-mono break-all">${errMsg}</div>
            <div class="flex items-center justify-center gap-2 pt-1">
              <button type="button" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs rounded border border-slate-700" onclick="window.location.reload()">새로고침</button>
            </div>
          </div>`;
        }
      });
  };

  // 인라인 HWP/HWPX 페이지 렌더링 핸들러
  const renderInlineHwp = useRef<(attId: string, customData?: string) => void>(() => {});
  renderInlineHwp.current = (attId: string, customData?: string) => {
    const preview = previewContainerRef.current;
    if (!preview) return;

    const viewer = preview.querySelector<HTMLElement>(`.hwp-inline-viewer[data-att-id="${attId}"]`);
    if (!viewer) return;

    const att = localAttachments[attId];
    const dataToUse = customData || att?.data;
    if (!dataToUse) return;

    const svgContainer = viewer.querySelector<HTMLElement>(`#hwp-svg-container-${attId}`);
    const loadingBox = viewer.querySelector<HTMLElement>(`#hwp-loading-${attId}`);
    const pageIndicator = viewer.querySelector<HTMLElement>(`#hwp-page-indicator-${attId}`);
    const prevBtn = viewer.querySelector<HTMLButtonElement>('.hwp-btn-prev');
    const nextBtn = viewer.querySelector<HTMLButtonElement>('.hwp-btn-next');

    const currentPage = hwpPagesRef.current.get(attId) || 1;

    loadHwpDocument(dataToUse, attId)
      .then((doc) => {
        const total = doc.pageCount();
        if (pageIndicator) pageIndicator.textContent = `${currentPage} / ${total}`;
        if (prevBtn) prevBtn.disabled = currentPage <= 1;
        if (nextBtn) nextBtn.disabled = currentPage >= total;

        if (svgContainer) {
          const pageSvg = renderHwpPageToSvg(doc, currentPage - 1);
          svgContainer.innerHTML = pageSvg;
          if (loadingBox) loadingBox.classList.add('hidden');
          svgContainer.classList.remove('hidden');
        }
      })
      .catch((err) => {
        console.error('인라인 HWP 렌더링 실패:', err);
        if (loadingBox) {
          loadingBox.innerHTML = '<span class="text-rose-400 font-semibold">⚠️ 한글(HWP) 문서를 로드하지 못했습니다. 상단 저장 버튼을 이용해 주세요.</span>';
        }
      });
  };

  // 미리보기 화면이 갱신되거나 활성화될 때 모든 인라인 PDF & HWP 자동 렌더링
  useEffect(() => {
    if (viewMode !== 'preview' && viewMode !== 'split') return;
    const preview = previewContainerRef.current;
    if (!preview) return;

    const viewers = preview.querySelectorAll<HTMLElement>('.pdf-inline-viewer');
    viewers.forEach((v) => {
      const attId = v.getAttribute('data-att-id');
      if (attId) {
        if (!pdfPagesRef.current.has(attId)) {
          pdfPagesRef.current.set(attId, 1);
        }
        renderInlinePdf.current(attId);
      }
    });

    const hwpViewers = preview.querySelectorAll<HTMLElement>('.hwp-inline-viewer');
    hwpViewers.forEach((v) => {
      const attId = v.getAttribute('data-att-id');
      if (attId) {
        if (!hwpPagesRef.current.has(attId)) {
          hwpPagesRef.current.set(attId, 1);
        }
        renderInlineHwp.current(attId);
      }
    });
  }, [renderedPreviewHtml, viewMode, localAttachments]);

  // 글자 색상 및 형광펜 팔레트 팝업 상태 (훅은 조건문 위에 위치)
  const [showColorPicker, setShowColorPicker] = useState(false);
  const [pickerPosition, setPickerPosition] = useState<{ top: number; left: number } | null>(null);
  const colorPickerRef = useRef<HTMLDivElement>(null);
  const paletteButtonRef = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    if (!showColorPicker) return;
    const handleClickOutside = (e: MouseEvent) => {
      if (
        colorPickerRef.current && 
        !colorPickerRef.current.contains(e.target as Node) &&
        !paletteButtonRef.current?.contains(e.target as Node)
      ) {
        setShowColorPicker(false);
      }
    };
    const handleClose = () => setShowColorPicker(false);
    document.addEventListener('mousedown', handleClickOutside);
    window.addEventListener('resize', handleClose);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      window.removeEventListener('resize', handleClose);
    };
  }, [showColorPicker]);

  if (!note) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center bg-slate-950 text-slate-500 p-6 text-center">
        <Edit3 size={40} className="mb-3 opacity-30 text-indigo-400" />
        <p className="text-sm font-medium">선택된 메모가 없습니다.</p>
        <p className="text-xs text-slate-600 mt-1">좌측 목록에서 메모를 선택하거나 새 메모를 작성해 보세요.</p>
      </div>
    );
  }

  // 마크다운 문법 빠른 삽입 함수
  const insertMarkdown = (prefix: string, suffix: string = '') => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const prevScrollTop = textarea.scrollTop;
    const prevScrollLeft = textarea.scrollLeft;

    const selected = localContent.substring(start, end);
    const replacement = `${prefix}${selected || '내용'}${suffix}`;

    // 실행 취소를 위해 변경 전 스냅샷 저장
    pushHistory(localContent, start);

    const newContent = 
      localContent.substring(0, start) + 
      replacement + 
      localContent.substring(end);

    setLocalContent(newContent);
    debouncedSyncToParent({ content: newContent });

    const newCursorStart = start + prefix.length;
    const newCursorEnd = start + prefix.length + (selected.length || 2);
    pushHistory(newContent, newCursorEnd);

    setTimeout(() => {
      textarea.focus({ preventScroll: true });
      textarea.setSelectionRange(newCursorStart, newCursorEnd);
      textarea.scrollTop = prevScrollTop;
      textarea.scrollLeft = prevScrollLeft;
    }, 0);
  };

  // 색상 팔레트 열기 / 닫기 토글 (버튼 위치 기반 고정 팝업)
  const toggleColorPicker = () => {
    if (!showColorPicker) {
      if (paletteButtonRef.current) {
        const rect = paletteButtonRef.current.getBoundingClientRect();
        const popupWidth = 224; // 14rem (w-56)
        const left = Math.min(window.innerWidth - popupWidth - 12, Math.max(10, rect.left));
        setPickerPosition({ top: rect.bottom + 6, left });
      }
      setShowColorPicker(true);
    } else {
      setShowColorPicker(false);
    }
  };

  // 글자 색상 또는 형광펜 배경 적용 함수
  const applyTextColor = (colorValue: string, isBackground: boolean = false) => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const prevScrollTop = textarea.scrollTop;
    const prevScrollLeft = textarea.scrollLeft;

    const selected = localContent.substring(start, end);
    const targetText = selected || '색상 텍스트';

    let replacement: string;
    if (isBackground) {
      replacement = `<mark style="background-color: ${colorValue}; color: #0f172a; padding: 2px 5px; border-radius: 4px;">${targetText}</mark>`;
    } else {
      replacement = `<span style="color: ${colorValue};">${targetText}</span>`;
    }

    pushHistory(localContent, start);

    const newContent = 
      localContent.substring(0, start) + 
      replacement + 
      localContent.substring(end);

    setLocalContent(newContent);
    debouncedSyncToParent({ content: newContent });

    const openTagLength = isBackground 
      ? `<mark style="background-color: ${colorValue}; color: #0f172a; padding: 2px 5px; border-radius: 4px;">`.length
      : `<span style="color: ${colorValue};">`.length;

    const newCursorStart = start + openTagLength;
    const newCursorEnd = newCursorStart + targetText.length;
    pushHistory(newContent, newCursorEnd);

    setShowColorPicker(false);

    setTimeout(() => {
      textarea.focus({ preventScroll: true });
      textarea.setSelectionRange(newCursorStart, newCursorEnd);
      textarea.scrollTop = prevScrollTop;
      textarea.scrollLeft = prevScrollLeft;
    }, 0);
  };

  type LinePrefixType = 'bullet' | 'ordered' | 'todo' | 'h1' | 'h2' | 'h3' | 'quote';

  // 줄(행) 단위 마크다운 서식 적용 함수 (글머리 기호, 번호 매기기, To-Do, 제목, 인용구 등)
  const applyLinePrefix = (type: LinePrefixType) => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = localContent;
    const prevScrollTop = textarea.scrollTop;
    const prevScrollLeft = textarea.scrollLeft;

    // 실행 취소를 위해 변경 전 스냅샷 저장
    pushHistory(localContent, start);

    // 선택 영역이 걸쳐 있는 모든 라인의 시작과 끝 인덱스 찾기
    const lineStart = text.lastIndexOf('\n', start - 1) + 1;
    let lineEnd = text.indexOf('\n', end);
    if (lineEnd === -1) lineEnd = text.length;

    const selectedBlock = text.substring(lineStart, lineEnd);
    const lines = selectedBlock.split('\n');

    // 기존 접두어 정규식 패턴들
    const todoRegex = /^(\s*)-\s*\[[ xX]\]\s*/;
    const bulletRegex = /^(\s*)[-*+]\s+/;
    const orderedRegex = /^(\s*)\d+\.\s+/;
    const headingRegex = /^(\s*)#{1,6}\s+/;
    const quoteRegex = /^(\s*)>\s+/;

    // 기존 접두어(체크박스, 글머리, 번호, 제목, 인용) 제거 후 { indent, content } 반환
    const stripExistingPrefix = (line: string): { indent: string; content: string } => {
      if (todoRegex.test(line)) {
        const match = line.match(todoRegex)!;
        return { indent: match[1], content: line.slice(match[0].length) };
      }
      if (bulletRegex.test(line)) {
        const match = line.match(bulletRegex)!;
        return { indent: match[1], content: line.slice(match[0].length) };
      }
      if (orderedRegex.test(line)) {
        const match = line.match(orderedRegex)!;
        return { indent: match[1], content: line.slice(match[0].length) };
      }
      if (headingRegex.test(line)) {
        const match = line.match(headingRegex)!;
        return { indent: match[1], content: line.slice(match[0].length) };
      }
      if (quoteRegex.test(line)) {
        const match = line.match(quoteRegex)!;
        return { indent: match[1], content: line.slice(match[0].length) };
      }
      const indentMatch = line.match(/^(\s*)(.*)$/)!;
      return { indent: indentMatch[1], content: indentMatch[2] };
    };

    // 선택된 모든 줄이 이미 해당 타입인지 확인 (전부 이미 해당 타입이면 토글하여 접두어 제거)
    const isAllSameType = lines.every((line) => {
      if (type === 'bullet') return bulletRegex.test(line) && !todoRegex.test(line);
      if (type === 'ordered') return orderedRegex.test(line);
      if (type === 'todo') return todoRegex.test(line);
      if (type === 'h1') return /^(\s*)#\s+/.test(line);
      if (type === 'h2') return /^(\s*)##\s+/.test(line);
      if (type === 'h3') return /^(\s*)###\s+/.test(line);
      if (type === 'quote') return quoteRegex.test(line);
      return false;
    });

    // 번호 매기기('ordered')인 경우, 이전 라인들에 이미 작성된 번호가 있으면 그 번호에 이어서 순차 번호 부여
    let orderedCounter = 1;
    if (type === 'ordered') {
      const textBefore = text.substring(0, lineStart);
      const prevLines = textBefore.split('\n');
      let emptyCount = 0;
      for (let i = prevLines.length - 2; i >= 0; i--) {
        const prev = prevLines[i];
        if (prev.trim() === '') {
          emptyCount++;
          if (emptyCount > 1) break; // 빈 줄이 2개 이상이면 새 목록으로 간주
          continue;
        }
        const match = prev.match(/^\s*(\d+)\.\s+/);
        if (match) {
          orderedCounter = parseInt(match[1], 10) + 1;
        }
        break;
      }
    }
    const newLines = lines.map((line) => {
      const { indent, content } = stripExistingPrefix(line);

      // 이미 전부 같은 타입이면 토글하여 일반 텍스트로 복원
      if (isAllSameType) {
        return `${indent}${content}`;
      }

      switch (type) {
        case 'bullet':
          return `${indent}- ${content}`;
        case 'ordered':
          return `${indent}${orderedCounter++}. ${content}`;
        case 'todo':
          return `${indent}- [ ] ${content}`;
        case 'h1':
          return `${indent}# ${content}`;
        case 'h2':
          return `${indent}## ${content}`;
        case 'h3':
          return `${indent}### ${content}`;
        case 'quote':
          return `${indent}> ${content}`;
        default:
          return line;
      }
    });

    const replacedBlock = newLines.join('\n');
    const newContent = text.substring(0, lineStart) + replacedBlock + text.substring(lineEnd);

    setLocalContent(newContent);
    debouncedSyncToParent({ content: newContent });

    // 커서 위치 스마트 복원 (줄 끝이나 문서 맨 아래로 튀지 않고 기존 상대 위치 정밀 유지)
    const diff = replacedBlock.length - selectedBlock.length;
    let newCursorStart: number;
    let newCursorEnd: number;

    if (start === end) {
      if (lines.length === 1 && lines[0].trim() === '') {
        // 빈 줄이었던 경우 -> 접두어 바로 뒤에 커서 위치 (바로 타이핑 가능)
        newCursorStart = lineStart + newLines[0].length;
      } else {
        // 원래 커서가 있던 위치를 prefix 길이 변화만큼 보정하여 정확히 유지
        newCursorStart = Math.max(lineStart, Math.min(newContent.length, start + diff));
      }
      newCursorEnd = newCursorStart;
    } else {
      newCursorStart = lineStart;
      newCursorEnd = lineStart + replacedBlock.length;
    }

    pushHistory(newContent, newCursorStart);

    setTimeout(() => {
      textarea.focus({ preventScroll: true });
      textarea.setSelectionRange(newCursorStart, newCursorEnd);
      textarea.scrollTop = prevScrollTop;
      textarea.scrollLeft = prevScrollLeft;
    }, 0);
  };

  // 체크박스(- [ ] <-> - [x]) 토글 함수 (현재 커서 라인 또는 선택된 여러 줄 일괄 토글)
  const toggleCheckboxAtCursor = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = localContent;
    const prevScrollTop = textarea.scrollTop;
    const prevScrollLeft = textarea.scrollLeft;

    // 실행 취소를 위해 변경 전 스냅샷 저장
    pushHistory(localContent, start);

    // 선택 영역이 걸쳐 있는 모든 라인의 시작과 끝 인덱스 찾기
    const lineStart = text.lastIndexOf('\n', start - 1) + 1;
    let lineEnd = text.indexOf('\n', end);
    if (lineEnd === -1) lineEnd = text.length;

    const selectedBlock = text.substring(lineStart, lineEnd);
    const lines = selectedBlock.split('\n');

    const newLines = lines.map((line) => {
      // 1. 이미 - [ ] 형태인 경우 -> - [x] 로 변경 (체크, 뒤에 반드시 공백 1칸 보장)
      if (/^(\s*-\s*)\[\s*\](.*)$/.test(line)) {
        return line.replace(/^(\s*-\s*)\[\s*\]\s*/, '$1[x] ');
      }
      // 2. 이미 - [x] 형태인 경우 -> - [ ] 로 변경 (해제, 뒤에 반드시 공백 1칸 보장)
      if (/^(\s*-\s*)\[[xX]\](.*)$/.test(line)) {
        return line.replace(/^(\s*-\s*)\[[xX]\]\s*/, '$1[ ] ');
      }
      // 3. - 글머리 기호만 있는 경우 -> - [ ] 추가
      if (/^(\s*-\s+)(.*)$/.test(line)) {
        return line.replace(/^(\s*-\s+)/, '$1[ ] ');
      }
      // 4. 일반 텍스트 줄인 경우 -> 줄 앞에 - [ ] 추가
      if (line.trim().length > 0) {
        return `- [ ] ${line}`;
      }
      return line;
    });

    const replacedBlock = newLines.join('\n');
    const newContent = text.substring(0, lineStart) + replacedBlock + text.substring(lineEnd);

    setLocalContent(newContent);
    debouncedSyncToParent({ content: newContent });
    pushHistory(newContent, start);

    // 커서 위치 복원
    setTimeout(() => {
      textarea.focus({ preventScroll: true });
      const lengthDiff = replacedBlock.length - selectedBlock.length;
      textarea.setSelectionRange(
        Math.max(lineStart, start),
        Math.max(lineStart, end + lengthDiff)
      );
      textarea.scrollTop = prevScrollTop;
      textarea.scrollLeft = prevScrollLeft;
    }, 0);
  };

  // 들여쓰기 2칸 함수 (아이콘 클릭 및 Tab 키 공용)
  const handleIndent = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = localContent;
    const prevScrollTop = textarea.scrollTop;
    const prevScrollLeft = textarea.scrollLeft;

    pushHistory(localContent, start);

    const lineStart = text.lastIndexOf('\n', start - 1) + 1;
    let lineEnd = text.indexOf('\n', end);
    if (lineEnd === -1) lineEnd = text.length;

    const selectedBlock = text.substring(lineStart, lineEnd);
    const lines = selectedBlock.split('\n');
    const newLines = lines.map((l) => '  ' + l);
    const replaced = newLines.join('\n');

    const newContent = text.substring(0, lineStart) + replaced + text.substring(lineEnd);
    setLocalContent(newContent);
    debouncedSyncToParent({ content: newContent });
    pushHistory(newContent, start + 2);

    setTimeout(() => {
      textarea.focus({ preventScroll: true });
      const added = lines.length * 2;
      textarea.setSelectionRange(start + 2, end + added);
      textarea.scrollTop = prevScrollTop;
      textarea.scrollLeft = prevScrollLeft;
    }, 0);
  };

  // 내어쓰기 2칸 함수 (아이콘 클릭 및 Shift+Tab 키 공용)
  const handleOutdent = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = localContent;
    const prevScrollTop = textarea.scrollTop;
    const prevScrollLeft = textarea.scrollLeft;

    pushHistory(localContent, start);

    const lineStart = text.lastIndexOf('\n', start - 1) + 1;
    let lineEnd = text.indexOf('\n', end);
    if (lineEnd === -1) lineEnd = text.length;

    const selectedBlock = text.substring(lineStart, lineEnd);
    const lines = selectedBlock.split('\n');
    let removedTotal = 0;
    const newLines = lines.map((l) => {
      if (l.startsWith('  ')) {
        removedTotal += 2;
        return l.substring(2);
      }
      if (l.startsWith(' ')) {
        removedTotal += 1;
        return l.substring(1);
      }
      if (l.startsWith('\t')) {
        removedTotal += 1;
        return l.substring(1);
      }
      return l;
    });

    const replaced = newLines.join('\n');
    const newContent = text.substring(0, lineStart) + replaced + text.substring(lineEnd);
    setLocalContent(newContent);
    debouncedSyncToParent({ content: newContent });
    pushHistory(newContent, Math.max(lineStart, start - 2));

    setTimeout(() => {
      textarea.focus({ preventScroll: true });
      textarea.setSelectionRange(
        Math.max(lineStart, start - 2),
        Math.max(lineStart, end - removedTotal)
      );
      textarea.scrollTop = prevScrollTop;
      textarea.scrollLeft = prevScrollLeft;
    }, 0);
  };

  // 실행 취소 (Undo) 함수
  const handleUndo = () => {
    if (historyIndexRef.current > 0) {
      historyIndexRef.current -= 1;
      const target = historyRef.current[historyIndexRef.current];
      isUndoRedoActionRef.current = true;
      setLocalContent(target.content);
      debouncedSyncToParent({ content: target.content });
      setTimeout(() => {
        isUndoRedoActionRef.current = false;
        const textarea = textareaRef.current;
        if (textarea && target.cursorPos !== undefined) {
          textarea.focus({ preventScroll: true });
          textarea.setSelectionRange(target.cursorPos, target.cursorPos);
        }
      }, 0);
    }
  };

  // 다시 실행 (Redo) 함수
  const handleRedo = () => {
    if (historyIndexRef.current < historyRef.current.length - 1) {
      historyIndexRef.current += 1;
      const target = historyRef.current[historyIndexRef.current];
      isUndoRedoActionRef.current = true;
      setLocalContent(target.content);
      debouncedSyncToParent({ content: target.content });
      setTimeout(() => {
        isUndoRedoActionRef.current = false;
        const textarea = textareaRef.current;
        if (textarea && target.cursorPos !== undefined) {
          textarea.focus({ preventScroll: true });
          textarea.setSelectionRange(target.cursorPos, target.cursorPos);
        }
      }, 0);
    }
  };

  // 즉시 저장 및 GitHub 동기화 실행 (아이콘 클릭 및 Ctrl+S 공용)
  const handleImmediateSync = () => {
    if (updateTimerRef.current) {
      clearTimeout(updateTimerRef.current);
    }
    onUpdateNote({ content: localContent, title: localTitle });
    onSync?.();
  };

  // 편집창 & 미리보기 분할 바 마우스 드래그 조절 핸들러
  const handleSplitResizeMouseDown = (e: React.MouseEvent) => {
    e.preventDefault();
    setIsResizingSplit(true);
    const container = splitContainerRef.current;
    if (!container) return;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const rect = container.getBoundingClientRect();
      if (rect.width <= 0) return;
      const rawRatio = ((moveEvent.clientX - rect.left) / rect.width) * 100;
      const clampedRatio = Math.max(15, Math.min(85, rawRatio));
      setSplitRatio(clampedRatio);
    };

    const onMouseUp = () => {
      setIsResizingSplit(false);
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      setSplitRatio((r) => {
        localStorage.setItem('mymemo_editor_split_ratio', String(r));
        return r;
      });
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  // 편집창 & 미리보기 분할 바 터치 드래그 조절 핸들러 (태블릿 / 모바일 지원)
  const handleSplitResizeTouchStart = (_e: React.TouchEvent) => {
    setIsResizingSplit(true);
    const container = splitContainerRef.current;
    if (!container) return;

    const onTouchMove = (moveEvent: TouchEvent) => {
      if (!moveEvent.touches || moveEvent.touches.length === 0) return;
      const clientX = moveEvent.touches[0].clientX;
      const rect = container.getBoundingClientRect();
      if (rect.width <= 0) return;
      const rawRatio = ((clientX - rect.left) / rect.width) * 100;
      const clampedRatio = Math.max(15, Math.min(85, rawRatio));
      setSplitRatio(clampedRatio);
    };

    const onTouchEnd = () => {
      setIsResizingSplit(false);
      window.removeEventListener('touchmove', onTouchMove);
      window.removeEventListener('touchend', onTouchEnd);
      setSplitRatio((r) => {
        localStorage.setItem('mymemo_editor_split_ratio', String(r));
        return r;
      });
    };

    window.addEventListener('touchmove', onTouchMove);
    window.addEventListener('touchend', onTouchEnd);
  };

  // 분할 바 더블 클릭 시 5:5 (50%) 균등 분할로 리셋
  const handleSplitDoubleClick = () => {
    setSplitRatio(50);
    localStorage.setItem('mymemo_editor_split_ratio', '50');
  };

  // 표 서식 템플릿 삽입
  const insertTable = () => {
    const tableTpl = '\n\n| 제목 | 내용 | 상태 |\n| :--- | :--- | :---: |\n| 항목1 | 내용1<br>(줄바꿈 예시) | 완료 |\n\n';
    insertMarkdown(tableTpl, '');
  };

  // 구분선 삽입
  const insertDivider = () => {
    insertMarkdown('\n\n---\n\n', '');
  };

  // 링크 삽입
  const insertLink = () => {
    insertMarkdown('[', '](https://)');
  };

  // 코드 블록 삽입
  const insertCodeBlock = () => {
    insertMarkdown('\n```\n', '\n```\n');
  };

  // 번호 목록 삽입
  const insertOrderedList = () => {
    applyLinePrefix('ordered');
  };

  // PC 작성 편의를 위한 키보드 단축키 핸들러
  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const isCtrlOrCmd = e.ctrlKey || e.metaKey;

    // Ctrl + S: 현재 작성 중인 내용 즉시 확정 저장 후 GitHub 동기화
    if (isCtrlOrCmd && e.key.toLowerCase() === 's') {
      e.preventDefault();
      handleImmediateSync();
      return;
    }

    // Ctrl + Enter 또는 Alt + Enter: 체크박스 토글 (- [ ] <-> - [x])
    if ((isCtrlOrCmd || e.altKey) && e.key === 'Enter') {
      e.preventDefault();
      toggleCheckboxAtCursor();
      return;
    }

    // Shift + Enter: 표(Table) 내부 셀 줄바꿈 (<br> 자동 삽입)
    if (e.key === 'Enter' && e.shiftKey && !isCtrlOrCmd && !e.altKey) {
      if (e.nativeEvent.isComposing || e.keyCode === 229) {
        return; // 한글 조합(IME) 중에는 인터셉트하지 않음
      }

      const textarea = textareaRef.current;
      if (!textarea) return;

      const start = textarea.selectionStart;
      const end = textarea.selectionEnd;
      const text = localContent;
      const lineStart = text.lastIndexOf('\n', start - 1) + 1;
      let lineEnd = text.indexOf('\n', end);
      if (lineEnd === -1) lineEnd = text.length;
      const currentFullLine = text.substring(lineStart, lineEnd);

      // 표(Table) 행 내부인 경우 (| 가 포함된 행) -> <br> 자동 삽입
      if (currentFullLine.includes('|')) {
        e.preventDefault();
        const insertText = '<br>';
        const newContent = text.substring(0, start) + insertText + text.substring(end);
        setLocalContent(newContent);
        debouncedSyncToParent({ content: newContent });
        const nextPos = start + insertText.length;
        pushHistory(newContent, nextPos);
        setTimeout(() => {
          textarea.focus({ preventScroll: true });
          textarea.setSelectionRange(nextPos, nextPos);
        }, 0);
        return;
      }
    }

    // Enter 키 목록(글머리 기호, 번호 매기기, To-Do) 자동 이어쓰기 및 빈 줄에서 탈출
    if (e.key === 'Enter' && !isCtrlOrCmd && !e.altKey && !e.shiftKey) {
      if (e.nativeEvent.isComposing || e.keyCode === 229) {
        return; // 한글 조합(IME) 중에는 엔터 인터셉트하지 않음
      }

      const textarea = textareaRef.current;
      if (textarea && textarea.selectionStart === textarea.selectionEnd) {
        const cursor = textarea.selectionStart;
        const text = localContent;
        const lineStart = text.lastIndexOf('\n', cursor - 1) + 1;
        const currentLine = text.substring(lineStart, cursor);

        // 1. To-Do 항목
        const todoMatch = currentLine.match(/^(\s*)-\s*\[[ xX]\]\s*(.*)$/);
        if (todoMatch) {
          e.preventDefault();
          const indent = todoMatch[1];
          const textAfter = todoMatch[2];
          if (textAfter.trim() === '') {
            // 내용 없는 빈 To-Do에서 엔터 -> To-Do 접두어 제거하고 일반 빈 줄로 복원
            const newContent = text.substring(0, lineStart) + text.substring(cursor);
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            pushHistory(newContent, lineStart);
            setTimeout(() => {
              textarea.focus({ preventScroll: true });
              textarea.setSelectionRange(lineStart, lineStart);
            }, 0);
          } else {
            // 내용이 있는 경우 다음 줄에 '- [ ] ' 자동 생성
            const nextPrefix = `\n${indent}- [ ] `;
            const newContent = text.substring(0, cursor) + nextPrefix + text.substring(cursor);
            const nextCursor = cursor + nextPrefix.length;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            pushHistory(newContent, nextCursor);
            setTimeout(() => {
              textarea.focus({ preventScroll: true });
              textarea.setSelectionRange(nextCursor, nextCursor);
            }, 0);
          }
          return;
        }

        // 2. 글머리 기호 항목 (- 또는 * 또는 +)
        const bulletMatch = currentLine.match(/^(\s*)([-*+])\s+(.*)$/);
        if (bulletMatch) {
          e.preventDefault();
          const indent = bulletMatch[1];
          const bulletChar = bulletMatch[2];
          const textAfter = bulletMatch[3];
          if (textAfter.trim() === '') {
            // 내용 없는 빈 글머리 기호에서 엔터 -> 접두어 제거하고 탈출
            const newContent = text.substring(0, lineStart) + text.substring(cursor);
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            pushHistory(newContent, lineStart);
            setTimeout(() => {
              textarea.focus({ preventScroll: true });
              textarea.setSelectionRange(lineStart, lineStart);
            }, 0);
          } else {
            // 다음 줄에 같은 불릿 기호 추가
            const nextPrefix = `\n${indent}${bulletChar} `;
            const newContent = text.substring(0, cursor) + nextPrefix + text.substring(cursor);
            const nextCursor = cursor + nextPrefix.length;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            pushHistory(newContent, nextCursor);
            setTimeout(() => {
              textarea.focus({ preventScroll: true });
              textarea.setSelectionRange(nextCursor, nextCursor);
            }, 0);
          }
          return;
        }

        // 3. 번호 매기기 목록 (1. , 2. ...)
        const orderedMatch = currentLine.match(/^(\s*)(\d+)\.\s+(.*)$/);
        if (orderedMatch) {
          e.preventDefault();
          const indent = orderedMatch[1];
          const currentNum = parseInt(orderedMatch[2], 10);
          const textAfter = orderedMatch[3];
          if (textAfter.trim() === '') {
            // 내용 없는 번호 목록에서 엔터 -> 번호 제거하고 탈출
            const newContent = text.substring(0, lineStart) + text.substring(cursor);
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            pushHistory(newContent, lineStart);
            setTimeout(() => {
              textarea.focus({ preventScroll: true });
              textarea.setSelectionRange(lineStart, lineStart);
            }, 0);
          } else {
            // 다음 번호 자동 증가 (예: 1. -> 2.)
            const nextPrefix = `\n${indent}${currentNum + 1}. `;
            const newContent = text.substring(0, cursor) + nextPrefix + text.substring(cursor);
            const nextCursor = cursor + nextPrefix.length;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            pushHistory(newContent, nextCursor);
            setTimeout(() => {
              textarea.focus({ preventScroll: true });
              textarea.setSelectionRange(nextCursor, nextCursor);
            }, 0);
          }
          return;
        }
      }
    }

    // Tab 키 처리 (들여쓰기 2칸)
    if (e.key === 'Tab') {
      e.preventDefault();
      if (e.shiftKey) {
        handleOutdent();
      } else {
        handleIndent();
      }
      return;
    }

    if (isCtrlOrCmd) {
      // Ctrl + Z: 실행 취소 (Undo)
      if (!e.shiftKey && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        handleUndo();
        return;
      }
      // Ctrl + Y 또는 Ctrl + Shift + Z: 다시 실행 (Redo)
      if (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z')) {
        e.preventDefault();
        handleRedo();
        return;
      }

      // Ctrl + B: 굵게
      if (e.key.toLowerCase() === 'b') {
        e.preventDefault();
        insertMarkdown('**', '**');
        return;
      }
      // Ctrl + I: 기울임
      if (e.key.toLowerCase() === 'i') {
        e.preventDefault();
        insertMarkdown('*', '*');
        return;
      }
      // Ctrl + K: 링크 삽입
      if (e.key.toLowerCase() === 'k' && !e.shiftKey) {
        e.preventDefault();
        insertMarkdown('[', '](https://)');
        return;
      }
      // Ctrl + Shift + K: 인라인 코드
      if (e.key.toLowerCase() === 'k' && e.shiftKey) {
        e.preventDefault();
        insertMarkdown('`', '`');
        return;
      }
      // Ctrl + Shift + X: 취소선
      if (e.key.toLowerCase() === 'x' && e.shiftKey) {
        e.preventDefault();
        insertMarkdown('~~', '~~');
        return;
      }
      // Ctrl + Shift + T 또는 Alt + T: To-Do 체크리스트 (행 단위 적용/해제)
      if ((e.key.toLowerCase() === 't' && e.shiftKey) || (e.altKey && e.key.toLowerCase() === 't')) {
        e.preventDefault();
        applyLinePrefix('todo');
        return;
      }
      // Ctrl + Shift + L: 글머리 기호 목록 (행 단위 적용/해제)
      if (e.key.toLowerCase() === 'l' && e.shiftKey) {
        e.preventDefault();
        applyLinePrefix('bullet');
        return;
      }
      // Ctrl + Shift + O: 번호 매기기 목록 (행 단위 적용/해제)
      if (e.key.toLowerCase() === 'o' && e.shiftKey) {
        e.preventDefault();
        applyLinePrefix('ordered');
        return;
      }
      // Ctrl + 1: 큰 제목 (# )
      if (e.key === '1') {
        e.preventDefault();
        applyLinePrefix('h1');
        return;
      }
      // Ctrl + 2: 중간 제목 (## )
      if (e.key === '2') {
        e.preventDefault();
        applyLinePrefix('h2');
        return;
      }
      // Ctrl + 3: 소제목 (### )
      if (e.key === '3') {
        e.preventDefault();
        applyLinePrefix('h3');
        return;
      }
      // Ctrl + / 또는 Ctrl + ?: 단축키 안내 창 토글
      if (e.key === '/' || e.key === '?') {
        e.preventDefault();
        setShowShortcutsModal((prev) => !prev);
        return;
      }
    }
  };

  // 커서 위치에 첨부파일을 깔끔한 단일 줄 태그로 삽입하고 data는 note.attachments에 보관
  const insertAttachmentAtCursor = (
    attName: string,
    dataUrl: string,
    mimeType: string = 'image/webp'
  ) => {
    const attId = `att-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
    const safeName = attName.replace(/`/g, "'").replace(/[\[\]]/g, '').trim();
    const updatedAttachments: Record<string, AttachmentItem> = {
      ...localAttachments,
      [attId]: {
        id: attId,
        name: attName,
        type: mimeType,
        data: dataUrl,
      },
    };

    setLocalAttachments(updatedAttachments);

    const isPdf = mimeType === 'application/pdf' || attName.toLowerCase().endsWith('.pdf');
    const isHwp = attName.toLowerCase().endsWith('.hwp') || attName.toLowerCase().endsWith('.hwpx') || mimeType.includes('hwp');
    const insertTag = isPdf
      ? `\n\n[📄 PDF: ${safeName}](attachment:${attId})\n\n`
      : isHwp
      ? `\n\n[📝 한글: ${safeName}](attachment:${attId})\n\n`
      : `\n\n![${safeName}](attachment:${attId})\n\n`;

    const textarea = textareaRef.current;
    if (!textarea) {
      pushHistory(localContent, localContent.length);
      const newContent = `${localContent}${insertTag}`;
      setLocalContent(newContent);
      onUpdateNote({ content: newContent, attachments: updatedAttachments });
      pushHistory(newContent, newContent.length);
      return;
    }

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;

    pushHistory(localContent, start);

    const newContent = 
      localContent.substring(0, start) + 
      insertTag + 
      localContent.substring(end);

    setLocalContent(newContent);
    onUpdateNote({ content: newContent, attachments: updatedAttachments });
    const nextPos = start + insertTag.length;
    pushHistory(newContent, nextPos);

    setTimeout(() => {
      textarea.focus({ preventScroll: true });
      textarea.setSelectionRange(nextPos, nextPos);
    }, 0);
  };

  // 1. 붙여넣기(Paste) 이벤트 핸들러 (Ctrl + V 이미지 캡처 지원)
  const handlePaste = async (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    const items = e.clipboardData?.items;
    if (!items) return;

    for (let i = 0; i < items.length; i++) {
      if (items[i].type.indexOf('image') !== -1) {
        e.preventDefault();
        const blob = items[i].getAsFile();
        if (!blob) return;

        setIsProcessingImage(true);
        try {
          const compressed = await compressImageToDataUrl(blob, 1000, 0.75);
          insertAttachmentAtCursor('캡처 이미지', compressed, 'image/webp');
        } catch (err) {
          console.error('이미지 압축 실패:', err);
          alert('이미지를 처리하는 중 오류가 발생했습니다.');
        } finally {
          setIsProcessingImage(false);
        }
        return;
      }
    }
  };

  // 2. 사진 파일 선택 버튼 핸들러 (모바일/PC 지원)
  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessingImage(true);
    try {
      // 사진 선택기에서 실수로 PDF를 선택했더라도 정상적으로 PDF 처리 지원
      if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
        const reader = new FileReader();
        reader.onload = (event) => {
          const dataUrl = event.target?.result as string;
          if (dataUrl) {
            insertAttachmentAtCursor(file.name, dataUrl, 'application/pdf');
          }
          setIsProcessingImage(false);
          if (fileInputRef.current) fileInputRef.current.value = '';
        };
        reader.onerror = () => {
          alert('PDF 파일을 읽지 못했습니다.');
          setIsProcessingImage(false);
          if (fileInputRef.current) fileInputRef.current.value = '';
        };
        reader.readAsDataURL(file);
        return;
      }

      const compressed = await compressImageToDataUrl(file, 1000, 0.75);
      insertAttachmentAtCursor(file.name.replace(/\.[^/.]+$/, ''), compressed, file.type || 'image/webp');
    } catch (err) {
      console.error('이미지 업로드 실패:', err);
      alert('이미지를 불러오지 못했습니다.');
    } finally {
      setIsProcessingImage(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  // 3. PDF 문서 전용 파일 선택 핸들러 (스마트폰 파일 관리자/다운로드 폴더 연동)
  const handlePdfFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 25 * 1024 * 1024) {
      alert('25MB 이하의 PDF 파일만 첨부할 수 있습니다.');
      if (pdfInputRef.current) pdfInputRef.current.value = '';
      return;
    }

    setIsProcessingPdf(true);
    try {
      const reader = new FileReader();
      reader.onload = (event) => {
        const dataUrl = event.target?.result as string;
        if (dataUrl) {
          insertAttachmentAtCursor(file.name, dataUrl, 'application/pdf');
        }
        setIsProcessingPdf(false);
        if (pdfInputRef.current) pdfInputRef.current.value = '';
      };
      reader.onerror = () => {
        alert('PDF 파일을 읽는 중 오류가 발생했습니다.');
        setIsProcessingPdf(false);
        if (pdfInputRef.current) pdfInputRef.current.value = '';
      };
      reader.readAsDataURL(file);
    } catch (err) {
      console.error('PDF 업로드 실패:', err);
      alert('PDF 파일을 불러오지 못했습니다.');
      setIsProcessingPdf(false);
      if (pdfInputRef.current) pdfInputRef.current.value = '';
    }
  };

  // 4. 한글(HWP/HWPX) 문서 전용 파일 선택 핸들러
  const handleHwpFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 30 * 1024 * 1024) {
      alert('30MB 이하의 한글(HWP) 파일만 첨부할 수 있습니다.');
      if (hwpInputRef.current) hwpInputRef.current.value = '';
      return;
    }

    setIsProcessingHwp(true);
    try {
      const reader = new FileReader();
      reader.onload = (event) => {
        const dataUrl = event.target?.result as string;
        if (dataUrl) {
          insertAttachmentAtCursor(file.name, dataUrl, 'application/x-hwp');
        }
        setIsProcessingHwp(false);
        if (hwpInputRef.current) hwpInputRef.current.value = '';
      };
      reader.onerror = () => {
        alert('한글(HWP) 파일을 읽는 중 오류가 발생했습니다.');
        setIsProcessingHwp(false);
        if (hwpInputRef.current) hwpInputRef.current.value = '';
      };
      reader.readAsDataURL(file);
    } catch (err) {
      console.error('HWP 업로드 실패:', err);
      alert('한글 파일을 불러오지 못했습니다.');
      setIsProcessingHwp(false);
      if (hwpInputRef.current) hwpInputRef.current.value = '';
    }
  };

  // 3. 에디터 본문 드래그 앤 드롭 핸들러
  const handleEditorDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOverEditor(false);

    const files = e.dataTransfer.files;
    if (!files || files.length === 0) return;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      setIsProcessingImage(true);
      try {
        const processed = await processFileForNote(file);
        const textarea = textareaRef.current;
        const insertText = `\n\n${processed.content}\n\n`;

        const mergedAttachments: Record<string, AttachmentItem> = {
          ...localAttachments,
          ...(processed.attachments || {}),
        };

        setLocalAttachments(mergedAttachments);

        if (!textarea) {
          const newContent = `${localContent}${insertText}`;
          setLocalContent(newContent);
          onUpdateNote({ content: newContent, attachments: mergedAttachments });
        } else {
          const start = textarea.selectionStart;
          const end = textarea.selectionEnd;
          const newContent = localContent.substring(0, start) + insertText + localContent.substring(end);
          setLocalContent(newContent);
          onUpdateNote({ content: newContent, attachments: mergedAttachments });
        }
      } catch (err) {
        console.error('파일 드롭 처리 실패:', err);
      } finally {
        setIsProcessingImage(false);
      }
    }
  };

  // 마크다운 미리보기 내 체크박스(- [ ] <-> - [x]) 클릭 토글 지원
  const toggleTaskCheckbox = (taskIndex: number) => {
    let currentIndex = 0;
    const newContent = localContent.replace(/^(\s*-\s*\[)([ xX])(\])/gm, (match, prefix, checkState, suffix) => {
      if (currentIndex === taskIndex) {
        currentIndex++;
        const nextState = checkState === ' ' ? 'x' : ' ';
        return `${prefix}${nextState}${suffix}`;
      }
      currentIndex++;
      return match;
    });

    setLocalContent(newContent);
    onUpdateNote({ content: newContent });
  };

  const handlePreviewClick = (e: React.MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;
    // 1. 체크박스 클릭 처리
    if (target && target.tagName === 'INPUT' && (target as HTMLInputElement).type === 'checkbox') {
      const taskIndex = target.getAttribute('data-task-index');
      if (taskIndex !== null) {
        const idx = parseInt(taskIndex, 10);
        toggleTaskCheckbox(idx);
      }
      return;
    }

    // 1-1. 이미지 [✏️ 편집] 버튼 클릭 시 이미지 편집기 모달 열기
    const editImgBtn = target.closest<HTMLElement>('.img-btn-edit');
    if (editImgBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = editImgBtn.getAttribute('data-att-id');
      const title = editImgBtn.getAttribute('data-file-name') || '첨부 이미지';
      const att = attId ? localAttachmentsRef.current[attId] : null;
      if (att && att.data) {
        setActiveImageEditorModal({
          attId: attId!,
          data: att.data,
          title,
        });
      }
      return;
    }

    // 2. 이미지 클릭 시 확대 및 크기 조절 라이트박스 열기
    if (target && target.tagName === 'IMG') {
      const img = target as HTMLImageElement;
      const attId = img.getAttribute('data-att-id') || '';
      if (img.src) {
        setLightboxImage({
          attId,
          src: img.src,
          alt: img.alt || '첨부 이미지',
        });
      }
      return;
    }

    // 3. PDF 인라인 이전/다음 페이지 버튼 클릭
    const prevPdfBtn = target.closest<HTMLButtonElement>('.pdf-btn-prev');
    if (prevPdfBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = prevPdfBtn.getAttribute('data-att-id');
      if (attId) {
        const curr = pdfPagesRef.current.get(attId) || 1;
        if (curr > 1) {
          pdfPagesRef.current.set(attId, curr - 1);
          renderInlinePdf.current(attId);
        }
      }
      return;
    }

    const nextPdfBtn = target.closest<HTMLButtonElement>('.pdf-btn-next');
    if (nextPdfBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = nextPdfBtn.getAttribute('data-att-id');
      if (attId) {
        const curr = pdfPagesRef.current.get(attId) || 1;
        pdfPagesRef.current.set(attId, curr + 1);
        renderInlinePdf.current(attId);
      }
      return;
    }

    // 4. PDF 캔버스 클릭 또는 [🔍 크게 보기] 버튼 클릭 시 전면 뷰어 모달 열기
    const fullscreenPdfBtn = target.closest<HTMLElement>('.pdf-btn-fullscreen') || target.closest<HTMLCanvasElement>('canvas.pdf-canvas');
    if (fullscreenPdfBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = fullscreenPdfBtn.getAttribute('data-att-id');
      const title = fullscreenPdfBtn.getAttribute('data-file-name') || 'PDF 문서';
      const att = attId ? localAttachmentsRef.current[attId] : null;
      if (att && att.data) {
        const curr = pdfPagesRef.current.get(attId!) || 1;
        setActivePdfModal({
          data: att.data,
          title,
          initialPage: curr,
        });
      }
      return;
    }

    // 4-1. PDF [✏️ 주석/서명] 버튼 클릭 시 PDF 주석 편집기 모달 열기
    const editPdfBtn = target.closest<HTMLElement>('.pdf-btn-edit');
    if (editPdfBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = editPdfBtn.getAttribute('data-att-id');
      const title = editPdfBtn.getAttribute('data-file-name') || 'PDF 문서';
      const att = attId ? localAttachmentsRef.current[attId] : null;
      if (attId && att && att.data) {
        const curr = pdfPagesRef.current.get(attId) || 1;
        setActivePdfEditorModal({
          attId,
          data: att.data,
          title,
          initialPage: curr,
        });
      }
      return;
    }

    // 5. HWP 인라인 이전/다음 페이지 버튼 클릭
    const prevHwpBtn = target.closest<HTMLButtonElement>('.hwp-btn-prev');
    if (prevHwpBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = prevHwpBtn.getAttribute('data-att-id');
      if (attId) {
        const curr = hwpPagesRef.current.get(attId) || 1;
        if (curr > 1) {
          hwpPagesRef.current.set(attId, curr - 1);
          renderInlineHwp.current(attId);
        }
      }
      return;
    }

    const nextHwpBtn = target.closest<HTMLButtonElement>('.hwp-btn-next');
    if (nextHwpBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = nextHwpBtn.getAttribute('data-att-id');
      if (attId) {
        const curr = hwpPagesRef.current.get(attId) || 1;
        hwpPagesRef.current.set(attId, curr + 1);
        renderInlineHwp.current(attId);
      }
      return;
    }

    // 5-1. HWP [✏️ 직접 편집] 버튼 클릭 시 웹 한글 에디터 모달 열기
    const editHwpBtn = target.closest<HTMLElement>('.hwp-btn-edit');
    if (editHwpBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = editHwpBtn.getAttribute('data-att-id');
      const title = editHwpBtn.getAttribute('data-file-name') || '한글 문서';
      if (attId && localAttachments[attId]) {
        setActiveHwpEditorModal({
          attId,
          data: localAttachments[attId].data,
          title,
        });
      }
      return;
    }

    // 6. HWP 뷰어 클릭 또는 [🔍 크게 보기] 클릭 시 전면 모달 열기
    const fullscreenHwpBtn = target.closest<HTMLElement>('.hwp-btn-fullscreen') || target.closest<HTMLElement>('.hwp-svg-container');
    if (fullscreenHwpBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = fullscreenHwpBtn.getAttribute('data-att-id');
      const title = fullscreenHwpBtn.getAttribute('data-file-name') || '한글 문서';
      if (attId && localAttachments[attId]) {
        const curr = hwpPagesRef.current.get(attId) || 1;
        setActiveHwpModal({
          data: localAttachments[attId].data,
          title,
          initialPage: curr,
        });
      }
      return;
    }

    // 7. HWP [📝 메모로 가져오기] 버튼 클릭 시 텍스트 추출 및 본문 반영
    const importHwpBtn = target.closest<HTMLElement>('.hwp-btn-import');
    if (importHwpBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = importHwpBtn.getAttribute('data-att-id');
      if (attId && localAttachments[attId]) {
        loadHwpDocument(localAttachments[attId].data, attId)
          .then((doc) => {
            const text = extractHwpText(doc);
            if (!text.trim()) {
              alert('추출할 수 있는 본문 텍스트가 없습니다.');
              return;
            }
            const insertBlock = `\n\n### 📝 ${localAttachments[attId].name} 본문\n\n${text}\n\n`;
            const newContent = `${localContent}${insertBlock}`;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            alert(`'${localAttachments[attId].name}' 문서의 본문 텍스트(${text.length}자)를 메모에 성공적으로 가져왔습니다!`);
          })
          .catch((err) => {
            console.error('HWP 텍스트 추출 실패:', err);
            alert('한글 문서 텍스트를 추출하지 못했습니다.');
          });
      }
      return;
    }

    // 7-2. PDF [📝 메모로 가져오기] 버튼 클릭 시 텍스트 추출 및 본문 반영
    const importPdfBtn = target.closest<HTMLElement>('.pdf-btn-import');
    if (importPdfBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = importPdfBtn.getAttribute('data-att-id');
      if (attId && localAttachments[attId]) {
        loadPdfDocument(localAttachments[attId].data, attId)
          .then(async (doc) => {
            const text = await extractPdfText(doc);
            if (!text.trim()) {
              alert('추출할 수 있는 텍스트가 없습니다. (스캔 이미지로만 구성된 PDF는 텍스트 레이어가 포함되어 있지 않습니다.)');
              return;
            }
            const insertBlock = `\n\n### 📄 ${localAttachments[attId].name} 본문\n\n${text}\n\n`;
            const newContent = `${localContent}${insertBlock}`;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            alert(`'${localAttachments[attId].name}' 문서의 본문 텍스트(${text.length}자)를 메모에 성공적으로 가져왔습니다!`);
          })
          .catch((err) => {
            console.error('PDF 텍스트 추출 실패:', err);
            alert('PDF 문서 텍스트를 추출하지 못했습니다.');
          });
      }
      return;
    }

    // 8. 링크(<a>) 클릭 처리: PWA 및 모바일/PC 어디서나 확실하게 새 창(새 브라우저 탭)으로 열기
    const anchor = target.closest('a');
    if (anchor && anchor.href) {
      if (anchor.hasAttribute('download')) {
        return; // 다운로드 링크는 기본 브라우저 다운로드 동작 허용
      }
      e.preventDefault();
      window.open(anchor.href, '_blank', 'noopener,noreferrer');
      return;
    }
  };

  // 본문 내 첨부 이미지 크기 조절 (|작게, |보통, |크게, |전체)
  const handleResizeImageInContent = (attId: string, sizeLabel: '작게' | '보통' | '크게' | '전체') => {
    if (!attId) return;
    const regex = new RegExp(`!\\[([^\\]]*?)\\]\\((?:attachment:)?${attId}\\)`, 'g');
    const newContent = localContent.replace(regex, (_match, rawAlt) => {
      const cleanAlt = (rawAlt.split('|')[0] || '').trim() || '이미지';
      return `![${cleanAlt}|${sizeLabel}](attachment:${attId})`;
    });

    setLocalContent(newContent);
    onUpdateNote({ content: newContent });
    setLightboxImage((prev) => (prev ? { ...prev, alt: `${prev.alt.split('|')[0]}|${sizeLabel}` } : null));
  };


  // 소속 폴더 변경 시 제목 접두어([상위-하위]) 자동 갱신
  const handleFolderChange = (newProjectId: string) => {
    if (!note) return;
    const newFolderPath = getFolderHierarchyPath(newProjectId, projects);
    const newTitle = updateTitleWithFolderPath(localTitle || note.title, newFolderPath);

    setLocalTitle(newTitle);
    lastSentTitleRef.current = newTitle;
    onUpdateNote({ 
      projectId: newProjectId, 
      title: newTitle 
    });
  };

  // 마크다운(.md) 파일로 내 컴퓨터에 다운로드 저장
  const handleExportToMarkdown = () => {
    if (!note) return;
    const rawFileName = (localTitle || note.title || '메모').trim();
    // 윈도우 파일 시스템에서 파일명으로 사용할 수 없는 특수문자(\ / : * ? " < > |) 치환
    const safeFileName = rawFileName.replace(/[\\/:*?"<>|]/g, '_');

    // attachment:attId 링크를 실제 Base64 데이터(data:image/..., data:application/pdf 등)로 치환하여 단일 md 파일만으로도 외부 마크다운 뷰어(Obsidian, VS Code 등)에서 이미지가 보이도록 완전 포함
    let exportContent = localContent;
    if (localAttachments && Object.keys(localAttachments).length > 0) {
      exportContent = exportContent.replace(/attachment:([a-zA-Z0-9_-]+)/g, (match, attId) => {
        const att = localAttachments[attId];
        if (att && att.data) {
          return att.data;
        }
        return match;
      });
    }

    const blob = new Blob([exportContent], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${safeFileName}.md`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  return (
    <div 
      className={`flex-1 flex flex-col h-full bg-slate-950 overflow-hidden transition-all ${
        isDragOverEditor ? 'ring-2 ring-indigo-500 ring-inset bg-indigo-950/20' : ''
      }`}
      onDragOver={(e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragOverEditor(true);
      }}
      onDragLeave={(e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragOverEditor(false);
      }}
      onDrop={handleEditorDrop}
    >
      {/* 숨겨진 사진 선택 인풋 (갤러리/카메라) */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileChange}
        accept="image/*"
        className="hidden"
      />

      {/* 숨겨진 PDF 문서 선택 인풋 (스마트폰 파일 관리자/다운로드 폴더) */}
      <input
        type="file"
        ref={pdfInputRef}
        onChange={handlePdfFileChange}
        accept="application/pdf,.pdf"
        className="hidden"
      />

      {/* 숨겨진 한글(HWP/HWPX) 문서 선택 인풋 */}
      <input
        type="file"
        ref={hwpInputRef}
        onChange={handleHwpFileChange}
        accept=".hwp,.hwpx,application/x-hwp,application/haansofthwp,application/vnd.hancom.hwp,application/vnd.hancom.hwpx"
        className="hidden"
      />

      {/* 충돌 경고 배너 */}
      {note.syncStatus === 'conflict' && (
        <div className="bg-amber-500/15 border-b border-amber-500/30 p-2.5 px-4 flex items-center justify-between text-xs text-amber-300">
          <div className="flex items-center gap-2">
            <AlertTriangle size={15} className="text-amber-400 shrink-0" />
            <span>원격 수정 내용과의 충돌이 감지되어 백업 사본이 보관되었습니다.</span>
          </div>
          <button
            onClick={() => onUpdateNote({ syncStatus: 'pending' })}
            className="px-2 py-1 bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold rounded text-[11px]"
          >
            충돌 해결 확인
          </button>
        </div>
      )}

      {/* 상단 툴바 1: 내비게이션 & 빵가루 경로 & 로컬 토글 */}
      <div className="p-2 md:p-3 border-b border-slate-800/80 flex flex-wrap items-center justify-between gap-2 bg-slate-950/70">
        <div className="flex items-center justify-between w-full md:w-auto gap-2 min-w-0">
          <div className="flex items-center gap-1.5 md:gap-2 min-w-0">
            <button
              onClick={onBackToList}
              className="md:hidden p-1.5 -ml-1 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 flex items-center gap-1 text-xs shrink-0"
            >
              <ArrowLeft size={16} />
              <span>목록</span>
            </button>

            {/* 데스크톱: 좌측 패널(사이드바, 메모목록) 접힘 상태일 때 펼치기 토글 버튼 */}
            <div className="hidden md:flex items-center gap-1.5 mr-1 shrink-0">
              {isSidebarCollapsed && onToggleSidebar && (
                <button
                  onClick={onToggleSidebar}
                  className="p-1.5 text-slate-400 hover:text-indigo-300 hover:bg-slate-900 border border-slate-800/90 rounded-lg transition-colors flex items-center gap-1 text-xs shadow-sm"
                  title="사이드바(폴더 목록) 펼치기"
                >
                  <Folder size={13} className="text-indigo-400" />
                  <span className="font-medium text-[11px]">폴더 펼치기</span>
                </button>
              )}
              {isNoteListCollapsed && onToggleNoteList && (
                <button
                  onClick={onToggleNoteList}
                  className="p-1.5 text-slate-400 hover:text-indigo-300 hover:bg-slate-900 border border-slate-800/90 rounded-lg transition-colors flex items-center gap-1 text-xs shadow-sm"
                  title="메모 목록 펼치기"
                >
                  <PanelLeftOpen size={14} className="text-indigo-400" />
                  <span className="font-medium text-[11px]">목록 펼치기</span>
                </button>
              )}
            </div>

            {/* 계층 경로 빵가루 드롭다운 (현재 메모의 소속 폴더 이동) */}
            <div 
              className="flex items-center gap-1.5 bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-300 max-w-[190px] sm:max-w-[240px] md:max-w-md truncate" 
              title="현재 메모의 소속 폴더를 다른 폴더로 이동할 수 있습니다."
            >
              <Folder size={12} className="text-indigo-400 shrink-0" />
              <span className="text-[11px] text-slate-400 font-medium shrink-0 hidden sm:inline">폴더:</span>
              <select
                value={note.projectId}
                onChange={(e) => handleFolderChange(e.target.value)}
                className="bg-transparent text-xs text-slate-200 focus:outline-none cursor-pointer truncate max-w-full font-medium"
                title="클릭하여 다른 폴더를 선택하면 메모가 해당 폴더로 이동합니다."
              >
                {hierarchicalProjects.map(({ project: p, depth }) => {
                  const indent = '　'.repeat(depth) + (depth > 0 ? '└ ' : '');
                  return (
                    <option key={p.id} value={p.id} className="bg-slate-900 text-slate-200">
                      {indent}{p.isLocalOnly ? `🔒 ${p.name}` : p.name}
                    </option>
                  );
                })}
              </select>
            </div>
          </div>

          {/* 모바일 상단 우측 전용: 잘리지 않고 바로 보이는 메모 삭제 버튼 */}
          {onDeleteNote && (
            <button
              onClick={() => onDeleteNote(note.id)}
              className="md:hidden p-1.5 text-slate-400 hover:text-rose-400 hover:bg-rose-950/40 border border-slate-800 hover:border-rose-800/50 rounded-lg transition-colors shrink-0 shadow-sm"
              title="현재 메모 삭제"
            >
              <Trash2 size={16} />
            </button>
          )}
        </div>

        {/* 오른쪽: 기기 전용(로컬) 토글 & 뷰 모드 & PDF/MD & 데스크톱 삭제 */}
        <div className="flex items-center gap-1.5 md:gap-2 overflow-x-auto scrollbar-none py-0.5 max-w-full shrink-0">
          <button
            onClick={() => onUpdateNote({ isLocalOnly: !note.isLocalOnly })}
            className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-medium border transition-all shrink-0 ${
              note.isLocalOnly
                ? 'bg-amber-500/20 text-amber-300 border-amber-500/50 shadow-sm shadow-amber-500/10'
                : 'bg-slate-900 text-slate-400 border-slate-800 hover:text-slate-200'
            }`}
            title={note.isLocalOnly ? '클라우드 동기화 안 됨 (이 기기에만 보관)' : 'GitHub 클라우드에 동기화됨'}
          >
            {note.isLocalOnly ? (
              <>
                <Lock size={13} className="text-amber-400 shrink-0" />
                <span className="font-semibold text-[11px] sm:text-xs">이 기기에만 저장됨</span>
              </>
            ) : (
              <>
                <Unlock size={13} className="text-slate-500 shrink-0" />
                <span className="text-[11px] sm:text-xs">GitHub 동기화 대상</span>
              </>
            )}
          </button>

          <div className="flex items-center bg-slate-900 border border-slate-800 rounded-lg p-0.5 text-slate-400 shrink-0">
            <button
              onClick={() => setViewMode('edit')}
              className={`p-1.5 md:p-1.5 rounded text-xs transition-colors flex items-center gap-1 ${
                viewMode === 'edit' ? 'bg-indigo-600 text-white font-semibold shadow-xs' : 'hover:text-slate-200'
              }`}
              title="편집 모드 (단축키: Alt + 1 또는 Ctrl + M)"
            >
              <Edit3 size={13} />
              <span className="md:hidden text-[11px]">편집</span>
            </button>
            <button
              onClick={() => setViewMode('preview')}
              className={`p-1.5 md:p-1.5 rounded text-xs transition-colors flex items-center gap-1 ${
                viewMode === 'preview' ? 'bg-indigo-600 text-white font-semibold shadow-xs' : 'hover:text-slate-200'
              }`}
              title="보기 모드 (단축키: Alt + 2 또는 Ctrl + M)"
            >
              <Eye size={13} />
              <span className="md:hidden text-[11px]">보기</span>
            </button>
            <button
              onClick={() => setViewMode('split')}
              className={`hidden md:block p-1.5 rounded text-xs transition-colors ${
                viewMode === 'split' ? 'bg-indigo-600 text-white font-semibold' : 'hover:text-slate-200'
              }`}
              title="분할 보기 모드 (단축키: Alt + 3 또는 Ctrl + M)"
            >
              <Columns size={13} />
            </button>
          </div>

          {/* 📄 PDF 내보내기 / 인쇄 버튼 */}
          <button
            onClick={async () => {
              if (isExportingPdf) return;
              setIsExportingPdf(true);
              try {
                const breadcrumbText = folderBreadcrumbs.map((b) => b.name).join(' > ');
                await exportNoteToPdf(
                  { ...note, title: localTitle, content: localContent, attachments: localAttachments },
                  breadcrumbText,
                  localContent
                );
              } catch (e) {
                console.error('PDF 내보내기 실패:', e);
                alert('PDF 내보내기 중 오류가 발생했습니다.');
              } finally {
                setIsExportingPdf(false);
              }
            }}
            disabled={isExportingPdf}
            className="flex items-center gap-1 px-2 py-1 md:px-2.5 bg-slate-900 hover:bg-slate-800 text-slate-300 hover:text-indigo-300 border border-slate-800 rounded-lg text-xs font-medium transition-all shadow-sm active:scale-95 shrink-0 disabled:opacity-50"
            title="현재 메모를 깔끔한 PDF 문서로 내보내기 / 인쇄 (첨부된 PDF/한글 문서 페이지 포함)"
          >
            {isExportingPdf ? (
              <Loader2 size={13} className="animate-spin text-indigo-400" />
            ) : (
              <FileDown size={13} className="text-indigo-400" />
            )}
            <span className="hidden sm:inline">{isExportingPdf ? '생성 중...' : 'PDF 내보내기'}</span>
          </button>

          {/* 📝 MD 파일로 내 컴퓨터에 다운로드 저장 */}
          <button
            onClick={handleExportToMarkdown}
            className="flex items-center gap-1 px-2 py-1 md:px-2.5 bg-slate-900 hover:bg-slate-800 text-slate-300 hover:text-emerald-300 border border-slate-800 rounded-lg text-xs font-medium transition-all shadow-sm active:scale-95 shrink-0"
            title="현재 메모를 마크다운(.md) 파일로 내 컴퓨터에 다운로드 저장 (이미지/첨부파일 Base64 포함)"
          >
            <Download size={13} className="text-emerald-400" />
            <span className="hidden sm:inline">MD 저장</span>
          </button>

          {/* 데스크톱 메모 삭제 버튼 */}
          {onDeleteNote && (
            <button
              onClick={() => onDeleteNote(note.id)}
              className="hidden md:block p-1.5 text-slate-400 hover:text-rose-400 hover:bg-slate-900 border border-slate-800 rounded-lg transition-colors shrink-0"
              title="현재 메모 삭제"
            >
              <Trash2 size={13} />
            </button>
          )}
        </div>
      </div>

      {/* 상단 툴바 2: 마크다운 단축키 툴바 (휴대폰 클릭 지원 및 가로 2줄 레이아웃) */}
      {(viewMode === 'edit' || viewMode === 'split') && (
        <div className="border-b border-slate-800/80 bg-slate-950/70 select-none">
          {/* 1번째 줄: 제목 서식, 텍스트 스타일, 목록 및 들여쓰기 */}
          <div className="px-3 py-1 border-b border-slate-800/40 flex items-center gap-1 overflow-x-auto text-slate-400 text-xs scrollbar-none">
            {/* H1 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyLinePrefix('h1')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="큰 제목 H1 (Ctrl + 1)"
            >
              <Heading1 size={14} />
            </button>
            {/* H2 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyLinePrefix('h2')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="중간 제목 H2 (Ctrl + 2)"
            >
              <Heading2 size={14} />
            </button>
            {/* H3 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyLinePrefix('h3')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="소제목 H3 (Ctrl + 3)"
            >
              <Heading3 size={14} />
            </button>

            <div className="w-[1px] h-3.5 bg-slate-800/80 mx-1 shrink-0" />

            {/* Bold */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => insertMarkdown('**', '**')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="굵게 Bold (Ctrl + B)"
            >
              <Bold size={13} />
            </button>
            {/* Italic */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => insertMarkdown('*', '*')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="기울임 Italic (Ctrl + I)"
            >
              <Italic size={13} />
            </button>
            {/* Strikethrough */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => insertMarkdown('~~', '~~')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="취소선 (Ctrl + Shift + X)"
            >
              <Strikethrough size={13} />
            </button>
            {/* Inline Code */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => insertMarkdown('`', '`')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="인라인 코드 (Ctrl + Shift + K)"
            >
              <Code size={13} />
            </button>

            {/* 글자 색상 및 형광펜 팔레트 버튼 */}
            <button
              ref={paletteButtonRef}
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={toggleColorPicker}
              className={`w-7 h-7 flex items-center justify-center rounded-md transition-colors shrink-0 active:scale-90 ${
                showColorPicker ? 'bg-amber-500/20 text-amber-300 ring-1 ring-amber-500/50' : 'hover:bg-slate-800 text-slate-400 hover:text-amber-300'
              }`}
              title="글자 색상 및 형광펜 배경 (HTML span/mark)"
            >
              <Palette size={13} />
            </button>

            <div className="w-[1px] h-3.5 bg-slate-800/80 mx-1 shrink-0" />

            {/* Bullet List */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyLinePrefix('bullet')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="글머리 기호 목록 (Ctrl + Shift + L)"
            >
              <List size={13} />
            </button>
            {/* Numbered List */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={insertOrderedList}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="번호 매기기 목록 (1. 2. 3.)"
            >
              <ListOrdered size={13} />
            </button>
            {/* Indent (Tab) */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={handleIndent}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="들여쓰기 2칸 (Tab)"
            >
              <Indent size={13} />
            </button>
            {/* Outdent (Shift + Tab) */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={handleOutdent}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="내어쓰기 (Shift + Tab)"
            >
              <Outdent size={13} />
            </button>

            <div className="w-[1px] h-3.5 bg-slate-800/80 mx-1 shrink-0" />

            {/* 실행 취소 (Undo) */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={handleUndo}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-amber-300 text-slate-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="실행 취소 Undo (Ctrl + Z)"
            >
              <Undo2 size={13} />
            </button>
            {/* 다시 실행 (Redo) */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={handleRedo}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-amber-300 text-slate-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="다시 실행 Redo (Ctrl + Y 또는 Ctrl + Shift + Z)"
            >
              <Redo2 size={13} />
            </button>
          </div>

          {/* 2번째 줄: To-Do 체크리스트, 체크토글, 표, 링크, 코드블록, 인용구, 구분선, 사진, 동기화, 안내 */}
          <div className="px-3 py-1 flex items-center gap-1 overflow-x-auto text-slate-400 text-xs scrollbar-none">
            {/* To-Do 할 일 추가 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyLinePrefix('todo')}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-indigo-300 text-indigo-400 bg-indigo-950/40 border border-indigo-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="To-Do 체크리스트 추가 (Ctrl + Shift + T)"
            >
              <CheckSquare size={12} />
              <span>To-Do</span>
            </button>
            {/* 체크박스 토글 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={toggleCheckboxAtCursor}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-emerald-300 text-emerald-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="체크박스 체크 / 해제 토글 (Ctrl + Enter)"
            >
              <CheckCheck size={14} />
            </button>

            <div className="w-[1px] h-3.5 bg-slate-800/80 mx-1 shrink-0" />

            {/* 표 (Table) */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={insertTable}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-cyan-300 text-cyan-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="마크다운 표(Table) 서식 삽입"
            >
              <Table size={13} />
            </button>
            {/* 링크 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={insertLink}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-sky-300 text-sky-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="하이퍼링크 삽입 (Ctrl + K)"
            >
              <Link2 size={13} />
            </button>
            {/* 코드 블록 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={insertCodeBlock}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-purple-300 text-purple-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="코드 블록(```) 삽입"
            >
              <FileCode size={13} />
            </button>
            {/* 인용구 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyLinePrefix('quote')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-amber-300 text-amber-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="인용구(>) 삽입"
            >
              <Quote size={13} />
            </button>
            {/* 구분선 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={insertDivider}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="구분선(---) 삽입"
            >
              <Minus size={13} />
            </button>

            <div className="w-[1px] h-3.5 bg-slate-800/80 mx-1 shrink-0" />

            {/* 사진 첨부 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => fileInputRef.current?.click()}
              disabled={isProcessingImage}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-indigo-300 text-indigo-400 bg-indigo-500/10 border border-indigo-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="커서 위치에 사진/이미지 첨부 (또는 복사 후 Ctrl + V)"
            >
              {isProcessingImage ? (
                <Loader2 size={12} className="animate-spin text-indigo-400" />
              ) : (
                <ImageIcon size={12} />
              )}
              <span className="hidden xs:inline">사진</span>
            </button>

            {/* PDF 문서 첨부 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => pdfInputRef.current?.click()}
              disabled={isProcessingPdf}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-rose-300 text-rose-400 bg-rose-500/10 border border-rose-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="커서 위치에 PDF 문서 파일 첨부 (휴대폰 파일 관리자 / 내 파일)"
            >
              {isProcessingPdf ? (
                <Loader2 size={12} className="animate-spin text-rose-400" />
              ) : (
                <FileText size={12} />
              )}
              <span className="hidden xs:inline">PDF</span>
            </button>

            {/* 한글(HWP) 문서 첨부 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => hwpInputRef.current?.click()}
              disabled={isProcessingHwp}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-blue-300 text-blue-400 bg-blue-500/10 border border-blue-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="커서 위치에 한글(HWP, HWPX) 문서 파일 첨부"
            >
              {isProcessingHwp ? (
                <Loader2 size={12} className="animate-spin text-blue-400" />
              ) : (
                <FileText size={12} className="text-blue-400" />
              )}
              <span className="hidden xs:inline">한글</span>
            </button>

            {/* GitHub 동기화 버튼 (Ctrl + S) */}
            {onSync && (
              <button
                type="button"
                onMouseDown={(e) => e.preventDefault()}
                onClick={handleImmediateSync}
                disabled={isSyncing}
                className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-emerald-300 text-emerald-400 bg-emerald-500/10 border border-emerald-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
                title="메모 즉시 저장 및 GitHub 동기화 (Ctrl + S)"
              >
                {isSyncing ? (
                  <Loader2 size={12} className="animate-spin text-emerald-400" />
                ) : (
                  <CloudUpload size={13} />
                )}
                <span className="hidden xs:inline">동기화</span>
              </button>
            )}

            {/* ⌨️ 단축키 안내 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setShowShortcutsModal(true)}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 text-slate-400 hover:text-indigo-400 rounded-md transition-colors ml-auto shrink-0"
              title="PC 마크다운 단축키 안내 (Ctrl + /)"
            >
              <Keyboard size={13} />
            </button>
          </div>
        </div>
      )}

      {/* 글자 색상 및 형광펜 팔레트 플로팅 팝업 (툴바의 overflow-x-auto에 잘리지 않도록 fixed 렌더링) */}
      {showColorPicker && pickerPosition && (
        <div 
          ref={colorPickerRef}
          style={{ top: `${pickerPosition.top}px`, left: `${pickerPosition.left}px` }}
          className="fixed z-50 bg-slate-900 border border-slate-700/90 rounded-2xl p-3 shadow-2xl w-60 animate-in fade-in zoom-in-95 backdrop-blur-md"
          onMouseDown={(e) => e.preventDefault()}
        >
          <div className="text-[11px] font-bold text-slate-300 mb-2 flex items-center justify-between">
            <span className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-amber-400"></span>
              글자 색상
            </span>
            <span className="text-[10px] text-slate-500 font-normal">선택 텍스트에 적용</span>
          </div>
          <div className="grid grid-cols-7 gap-1.5 mb-3">
            {TEXT_COLORS.map((c) => (
              <button
                key={c.name}
                type="button"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => applyTextColor(c.value, false)}
                className="w-6 h-6 rounded-lg border border-slate-700/80 hover:scale-115 transition-transform shadow-sm cursor-pointer flex items-center justify-center"
                style={{ backgroundColor: c.value }}
                title={`${c.name} 글자색`}
              />
            ))}
          </div>

          <div className="text-[11px] font-bold text-slate-300 mb-2 flex items-center justify-between border-t border-slate-800 pt-2.5">
            <span className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-yellow-300"></span>
              형광펜 배경색
            </span>
            <span className="text-[10px] text-slate-500 font-normal">==문법 지원==</span>
          </div>
          <div className="grid grid-cols-4 gap-1.5">
            {HIGHLIGHT_COLORS.map((c) => (
              <button
                key={c.name}
                type="button"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => applyTextColor(c.value, true)}
                className="py-1 rounded-md text-[11px] font-bold text-slate-950 border border-slate-600/40 hover:scale-105 transition-transform cursor-pointer shadow-xs text-center"
                style={{ backgroundColor: c.value }}
                title={`${c.name} 형광펜`}
              >
                {c.name}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* 제목 입력 영역 */}
      <div className="px-6 pt-4 pb-2">
        <input
          type="text"
          value={localTitle}
          onChange={handleTitleChange}
          placeholder="메모 제목을 입력하세요..."
          className="w-full bg-transparent text-xl md:text-2xl font-bold text-slate-100 placeholder-slate-600 focus:outline-none"
        />
      </div>

      {/* 본문 영역: 편집창 / 미리보기 / 분할(드래그 너비 조절 가능) */}
      <div 
        ref={splitContainerRef}
        className={`flex-1 flex overflow-hidden p-4 md:px-6 relative ${
          viewMode === 'split' ? 'gap-0' : 'gap-6'
        } ${isResizingSplit ? 'select-none cursor-col-resize' : ''}`}
      >
        {(viewMode === 'edit' || viewMode === 'split') && (
          <div 
            style={viewMode === 'split' ? { width: `${splitRatio}%` } : undefined}
            className={`${viewMode === 'split' ? 'shrink-0 pr-3' : 'flex-1'} flex flex-col h-full overflow-hidden`}
          >
            <textarea
              ref={textareaRef}
              value={localContent}
              onChange={handleContentChange}
              onKeyDown={handleKeyDown}
              onPaste={handlePaste}
              onScroll={handleEditorScroll}
              placeholder="마크다운 문법으로 자유롭게 작성해 보세요... 이미지는 Ctrl + V 로 바로 붙여넣을 수 있습니다!"
              className="w-full flex-1 bg-transparent text-slate-200 text-sm md:text-base leading-relaxed resize-none focus:outline-none font-mono placeholder-slate-600 scrollbar-thin"
              spellCheck={false}
            />
          </div>
        )}

        {/* 편집창과 미리보기 사이 좌우 조절 드래그 바 (분할 모드에서만 표시) */}
        {viewMode === 'split' && (
          <div
            onMouseDown={handleSplitResizeMouseDown}
            onTouchStart={handleSplitResizeTouchStart}
            onDoubleClick={handleSplitDoubleClick}
            title="마우스로 좌우 드래그하여 분할 너비 조절 (더블 클릭 시 5:5 초기화)"
            className={`group relative flex items-center justify-center w-4 -mx-2 cursor-col-resize select-none shrink-0 z-30 transition-colors ${
              isResizingSplit ? 'bg-indigo-600/20' : 'hover:bg-slate-800/40'
            }`}
          >
            {/* 세로 분할 선 */}
            <div className={`w-[1px] h-full transition-colors ${
              isResizingSplit ? 'bg-indigo-500' : 'bg-slate-800/90 group-hover:bg-indigo-500/70'
            }`} />

            {/* 손잡이 그립 핀 (가운데 점 3개) */}
            <div className={`absolute top-1/2 -translate-y-1/2 w-4 h-8 rounded-full flex flex-col items-center justify-center gap-1 border shadow-md transition-all ${
              isResizingSplit 
                ? 'bg-indigo-600 border-indigo-400 text-white scale-110 shadow-indigo-500/30' 
                : 'bg-slate-900 border-slate-700/80 group-hover:border-indigo-400 group-hover:bg-slate-800 text-slate-500 group-hover:text-indigo-400'
            }`}>
              <div className="w-1.5 h-[1.5px] bg-current rounded-full" />
              <div className="w-1.5 h-[1.5px] bg-current rounded-full" />
              <div className="w-1.5 h-[1.5px] bg-current rounded-full" />
            </div>
          </div>
        )}

        {(viewMode === 'preview' || viewMode === 'split') && (
          <div 
            ref={previewContainerRef}
            onScroll={handlePreviewScroll}
            style={viewMode === 'split' ? { width: `calc(${100 - splitRatio}% - 0px)` } : undefined}
            className={`flex-1 h-full overflow-y-auto markdown-preview prose prose-invert max-w-none ${
              viewMode === 'split' ? 'pl-3 pr-2' : 'pr-2'
            } ${isResizingSplit ? 'pointer-events-none' : ''}`}
            onClick={handlePreviewClick}
          >
            <div dangerouslySetInnerHTML={renderedPreviewHtml} />
          </div>
        )}
      </div>

      {/* 모바일 보기 모드일 때 우측 하단 빠른 편집 플로팅 버튼 (FAB) */}
      {viewMode === 'preview' && (
        <button
          type="button"
          onClick={() => setViewMode('edit')}
          className="md:hidden fixed bottom-6 right-6 px-4 py-2.5 bg-indigo-600 hover:bg-indigo-500 active:scale-95 text-white rounded-full shadow-2xl shadow-indigo-950/80 flex items-center gap-1.5 z-40 transition-all font-semibold text-xs border border-indigo-400/40"
          title="편집 모드로 전환"
        >
          <Edit3 size={15} />
          <span>편집</span>
        </button>
      )}

      {/* 이미지 확대 및 크기 조절 모달 (Lightbox) */}
      {lightboxImage && (
        <div 
          className="fixed inset-0 z-50 flex flex-col items-center justify-center p-3 md:p-6 bg-black/85 backdrop-blur-md animate-in fade-in"
          onClick={() => setLightboxImage(null)}
        >
          {/* 상단 컨트롤 바 */}
          <div 
            className="w-full max-w-xl bg-slate-900 border border-slate-700/80 rounded-2xl p-3 mb-3 shadow-2xl flex flex-wrap items-center justify-between gap-2 z-10"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-2 min-w-0">
              <ImageIcon size={16} className="text-indigo-400 shrink-0" />
              <span className="text-xs font-semibold text-slate-200 truncate">{lightboxImage.alt.split('|')[0]}</span>
            </div>

            {/* 크기 조절 원클릭 버튼 그룹 */}
            {lightboxImage.attId && (
              <div className="flex items-center gap-1 bg-slate-950 p-1 rounded-xl border border-slate-800">
                <span className="text-[11px] text-slate-400 px-1 font-medium">크기:</span>
                {(['작게', '보통', '크게', '전체'] as const).map((size) => {
                  const isCurrent = lightboxImage.alt.includes(`|${size}`);
                  return (
                    <button
                      key={size}
                      type="button"
                      onClick={() => handleResizeImageInContent(lightboxImage.attId, size)}
                      className={`px-2.5 py-1 text-xs font-semibold rounded-lg transition-all active:scale-95 ${
                        isCurrent
                          ? 'bg-indigo-600 text-white ring-1 ring-indigo-400 shadow-sm'
                          : 'bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white'
                      }`}
                      title={`본문에서 이 이미지 크기를 '${size}'으로 변경 (가로/세로 동시 조절)`}
                    >
                      {size}
                    </button>
                  );
                })}
              </div>
            )}

            {/* 이미지 직접 편집 버튼 */}
            {lightboxImage.attId && (
              <button
                type="button"
                onClick={() => {
                  const att = localAttachmentsRef.current[lightboxImage.attId];
                  if (att && att.data) {
                    setActiveImageEditorModal({
                      attId: lightboxImage.attId,
                      data: att.data,
                      title: lightboxImage.alt.split('|')[0] || '첨부 이미지',
                    });
                    setLightboxImage(null);
                  }
                }}
                className="px-2.5 py-1 text-xs font-semibold rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white flex items-center gap-1 shadow-sm transition-all"
                title="이미지 편집기 열기 (그리기, 자르기, 영역 이동/지우기, 모자이크, 회전)"
              >
                <Edit3 size={13} />
                <span>편집</span>
              </button>
            )}

            {/* 닫기 버튼 */}
            <button
              type="button"
              onClick={() => setLightboxImage(null)}
              className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors ml-auto"
              title="닫기"
            >
              <X size={18} />
            </button>
          </div>

          {/* 원본 이미지 뷰어 */}
          <div 
            className="relative max-w-full max-h-[75vh] flex items-center justify-center overflow-auto rounded-2xl shadow-2xl"
            onClick={(e) => e.stopPropagation()}
          >
            <img
              src={lightboxImage.src}
              alt={lightboxImage.alt}
              className="max-w-full max-h-[75vh] object-contain rounded-xl border border-slate-700/60"
            />
          </div>
        </div>
      )}

      {/* PDF 전면 뷰어 모달 (모바일/PC 어디서나 다운로드 없이 고해상도 열람 & 줌) */}
      {activePdfModal && (
        <PdfViewerModal
          isOpen={true}
          onClose={() => setActivePdfModal(null)}
          data={activePdfModal.data}
          title={activePdfModal.title}
          initialPage={activePdfModal.initialPage}
          onOpenEditor={() => {
            const attId = Object.keys(localAttachments).find(
              (id) => localAttachments[id].data === activePdfModal.data
            ) || '';
            setActivePdfEditorModal({
              attId,
              data: activePdfModal.data,
              title: activePdfModal.title,
              initialPage: activePdfModal.initialPage || 1,
            });
            setActivePdfModal(null);
          }}
        />
      )}

      {/* PDF 주석 및 서명 직접 편집기 모달 */}
      {activePdfEditorModal && (
        <PdfEditorModal
          isOpen={true}
          onClose={() => setActivePdfEditorModal(null)}
          data={activePdfEditorModal.data}
          title={activePdfEditorModal.title}
          attId={activePdfEditorModal.attId}
          initialPage={activePdfEditorModal.initialPage}
          onSave={async (updatedDataUrl) => {
            const attId = activePdfEditorModal.attId;
            if (attId) {
              // 1. 기존 PDF 캐시 및 Blob URL 캐시 즉시 무효화 (새 파일 재로드 보장)
              invalidatePdfCache(attId);
              invalidateAttachmentBlobUrl(attId);

              // 2. 첨부파일 목록 상태 업데이트 및 상위 메모 저장
              const updatedAtts = {
                ...localAttachmentsRef.current,
                [attId]: {
                  ...(localAttachmentsRef.current[attId] || {}),
                  data: updatedDataUrl,
                },
              };
              setLocalAttachments(updatedAtts);
              localAttachmentsRef.current = updatedAtts;
              onUpdateNote({ attachments: updatedAtts });

              // 3. 편집 모달의 내부 데이터 상태도 최신으로 동기화
              setActivePdfEditorModal((prev) => prev ? { ...prev, data: updatedDataUrl } : null);

              // 4. 인라인 뷰어에 새 데이터 URL을 즉시 전달하여 재로드 & 재렌더링
              renderInlinePdf.current(attId, updatedDataUrl);

              // 5. 마크다운 HTML 재계산 및 DOM 갱신 후 한 번 더 보장 렌더링
              setTimeout(() => {
                renderInlinePdf.current(attId, updatedDataUrl);
              }, 120);
            }
          }}
        />
      )}

      {/* 한글(HWP/HWPX) 전면 뷰어 모달 (모바일/PC 어디서나 열람 & 텍스트 추출 & 직접 편집 연동) */}
      {activeHwpModal && (
        <HwpViewerModal
          isOpen={true}
          onClose={() => setActiveHwpModal(null)}
          data={activeHwpModal.data}
          title={activeHwpModal.title}
          initialPage={activeHwpModal.initialPage}
          onOpenEditor={() => {
            const attId = Object.keys(localAttachments).find(
              (id) => localAttachments[id].data === activeHwpModal.data
            ) || '';
            setActiveHwpEditorModal({
              attId,
              data: activeHwpModal.data,
              title: activeHwpModal.title,
            });
            setActiveHwpModal(null);
          }}
        />
      )}

      {/* 한글(HWP/HWPX) 웹 직접 편집기 모달 (문서 원본 직접 수정 및 메모 자동 저장) */}
      {activeHwpEditorModal && (
        <HwpEditorModal
          isOpen={true}
          onClose={() => setActiveHwpEditorModal(null)}
          data={activeHwpEditorModal.data}
          title={activeHwpEditorModal.title}
          attId={activeHwpEditorModal.attId}
          onSave={async (updatedDataUrl) => {
            const attId = activeHwpEditorModal.attId;
            if (attId) {
              // 1. 기존 HWP 캐시 및 Blob URL 캐시 즉시 무효화 (새 파일 재로드 보장)
              invalidateHwpCache(attId);
              invalidateAttachmentBlobUrl(attId);

              // 2. 첨부파일 목록 상태 업데이트 및 상위 메모 저장
              const updatedAtts = {
                ...localAttachmentsRef.current,
                [attId]: {
                  ...(localAttachmentsRef.current[attId] || {}),
                  data: updatedDataUrl,
                },
              };
              setLocalAttachments(updatedAtts);
              localAttachmentsRef.current = updatedAtts;
              onUpdateNote({ attachments: updatedAtts });

              // 3. 편집 모달의 내부 데이터 상태도 최신으로 동기화
              setActiveHwpEditorModal((prev) => prev ? { ...prev, data: updatedDataUrl } : null);

              // 4. 인라인 뷰어에 새 데이터 URL을 즉시 전달하여 재로드 & 재렌더링
              renderInlineHwp.current(attId, updatedDataUrl);

              // 5. 마크다운 HTML 재계산 및 DOM 갱신 후 한 번 더 보장 렌더링
              setTimeout(() => {
                renderInlineHwp.current(attId, updatedDataUrl);
              }, 120);
            }
          }}
        />
      )}

      {/* 첨부 이미지 웹 직접 편집기 모달 */}
      {activeImageEditorModal && (
        <ImageEditorModal
          isOpen={true}
          onClose={() => setActiveImageEditorModal(null)}
          data={activeImageEditorModal.data}
          title={activeImageEditorModal.title}
          attId={activeImageEditorModal.attId}
          onSave={async (updatedDataUrl) => {
            const attId = activeImageEditorModal.attId;
            if (attId) {
              // 1. Blob URL 캐시 무효화 (새 이미지 즉시 반영)
              invalidateAttachmentBlobUrl(attId);

              // 2. 첨부파일 목록 상태 업데이트 및 상위 메모 저장
              const updatedAtts = {
                ...localAttachmentsRef.current,
                [attId]: {
                  ...(localAttachmentsRef.current[attId] || {}),
                  data: updatedDataUrl,
                },
              };
              setLocalAttachments(updatedAtts);
              localAttachmentsRef.current = updatedAtts;
              onUpdateNote({ attachments: updatedAtts });

              // 3. 편집 모달 내부 데이터 동기화
              setActiveImageEditorModal((prev) => prev ? { ...prev, data: updatedDataUrl } : null);
            }
          }}
        />
      )}

      {/* 단축키 안내 팝업 모달 */}
      {showShortcutsModal && (
        <div 
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4 animate-in fade-in"
          onClick={() => setShowShortcutsModal(false)}
        >
          <div 
            className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-md shadow-2xl p-5 text-slate-200"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between pb-3 border-b border-slate-800">
              <div className="flex items-center gap-2">
                <Keyboard size={18} className="text-indigo-400" />
                <h3 className="font-bold text-sm text-slate-100">PC 마크다운 단축키 안내</h3>
              </div>
              <button
                onClick={() => setShowShortcutsModal(false)}
                className="p-1 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 transition-colors"
              >
                <X size={16} />
              </button>
            </div>
            
            <div className="py-3 space-y-2 text-xs">
              <div className="grid grid-cols-2 gap-2">
                <div className="flex items-center justify-between p-2 bg-indigo-950/40 rounded-lg border border-indigo-500/30 col-span-2">
                  <span className="text-indigo-300 font-semibold">보기 모드 순환 (편집 ➔ 미리보기 ➔ 분할)</span>
                  <kbd className="px-2 py-0.5 bg-indigo-600 border border-indigo-400 text-white rounded font-mono text-[11px] font-bold">Ctrl + M / Alt + V</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-indigo-950/40 rounded-lg border border-indigo-500/30 col-span-2">
                  <span className="text-indigo-300 font-semibold">모드 직접 전환: 편집 / 미리보기 / 분할</span>
                  <div className="flex gap-1">
                    <kbd className="px-1.5 py-0.5 bg-indigo-800 border border-indigo-400 text-white rounded font-mono text-[10px] font-bold">Alt+1</kbd>
                    <kbd className="px-1.5 py-0.5 bg-indigo-800 border border-indigo-400 text-white rounded font-mono text-[10px] font-bold">Alt+2</kbd>
                    <kbd className="px-1.5 py-0.5 bg-indigo-800 border border-indigo-400 text-white rounded font-mono text-[10px] font-bold">Alt+3</kbd>
                  </div>
                </div>
                <div className="flex items-center justify-between p-2 bg-indigo-950/40 rounded-lg border border-indigo-500/30 col-span-2">
                  <span className="text-indigo-300 font-semibold">GitHub 동기화 (클라우드 저장)</span>
                  <kbd className="px-2 py-0.5 bg-indigo-600 border border-indigo-400 text-white rounded font-mono text-[11px] font-bold">Ctrl + S</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-indigo-950/40 rounded-lg border border-indigo-500/30 col-span-2">
                  <span className="text-indigo-300 font-semibold">체크박스 체크 / 해제 토글</span>
                  <kbd className="px-2 py-0.5 bg-indigo-600 border border-indigo-400 text-white rounded font-mono text-[11px] font-bold">Ctrl + Enter</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-indigo-950/40 rounded-lg border border-indigo-500/30 col-span-2">
                  <span className="text-indigo-300 font-semibold">표(Table) 셀 내부 줄바꿈</span>
                  <div className="flex items-center gap-1.5">
                    <kbd className="px-2 py-0.5 bg-indigo-600 border border-indigo-400 text-white rounded font-mono text-[11px] font-bold">Shift + Enter</kbd>
                    <span className="text-slate-400 text-[10px]">(또는 &lt;br&gt;)</span>
                  </div>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">실행 취소 (Undo)</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + Z</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">다시 실행 (Redo)</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + Y</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">굵게 (Bold)</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + B</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">기울임 (Italic)</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + I</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">링크 걸기</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + K</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">인라인 코드</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl+Shift+K</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">To-Do 체크</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl+Shift+T</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">글머리 목록</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl+Shift+L</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">번호 매기기 목록</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl+Shift+O</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">취소선</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl+Shift+X</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">큰 제목 (H1)</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + 1</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">중간 제목 (H2)</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + 2</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">소제목 (H3)</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + 3</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">단축키 안내 창</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + /</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">들여쓰기 2칸</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Tab</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80">
                  <span className="text-slate-400">들여쓰기 취소</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Shift + Tab</kbd>
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-950/60 rounded-lg border border-slate-800/80 col-span-2">
                  <span className="text-slate-400">클립보드 스크린샷 붙여넣기</span>
                  <kbd className="px-1.5 py-0.5 bg-slate-800 border border-slate-700 rounded font-mono text-[11px] text-indigo-300">Ctrl + V</kbd>
                </div>
              </div>
            </div>

            <div className="pt-2 text-right">
              <button
                onClick={() => setShowShortcutsModal(false)}
                className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold"
              >
                닫기
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
