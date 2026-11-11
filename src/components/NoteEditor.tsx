import React, { useState, useRef, useEffect, useLayoutEffect, useMemo } from 'react';
import type { Note, Project } from '../types';
import { marked } from 'marked';
import { compressImageToDataUrl } from '../utils/image';
import { processFileForNote } from '../utils/file';
import { cleanLegacyAttachments, renderAttachmentsInHtml, invalidateAttachmentBlobUrl } from '../utils/attachment';
import { exportNoteToPdf } from '../utils/exportPdf';
import { getFolderHierarchyPath, updateTitleWithFolderPath, stripFolderPrefixFromTitle } from '../utils/folderPath';
import { PdfViewerModal } from './PdfViewerModal';
import { PdfEditorModal } from './PdfEditorModal';
import { ImageEditorModal } from './ImageEditorModal';
import { loadPdfDocument, renderPdfPageToCanvas, extractPdfText, invalidatePdfCache } from '../utils/pdfViewer';
import { HwpViewerModal } from './HwpViewerModal';
import { HwpEditorModal } from './HwpEditorModal';
import { loadHwpDocument, renderHwpPageToSvg, extractHwpText, invalidateHwpCache } from '../utils/hwpViewer';
import { ExcelViewerModal } from './ExcelViewerModal';
import { ExcelEditorModal } from './ExcelEditorModal';
import { parseExcelWorkbook, sheetRowsToMarkdownTable, invalidateExcelCache, generateExcelDataUrl, buildMergeMap, getColumnLabel } from '../utils/excelViewer';
import { DocxViewerModal } from './DocxViewerModal';
import { DocxEditorModal } from './DocxEditorModal';
import { renderDocxToContainer, extractDocxTextAndHtml, invalidateDocxCache } from '../utils/docxViewer';
import { PptxViewerModal } from './PptxViewerModal';
import { PptxEditorModal } from './PptxEditorModal';
import { extractPptxTextAndMarkdown, invalidatePptxCache } from '../utils/pptxViewer';
import { HexEditorModal } from './HexEditorModal';
import { SlashCommandMenu, getSlashCommands, type SlashCommandItem } from './SlashCommandMenu';
import { TemplateModal } from './TemplateModal';
import { getNoteTemplates, type NoteTemplate } from '../utils/templates';
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
  ChevronRight,
  Info,
  Sparkles,
  FileSpreadsheet,
  Palette,
  LayoutTemplate,
  Download,
  Plus,
  Layers,
  ExternalLink
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
  allNotes?: Note[];
  onSelectNote?: (id: string) => void;
  onCreateNote?: () => void;
  onCreateNoteWithTitle?: (title: string) => void;
  onSearchTag?: (tag: string) => void;
  onOpenQuickSwitcher?: () => void;
  onUpdateNote: (updated: Partial<Note>) => void;
  onBackToList: () => void;
  onDeleteNote?: (id: string) => void;
  isSidebarCollapsed?: boolean;
  onToggleSidebar?: () => void;
  isNoteListCollapsed?: boolean;
  onToggleNoteList?: () => void;
  onSync?: () => void;
  isSyncing?: boolean;
  onOpenFloatingNote?: (id: string, isPopout?: boolean) => void;
}

export const NoteEditor: React.FC<NoteEditorProps> = ({
  note,
  projects,
  allNotes = [],
  onSelectNote,
  onCreateNote,
  onCreateNoteWithTitle,
  onSearchTag,
  onOpenQuickSwitcher,
  onUpdateNote,
  onBackToList,
  onDeleteNote,
  isSidebarCollapsed,
  onToggleSidebar,
  isNoteListCollapsed,
  onToggleNoteList,
  onSync,
  isSyncing,
  onOpenFloatingNote,
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
  const [isProcessingExcel, setIsProcessingExcel] = useState(false);
  const [isProcessingDocx, setIsProcessingDocx] = useState(false);
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
    attId?: string;
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
  const [activeExcelModal, setActiveExcelModal] = useState<{
    attId?: string;
    data: string;
    title: string;
  } | null>(null);
  const [activeExcelEditorModal, setActiveExcelEditorModal] = useState<{
    attId: string;
    data: string;
    title: string;
  } | null>(null);
  const [activeDocxModal, setActiveDocxModal] = useState<{
    attId?: string;
    data: string;
    title: string;
  } | null>(null);
  const [activeDocxEditorModal, setActiveDocxEditorModal] = useState<{
    attId: string;
    data: string;
    title: string;
  } | null>(null);
  const [activePptxModal, setActivePptxModal] = useState<{
    attId?: string;
    data: string;
    title: string;
  } | null>(null);
  const [activePptxEditorModal, setActivePptxEditorModal] = useState<{
    attId: string;
    data: string;
    title: string;
  } | null>(null);
  const [activeHexModal, setActiveHexModal] = useState<{
    attId?: string;
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

  // 노션 스타일 슬래시(/) 빠른 입력 메뉴 상태
  const [slashMenuState, setSlashMenuState] = useState<{
    isOpen: boolean;
    query: string;
    matchStart: number;
    matchEnd: number;
    position?: { top: number; left: number };
  } | null>(null);
  const [slashSelectedIndex, setSlashSelectedIndex] = useState<number>(0);

  // 지식 연결망 & 백링크(Backlinks) 패널 펼침 상태
  const [isBacklinksExpanded, setIsBacklinksExpanded] = useState(false);

  // 현재 노트를 인용하고 있는 다른 메모들(백링크) 실시간 계산
  const backlinks = useMemo(() => {
    if (!note || !allNotes) return [];
    const currentTitle = (note.title || '').trim();
    if (!currentTitle) return [];
    const strippedCurrent = stripFolderPrefixFromTitle(currentTitle);
    const escaped = currentTitle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const escapedStripped = strippedCurrent.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const wikiRegex = new RegExp(`\\[\\[(?:${escaped}|${escapedStripped})(?:\\|[^\\]]*)?\\]\\]`, 'i');
    return allNotes.filter((n) => n.id !== note.id && wikiRegex.test(n.content));
  }, [note, allNotes]);

  // 현재 노트에서 [[...]] 형태로 연결한 나가는 링크들 계산
  const outlinks = useMemo(() => {
    if (!localContent) return [];
    const matches = Array.from(localContent.matchAll(/\[\[((?:(?!\[\[|\]\])[^\n\r])+?)\]\]/g));
    const titles = Array.from(new Set(matches.map((m) => m[1].split('|')[0].trim()))).filter(Boolean);
    return titles.map((title) => {
      const cleanTarget = title.toLowerCase();
      const found = allNotes.find((n) => {
        const rawTitle = n.title.trim().toLowerCase();
        const strippedTitle = stripFolderPrefixFromTitle(n.title).toLowerCase();
        return rawTitle === cleanTarget || strippedTitle === cleanTarget;
      });
      return { title, noteId: found?.id, exists: !!found };
    });
  }, [localContent, allNotes]);

  // 현재 본문에 포함된 #해시태그 목록 추출
  const tagsInNote = useMemo(() => {
    if (!localContent) return [];
    const matches = Array.from(localContent.matchAll(/(?:^|[\s(（\[{<])#([a-zA-Z0-9가-힣_-]{1,30})/g));
    return Array.from(new Set(matches.map((m) => m[1]))).slice(0, 20);
  }, [localContent]);

  // 실무 서식 템플릿 모달 열림 상태
  const [isTemplateModalOpen, setIsTemplateModalOpen] = useState(false);

  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);
  const hwpInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);
  const docxInputRef = useRef<HTMLInputElement>(null);
  const updateTimerRef = useRef<any>(null);
  const pendingFieldsRef = useRef<Partial<Note>>({});
  const currentNoteIdRef = useRef(note?.id);
  const lastSentContentRef = useRef<string | null>(null);
  const lastSentTitleRef = useRef<string | null>(null);
  const sentContentsHistoryRef = useRef<Set<string>>(new Set());

  // 실무 템플릿 적용 함수
  const insertTemplate = (tpl: NoteTemplate) => {
    const isBlank = !localContent.trim();

    if (isBlank) {
      // 본문이 비어있으면 제목과 본문을 템플릿 기본값으로 전체 설정
      const newContent = tpl.content;
      setLocalContent(newContent);
      pushHistory(newContent, newContent.length);

      let newTitle: string | undefined = undefined;
      // 제목이 비어있거나 '새 메모'인 경우 템플릿 제목으로 자동 보정
      if (!localTitle.trim() || localTitle.includes('새 메모')) {
        const cleanedTitle = tpl.title.replace(/템플릿/g, '').trim();
        const folderPath = note?.projectId ? getFolderHierarchyPath(note.projectId, projects) : '';
        newTitle = folderPath ? `[${folderPath}] ${cleanedTitle}` : cleanedTitle;
        setLocalTitle(newTitle);
      }

      // 동기화할 필드 객체 (본문과 제목을 한 번에 동기화!)
      const fieldsToSync: Partial<Note> = {
        content: newContent,
        ...(newTitle !== undefined ? { title: newTitle } : {}),
      };

      // 상위 에코 덮어쓰기 방지 기록
      lastSentContentRef.current = newContent;
      sentContentsHistoryRef.current.add(newContent);
      if (newTitle !== undefined) {
        lastSentTitleRef.current = newTitle;
      }

      // 대기 중이던 디바운스 타이머가 있다면 취소하고 즉시 상위에 저장!
      if (updateTimerRef.current) {
        clearTimeout(updateTimerRef.current);
      }
      pendingFieldsRef.current = {};
      onUpdateNote(fieldsToSync);

      // textarea에 포커스
      setTimeout(() => {
        const ta = textareaRef.current;
        if (ta) {
          ta.focus();
          ta.setSelectionRange(newContent.length, newContent.length);
        }
      }, 50);
    } else {
      // 이미 내용이 있으면 커서 위치에 템플릿 삽입
      insertMarkdown(`\n\n${tpl.content}\n\n`, '');
    }

    setIsTemplateModalOpen(false);
  };

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

        // 로컬에 이미 작성 중이거나 템플릿이 들어가 내용이 있는데 상위에서 빈 문자열("")이 내려오는 경우 덮어쓰기 방지
        const isParentEmptyReset = !note.content && localContent.trim().length > 0;

        if (note.content !== localContent && !isSentByMe && !isEditingNow && !isParentEmptyReset) {
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

  // 부모 컴포넌트에 안전하게 디바운스 업데이트 전송 (연속 호출 시 필드 유실 방지 병합)
  const debouncedSyncToParent = (updatedFields: Partial<Note>) => {
    pendingFieldsRef.current = {
      ...pendingFieldsRef.current,
      ...updatedFields,
    };

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
      const toSend = { ...pendingFieldsRef.current };
      pendingFieldsRef.current = {};
      if (Object.keys(toSend).length > 0) {
        onUpdateNote(toSend);
      }
    }, 250);
  };

  // 제목 변경 핸들러 (한글 완성형 완벽 지원)
  const handleTitleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newTitle = e.target.value;
    setLocalTitle(newTitle);
    debouncedSyncToParent({ title: newTitle });
  };

  // 슬래시(/) 명령어 감지 및 위치 계산
  const checkSlashCommand = (textarea: HTMLTextAreaElement) => {
    const cursor = textarea.selectionStart;
    if (cursor !== textarea.selectionEnd) {
      if (slashMenuState?.isOpen) setSlashMenuState(null);
      return;
    }
    const val = textarea.value;
    const lineStart = val.lastIndexOf('\n', cursor - 1) + 1;
    const lineToCursor = val.substring(lineStart, cursor);
    const slashMatch = lineToCursor.match(/(?:^|\s)\/([a-zA-Z0-9가-힣ㄱ-ㅎㅏ-ㅣ_-]*)$/);

    if (slashMatch) {
      const query = slashMatch[1];
      const matchStart = cursor - slashMatch[1].length - 1; // '/'의 시작 위치
      const textBefore = val.substring(0, cursor);
      const lines = textBefore.split('\n');
      const lineIndex = lines.length - 1;
      const currentLine = lines[lineIndex];

      const approxTop = Math.max(10, Math.min(textarea.clientHeight - 340, (lineIndex + 1) * 24 - textarea.scrollTop + 8));
      const approxLeft = Math.max(16, Math.min(textarea.clientWidth - 320, currentLine.length * 8.5 + 16));

      setSlashMenuState({
        isOpen: true,
        query,
        matchStart,
        matchEnd: cursor,
        position: { top: approxTop, left: approxLeft },
      });
      setSlashSelectedIndex(0);
    } else {
      if (slashMenuState?.isOpen) {
        setSlashMenuState(null);
      }
    }
  };

  // 슬래시 명령어 실행 및 텍스트 치환
  const executeSlashCommand = (cmd: SlashCommandItem) => {
    if (!slashMenuState) return;
    const textarea = textareaRef.current;
    if (!textarea) return;

    if (cmd.id === 'blank-excel') {
      const text = localContent;
      const before = text.substring(0, slashMenuState.matchStart);
      const after = text.substring(slashMenuState.matchEnd);
      const newBaseContent = before + after;
      const insertPos = slashMenuState.matchStart;
      setSlashMenuState(null);
      handleCreateBlankExcel(newBaseContent, insertPos);
      return;
    }

    const text = localContent;
    const before = text.substring(0, slashMenuState.matchStart);
    const after = text.substring(slashMenuState.matchEnd);
    const newContent = before + cmd.snippet + after;
    const nextPos = slashMenuState.matchStart + (cmd.cursorOffset !== undefined ? cmd.cursorOffset : cmd.snippet.length);

    setLocalContent(newContent);
    debouncedSyncToParent({ content: newContent });
    pushHistory(newContent, nextPos);
    setSlashMenuState(null);

    setTimeout(() => {
      textarea.focus({ preventScroll: true });
      textarea.setSelectionRange(nextPos, nextPos);
    }, 0);
  };

  // 본문 변경 핸들러 (한글 조합 중 리렌더링 차단)
  const handleContentChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const newContent = e.target.value;
    setLocalContent(newContent);
    debouncedSyncToParent({ content: newContent });
    debouncedPushHistory(newContent, e.target.selectionStart);
    checkSlashCommand(e.target);
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
      // <summary> 태그 뒤에 빈 줄이 없으면 마크다운 파싱을 위해 자동 줄바꿈 보정
      normalizedContent = normalizedContent.replace(/(<summary>[\s\S]*?<\/summary>)\s*([^\s\n<])/gi, '$1\n\n$2');
      let rawHtml = marked.parse(normalizedContent, { breaks: true, gfm: true }) as string;

      // 표(table)를 모바일 가로 스크롤 컨테이너로 감싸 좁은 화면에서 컬럼이 찌그러지는 현상 방지
      rawHtml = rawHtml.replace(/<table>([\s\S]*?)<\/table>/g, '<div class="table-scroll-container"><table>$1</table></div>');

      // 노션 및 GitHub 스타일 콜아웃(Callout) 서식 변환 (> [!NOTE], > [!TIP], > [!WARNING], > [!IMPORTANT], > [!CAUTION])
      const alertConfig: Record<string, { title: string; icon: string; bg: string; border: string; borderLeft: string; headerColor: string }> = {
        NOTE: {
          title: '참고 (Note)',
          icon: 'ℹ️',
          bg: 'rgba(30, 58, 138, 0.22)',
          border: 'rgba(59, 130, 246, 0.4)',
          borderLeft: '#3b82f6',
          headerColor: '#60a5fa',
        },
        TIP: {
          title: '팁 (Tip)',
          icon: '💡',
          bg: 'rgba(6, 78, 59, 0.22)',
          border: 'rgba(16, 185, 129, 0.4)',
          borderLeft: '#10b981',
          headerColor: '#34d399',
        },
        IMPORTANT: {
          title: '중요 (Important)',
          icon: '📌',
          bg: 'rgba(76, 29, 149, 0.25)',
          border: 'rgba(139, 92, 246, 0.4)',
          borderLeft: '#8b5cf6',
          headerColor: '#a78bfa',
        },
        WARNING: {
          title: '주의 (Warning)',
          icon: '⚠️',
          bg: 'rgba(120, 53, 15, 0.25)',
          border: 'rgba(245, 158, 11, 0.4)',
          borderLeft: '#f59e0b',
          headerColor: '#fbbf24',
        },
        CAUTION: {
          title: '위험 (Caution)',
          icon: '🛑',
          bg: 'rgba(127, 29, 29, 0.25)',
          border: 'rgba(239, 68, 68, 0.4)',
          borderLeft: '#ef4444',
          headerColor: '#f87171',
        },
      };

      rawHtml = rawHtml.replace(
        /<blockquote>\s*<p>\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\](?:\s*<br\s*\/?>)?([\s\S]*?)<\/blockquote>/gi,
        (_match, type, content) => {
          const key = type.toUpperCase();
          const config = alertConfig[key] || alertConfig.NOTE;
          let body = (content || '').trim();
          if (body.startsWith('</p>')) body = body.substring(4);
          if (body.endsWith('</p>')) body = body.substring(0, body.length - 4);
          return `
            <div class="callout-box" style="background-color: ${config.bg}; border-color: ${config.border}; border-left-color: ${config.borderLeft};">
              <div style="display: flex; align-items: center; gap: 7px; font-weight: 700; color: ${config.headerColor}; margin-bottom: 6px; font-size: 0.9em; text-transform: uppercase; letter-spacing: 0.05em;">
                <span style="font-size: 1.1em;">${config.icon}</span>
                <span>${config.title}</span>
              </div>
              <div style="color: #cbd5e1; font-size: 0.95em; line-height: 1.65;">
                <p style="margin: 0;">${body.trim()}</p>
              </div>
            </div>
          `;
        }
      );

      // 마크다운 체크박스를 보기 모드에서도 직접 클릭해 체크/해제할 수 있도록 속성 변환
      let taskCounter = 0;
      rawHtml = rawHtml.replace(/<input\s+([^>]*?)type="checkbox"([^>]*?)>/g, (_match, before, after) => {
        const currentIndex = taskCounter++;
        const cleanBefore = before.replace(/disabled(?:="[^"]*")?/g, '');
        const cleanAfter = after.replace(/disabled(?:="[^"]*")?/g, '');
        return `<input ${cleanBefore}type="checkbox"${cleanAfter} data-task-index="${currentIndex}" class="task-checkbox cursor-pointer">`;
      });

      // 위키링크([[노트제목]]) 및 해시태그(#태그) 인터랙티브 배지 변환
      // HTML 태그 속성 및 code/pre 블록 내부를 안전하게 보존하기 위해 태그 기준으로 분할 처리
      const htmlTokens = rawHtml.split(/(<[^>]+>)/g);
      let inCodeBlock = false;
      for (let i = 0; i < htmlTokens.length; i++) {
        const token = htmlTokens[i];
        if (token.startsWith('<')) {
          if (/^<(?:code|pre|textarea|input|kbd)\b/i.test(token)) inCodeBlock = true;
          if (/^<\/(?:code|pre|textarea|input|kbd)>/i.test(token)) inCodeBlock = false;
          continue;
        }
        if (inCodeBlock || !token) continue;

        let processed = token;

        // 1. 위키링크 변환: [[문서명]] 또는 [[문서명|별칭]] (내부 [폴더명] 대괄호 지원)
        processed = processed.replace(/\[\[((?:(?!\[\[|\]\])[^\n\r])+?)\]\]/g, (match, inner) => {
          const parts = inner.split('|');
          const targetTitle = parts[0].trim();
          if (!targetTitle) return match;

          const alias = (parts[1] !== undefined ? parts[1].trim() : targetTitle) || targetTitle;
          const cleanTarget = targetTitle.toLowerCase();
          const targetNote = allNotes.find((n) => {
            const rawTitle = n.title.trim().toLowerCase();
            const strippedTitle = stripFolderPrefixFromTitle(n.title).toLowerCase();
            return rawTitle === cleanTarget || strippedTitle === cleanTarget;
          });

          const safeTarget = targetTitle.replace(/"/g, '&quot;');
          const safeAlias = alias.replace(/</g, '&lt;').replace(/>/g, '&gt;');

          if (targetNote) {
            return `<button type="button" class="wikilink-pill wikilink-exists" data-wikilink-id="${targetNote.id}" data-wikilink-title="${safeTarget}" title="메모 바로가기: ${safeTarget}">🔗 ${safeAlias}</button>`;
          } else {
            return `<button type="button" class="wikilink-pill wikilink-missing" data-wikilink-create="true" data-wikilink-title="${safeTarget}" title="클릭 시 새 메모 생성: ${safeTarget}">📄 ${safeAlias} <span class="text-[10px] text-indigo-400 opacity-90 ml-0.5 font-normal">+생성</span></button>`;
          }
        });

        // 2. 해시태그 변환: #태그명 (공백이나 문장부호 경계)
        processed = processed.replace(/(?<=^|[\s(（\[{<])#([a-zA-Z0-9가-힣_-]{1,30})(?=[\s)）\]}>,.:;!?~]|$)/g, (_m, tagName) => {
          const safeTag = tagName.replace(/"/g, '&quot;');
          return `<button type="button" class="hashtag-pill" data-tag-name="${safeTag}" title="#${safeTag} 태그로 메모 검색">#${safeTag}</button>`;
        });

        htmlTokens[i] = processed;
      }
      rawHtml = htmlTokens.join('');

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

  // 인라인 Word(DOCX) 페이지 렌더링 핸들러
  const renderInlineDocx = useRef<(attId: string, customData?: string) => void>(() => {});
  renderInlineDocx.current = (attId: string, customData?: string) => {
    const preview = previewContainerRef.current;
    if (!preview) return;

    const viewer = preview.querySelector<HTMLElement>(`.docx-inline-viewer[data-att-id="${attId}"]`);
    if (!viewer) return;

    const att = localAttachments[attId];
    const dataToUse = customData || att?.data;
    if (!dataToUse) return;

    const docxContainer = viewer.querySelector<HTMLElement>(`#docx-container-${attId}`);
    const loadingBox = viewer.querySelector<HTMLElement>(`#docx-loading-${attId}`);

    if (docxContainer) {
      renderDocxToContainer(dataToUse, docxContainer)
        .then(() => {
          if (loadingBox) loadingBox.classList.add('hidden');
          docxContainer.classList.remove('hidden');
          const scrollBox = viewer.querySelector<HTMLElement>('.docx-scroll-container');
          if (scrollBox) scrollBox.scrollTop = 0;
        })
        .catch((err) => {
          console.error('인라인 워드 렌더링 실패:', err);
          if (loadingBox) {
            loadingBox.innerHTML = '<span class="text-rose-400 font-semibold text-xs">⚠️ 워드(DOCX) 문서를 로드하지 못했습니다. 상단 저장 버튼을 이용해 주세요.</span>';
          }
        });
    }
  };

  // 미리보기 화면이 갱신되거나 활성화될 때 모든 인라인 PDF & HWP & DOCX 자동 렌더링
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

    const docxViewers = preview.querySelectorAll<HTMLElement>('.docx-inline-viewer');
    docxViewers.forEach((v) => {
      const attId = v.getAttribute('data-att-id');
      if (attId) {
        renderInlineDocx.current(attId);
      }
    });

    // <details> 토글 목록이 펼쳐질 때 내부의 워드/한글/PDF 뷰어 자동 재렌더링
    const handleDetailsToggle = (e: Event) => {
      const details = e.target as HTMLDetailsElement;
      if (!details || !details.open) return;
      const dViewers = details.querySelectorAll<HTMLElement>('.docx-inline-viewer');
      dViewers.forEach((v) => {
        const attId = v.getAttribute('data-att-id');
        if (attId) renderInlineDocx.current(attId);
      });
      const hViewers = details.querySelectorAll<HTMLElement>('.hwp-inline-viewer');
      hViewers.forEach((v) => {
        const attId = v.getAttribute('data-att-id');
        if (attId) renderInlineHwp.current(attId);
      });
      const pViewers = details.querySelectorAll<HTMLElement>('.pdf-inline-viewer');
      pViewers.forEach((v) => {
        const attId = v.getAttribute('data-att-id');
        if (attId) renderInlinePdf.current(attId);
      });
    };
    preview.addEventListener('toggle', handleDetailsToggle, true);
    return () => {
      preview.removeEventListener('toggle', handleDetailsToggle, true);
    };
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
      <div className="flex-1 flex flex-col items-center justify-center bg-slate-950 text-slate-500 p-6 text-center select-none">
        <div className="w-16 h-16 rounded-2xl bg-slate-900/80 border border-slate-800/80 flex items-center justify-center mb-4 text-indigo-400/70 shadow-inner">
          <Edit3 size={32} />
        </div>
        <p className="text-sm font-semibold text-slate-300">선택된 메모가 없습니다</p>
        <p className="text-xs text-slate-500 mt-1.5 max-w-sm leading-relaxed">
          좌측 목록에서 메모를 선택하거나, 아래 버튼을 눌러 새 메모를 작성해 보세요.
        </p>
        {onCreateNote && (
          <button
            onClick={onCreateNote}
            className="mt-5 px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-medium transition-all shadow-md hover:shadow-indigo-500/20 flex items-center gap-1.5 active:scale-95 cursor-pointer"
          >
            <Plus size={15} />
            새 메모 작성
          </button>
        )}
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

  // 노션 토글(Toggle) 삽입
  const insertToggle = () => {
    insertMarkdown('\n<details>\n<summary>토글 제목을 입력하세요</summary>\n\n숨겨둘 상세 내용을 여기에 작성하세요.\n</details>\n\n');
  };

  // 노션 콜아웃(Callout) 삽입
  const insertCallout = (type: 'NOTE' | 'TIP' | 'WARNING' | 'IMPORTANT' = 'NOTE') => {
    insertMarkdown(`\n> [!${type}]\n> 강조할 내용을 여기에 입력하세요.\n\n`);
  };

  // 위키링크([[ ]]) 삽입
  const insertWikiLink = () => {
    insertMarkdown('[[', ']]');
  };

  // PC 작성 편의를 위한 키보드 단축키 핸들러
  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    // 슬래시(/) 빠른 서식 메뉴 키보드 내비게이션
    if (slashMenuState?.isOpen) {
      const allCommands = getSlashCommands();
      const filteredCommands = allCommands.filter((cmd) => {
        if (!slashMenuState.query) return true;
        const cleanQ = slashMenuState.query.toLowerCase().trim();
        return (
          cmd.title.toLowerCase().includes(cleanQ) ||
          cmd.description.toLowerCase().includes(cleanQ) ||
          cmd.keywords.some((k) => k.toLowerCase().includes(cleanQ))
        );
      });

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        if (filteredCommands.length > 0) {
          setSlashSelectedIndex((prev) => (prev + 1) % filteredCommands.length);
        }
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        if (filteredCommands.length > 0) {
          setSlashSelectedIndex((prev) => (prev - 1 + filteredCommands.length) % filteredCommands.length);
        }
        return;
      }
      if (e.key === 'Enter') {
        if (filteredCommands.length > 0 && filteredCommands[slashSelectedIndex]) {
          e.preventDefault();
          executeSlashCommand(filteredCommands[slashSelectedIndex]);
          return;
        }
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        setSlashMenuState(null);
        return;
      }
    }

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
      // Ctrl + P: 초고속 퀵 스위처 열기
      if (e.key.toLowerCase() === 'p') {
        if (onOpenQuickSwitcher) {
          e.preventDefault();
          onOpenQuickSwitcher();
          return;
        }
      }
      // Ctrl + K: 텍스트가 선택되어 있으면 링크 삽입, 미선택 시 퀵 스위처 열기
      if (e.key.toLowerCase() === 'k' && !e.shiftKey) {
        e.preventDefault();
        const textarea = textareaRef.current;
        if (textarea && textarea.selectionStart !== textarea.selectionEnd) {
          insertMarkdown('[', '](https://)');
        } else if (onOpenQuickSwitcher) {
          onOpenQuickSwitcher();
        } else {
          insertMarkdown('[', '](https://)');
        }
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

  // 빈 엑셀 스프레드시트 생성 및 즉시 편집 모달 열기
  const handleCreateBlankExcel = (baseContent?: string, targetPos?: number) => {
    const now = new Date();
    const timeStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
    const fileName = `스프레드시트_${timeStr}.xlsx`;
    const defaultSheets: Record<string, (string | number)[][]> = {
      'Sheet1': [
        ['항목', '구분', '금액', '비고'],
        ['', '', '', ''],
        ['', '', '', ''],
        ['', '', '', ''],
      ],
    };
    const dataUrl = generateExcelDataUrl(defaultSheets, 'xlsx');
    const attId = `att-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
    const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    const updatedAttachments: Record<string, AttachmentItem> = {
      ...localAttachmentsRef.current,
      [attId]: {
        id: attId,
        name: fileName,
        type: mimeType,
        data: dataUrl,
      },
    };
    setLocalAttachments(updatedAttachments);
    localAttachmentsRef.current = updatedAttachments;

    const insertTag = `\n\n[📊 엑셀: ${fileName}](attachment:${attId})\n\n`;
    const textarea = textareaRef.current;
    const content = baseContent !== undefined ? baseContent : localContent;

    let newContent = '';
    let nextPos = 0;

    if (targetPos !== undefined) {
      newContent = content.substring(0, targetPos) + insertTag + content.substring(targetPos);
      nextPos = targetPos + insertTag.length;
    } else if (textarea) {
      const start = textarea.selectionStart;
      const end = textarea.selectionEnd;
      newContent = content.substring(0, start) + insertTag + content.substring(end);
      nextPos = start + insertTag.length;
    } else {
      newContent = `${content}${insertTag}`;
      nextPos = newContent.length;
    }

    setLocalContent(newContent);
    onUpdateNote({ content: newContent, attachments: updatedAttachments });
    pushHistory(newContent, nextPos);

    if (textarea) {
      setTimeout(() => {
        textarea.focus({ preventScroll: true });
        textarea.setSelectionRange(nextPos, nextPos);
      }, 0);
    }

    // 엑셀 웹 스프레드시트 편집기 모달 즉시 열기
    setActiveExcelEditorModal({
      attId,
      data: dataUrl,
      title: fileName,
    });
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
      ...localAttachmentsRef.current,
      [attId]: {
        id: attId,
        name: attName,
        type: mimeType,
        data: dataUrl,
      },
    };

    setLocalAttachments(updatedAttachments);
    localAttachmentsRef.current = updatedAttachments;

    const isPdf = mimeType === 'application/pdf' || attName.toLowerCase().endsWith('.pdf');
    const isHwp = attName.toLowerCase().endsWith('.hwp') || attName.toLowerCase().endsWith('.hwpx') || mimeType.includes('hwp');
    const isExcel =
      attName.toLowerCase().endsWith('.xlsx') ||
      attName.toLowerCase().endsWith('.xls') ||
      attName.toLowerCase().endsWith('.csv') ||
      mimeType.includes('spreadsheet') ||
      mimeType === 'text/csv';
    const isDocx =
      attName.toLowerCase().endsWith('.docx') ||
      attName.toLowerCase().endsWith('.doc') ||
      mimeType.includes('wordprocessingml') ||
      mimeType === 'application/msword';
    const insertTag = isPdf
      ? `\n\n[📄 PDF: ${safeName}](attachment:${attId})\n\n`
      : isHwp
      ? `\n\n[📝 한글: ${safeName}](attachment:${attId})\n\n`
      : isExcel
      ? `\n\n[📊 엑셀: ${safeName}](attachment:${attId})\n\n`
      : isDocx
      ? `\n\n[📄 워드: ${safeName}](attachment:${attId})\n\n`
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

  // 4-2. 엑셀/CSV(XLSX, XLS, CSV) 문서 전용 파일 선택 핸들러
  const handleExcelFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 30 * 1024 * 1024) {
      alert('30MB 이하의 스프레드시트 파일만 첨부할 수 있습니다.');
      if (excelInputRef.current) excelInputRef.current.value = '';
      return;
    }

    setIsProcessingExcel(true);
    try {
      const reader = new FileReader();
      reader.onload = (event) => {
        const dataUrl = event.target?.result as string;
        if (dataUrl) {
          const mime =
            file.type ||
            (file.name.toLowerCase().endsWith('.csv')
              ? 'text/csv'
              : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          insertAttachmentAtCursor(file.name, dataUrl, mime);
        }
        setIsProcessingExcel(false);
        if (excelInputRef.current) excelInputRef.current.value = '';
      };
      reader.onerror = () => {
        alert('스프레드시트 파일을 읽는 중 오류가 발생했습니다.');
        setIsProcessingExcel(false);
        if (excelInputRef.current) excelInputRef.current.value = '';
      };
      reader.readAsDataURL(file);
    } catch (err) {
      console.error('엑셀 업로드 실패:', err);
      alert('엑셀 파일을 불러오지 못했습니다.');
      setIsProcessingExcel(false);
      if (excelInputRef.current) excelInputRef.current.value = '';
    }
  };

  // 4-3. 워드(DOCX, DOC) 문서 전용 파일 선택 핸들러
  const handleDocxFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 30 * 1024 * 1024) {
      alert('30MB 이하의 워드 파일만 첨부할 수 있습니다.');
      if (docxInputRef.current) docxInputRef.current.value = '';
      return;
    }

    setIsProcessingDocx(true);
    try {
      const reader = new FileReader();
      reader.onload = (event) => {
        const dataUrl = event.target?.result as string;
        if (dataUrl) {
          const mime = file.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
          insertAttachmentAtCursor(file.name, dataUrl, mime);
        }
        setIsProcessingDocx(false);
        if (docxInputRef.current) docxInputRef.current.value = '';
      };
      reader.onerror = () => {
        alert('워드 파일을 읽는 중 오류가 발생했습니다.');
        setIsProcessingDocx(false);
        if (docxInputRef.current) docxInputRef.current.value = '';
      };
      reader.readAsDataURL(file);
    } catch (err) {
      console.error('워드 업로드 실패:', err);
      alert('워드 파일을 불러오지 못했습니다.');
      setIsProcessingDocx(false);
      if (docxInputRef.current) docxInputRef.current.value = '';
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

    // 1-2. 위키링크([[문서명]]) 클릭 처리
    const wikiBtn = target.closest<HTMLElement>('.wikilink-pill');
    if (wikiBtn) {
      e.preventDefault();
      e.stopPropagation();
      const noteId = wikiBtn.getAttribute('data-wikilink-id');
      const targetTitle = wikiBtn.getAttribute('data-wikilink-title') || '';
      const isCreate = wikiBtn.getAttribute('data-wikilink-create') === 'true';

      if (noteId && onSelectNote) {
        onSelectNote(noteId);
      } else if (isCreate && targetTitle && onCreateNoteWithTitle) {
        if (window.confirm(`'${targetTitle}' 메모가 아직 없습니다. 새 메모로 생성하시겠습니까?`)) {
          onCreateNoteWithTitle(targetTitle);
        }
      }
      return;
    }

    // 1-3. 해시태그(#태그) 클릭 처리
    const tagBtn = target.closest<HTMLElement>('.hashtag-pill');
    if (tagBtn) {
      e.preventDefault();
      e.stopPropagation();
      const tagName = tagBtn.getAttribute('data-tag-name');
      if (tagName && onSearchTag) {
        onSearchTag(`#${tagName}`);
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
          attId,
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

    // 7-3. 엑셀 [✏️ 직접 편집] 버튼 클릭 시 웹 스프레드시트 편집기 모달 열기
    const editExcelBtn = target.closest<HTMLElement>('.excel-btn-edit');
    if (editExcelBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = editExcelBtn.getAttribute('data-att-id') || '';
      const title = editExcelBtn.getAttribute('data-file-name') || '엑셀 문서';

      const attachmentsMap = localAttachmentsRef.current || localAttachments;
      const targetAtt = 
        (attId && attachmentsMap[attId])
        || (attId && localAttachments[attId])
        || Object.values(attachmentsMap).find((a) => a.name === title)
        || Object.values(localAttachments).find((a) => a.name === title)
        || Object.values(attachmentsMap).find((a) => {
          const n = (a.name || '').toLowerCase();
          return n.endsWith('.xlsx') || n.endsWith('.xls') || n.endsWith('.csv');
        });

      if (targetAtt && targetAtt.data) {
        const resolvedAttId = targetAtt.id || attId || Object.keys(attachmentsMap).find(k => attachmentsMap[k] === targetAtt) || '';
        setActiveExcelEditorModal({
          attId: resolvedAttId,
          data: targetAtt.data,
          title: targetAtt.name || title,
        });
      }
      return;
    }

    // 7-4. 엑셀 [🔍 크게 보기] 클릭 시 전면 뷰어 모달 열기
    const fullscreenExcelBtn = target.closest<HTMLElement>('.excel-btn-fullscreen');
    if (fullscreenExcelBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = fullscreenExcelBtn.getAttribute('data-att-id') || '';
      const title = fullscreenExcelBtn.getAttribute('data-file-name') || '엑셀 문서';

      const attachmentsMap = localAttachmentsRef.current || localAttachments;
      const targetAtt = 
        (attId && attachmentsMap[attId])
        || (attId && localAttachments[attId])
        || Object.values(attachmentsMap).find((a) => a.name === title)
        || Object.values(localAttachments).find((a) => a.name === title)
        || Object.values(attachmentsMap).find((a) => {
          const n = (a.name || '').toLowerCase();
          return n.endsWith('.xlsx') || n.endsWith('.xls') || n.endsWith('.csv');
        });

      if (targetAtt && targetAtt.data) {
        const resolvedAttId = targetAtt.id || attId || Object.keys(attachmentsMap).find(k => attachmentsMap[k] === targetAtt) || '';
        setActiveExcelModal({
          attId: resolvedAttId,
          data: targetAtt.data,
          title: targetAtt.name || title,
        });
      }
      return;
    }

    // 7-5. 엑셀 [📝 표 가져오기] 클릭 시 마크다운 표로 본문에 가져오기
    const importExcelBtn = target.closest<HTMLElement>('.excel-btn-import');
    if (importExcelBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = importExcelBtn.getAttribute('data-att-id');
      if (attId && localAttachments[attId]) {
        const parsed = parseExcelWorkbook(localAttachments[attId].data, attId);
        const firstSheet = parsed.sheetNames[0] || 'Sheet1';
        const rows = parsed.sheets[firstSheet] || [];
        const mdTable = sheetRowsToMarkdownTable(rows);
        if (!mdTable.trim()) {
          alert('가져올 수 있는 표 데이터가 없습니다.');
          return;
        }
        const insertBlock = `\n\n### 📊 ${localAttachments[attId].name} (${firstSheet})\n\n${mdTable}\n\n`;
        const newContent = `${localContent}${insertBlock}`;
        setLocalContent(newContent);
        debouncedSyncToParent({ content: newContent });
        alert(`'${localAttachments[attId].name}' 표 데이터(${rows.length}행)를 마크다운 표로 메모에 성공적으로 가져왔습니다!`);
      }
      return;
    }

    // 7-6. 엑셀 인라인 시트 탭 전환 클릭
    const sheetPill = target.closest<HTMLElement>('.excel-sheet-pill');
    if (sheetPill) {
      e.preventDefault();
      e.stopPropagation();
      const attId = sheetPill.getAttribute('data-att-id');
      const sheetName = sheetPill.getAttribute('data-sheet-name');
      const preview = previewContainerRef.current;
      if (attId && sheetName && preview && localAttachments[attId]) {
        const viewer = preview.querySelector<HTMLElement>(`.excel-inline-viewer[data-att-id="${attId}"]`);
        if (viewer) {
          const parsed = parseExcelWorkbook(localAttachments[attId].data, attId);
          const rows = parsed.sheets[sheetName] || [];
          const sheetMerges = parsed.merges?.[sheetName] || [];
          const tableEl = viewer.querySelector<HTMLTableElement>(`#excel-table-${attId}`);
          if (tableEl) {
            let maxCols = 0;
            rows.forEach((r) => { maxCols = Math.max(maxCols, r.length); });
            sheetMerges.forEach((m) => { maxCols = Math.max(maxCols, m.e.c + 1); });
            const colCount = Math.max(maxCols, 1);
            const previewRows = rows.slice(0, 12);
            const mergeMap = buildMergeMap(sheetMerges);

            const thead = tableEl.querySelector('thead');
            if (thead) {
              const colHeaders = Array.from({ length: colCount }).map((_, cIdx) => {
                return `<th class="px-3 py-1.5 text-[11px] font-mono font-semibold text-[#334155] border-r border-[#cbd5e1] min-w-[85px] bg-[#f1f5f9] text-center select-none">${getColumnLabel(cIdx)}</th>`;
              }).join('');
              thead.innerHTML = `<tr><th class="w-10 px-2 py-1.5 text-center text-[10px] font-bold text-[#64748b] bg-[#e2e8f0] border-r border-[#cbd5e1] sticky left-0 z-20 select-none">#</th>${colHeaders}</tr>`;
            }

            const tbody = tableEl.querySelector('tbody');
            if (tbody) {
              tbody.innerHTML = previewRows.map((row, rIdx) => {
                const cellsHtml: string[] = [];
                for (let cIdx = 0; cIdx < colCount; cIdx++) {
                  const mergeInfo = mergeMap.get(`${rIdx},${cIdx}`);
                  if (mergeInfo?.isCovered) {
                    continue;
                  }
                  const val = row[cIdx] !== undefined && row[cIdx] !== null ? String(row[cIdx]) : '';
                  const colSpanAttr = mergeInfo && mergeInfo.colSpan > 1 ? ` colspan="${mergeInfo.colSpan}"` : '';
                  const rowSpanAttr = mergeInfo && mergeInfo.rowSpan > 1 ? ` rowspan="${mergeInfo.rowSpan}"` : '';
                  const isMerged = mergeInfo && (mergeInfo.colSpan > 1 || mergeInfo.rowSpan > 1);

                  cellsHtml.push(
                    `<td${colSpanAttr}${rowSpanAttr} class="px-3 py-1.5 text-xs text-[#0f172a] border-r border-[#e2e8f0] break-words whitespace-pre-wrap max-w-xs font-sans leading-relaxed ${isMerged ? 'bg-slate-50/70 font-semibold' : ''}">${val}</td>`
                  );
                }

                return `
                  <tr class="hover:bg-[#ecfdf5] transition-colors border-b border-[#e2e8f0]">
                    <td class="px-2 py-1.5 text-center text-[10px] font-mono font-semibold text-[#64748b] bg-[#f8fafc] border-r border-[#cbd5e1] sticky left-0 z-10 select-none">${rIdx + 1}</td>
                    ${cellsHtml.join('')}
                  </tr>
                `;
              }).join('');
            }

            viewer.querySelectorAll('.excel-sheet-pill').forEach((pill) => {
              if (pill.getAttribute('data-sheet-name') === sheetName) {
                pill.className = 'excel-sheet-pill px-2 py-0.5 text-[11px] font-semibold rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 cursor-pointer';
              } else {
                pill.className = 'excel-sheet-pill px-2 py-0.5 text-[11px] font-semibold rounded bg-slate-800 text-slate-400 hover:text-slate-200 border border-slate-700/60 cursor-pointer';
              }
            });
          }
        }
      }
      return;
    }

    // 7-7. 워드 [✏️ 직접 편집] 버튼 클릭 시 웹 에디터 모달 열기
    const editDocxBtn = target.closest<HTMLElement>('.docx-btn-edit');
    if (editDocxBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = editDocxBtn.getAttribute('data-att-id');
      const title = editDocxBtn.getAttribute('data-file-name') || '워드 문서';
      if (attId && localAttachments[attId]) {
        setActiveDocxEditorModal({
          attId,
          data: localAttachments[attId].data,
          title,
        });
      }
      return;
    }

    // 7-8. 워드 [🔍 크게 보기] 또는 워드 용지 클릭 시 전면 뷰어 모달 열기
    const fullscreenDocxBtn = target.closest<HTMLElement>('.docx-btn-fullscreen') || target.closest<HTMLElement>('.docx-container');
    if (fullscreenDocxBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = fullscreenDocxBtn.getAttribute('data-att-id');
      const title = fullscreenDocxBtn.getAttribute('data-file-name') || '워드 문서';
      if (attId && localAttachments[attId]) {
        setActiveDocxModal({
          attId,
          data: localAttachments[attId].data,
          title,
        });
      }
      return;
    }

    // 7-9. 워드 [📝 메모로 가져오기] 클릭 시 마크다운으로 본문 반영
    const importDocxBtn = target.closest<HTMLElement>('.docx-btn-import');
    if (importDocxBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = importDocxBtn.getAttribute('data-att-id');
      if (attId && localAttachments[attId]) {
        extractDocxTextAndHtml(localAttachments[attId].data)
          .then(({ markdown, text }) => {
            const contentToImport = markdown || text;
            if (!contentToImport.trim()) {
              alert('추출할 수 있는 문서 내용이 없습니다.');
              return;
            }
            const insertBlock = `\n\n### 📄 ${localAttachments[attId].name} 내용\n\n${contentToImport}\n\n`;
            const newContent = `${localContent}${insertBlock}`;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            alert(`'${localAttachments[attId].name}' 문서 내용(${contentToImport.length}자)을 메모에 성공적으로 가져왔습니다!`);
          })
          .catch((err) => {
            console.error('워드 본문 추출 실패:', err);
            alert('워드 문서 내용을 추출하지 못했습니다.');
          });
      }
      return;
    }

    // 7-10. 파워포인트 [✏️ 직접 편집] 클릭 시 웹 슬라이드 에디터 열기
    const editPptxBtn = target.closest<HTMLElement>('.pptx-btn-edit');
    if (editPptxBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = editPptxBtn.getAttribute('data-att-id');
      const title = editPptxBtn.getAttribute('data-file-name') || '파워포인트 발표 자료';
      if (attId && localAttachments[attId]) {
        setActivePptxEditorModal({
          attId,
          data: localAttachments[attId].data,
          title,
        });
      }
      return;
    }

    // 7-11. 파워포인트 [🔍 크게 보기] 또는 슬라이드 카드 클릭 시 전면 슬라이드 뷰어/쇼 열기
    const fullscreenPptxBtn = target.closest<HTMLElement>('.pptx-btn-fullscreen') || target.closest<HTMLElement>('.pptx-container');
    if (fullscreenPptxBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = fullscreenPptxBtn.getAttribute('data-att-id');
      const title = fullscreenPptxBtn.getAttribute('data-file-name') || '파워포인트 발표 자료';
      if (attId && localAttachments[attId]) {
        setActivePptxModal({
          attId,
          data: localAttachments[attId].data,
          title,
        });
      }
      return;
    }

    // 7-12. 파워포인트 [📝 메모로 가져오기] 클릭 시 마크다운으로 본문 반영
    const importPptxBtn = target.closest<HTMLElement>('.pptx-btn-import');
    if (importPptxBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = importPptxBtn.getAttribute('data-att-id');
      if (attId && localAttachments[attId]) {
        extractPptxTextAndMarkdown(localAttachments[attId].data)
          .then(({ markdown, text }) => {
            const contentToImport = markdown || text;
            if (!contentToImport.trim()) {
              alert('추출할 수 있는 슬라이드 내용이 없습니다.');
              return;
            }
            const insertBlock = `\n\n${contentToImport}\n\n`;
            const newContent = `${localContent}${insertBlock}`;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            alert(`'${localAttachments[attId].name}' 발표 자료 내용을 메모에 성공적으로 가져왔습니다!`);
          })
          .catch((err) => {
            console.error('파워포인트 본문 추출 실패:', err);
            alert('파워포인트 슬라이드 내용을 추출하지 못했습니다.');
          });
      }
      return;
    }

    // 7-13. 바이너리/모든 첨부파일 [🔍 헥스 편집기] 또는 바이너리 뷰어 클릭 시 울트라에디트 모달 열기
    const hexBtn = target.closest<HTMLElement>('.hex-btn-edit') || target.closest<HTMLElement>('.binary-container');
    if (hexBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = hexBtn.getAttribute('data-att-id');
      const title = hexBtn.getAttribute('data-file-name') || '바이너리 파일';
      if (attId && localAttachments[attId]) {
        setActiveHexModal({
          attId,
          data: localAttachments[attId].data,
          title,
        });
      }
      return;
    }

    // 7-14. 첨부파일 [📁 접기 / 📂 접기 해제] 원클릭 토글 변환
    const toggleAttBtn = target.closest<HTMLElement>('.att-btn-toggle');
    if (toggleAttBtn) {
      e.preventDefault();
      e.stopPropagation();
      const attId = toggleAttBtn.getAttribute('data-att-id');
      const fileName = toggleAttBtn.getAttribute('data-file-name') || '문서';
      if (!attId) return;

      const content = localContent;
      // 1. 이미 <details> 안에 감싸져 있는지 확인
      const wrappedRegex = new RegExp(
        `[\\r\\n]*<details[^>]*>\\s*<summary>[\\s\\S]*?<\\/summary>\\s*([\\s\\S]*?(!?\\[[^\\]]*\\]\\((?:attachment:)?${attId}\\))[\\s\\S]*?)\\s*<\\/details>[\\r\\n]*`,
        'i'
      );
      const match = content.match(wrappedRegex);

      if (match) {
        // 이미 토글 상태 -> 토글 해제 (내부 첨부 태그만 유지)
        const innerContent = match[1].trim();
        const newContent = content.replace(wrappedRegex, `\n\n${innerContent}\n\n`);
        setLocalContent(newContent);
        debouncedSyncToParent({ content: newContent });
        pushHistory(newContent);
      } else {
        // 일반 상태 -> <details><summary>... 로 감싸기
        const plainAttRegex = new RegExp(
          `[\\r\\n]*(!?\\[[^\\]]*\\]\\((?:attachment:)?${attId}\\))[\\r\\n]*`,
          'i'
        );
        const plainMatch = content.match(plainAttRegex);
        if (plainMatch) {
          const rawTag = plainMatch[1].trim();
          const wrappedBlock = `\n\n<details>\n<summary>📄 ${fileName}</summary>\n\n${rawTag}\n\n</details>\n\n`;
          const newContent = content.replace(plainAttRegex, wrappedBlock);
          setLocalContent(newContent);
          debouncedSyncToParent({ content: newContent });
          pushHistory(newContent);
        }
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

      {/* 숨겨진 엑셀/CSV(XLSX, XLS, CSV) 문서 선택 인풋 */}
      <input
        type="file"
        ref={excelInputRef}
        onChange={handleExcelFileChange}
        accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv"
        className="hidden"
      />

      {/* 숨겨진 워드(DOCX, DOC) 문서 선택 인풋 */}
      <input
        type="file"
        ref={docxInputRef}
        onChange={handleDocxFileChange}
        accept=".docx,.doc,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/msword"
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
                onWheel={(e) => (e.target as HTMLElement).blur()}
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

          {/* 플로팅 창으로 열기 (창 분리 / 새 창 팝업) 버튼 */}
          {onOpenFloatingNote && note && (
            <div className="flex items-center gap-1 shrink-0">
              <button
                type="button"
                onClick={() => onOpenFloatingNote(note.id, false)}
                className="p-1.5 bg-slate-900 hover:bg-slate-800 text-slate-300 hover:text-indigo-300 border border-slate-800 rounded-lg text-xs font-medium transition-all shadow-sm active:scale-95 shrink-0 flex items-center gap-1"
                title="이 메모를 화면 내 플로팅(창 분리) 창으로 열기"
              >
                <Layers size={13} className="text-indigo-400" />
                <span className="hidden sm:inline text-[11px]">창 분리</span>
              </button>
              <button
                type="button"
                onClick={() => onOpenFloatingNote(note.id, true)}
                className="p-1.5 bg-slate-900 hover:bg-slate-800 text-indigo-300 hover:text-white border border-indigo-500/30 rounded-lg text-xs font-medium transition-all shadow-sm active:scale-95 shrink-0 flex items-center gap-1"
                title="이 메모를 브라우저 밖 독립 새 창(OS 팝업)으로 분리 (바탕화면/다른 모니터로 이동)"
              >
                <ExternalLink size={13} className="text-indigo-400" />
                <span className="hidden sm:inline text-[11px]">새 창</span>
              </button>
            </div>
          )}

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
            {/* 접이식 토글 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={insertToggle}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-amber-300 text-amber-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="접이식 토글 목록 삽입"
            >
              <ChevronRight size={14} />
            </button>
            {/* 콜아웃 박스 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => insertCallout('NOTE')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-blue-300 text-blue-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="콜아웃 강조 박스 삽입"
            >
              <Info size={13} />
            </button>
            {/* 위키링크 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={insertWikiLink}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-indigo-300 text-indigo-400 bg-indigo-500/10 border border-indigo-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="메모 연결 링크([[ ]]) 삽입"
            >
              <Link2 size={12} />
              <span>[[ ]]</span>
            </button>
            {/* 슬래시 빠른 서식 메뉴 열기 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                const textarea = textareaRef.current;
                if (!textarea) return;
                insertMarkdown('/', '');
                setTimeout(() => {
                  checkSlashCommand(textarea);
                }, 50);
              }}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-orange-300 text-orange-400 bg-orange-500/10 border border-orange-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="빠른 서식 메뉴 열기 (/)"
            >
              <Sparkles size={12} />
              <span>/ 서식</span>
            </button>
            {/* 실무 서식 템플릿 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setIsTemplateModalOpen(true)}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-violet-300 text-violet-400 bg-violet-500/10 border border-violet-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="실무 서식 템플릿 삽입"
            >
              <LayoutTemplate size={12} />
              <span>템플릿</span>
            </button>
            {/* 지식 연결망 & 백링크 토글 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setIsBacklinksExpanded((prev) => !prev)}
              className={`h-7 px-2 flex items-center gap-1 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90 ${
                isBacklinksExpanded
                  ? 'bg-indigo-600 text-white shadow-xs'
                  : backlinks.length > 0
                    ? 'hover:bg-slate-800 text-indigo-300 bg-indigo-500/15 border border-indigo-500/30'
                    : 'hover:bg-slate-800 text-slate-400 hover:text-slate-200'
              }`}
              title="지식 연결망 및 백링크 목록 열기/닫기"
            >
              <Link2 size={12} />
              <span>백링크{backlinks.length > 0 ? ` (${backlinks.length})` : ''}</span>
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

            {/* 빈 엑셀 스프레드시트 삽입 및 바로 편집 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleCreateBlankExcel()}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-emerald-300 text-emerald-400 bg-emerald-500/10 border border-emerald-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="빈 엑셀 스프레드시트 삽입 및 바로 편집"
            >
              <Plus size={11} className="text-emerald-400" />
              <FileSpreadsheet size={12} className="text-emerald-400" />
              <span className="hidden xs:inline">새 엑셀</span>
            </button>

            {/* 엑셀/CSV 문서 첨부 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => excelInputRef.current?.click()}
              disabled={isProcessingExcel}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-emerald-300 text-emerald-400 bg-emerald-500/10 border border-emerald-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="커서 위치에 기존 엑셀/CSV(XLSX, XLS, CSV) 파일 불러오기 첨부"
            >
              {isProcessingExcel ? (
                <Loader2 size={12} className="animate-spin text-emerald-400" />
              ) : (
                <FileSpreadsheet size={12} className="text-emerald-400" />
              )}
              <span className="hidden xs:inline">엑셀 첨부</span>
            </button>

            {/* 워드(DOCX) 문서 첨부 버튼 */}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => docxInputRef.current?.click()}
              disabled={isProcessingDocx}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-blue-300 text-blue-400 bg-blue-500/10 border border-blue-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="커서 위치에 워드(DOCX, DOC) 문서 파일 첨부"
            >
              {isProcessingDocx ? (
                <Loader2 size={12} className="animate-spin text-blue-400" />
              ) : (
                <FileText size={12} className="text-blue-400" />
              )}
              <span className="hidden xs:inline">워드</span>
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

      {/* 빈 메모일 때 원클릭 추천 템플릿 칩 */}
      {!localContent.trim() && (
        <div className="px-6 py-1.5 flex flex-wrap items-center gap-1.5 text-xs text-slate-400 select-none animate-fadeIn">
          <span className="text-[11px] text-slate-500 flex items-center gap-1 mr-1">
            <Sparkles size={11} className="text-violet-400" />
            <span>추천 서식:</span>
          </span>
          {getNoteTemplates().slice(0, 4).map((tpl) => (
            <button
              key={tpl.id}
              type="button"
              onClick={() => insertTemplate(tpl)}
              className="px-2.5 py-1 rounded-lg bg-slate-900 border border-slate-700/80 hover:border-violet-500/50 hover:bg-violet-950/20 text-slate-300 hover:text-violet-300 transition-all text-xs font-medium cursor-pointer shadow-xs active:scale-95"
            >
              {tpl.title}
            </button>
          ))}
          <button
            type="button"
            onClick={() => setIsTemplateModalOpen(true)}
            className="px-2.5 py-1 rounded-lg bg-slate-800/60 hover:bg-slate-800 text-slate-400 hover:text-slate-200 transition-colors text-[11px] cursor-pointer"
          >
            + 전체 서식
          </button>
        </div>
      )}

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
            className={`${viewMode === 'split' ? 'shrink-0 pr-3' : 'flex-1'} flex flex-col h-full overflow-hidden relative`}
          >
            <textarea
              ref={textareaRef}
              value={localContent}
              onChange={handleContentChange}
              onKeyDown={handleKeyDown}
              onKeyUp={(e) => {
                if (['ArrowUp', 'ArrowDown', 'Enter', 'Escape'].includes(e.key)) return;
                checkSlashCommand(e.currentTarget);
              }}
              onClick={(e) => checkSlashCommand(e.currentTarget)}
              onPaste={handlePaste}
              onScroll={handleEditorScroll}
              placeholder="마크다운 문법으로 자유롭게 작성해 보세요... ('/'를 입력하면 빠른 서식 메뉴가 열립니다)"
              className="w-full flex-1 bg-transparent text-slate-200 text-sm md:text-base leading-relaxed resize-none focus:outline-none font-mono placeholder-slate-600 scrollbar-thin"
              spellCheck={false}
            />

            {/* 슬래시(/) 명령어 자동완성 팝업 메뉴 */}
            {slashMenuState?.isOpen && (
              <SlashCommandMenu
                isOpen={slashMenuState.isOpen}
                query={slashMenuState.query}
                selectedIndex={slashSelectedIndex}
                onSelectCommand={executeSlashCommand}
                onClose={() => setSlashMenuState(null)}
                position={slashMenuState.position}
              />
            )}
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

      {/* 4단: 지식 연결망 & 백링크(Backlinks) 패널 */}
      {note && (
        <div className="border-t border-slate-800/80 bg-slate-950/70 backdrop-blur-xs select-none shrink-0 z-20">
          {/* 패널 헤더 바 */}
          <div className="px-4 md:px-6 py-2 flex items-center justify-between hover:bg-slate-900/60 transition-colors text-xs text-slate-400">
            <button
              type="button"
              onClick={() => setIsBacklinksExpanded((prev) => !prev)}
              className="flex items-center gap-2 font-medium group cursor-pointer focus:outline-none"
            >
              <Link2 size={13} className="text-indigo-400 group-hover:scale-110 transition-transform" />
              <span className="text-slate-300 font-semibold">지식 연결망</span>

              {/* 백링크 뱃지 */}
              <span className={`px-2 py-0.5 rounded-full text-[11px] font-semibold transition-colors ${
                backlinks.length > 0 
                  ? 'bg-indigo-500/25 text-indigo-300 border border-indigo-500/40' 
                  : 'bg-slate-800/80 text-slate-500'
              }`}>
                백링크 {backlinks.length}
              </span>

              {/* 나가는 링크 뱃지 */}
              {outlinks.length > 0 && (
                <span className="px-2 py-0.5 rounded-full text-[11px] font-medium bg-slate-800/80 text-slate-400 border border-slate-700/60">
                  링크 {outlinks.length}
                </span>
              )}

              {/* 태그 뱃지 */}
              {tagsInNote.length > 0 && (
                <span className="px-2 py-0.5 rounded-full text-[11px] font-medium bg-emerald-500/15 text-emerald-300 border border-emerald-500/30">
                  #{tagsInNote.length}
                </span>
              )}
            </button>

            <button
              type="button"
              onClick={() => setIsBacklinksExpanded((prev) => !prev)}
              className="flex items-center gap-1 text-[11px] text-slate-500 hover:text-slate-300 transition-colors cursor-pointer"
            >
              <span>{isBacklinksExpanded ? '접기' : '펼쳐보기'}</span>
              <ChevronRight
                size={14}
                className={`transition-transform duration-200 ${isBacklinksExpanded ? 'rotate-90 text-indigo-400' : ''}`}
              />
            </button>
          </div>

          {/* 확장된 상세 정보 영역 */}
          {isBacklinksExpanded && (
            <div className="px-4 md:px-6 py-3 border-t border-slate-800/50 bg-slate-900/60 max-h-56 overflow-y-auto space-y-3 scrollbar-thin">
              {/* 1. 백링크 (이 메모를 인용하는 메모들) */}
              <div>
                <div className="text-[11px] font-semibold text-indigo-300 mb-1.5 flex items-center gap-1.5">
                  <span>🔗 이 메모를 참조하는 메모 (백링크 {backlinks.length}개)</span>
                </div>
                {backlinks.length === 0 ? (
                  <p className="text-xs text-slate-500 pl-1">
                    다른 메모에서 <code className="text-indigo-400 bg-slate-800/80 px-1 py-0.5 rounded font-mono">[[{note.title}]]</code> 형태로 이 메모를 인용하면 여기에 연결 목록이 표시됩니다.
                  </p>
                ) : (
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                    {backlinks.map((bNote) => (
                      <button
                        key={bNote.id}
                        type="button"
                        onClick={() => onSelectNote?.(bNote.id)}
                        className="p-2.5 rounded-xl bg-slate-900/90 border border-indigo-500/25 hover:border-indigo-400/60 hover:bg-slate-800/90 text-left transition-all group/item cursor-pointer shadow-xs"
                      >
                        <div className="text-xs font-semibold text-slate-200 group-hover/item:text-indigo-300 flex items-center justify-between truncate">
                          <span className="truncate">{bNote.title || '제목 없음'}</span>
                          <span className="text-[10px] text-indigo-400 opacity-70 shrink-0 ml-1">이동 ↗</span>
                        </div>
                        <div className="text-[11px] text-slate-500 truncate mt-1">
                          {bNote.content.replace(/[#*`~>-]/g, '').slice(0, 50) || '내용 없음'}
                        </div>
                      </button>
                    ))}
                  </div>
                )}
              </div>

              {/* 2. 나가는 링크 (이 메모에서 인용 중인 링크들) */}
              {outlinks.length > 0 && (
                <div className="pt-2 border-t border-slate-800/40">
                  <div className="text-[11px] font-semibold text-slate-400 mb-1.5">
                    ↗️ 이 메모에서 연결한 링크 ({outlinks.length}개)
                  </div>
                  <div className="flex flex-wrap gap-1.5">
                    {outlinks.map((link, idx) => (
                      <button
                        key={idx}
                        type="button"
                        onClick={() => {
                          if (link.exists && link.noteId) {
                            onSelectNote?.(link.noteId);
                          } else if (onCreateNoteWithTitle) {
                            if (window.confirm(`'${link.title}' 메모가 아직 없습니다. 새 메모로 생성하시겠습니까?`)) {
                              onCreateNoteWithTitle(link.title);
                            }
                          }
                        }}
                        className={`px-2.5 py-1 rounded-lg text-xs font-medium flex items-center gap-1 transition-colors cursor-pointer ${
                          link.exists
                            ? 'bg-indigo-500/15 text-indigo-300 border border-indigo-500/30 hover:bg-indigo-500/25'
                            : 'bg-slate-800/80 text-slate-400 border border-dashed border-slate-600 hover:text-indigo-300 hover:border-indigo-400'
                        }`}
                      >
                        <span>{link.title}</span>
                        {!link.exists && <span className="text-[10px] text-indigo-400">+생성</span>}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {/* 3. 본문 내 해시태그 목록 */}
              {tagsInNote.length > 0 && (
                <div className="pt-2 border-t border-slate-800/40">
                  <div className="text-[11px] font-semibold text-emerald-400/90 mb-1.5">
                    🏷️ 이 메모의 태그 ({tagsInNote.length}개)
                  </div>
                  <div className="flex flex-wrap gap-1.5">
                    {tagsInNote.map((t) => (
                      <button
                        key={t}
                        type="button"
                        onClick={() => onSearchTag?.(`#${t}`)}
                        className="px-2 py-0.5 rounded-full text-xs font-medium bg-emerald-500/15 text-emerald-300 border border-emerald-500/30 hover:bg-emerald-500/30 transition-colors cursor-pointer"
                        title={`#${t} 태그로 전체 메모 검색`}
                      >
                        #{t}
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}

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
          attId={activeHwpModal.attId}
          initialPage={activeHwpModal.initialPage}
          onOpenEditor={() => {
            const attId = activeHwpModal.attId || Object.keys(localAttachments).find(
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
            // attId를 다양한 경로로 완벽하게 보장
            let attId = activeHwpEditorModal.attId;
            if (!attId || !localAttachmentsRef.current[attId]) {
              attId = Object.keys(localAttachmentsRef.current).find(
                (id) => localAttachmentsRef.current[id].name === activeHwpEditorModal.title
              ) || Object.keys(localAttachmentsRef.current).find(
                (id) => {
                  const name = (localAttachmentsRef.current[id].name || '').toLowerCase();
                  return name.endsWith('.hwp') || name.endsWith('.hwpx');
                }
              ) || '';
            }

            if (attId) {
              // 1. 기존 HWP 캐시(전체 및 ID별) 및 Blob URL 캐시 즉시 무효화 (새 파일 재로드 100% 보장)
              invalidateHwpCache(attId);
              invalidateHwpCache(); // 전체 캐시도 깨끗하게 초기화
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

              // 4. 만약 전면 보기 모달(activeHwpModal)이 열려있다면 해당 데이터도 최신으로 즉시 갱신
              setActiveHwpModal((prev) => prev && (prev.attId === attId || prev.title === activeHwpEditorModal.title) ? { ...prev, data: updatedDataUrl } : prev);

              // 5. 인라인 뷰어에 새 데이터 URL을 즉시 전달하여 재로드 & 재렌더링
              renderInlineHwp.current(attId, updatedDataUrl);

              // 6. 마크다운 HTML 재계산 및 DOM 갱신 후 한 번 더 보장 렌더링
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

      {/* 엑셀/CSV 스프레드시트 전면 뷰어 모달 */}
      {activeExcelModal && (
        <ExcelViewerModal
          isOpen={true}
          onClose={() => setActiveExcelModal(null)}
          data={activeExcelModal.data}
          title={activeExcelModal.title}
          attId={activeExcelModal.attId}
          onOpenEditor={() => {
            const attId = activeExcelModal.attId || Object.keys(localAttachments).find(
              (id) => localAttachments[id].data === activeExcelModal.data
            ) || '';
            setActiveExcelEditorModal({
              attId,
              data: activeExcelModal.data,
              title: activeExcelModal.title,
            });
            setActiveExcelModal(null);
          }}
          onImportTable={(mdTable) => {
            const insertBlock = `\n\n### 📊 ${activeExcelModal.title} 표\n\n${mdTable}\n\n`;
            const newContent = `${localContent}${insertBlock}`;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            alert(`'${activeExcelModal.title}' 표 데이터를 마크다운 표로 메모에 성공적으로 가져왔습니다!`);
          }}
        />
      )}

      {/* 엑셀/CSV 웹 직접 편집기 모달 */}
      {activeExcelEditorModal && (
        <ExcelEditorModal
          isOpen={true}
          onClose={() => setActiveExcelEditorModal(null)}
          data={activeExcelEditorModal.data}
          title={activeExcelEditorModal.title}
          attId={activeExcelEditorModal.attId}
          onSave={async (updatedDataUrl) => {
            let attId = activeExcelEditorModal.attId;
            if (!attId || !localAttachmentsRef.current[attId]) {
              attId = Object.keys(localAttachmentsRef.current).find(
                (id) => localAttachmentsRef.current[id].name === activeExcelEditorModal.title
              ) || Object.keys(localAttachmentsRef.current).find(
                (id) => {
                  const name = (localAttachmentsRef.current[id].name || '').toLowerCase();
                  return name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv');
                }
              ) || '';
            }

            if (attId) {
              // 1. 엑셀 캐시(전체 및 ID별) 및 Blob URL 캐시 즉시 무효화 (새 파일 재로드 100% 보장)
              invalidateExcelCache(attId);
              invalidateExcelCache();
              invalidateAttachmentBlobUrl(attId);

              // 2. 첨부파일 목록 상태 업데이트 및 상위 메모 저장
              const updatedAtts = {
                ...localAttachmentsRef.current,
                [attId]: {
                  ...(localAttachmentsRef.current[attId] || {}),
                  data: updatedDataUrl,
                },
              };
              setLocalAttachments({ ...updatedAtts });
              localAttachmentsRef.current = updatedAtts;
              onUpdateNote({ attachments: updatedAtts });

              // 3. 편집 모달 내부 데이터 동기화
              setActiveExcelEditorModal((prev) => prev ? { ...prev, data: updatedDataUrl } : null);

              // 4. 전면 뷰어 모달(activeExcelModal)이 열려있다면 최신 데이터 동기화
              setActiveExcelModal((prev) => prev && (prev.attId === attId || prev.title === activeExcelEditorModal.title) ? { ...prev, data: updatedDataUrl } : prev);
            }
          }}
        />
      )}

      {/* 워드(DOCX) 전면 뷰어 모달 */}
      {activeDocxModal && (
        <DocxViewerModal
          isOpen={true}
          onClose={() => setActiveDocxModal(null)}
          data={activeDocxModal.data}
          title={activeDocxModal.title}
          attId={activeDocxModal.attId}
          onOpenEditor={() => {
            const attId = activeDocxModal.attId || Object.keys(localAttachments).find(
              (id) => localAttachments[id].data === activeDocxModal.data
            ) || '';
            setActiveDocxEditorModal({
              attId,
              data: activeDocxModal.data,
              title: activeDocxModal.title,
            });
            setActiveDocxModal(null);
          }}
          onImportText={(text) => {
            const insertBlock = `\n\n### 📄 ${activeDocxModal.title} 내용\n\n${text}\n\n`;
            const newContent = `${localContent}${insertBlock}`;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            alert(`'${activeDocxModal.title}' 문서 내용(${text.length}자)을 메모에 성공적으로 가져왔습니다!`);
          }}
        />
      )}

      {/* 워드(DOCX) 웹 직접 편집기 모달 */}
      {activeDocxEditorModal && (
        <DocxEditorModal
          isOpen={true}
          onClose={() => setActiveDocxEditorModal(null)}
          data={activeDocxEditorModal.data}
          title={activeDocxEditorModal.title}
          attId={activeDocxEditorModal.attId}
          onSave={async (updatedDataUrl) => {
            const attId = activeDocxEditorModal.attId;
            if (attId) {
              // 1. 워드 캐시 및 Blob URL 캐시 무효화
              invalidateDocxCache(attId);
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
              setActiveDocxEditorModal((prev) => prev ? { ...prev, data: updatedDataUrl } : null);

              // 4. 인라인 뷰어 재렌더링
              renderInlineDocx.current(attId, updatedDataUrl);
              setTimeout(() => {
                renderInlineDocx.current(attId, updatedDataUrl);
              }, 120);
            }
          }}
        />
      )}

      {/* 파워포인트(PPTX) 슬라이드 뷰어 및 슬라이드 쇼 모달 */}
      {activePptxModal && (
        <PptxViewerModal
          isOpen={true}
          onClose={() => setActivePptxModal(null)}
          data={activePptxModal.data}
          title={activePptxModal.title}
          attId={activePptxModal.attId}
          onOpenEditor={() => {
            const attId = activePptxModal.attId || Object.keys(localAttachments).find(
              (id) => localAttachments[id].data === activePptxModal.data
            ) || '';
            setActivePptxEditorModal({
              attId,
              data: activePptxModal.data,
              title: activePptxModal.title,
            });
            setActivePptxModal(null);
          }}
          onImportMarkdown={(markdown) => {
            const insertBlock = `\n\n${markdown}\n\n`;
            const newContent = `${localContent}${insertBlock}`;
            setLocalContent(newContent);
            debouncedSyncToParent({ content: newContent });
            alert(`'${activePptxModal.title}' 슬라이드 내용을 메모 본문에 삽입했습니다!`);
          }}
        />
      )}

      {/* 파워포인트(PPTX) 슬라이드 웹 직접 편집기 모달 */}
      {activePptxEditorModal && (
        <PptxEditorModal
          isOpen={true}
          onClose={() => setActivePptxEditorModal(null)}
          data={activePptxEditorModal.data}
          title={activePptxEditorModal.title}
          attId={activePptxEditorModal.attId}
          onSave={async (updatedDataUrl) => {
            const attId = activePptxEditorModal.attId;
            if (attId) {
              invalidatePptxCache(attId);
              invalidateAttachmentBlobUrl(attId);

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
              setActivePptxEditorModal((prev) => prev ? { ...prev, data: updatedDataUrl } : null);
            }
          }}
        />
      )}

      {/* 바이너리/HEX 울트라에디트 모달 */}
      {activeHexModal && (
        <HexEditorModal
          isOpen={true}
          onClose={() => setActiveHexModal(null)}
          data={activeHexModal.data}
          title={activeHexModal.title}
          attId={activeHexModal.attId}
          onSave={async (updatedDataUrl) => {
            const attId = activeHexModal.attId;
            if (attId) {
              invalidateAttachmentBlobUrl(attId);
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
              setActiveHexModal((prev) => prev ? { ...prev, data: updatedDataUrl } : null);
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

      {/* 실무 서식 템플릿 선택 모달 */}
      <TemplateModal
        isOpen={isTemplateModalOpen}
        onClose={() => setIsTemplateModalOpen(false)}
        onSelectTemplate={insertTemplate}
      />
    </div>
  );
};
