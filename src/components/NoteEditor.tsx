import React, { useState, useRef, useEffect, useMemo } from 'react';
import type { Note, Project } from '../types';
import { marked } from 'marked';
import { compressImageToDataUrl } from '../utils/image';
import { processFileForNote } from '../utils/file';
import { cleanLegacyAttachments, renderAttachmentsInHtml } from '../utils/attachment';
import { exportNoteToPdf } from '../utils/exportPdf';
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
  FileText
} from 'lucide-react';

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
  const [viewMode, setViewMode] = useState<'edit' | 'preview' | 'split'>('edit');
  const [isProcessingImage, setIsProcessingImage] = useState(false);
  const [isProcessingPdf, setIsProcessingPdf] = useState(false);
  const [isDragOverEditor, setIsDragOverEditor] = useState(false);
  const [showShortcutsModal, setShowShortcutsModal] = useState(false);

  // 한글 IME 조합 깨짐 방지를 위한 로컬 버퍼 상태
  const [localTitle, setLocalTitle] = useState(note?.title || '');
  const [localContent, setLocalContent] = useState(note?.content || '');
  const [localAttachments, setLocalAttachments] = useState<Record<string, AttachmentItem>>(note?.attachments || {});

  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);
  const updateTimerRef = useRef<any>(null);
  const currentNoteIdRef = useRef(note?.id);
  const lastSentContentRef = useRef<string | null>(null);
  const lastSentTitleRef = useRef<string | null>(null);

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
        // 내가 직접 타이핑 중인 내용이 아닐 때(외부 동기화로 변경되었을 때)만 덮어쓰기
        if (note.title !== localTitle && note.title !== lastSentTitleRef.current) {
          setLocalTitle(note.title);
        }
        if (note.content !== localContent && note.content !== lastSentContentRef.current) {
          const { content: cleanedContent, attachments: cleanedAttachments } = cleanLegacyAttachments(
            note.content,
            note.attachments
          );
          setLocalContent(cleanedContent);
          setLocalAttachments(cleanedAttachments);
          pushHistory(cleanedContent, 0);
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

  // 화면 너비에 따라 기본 뷰 모드 자동 조정
  useEffect(() => {
    if (window.innerWidth >= 1280) {
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
        textareaRef.current?.focus();
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
    pushHistory(newContent, start + prefix.length + (selected.length || 2));

    setTimeout(() => {
      textarea.focus();
      textarea.setSelectionRange(
        start + prefix.length,
        start + prefix.length + (selected.length || 2)
      );
    }, 0);
  };

  // 체크박스(- [ ] <-> - [x]) 토글 함수 (현재 커서 라인 또는 선택된 여러 줄 일괄 토글)
  const toggleCheckboxAtCursor = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = localContent;

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
      textarea.focus();
      const lengthDiff = replacedBlock.length - selectedBlock.length;
      textarea.setSelectionRange(
        Math.max(lineStart, start),
        Math.max(lineStart, end + lengthDiff)
      );
    }, 0);
  };

  // 들여쓰기 2칸 함수 (아이콘 클릭 및 Tab 키 공용)
  const handleIndent = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = localContent;

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
      textarea.focus();
      const added = lines.length * 2;
      textarea.setSelectionRange(start + 2, end + added);
    }, 0);
  };

  // 내어쓰기 2칸 함수 (아이콘 클릭 및 Shift+Tab 키 공용)
  const handleOutdent = () => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = localContent;

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
      textarea.focus();
      textarea.setSelectionRange(
        Math.max(lineStart, start - 2),
        Math.max(lineStart, end - removedTotal)
      );
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
          textarea.focus();
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
          textarea.focus();
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

  // 표 서식 템플릿 삽입
  const insertTable = () => {
    const tableTpl = '\n\n| 제목 | 내용 | 상태 |\n| :--- | :--- | :---: |\n| 항목1 | 내용1 | 완료 |\n\n';
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
    insertMarkdown('1. ');
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
      // Ctrl + Shift + T 또는 Alt + T: To-Do 체크리스트
      if ((e.key.toLowerCase() === 't' && e.shiftKey) || (e.altKey && e.key.toLowerCase() === 't')) {
        e.preventDefault();
        insertMarkdown('- [ ] ');
        return;
      }
      // Ctrl + Shift + L: 글머리 기호 목록
      if (e.key.toLowerCase() === 'l' && e.shiftKey) {
        e.preventDefault();
        insertMarkdown('- ');
        return;
      }
      // Ctrl + 1: 큰 제목 (# )
      if (e.key === '1') {
        e.preventDefault();
        insertMarkdown('# ');
        return;
      }
      // Ctrl + 2: 중간 제목 (## )
      if (e.key === '2') {
        e.preventDefault();
        insertMarkdown('## ');
        return;
      }
      // Ctrl + 3: 소제목 (### )
      if (e.key === '3') {
        e.preventDefault();
        insertMarkdown('### ');
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

    const isPdf = mimeType === 'application/pdf' || attName.endsWith('.pdf');
    const insertTag = isPdf
      ? `\n\n[📄 PDF: ${safeName}](attachment:${attId})\n\n`
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
      textarea.focus();
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

  const renderMarkdown = (content: string) => {
    try {
      // 마크다운 체크박스(- [ ] 또는 - [x]) 뒤에 띄어쓰기를 빠뜨린 경우에도 정상 체크박스로 렌더링되도록 자동 보정
      const normalizedContent = (content || '').replace(/^(\s*-\s*\[[ xX]\])([^\s\n])/gm, '$1 $2');
      let rawHtml = marked.parse(normalizedContent, { breaks: true, gfm: true }) as string;

      // 표(table)를 모바일 가로 스크롤 컨테이너로 감싸 좁은 화면에서 컬럼이 찌그러지는 현상 방지
      rawHtml = rawHtml.replace(/<table>([\s\S]*?)<\/table>/g, '<div class="table-scroll-container"><table>$1</table></div>');

      return { __html: renderAttachmentsInHtml(rawHtml, localAttachments) };
    } catch {
      return { __html: '<p>마크다운을 렌더링할 수 없습니다.</p>' };
    }
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
      <div className="p-3 border-b border-slate-800/80 flex flex-wrap items-center justify-between gap-2 bg-slate-950/70">
        <div className="flex items-center gap-2 min-w-0">
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
            className="flex items-center gap-1.5 bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-300 max-w-[240px] md:max-w-md truncate" 
            title="현재 메모의 소속 폴더를 다른 폴더로 이동할 수 있습니다."
          >
            <Folder size={12} className="text-indigo-400 shrink-0" />
            <span className="text-[11px] text-slate-400 font-medium shrink-0">폴더:</span>
            <select
              value={note.projectId}
              onChange={(e) => onUpdateNote({ projectId: e.target.value })}
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

        {/* 오른쪽: 기기 전용(로컬) 토글 & 뷰 모드 */}
        <div className="flex items-center gap-2 shrink-0">
          <button
            onClick={() => onUpdateNote({ isLocalOnly: !note.isLocalOnly })}
            className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-medium border transition-all ${
              note.isLocalOnly
                ? 'bg-amber-500/20 text-amber-300 border-amber-500/50 shadow-sm shadow-amber-500/10'
                : 'bg-slate-900 text-slate-400 border-slate-800 hover:text-slate-200'
            }`}
            title={note.isLocalOnly ? '클라우드 동기화 안 됨 (이 기기에만 보관)' : 'GitHub 클라우드에 동기화됨'}
          >
            {note.isLocalOnly ? (
              <>
                <Lock size={13} className="text-amber-400" />
                <span className="font-semibold">이 기기에만 저장됨</span>
              </>
            ) : (
              <>
                <Unlock size={13} className="text-slate-500" />
                <span>GitHub 동기화 대상</span>
              </>
            )}
          </button>

          <div className="flex items-center bg-slate-900 border border-slate-800 rounded-lg p-0.5 text-slate-400">
            <button
              onClick={() => setViewMode('edit')}
              className={`p-1.5 rounded text-xs transition-colors ${
                viewMode === 'edit' ? 'bg-indigo-600 text-white' : 'hover:text-slate-200'
              }`}
              title="편집 모드 (단축키: Alt + 1 또는 Ctrl + M)"
            >
              <Edit3 size={13} />
            </button>
            <button
              onClick={() => setViewMode('preview')}
              className={`p-1.5 rounded text-xs transition-colors ${
                viewMode === 'preview' ? 'bg-indigo-600 text-white' : 'hover:text-slate-200'
              }`}
              title="미리보기 모드 (단축키: Alt + 2 또는 Ctrl + M)"
            >
              <Eye size={13} />
            </button>
            <button
              onClick={() => setViewMode('split')}
              className={`hidden lg:block p-1.5 rounded text-xs transition-colors ${
                viewMode === 'split' ? 'bg-indigo-600 text-white' : 'hover:text-slate-200'
              }`}
              title="분할 보기 모드 (단축키: Alt + 3 또는 Ctrl + M)"
            >
              <Columns size={13} />
            </button>
          </div>

          {/* 📄 PDF 내보내기 / 인쇄 버튼 */}
          <button
            onClick={() => {
              const breadcrumbText = folderBreadcrumbs.map((b) => b.name).join(' > ');
              exportNoteToPdf(
                { ...note, title: localTitle, content: localContent, attachments: localAttachments },
                breadcrumbText,
                localContent
              );
            }}
            className="flex items-center gap-1 px-2.5 py-1 bg-slate-900 hover:bg-slate-800 text-slate-300 hover:text-indigo-300 border border-slate-800 rounded-lg text-xs font-medium transition-all shadow-sm active:scale-95"
            title="현재 메모를 깔끔한 PDF 문서로 내보내기 / 인쇄"
          >
            <FileDown size={13} className="text-indigo-400" />
            <span className="hidden sm:inline">PDF 내보내기</span>
          </button>

          {/* 메모 삭제 버튼 */}
          {onDeleteNote && (
            <button
              onClick={() => onDeleteNote(note.id)}
              className="p-1.5 text-slate-400 hover:text-rose-400 hover:bg-slate-900 border border-slate-800 rounded-lg transition-colors"
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
              onClick={() => insertMarkdown('# ')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="큰 제목 H1 (Ctrl + 1)"
            >
              <Heading1 size={14} />
            </button>
            {/* H2 */}
            <button
              type="button"
              onClick={() => insertMarkdown('## ')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="중간 제목 H2 (Ctrl + 2)"
            >
              <Heading2 size={14} />
            </button>
            {/* H3 */}
            <button
              type="button"
              onClick={() => insertMarkdown('### ')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="소제목 H3 (Ctrl + 3)"
            >
              <Heading3 size={14} />
            </button>

            <div className="w-[1px] h-3.5 bg-slate-800/80 mx-1 shrink-0" />

            {/* Bold */}
            <button
              type="button"
              onClick={() => insertMarkdown('**', '**')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="굵게 Bold (Ctrl + B)"
            >
              <Bold size={13} />
            </button>
            {/* Italic */}
            <button
              type="button"
              onClick={() => insertMarkdown('*', '*')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="기울임 Italic (Ctrl + I)"
            >
              <Italic size={13} />
            </button>
            {/* Strikethrough */}
            <button
              type="button"
              onClick={() => insertMarkdown('~~', '~~')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="취소선 (Ctrl + Shift + X)"
            >
              <Strikethrough size={13} />
            </button>
            {/* Inline Code */}
            <button
              type="button"
              onClick={() => insertMarkdown('`', '`')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="인라인 코드 (Ctrl + Shift + K)"
            >
              <Code size={13} />
            </button>

            <div className="w-[1px] h-3.5 bg-slate-800/80 mx-1 shrink-0" />

            {/* Bullet List */}
            <button
              type="button"
              onClick={() => insertMarkdown('- ')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="글머리 기호 목록 (Ctrl + Shift + L)"
            >
              <List size={13} />
            </button>
            {/* Numbered List */}
            <button
              type="button"
              onClick={insertOrderedList}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="번호 매기기 목록 (1. 2. 3.)"
            >
              <ListOrdered size={13} />
            </button>
            {/* Indent (Tab) */}
            <button
              type="button"
              onClick={handleIndent}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-slate-100 rounded-md transition-colors shrink-0 active:scale-90"
              title="들여쓰기 2칸 (Tab)"
            >
              <Indent size={13} />
            </button>
            {/* Outdent (Shift + Tab) */}
            <button
              type="button"
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
              onClick={handleUndo}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-amber-300 text-slate-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="실행 취소 Undo (Ctrl + Z)"
            >
              <Undo2 size={13} />
            </button>
            {/* 다시 실행 (Redo) */}
            <button
              type="button"
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
              onClick={() => insertMarkdown('- [ ] ')}
              className="h-7 px-2 flex items-center gap-1 hover:bg-slate-800 hover:text-indigo-300 text-indigo-400 bg-indigo-950/40 border border-indigo-500/20 rounded-md text-[11px] font-medium transition-colors shrink-0 active:scale-90"
              title="To-Do 체크리스트 추가 (Ctrl + Shift + T)"
            >
              <CheckSquare size={12} />
              <span>To-Do</span>
            </button>
            {/* 체크박스 토글 */}
            <button
              type="button"
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
              onClick={insertTable}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-cyan-300 text-cyan-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="마크다운 표(Table) 서식 삽입"
            >
              <Table size={13} />
            </button>
            {/* 링크 */}
            <button
              type="button"
              onClick={insertLink}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-sky-300 text-sky-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="하이퍼링크 삽입 (Ctrl + K)"
            >
              <Link2 size={13} />
            </button>
            {/* 코드 블록 */}
            <button
              type="button"
              onClick={insertCodeBlock}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-purple-300 text-purple-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="코드 블록(```) 삽입"
            >
              <FileCode size={13} />
            </button>
            {/* 인용구 */}
            <button
              type="button"
              onClick={() => insertMarkdown('> ')}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 hover:text-amber-300 text-amber-400 rounded-md transition-colors shrink-0 active:scale-90"
              title="인용구(>) 삽입"
            >
              <Quote size={13} />
            </button>
            {/* 구분선 */}
            <button
              type="button"
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

            {/* GitHub 동기화 버튼 (Ctrl + S) */}
            {onSync && (
              <button
                type="button"
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
              onClick={() => setShowShortcutsModal(true)}
              className="w-7 h-7 flex items-center justify-center hover:bg-slate-800 text-slate-400 hover:text-indigo-400 rounded-md transition-colors ml-auto shrink-0"
              title="PC 마크다운 단축키 안내 (Ctrl + /)"
            >
              <Keyboard size={13} />
            </button>
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

      {/* 본문 영역: 편집창 / 미리보기 / 분할 */}
      <div className="flex-1 flex overflow-hidden p-4 md:px-6 gap-6">
        {(viewMode === 'edit' || viewMode === 'split') && (
          <div className={`flex-1 flex flex-col h-full ${viewMode === 'split' ? 'border-r border-slate-800/80 pr-6' : ''}`}>
            <textarea
              ref={textareaRef}
              value={localContent}
              onChange={handleContentChange}
              onKeyDown={handleKeyDown}
              onPaste={handlePaste}
              placeholder="마크다운 문법으로 자유롭게 작성해 보세요... 이미지는 Ctrl + V 로 바로 붙여넣을 수 있습니다!"
              className="w-full flex-1 bg-transparent text-slate-200 text-sm md:text-base leading-relaxed resize-none focus:outline-none font-mono placeholder-slate-600 scrollbar-thin"
              spellCheck={false}
            />
          </div>
        )}

        {(viewMode === 'preview' || viewMode === 'split') && (
          <div className="flex-1 h-full overflow-y-auto markdown-preview prose prose-invert max-w-none pr-2">
            <div dangerouslySetInnerHTML={renderMarkdown(localContent)} />
          </div>
        )}
      </div>

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
