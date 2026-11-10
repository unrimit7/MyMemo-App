import React, { useState, useEffect, useRef } from 'react';
import type { Note, Project } from '../types';
import { marked } from 'marked';
import { renderAttachmentsInHtml } from '../utils/attachment';
import { stripFolderPrefixFromTitle } from '../utils/folderPath';
import { PopoutWindow } from './PopoutWindow';
import {
  X,
  Maximize2,
  Minimize2,
  Eye,
  Edit3,
  Columns,
  FileText,
  Lock,
  Copy,
  Check,
  ExternalLink,
} from 'lucide-react';

export interface FloatingNoteWindowProps {
  note: Note;
  allNotes?: Note[];
  projects?: Project[];
  onClose: () => void;
  onUpdateNote: (updated: Partial<Note>, specificNoteId?: string) => void;
  initialX?: number;
  initialY?: number;
  zIndex?: number;
  onFocus?: () => void;
  initialPopout?: boolean;
}

export const FloatingNoteWindow: React.FC<FloatingNoteWindowProps> = ({
  note,
  allNotes = [],
  onClose,
  onUpdateNote,
  initialX = 80,
  initialY = 70,
  zIndex = 60,
  onFocus,
  initialPopout = false,
}) => {
  const [isPopout, setIsPopout] = useState<boolean>(initialPopout);
  const [localContent, setLocalContent] = useState(note.content);
  const [viewMode, setViewMode] = useState<'preview' | 'edit' | 'split'>('preview');
  const [isMaximized, setIsMaximized] = useState(false);
  const [copied, setCopied] = useState(false);

  // 위치 및 크기 상태 (인페이지 플로팅 창 전용)
  const [position, setPosition] = useState({ x: initialX, y: initialY });
  const [size, setSize] = useState({
    width: Math.min(760, typeof window !== 'undefined' ? window.innerWidth - 60 : 760),
    height: Math.min(680, typeof window !== 'undefined' ? window.innerHeight - 100 : 680),
  });

  const [isDragging, setIsDragging] = useState(false);
  const [isResizing, setIsResizing] = useState(false);

  const dragRef = useRef({ startX: 0, startY: 0, initialPosX: 0, initialPosY: 0 });
  const resizeRef = useRef({ startX: 0, startY: 0, initialW: 0, initialH: 0 });
  const syncTimerRef = useRef<any>(null);

  // 외부(메인 에디터)에서 note.content가 변경되었을 때 동기화
  useEffect(() => {
    setLocalContent(note.content);
  }, [note.content]);

  // 상위 메모 업데이트 디바운스
  const handleContentChange = (newContent: string) => {
    setLocalContent(newContent);
    if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    syncTimerRef.current = setTimeout(() => {
      onUpdateNote(
        {
          content: newContent,
          updatedAt: new Date().toISOString(),
        },
        note.id
      );
    }, 400);
  };

  // 창 드래그 시작 (인페이지 플로팅 전용)
  const handleHeaderPointerDown = (e: React.PointerEvent) => {
    if (isMaximized || isPopout) return;
    const target = e.target as HTMLElement;
    if (target.closest('button, input, select, textarea, a, .no-drag')) return;
    if (onFocus) onFocus();

    e.preventDefault();
    setIsDragging(true);
    dragRef.current = {
      startX: e.clientX,
      startY: e.clientY,
      initialPosX: position.x,
      initialPosY: position.y,
    };
  };

  // 창 크기 조절 시작 (인페이지 플로팅 전용)
  const handleResizePointerDown = (e: React.PointerEvent) => {
    if (isMaximized || isPopout) return;
    if (onFocus) onFocus();

    e.preventDefault();
    e.stopPropagation();
    setIsResizing(true);
    resizeRef.current = {
      startX: e.clientX,
      startY: e.clientY,
      initialW: size.width,
      initialH: size.height,
    };
  };

  // 전역 마우스/터치 리스너
  useEffect(() => {
    if (!isDragging && !isResizing) return;

    const handlePointerMove = (e: PointerEvent) => {
      if (isDragging) {
        const deltaX = e.clientX - dragRef.current.startX;
        const deltaY = e.clientY - dragRef.current.startY;
        const newX = dragRef.current.initialPosX + deltaX;
        const newY = dragRef.current.initialPosY + deltaY;

        const maxW = typeof window !== 'undefined' ? window.innerWidth : 1200;
        const maxH = typeof window !== 'undefined' ? window.innerHeight : 900;
        const clampedX = Math.max(-size.width + 100, Math.min(maxW - 80, newX));
        const clampedY = Math.max(0, Math.min(maxH - 50, newY));

        setPosition({ x: clampedX, y: clampedY });
      }

      if (isResizing) {
        const deltaX = e.clientX - resizeRef.current.startX;
        const deltaY = e.clientY - resizeRef.current.startY;
        const newW = resizeRef.current.initialW + deltaX;
        const newH = resizeRef.current.initialH + deltaY;

        const maxW = typeof window !== 'undefined' ? window.innerWidth - position.x - 10 : 1200;
        const maxH = typeof window !== 'undefined' ? window.innerHeight - position.y - 10 : 900;

        setSize({
          width: Math.max(360, Math.min(maxW, newW)),
          height: Math.max(280, Math.min(maxH, newH)),
        });
      }
    };

    const handlePointerUp = () => {
      setIsDragging(false);
      setIsResizing(false);
    };

    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerUp);

    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerUp);
    };
  }, [isDragging, isResizing, size.width, position.x, position.y]);

  // 마크다운 HTML 렌더링
  const renderedHtml = React.useMemo(() => {
    try {
      let normalized = (localContent || '').replace(/^(\s*-\s*\[[ xX]\])([^\s\n])/gm, '$1 $2');
      normalized = normalized.replace(/==([^=\n]+)==/g, '<mark style="background-color: #fef08a; color: #0f172a; padding: 2px 5px; border-radius: 4px;">$1</mark>');
      normalized = normalized.replace(/(<summary>[\s\S]*?<\/summary>)\s*([^\s\n<])/gi, '$1\n\n$2');
      let raw = marked.parse(normalized, { breaks: true, gfm: true }) as string;

      // 위키링크 치환
      raw = raw.replace(/\[\[([^\[\]\n\r]+)\]\]/g, (match, inner) => {
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
          return `<button type="button" class="wikilink-pill wikilink-missing" data-wikilink-create="true" data-wikilink-title="${safeTarget}" title="새 메모: ${safeTarget}">📄 ${safeAlias} <span class="text-[10px] text-indigo-400 opacity-90 ml-0.5 font-normal">+생성</span></button>`;
        }
      });

      return renderAttachmentsInHtml(raw, note.attachments || {});
    } catch {
      return '<p class="text-rose-400">마크다운을 렌더링할 수 없습니다.</p>';
    }
  }, [localContent, note.attachments, allNotes]);

  // 미리보기 내 체크박스 원클릭 토글
  const handlePreviewClick = (e: React.MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;
    if (target.tagName === 'INPUT' && (target as HTMLInputElement).type === 'checkbox') {
      const checkbox = target as HTMLInputElement;
      const isChecked = checkbox.checked;
      const checkboxes = Array.from(e.currentTarget.querySelectorAll('input[type="checkbox"]'));
      const index = checkboxes.indexOf(checkbox);
      if (index >= 0) {
        let count = 0;
        const newContent = localContent.replace(/^(\s*-\s*\[)([ xX])(\])/gm, (match, prefix, _mark, suffix) => {
          if (count === index) {
            count++;
            return `${prefix}${isChecked ? 'x' : ' '}${suffix}`;
          }
          count++;
          return match;
        });
        handleContentChange(newContent);
      }
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(localContent);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  // 본문 공통 렌더러 (에디터 / 미리보기 / 분할)
  const renderEditorBody = () => (
    <div className="flex-1 min-h-0 flex overflow-hidden bg-slate-950">
      {/* 편집기 (에디터) */}
      {(viewMode === 'edit' || viewMode === 'split') && (
        <div className="flex-1 h-full p-3 sm:p-4 border-r border-slate-800/80 flex flex-col min-w-0">
          <textarea
            value={localContent}
            onChange={(e) => handleContentChange(e.target.value)}
            placeholder="마크다운으로 자유롭게 작성해 보세요..."
            className="w-full flex-1 bg-transparent text-slate-100 text-xs sm:text-sm font-mono leading-relaxed outline-none resize-none placeholder-slate-600 select-text"
            spellCheck={false}
          />
        </div>
      )}

      {/* 마크다운 미리보기 */}
      {(viewMode === 'preview' || viewMode === 'split') && (
        <div
          onClick={handlePreviewClick}
          className="flex-1 h-full p-4 overflow-y-auto select-text min-w-0 markdown-preview"
        >
          <div
            className="prose prose-invert prose-slate max-w-none text-xs sm:text-sm leading-relaxed"
            dangerouslySetInnerHTML={{ __html: renderedHtml }}
          />
        </div>
      )}
    </div>
  );

  // 1. 브라우저 밖 독립 OS 팝업 창 모드
  if (isPopout) {
    return (
      <PopoutWindow
        title={note.title || '메모'}
        onClose={onClose}
      >
        <div className="w-full h-full flex flex-col bg-slate-950 text-slate-100 overflow-hidden">
          {/* 상단 툴바 헤더 */}
          <div className="px-3.5 py-2.5 bg-slate-900 border-b border-slate-800 flex items-center justify-between gap-2 shrink-0 select-none">
            {/* 제목 & 뱃지 */}
            <div className="flex items-center gap-2 min-w-0 max-w-[50%]">
              <span className="p-1 rounded-md bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 shrink-0">
                <FileText size={15} />
              </span>
              <span className="font-bold text-xs sm:text-sm text-slate-100 truncate" title={note.title}>
                {note.title || '제목 없음'}
              </span>
              {note.isLocalOnly && (
                <span className="text-[10px] bg-amber-500/10 text-amber-400 px-1.5 py-0.5 rounded border border-amber-500/20 shrink-0 flex items-center gap-0.5">
                  <Lock size={9} /> 로컬
                </span>
              )}
              <span className="text-[10px] bg-indigo-500/10 text-indigo-300 px-1.5 py-0.5 rounded border border-indigo-500/20 shrink-0 hidden xs:inline-block font-semibold">
                독립 새 창 (OS 팝업)
              </span>
            </div>

            {/* 컨트롤 버튼 그룹 */}
            <div className="flex items-center gap-1 sm:gap-1.5 shrink-0">
              {/* 뷰 모드 토글 */}
              <div className="flex items-center bg-slate-950 p-0.5 rounded-lg border border-slate-800 text-xs">
                <button
                  type="button"
                  onClick={() => setViewMode('preview')}
                  className={`p-1 sm:px-2 rounded flex items-center gap-1 transition-colors ${
                    viewMode === 'preview' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
                  }`}
                  title="미리보기 모드"
                >
                  <Eye size={13} />
                  <span className="hidden sm:inline text-[11px]">보기</span>
                </button>
                <button
                  type="button"
                  onClick={() => setViewMode('edit')}
                  className={`p-1 sm:px-2 rounded flex items-center gap-1 transition-colors ${
                    viewMode === 'edit' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
                  }`}
                  title="편집 모드"
                >
                  <Edit3 size={13} />
                  <span className="hidden sm:inline text-[11px]">편집</span>
                </button>
                <button
                  type="button"
                  onClick={() => setViewMode('split')}
                  className={`p-1 sm:px-2 rounded flex items-center gap-1 transition-colors ${
                    viewMode === 'split' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
                  }`}
                  title="양방향 분할 모드"
                >
                  <Columns size={13} />
                  <span className="hidden sm:inline text-[11px]">분할</span>
                </button>
              </div>

              {/* 본문 복사 */}
              <button
                type="button"
                onClick={handleCopy}
                className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-400 hover:text-slate-200 rounded-lg border border-slate-700/80 transition-colors"
                title="본문 전체 복사"
              >
                {copied ? <Check size={13} className="text-emerald-400" /> : <Copy size={13} />}
              </button>

              {/* 브라우저 내부 창으로 복귀 버튼 */}
              <button
                type="button"
                onClick={() => setIsPopout(false)}
                className="p-1.5 sm:px-2 bg-slate-800 hover:bg-slate-700 text-indigo-300 hover:text-indigo-200 rounded-lg border border-slate-700/80 transition-colors flex items-center gap-1 text-xs"
                title="브라우저 화면 안으로 다시 넣기"
              >
                <Minimize2 size={13} />
                <span className="hidden sm:inline text-[11px]">내부 창으로</span>
              </button>

              {/* 닫기 버튼 */}
              <button
                type="button"
                onClick={onClose}
                className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700/80 transition-colors"
                title="창 닫기"
              >
                <X size={15} />
              </button>
            </div>
          </div>

          {/* 본문 에디터/미리보기 */}
          {renderEditorBody()}
        </div>
      </PopoutWindow>
    );
  }

  // 2. 브라우저 화면 내 가상 플로팅 창 모드
  return (
    <div
      onClick={onFocus}
      style={{
        position: 'fixed',
        left: isMaximized ? '0px' : `${position.x}px`,
        top: isMaximized ? '0px' : `${position.y}px`,
        width: isMaximized ? '100vw' : `${size.width}px`,
        height: isMaximized ? '100vh' : `${size.height}px`,
        zIndex,
      }}
      className={`pointer-events-auto flex flex-col bg-slate-900 border border-slate-700/80 shadow-[0_25px_60px_-15px_rgba(0,0,0,0.85)] ring-1 ring-white/10 select-none overflow-hidden transition-[box-shadow] duration-150 ${
        isMaximized ? 'rounded-none' : 'rounded-2xl'
      }`}
    >
      {/* 플로팅 헤더 바 */}
      <div
        onPointerDown={handleHeaderPointerDown}
        style={{ cursor: isMaximized ? 'default' : isDragging ? 'grabbing' : 'grab' }}
        className="px-3.5 py-2.5 bg-slate-900/95 border-b border-slate-800 flex items-center justify-between gap-2 shrink-0 select-none"
      >
        {/* 좌측: 메모 제목 & 로컬 배지 */}
        <div className="flex items-center gap-2 min-w-0 max-w-[45%]">
          <span className="p-1 rounded-md bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 shrink-0">
            <FileText size={15} />
          </span>
          <span className="font-bold text-xs sm:text-sm text-slate-100 truncate" title={note.title}>
            {note.title || '제목 없음'}
          </span>
          {note.isLocalOnly && (
            <span className="text-[10px] bg-amber-500/10 text-amber-400 px-1.5 py-0.5 rounded border border-amber-500/20 shrink-0 flex items-center gap-0.5">
              <Lock size={9} /> 로컬
            </span>
          )}
        </div>

        {/* 우측: 보기/편집 모드 전환 & 새 창 분리 & 전체화면 & 닫기 */}
        <div className="flex items-center gap-1 shrink-0 no-drag">
          {/* 뷰 모드 토글 */}
          <div className="flex items-center bg-slate-950 p-0.5 rounded-lg border border-slate-800 text-xs">
            <button
              type="button"
              onClick={() => setViewMode('preview')}
              className={`p-1 rounded transition-colors ${
                viewMode === 'preview' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
              }`}
              title="미리보기 모드"
            >
              <Eye size={13} />
            </button>
            <button
              type="button"
              onClick={() => setViewMode('edit')}
              className={`p-1 rounded transition-colors ${
                viewMode === 'edit' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
              }`}
              title="편집 모드"
            >
              <Edit3 size={13} />
            </button>
            <button
              type="button"
              onClick={() => setViewMode('split')}
              className={`p-1 rounded transition-colors ${
                viewMode === 'split' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-400 hover:text-slate-200'
              }`}
              title="양방향 분할 모드"
            >
              <Columns size={13} />
            </button>
          </div>

          <button
            type="button"
            onClick={handleCopy}
            className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-400 hover:text-slate-200 rounded-lg border border-slate-700/80 transition-colors"
            title="본문 전체 복사"
          >
            {copied ? <Check size={13} className="text-emerald-400" /> : <Copy size={13} />}
          </button>

          {/* 브라우저 밖 독립 새 창(OS 팝업)으로 분리 버튼 */}
          <button
            type="button"
            onClick={() => setIsPopout(true)}
            className="p-1.5 px-2 bg-indigo-950/60 hover:bg-indigo-900 text-indigo-300 hover:text-white rounded-lg border border-indigo-500/40 transition-colors flex items-center gap-1"
            title="브라우저 밖 독립 새 창(OS 팝업)으로 분리 (바탕화면/다른 모니터로 이동 가능)"
          >
            <ExternalLink size={13} className="text-indigo-400" />
            <span className="hidden sm:inline text-[11px] font-semibold">새 창</span>
          </button>

          <button
            type="button"
            onClick={() => setIsMaximized((prev) => !prev)}
            className="p-1.5 bg-slate-800 hover:bg-slate-750 text-slate-400 hover:text-slate-200 rounded-lg border border-slate-700/80 transition-colors"
            title={isMaximized ? '원래 크기로 복원' : '전체화면'}
          >
            {isMaximized ? <Minimize2 size={13} /> : <Maximize2 size={13} />}
          </button>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700/80 transition-colors"
            title="플로팅 창 닫기"
          >
            <X size={15} />
          </button>
        </div>
      </div>

      {/* 본문 콘텐츠 영역 */}
      {renderEditorBody()}

      {/* 우측 하단 크기 조절 손잡이 */}
      {!isMaximized && (
        <div
          onPointerDown={handleResizePointerDown}
          className="absolute bottom-0 right-0 w-5 h-5 cursor-nwse-resize z-30 flex items-center justify-center text-slate-500 hover:text-indigo-400 pointer-events-auto select-none"
          title="드래그하여 창 크기 조절"
        >
          <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
            <circle cx="10" cy="10" r="1.2" />
            <circle cx="6" cy="10" r="1.2" />
            <circle cx="10" cy="6" r="1.2" />
            <circle cx="2" cy="10" r="1.2" />
            <circle cx="6" cy="6" r="1.2" />
            <circle cx="10" cy="2" r="1.2" />
          </svg>
        </div>
      )}
    </div>
  );
};
