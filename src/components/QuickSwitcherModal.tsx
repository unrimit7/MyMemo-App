import React, { useState, useEffect, useRef, useMemo } from 'react';
import type { Note, Project } from '../types';
import { Search, FileText, Plus, X, ArrowRight } from 'lucide-react';

interface QuickSwitcherModalProps {
  isOpen: boolean;
  onClose: () => void;
  notes: Note[];
  projects: Project[];
  onSelectNote: (noteId: string) => void;
  onCreateNoteWithTitle: (title: string) => void;
}

export const QuickSwitcherModal: React.FC<QuickSwitcherModalProps> = ({
  isOpen,
  onClose,
  notes,
  projects,
  onSelectNote,
  onCreateNoteWithTitle,
}) => {
  const [query, setQuery] = useState('');
  const [selectedIndex, setSelectedIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);

  // 모달 열릴 때 검색어 초기화 및 포커스
  useEffect(() => {
    if (isOpen) {
      setQuery('');
      setSelectedIndex(0);
      setTimeout(() => {
        inputRef.current?.focus();
      }, 50);
    }
  }, [isOpen]);

  // 프로젝트(폴더) 맵 생성
  const projectMap = useMemo(() => {
    const map = new Map<string, Project>();
    projects.forEach((p) => map.set(p.id, p));
    return map;
  }, [projects]);

  // 검색 결과 계산 (제목 일치 우선, 그 다음 본문 일치)
  const searchResults = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) {
      // 최근 수정된 순서 15개 반환
      return [...notes]
        .sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime())
        .slice(0, 15);
    }

    const titleMatches: Note[] = [];
    const contentMatches: Note[] = [];

    for (const note of notes) {
      const titleLower = (note.title || '').toLowerCase();
      const contentLower = (note.content || '').toLowerCase();

      if (titleLower.includes(q)) {
        titleMatches.push(note);
      } else if (contentLower.includes(q)) {
        contentMatches.push(note);
      }
    }

    return [...titleMatches, ...contentMatches].slice(0, 20);
  }, [notes, query]);

  // 새 메모 생성 옵션 표시 여부
  const showCreateOption = useMemo(() => {
    const q = query.trim();
    if (!q) return false;
    // 정확히 같은 제목의 메모가 없는 경우
    return !notes.some((n) => n.title.trim().toLowerCase() === q.toLowerCase());
  }, [notes, query]);

  // 총 선택 가능한 아이템 개수 (새 메모 생성 옵션 포함)
  const totalItems = searchResults.length + (showCreateOption ? 1 : 0);

  // selectedIndex 범위 보정
  useEffect(() => {
    if (selectedIndex >= totalItems) {
      setSelectedIndex(Math.max(0, totalItems - 1));
    }
  }, [totalItems, selectedIndex]);

  // 키보드 내비게이션 핸들러
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setSelectedIndex((prev) => (prev + 1) % Math.max(1, totalItems));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setSelectedIndex((prev) => (prev - 1 + totalItems) % Math.max(1, totalItems));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (showCreateOption && selectedIndex === searchResults.length) {
        // 새 메모 생성
        onCreateNoteWithTitle(query.trim());
        onClose();
      } else if (searchResults[selectedIndex]) {
        // 기존 메모 선택
        onSelectNote(searchResults[selectedIndex].id);
        onClose();
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      onClose();
    }
  };

  // 선택된 아이템이 스크롤 화면에 보이도록 자동 스크롤
  useEffect(() => {
    if (!listRef.current) return;
    const activeEl = listRef.current.querySelector(`[data-index="${selectedIndex}"]`) as HTMLElement;
    if (activeEl) {
      activeEl.scrollIntoView({ block: 'nearest' });
    }
  }, [selectedIndex]);

  if (!isOpen) return null;

  return (
    <div 
      className="fixed inset-0 z-50 flex items-start justify-center pt-16 sm:pt-24 px-3 sm:px-4 bg-black/75 backdrop-blur-sm animate-in fade-in duration-150"
      onClick={onClose}
    >
      <div 
        className="w-full max-w-xl bg-slate-900 border border-slate-700/80 rounded-2xl shadow-2xl overflow-hidden flex flex-col text-slate-100 animate-in zoom-in-95 duration-150"
        onClick={(e) => e.stopPropagation()}
        onKeyDown={handleKeyDown}
      >
        {/* 상단 검색 입력창 */}
        <div className="relative flex items-center px-4 py-3.5 border-b border-slate-800 bg-slate-950/60">
          <Search size={18} className="text-indigo-400 shrink-0 mr-3" />
          <input
            ref={inputRef}
            type="text"
            value={query}
            onChange={(e) => {
              setQuery(e.target.value);
              setSelectedIndex(0);
            }}
            placeholder="메모 제목 또는 내용을 검색하세요... (Ctrl+K)"
            className="flex-1 bg-transparent text-sm sm:text-base text-slate-100 placeholder-slate-500 focus:outline-none"
            spellCheck={false}
          />
          {query ? (
            <button
              type="button"
              onClick={() => {
                setQuery('');
                inputRef.current?.focus();
              }}
              className="p-1 text-slate-500 hover:text-slate-300 rounded-lg"
            >
              <X size={15} />
            </button>
          ) : (
            <kbd className="hidden sm:inline-block px-1.5 py-0.5 text-[10px] font-mono text-slate-400 bg-slate-800 border border-slate-700 rounded">
              ESC 닫기
            </kbd>
          )}
        </div>

        {/* 결과 목록 */}
        <div 
          ref={listRef}
          className="max-h-[380px] overflow-y-auto p-2 space-y-1 scrollbar-thin"
        >
          {searchResults.length === 0 && !showCreateOption ? (
            <div className="py-10 text-center text-slate-500 text-xs sm:text-sm">
              검색된 메모가 없습니다.
            </div>
          ) : (
            searchResults.map((note, index) => {
              const isSelected = index === selectedIndex;
              const project = projectMap.get(note.projectId);
              const snippet = note.content.replace(/[#*`~>-]/g, '').trim().slice(0, 70);

              return (
                <button
                  key={note.id}
                  type="button"
                  data-index={index}
                  onClick={() => {
                    onSelectNote(note.id);
                    onClose();
                  }}
                  onMouseEnter={() => setSelectedIndex(index)}
                  className={`w-full px-3 py-2.5 rounded-xl text-left flex items-center justify-between gap-3 transition-colors cursor-pointer ${
                    isSelected
                      ? 'bg-indigo-600/25 border border-indigo-500/50 text-white'
                      : 'hover:bg-slate-800/80 text-slate-300 border border-transparent'
                  }`}
                >
                  <div className="flex items-center gap-2.5 min-w-0 flex-1">
                    <FileText 
                      size={16} 
                      className={`shrink-0 ${isSelected ? 'text-indigo-300' : 'text-slate-500'}`} 
                    />
                    <div className="min-w-0 flex-1">
                      <div className="text-xs sm:text-sm font-semibold truncate flex items-center gap-2">
                        <span className="truncate">{note.title || '제목 없음'}</span>
                        {project && (
                          <span 
                            className="text-[10px] px-1.5 py-0.2 rounded-md font-normal shrink-0 border"
                            style={{ 
                              color: project.color || '#94a3b8', 
                              borderColor: `${project.color || '#94a3b8'}40`,
                              backgroundColor: `${project.color || '#94a3b8'}15` 
                            }}
                          >
                            {project.name}
                          </span>
                        )}
                      </div>
                      {snippet && (
                        <div className="text-[11px] text-slate-400 truncate mt-0.5">
                          {snippet}
                        </div>
                      )}
                    </div>
                  </div>

                  <ArrowRight 
                    size={14} 
                    className={`shrink-0 transition-transform ${isSelected ? 'translate-x-0.5 text-indigo-400' : 'opacity-0'}`} 
                  />
                </button>
              );
            })
          )}

          {/* 새 메모 생성 옵션 */}
          {showCreateOption && (
            <button
              type="button"
              data-index={searchResults.length}
              onClick={() => {
                onCreateNoteWithTitle(query.trim());
                onClose();
              }}
              onMouseEnter={() => setSelectedIndex(searchResults.length)}
              className={`w-full px-3 py-2.5 rounded-xl text-left flex items-center gap-2.5 transition-colors cursor-pointer border ${
                selectedIndex === searchResults.length
                  ? 'bg-indigo-600/25 border-indigo-500/50 text-white'
                  : 'hover:bg-slate-800/80 text-slate-300 border-dashed border-slate-700'
              }`}
            >
              <div className="p-1 rounded-lg bg-indigo-500/20 text-indigo-400">
                <Plus size={15} />
              </div>
              <div className="text-xs sm:text-sm">
                <span className="text-indigo-300 font-semibold">"{query.trim()}"</span>
                <span className="text-slate-400 ml-1.5">새 메모 바로 만들기</span>
              </div>
            </button>
          )}
        </div>

        {/* 하단 힌트 바 */}
        <div className="px-4 py-2 bg-slate-950/80 border-t border-slate-800 flex items-center justify-between text-[11px] text-slate-500 select-none">
          <div className="flex items-center gap-3">
            <span>↑↓ 이동</span>
            <span>↵ 열기</span>
            <span>ESC 닫기</span>
          </div>
          <span>총 {notes.length}개의 메모</span>
        </div>
      </div>
    </div>
  );
};
