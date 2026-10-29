import React, { useState, useMemo } from 'react';
import type { Note, Project } from '../types';
import { 
  Plus, 
  Search, 
  Lock, 
  Menu, 
  AlertTriangle, 
  CloudCheck, 
  CloudUpload,
  Calendar,
  Trash2,
  FolderPlus,
  Folder,
  ChevronRight,
  ChevronLeft
} from 'lucide-react';

interface NoteListProps {
  notes: Note[];
  projects: Project[];
  activeProjectId: string;
  activeNoteId: string | null;
  noteCounts?: Record<string, number>;
  onSelectProject?: (id: string) => void;
  onSelectNote: (id: string) => void;
  onCreateNote: () => void;
  onDeleteNote: (id: string) => void;
  onOpenSidebar: () => void;
  onRequestCreateSubfolder?: (parentId: string) => void;
  onDropFiles?: (files: FileList) => void;
}

export const NoteList: React.FC<NoteListProps> = ({
  notes,
  projects,
  activeProjectId,
  activeNoteId,
  noteCounts = {},
  onSelectProject,
  onSelectNote,
  onCreateNote,
  onDeleteNote,
  onOpenSidebar,
  onRequestCreateSubfolder,
  onDropFiles,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [isDragOverList, setIsDragOverList] = useState(false);

  // 현재 활성화된 프로젝트 정보 및 부모/자식 폴더
  const currentProject = projects.find((p) => p.id === activeProjectId);
  const parentProject = currentProject?.parentId 
    ? projects.find((p) => p.id === currentProject.parentId) 
    : null;

  // 현재 폴더 직속 하위 폴더 목록
  const subprojects = useMemo(() => {
    if (activeProjectId === 'ALL') return [];
    return projects.filter((p) => p.parentId === activeProjectId);
  }, [activeProjectId, projects]);

  const projectTitle = activeProjectId === 'ALL' ? '전체 메모' : currentProject?.name || '메모';

  // 필터링: 메인 폴더 접속 시 하위 폴더의 메모가 섞이지 않고 오직 직속 메모만 표시
  const filteredNotes = notes.filter((note) => {
    const matchesProject = activeProjectId === 'ALL' || note.projectId === activeProjectId;
    if (!matchesProject) return false;

    if (!searchQuery.trim()) return true;
    const q = searchQuery.toLowerCase();
    return note.title.toLowerCase().includes(q) || note.content.toLowerCase().includes(q);
  });

  const formatDate = (isoString: string) => {
    try {
      const date = new Date(isoString);
      const now = new Date();
      if (date.toDateString() === now.toDateString()) {
        return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      }
      return `${date.getMonth() + 1}월 ${date.getDate()}일`;
    } catch {
      return '';
    }
  };

  return (
    <div 
      className={`w-full md:w-80 lg:w-96 bg-slate-900 border-r border-slate-800 flex flex-col h-full shrink-0 transition-colors ${
        isDragOverList ? 'ring-2 ring-indigo-500 ring-inset bg-indigo-950/20' : ''
      }`}
      onDragOver={(e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragOverList(true);
      }}
      onDragLeave={(e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragOverList(false);
      }}
      onDrop={(e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragOverList(false);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0 && onDropFiles) {
          onDropFiles(e.dataTransfer.files);
        }
      }}
    >
      {/* 상단 헤더 & 모바일 햄버거 & 새 메모 버튼 */}
      <div className="p-3.5 border-b border-slate-800 flex items-center justify-between gap-2">
        <div className="flex items-center gap-1.5 truncate">
          <button
            onClick={onOpenSidebar}
            className="md:hidden p-1.5 -ml-1 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 shrink-0"
            title="프로젝트 목록 열기"
          >
            <Menu size={20} />
          </button>

          {/* 상위 폴더가 있을 경우 뒤로가기 버튼 */}
          {parentProject && (
            <button
              onClick={() => onSelectProject?.(parentProject.id)}
              className="p-1 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 flex items-center gap-0.5 text-xs shrink-0"
              title={`${parentProject.name} 폴더로 이동`}
            >
              <ChevronLeft size={16} />
              <span className="truncate max-w-[70px] hidden sm:inline">{parentProject.name}</span>
            </button>
          )}

          <div className="truncate">
            <h2 className="font-semibold text-slate-100 text-sm truncate flex items-center gap-1.5">
              {currentProject?.isLocalOnly && <Lock size={13} className="text-amber-400 shrink-0" />}
              <span>{projectTitle}</span>
            </h2>
            <span className="text-xs text-slate-400 font-mono">
              {filteredNotes.length}개의 메모
            </span>
          </div>
        </div>

        <div className="flex items-center gap-1.5 shrink-0">
          {currentProject && onRequestCreateSubfolder && (
            <button
              onClick={() => onRequestCreateSubfolder(currentProject.id)}
              className="flex items-center gap-1 px-2.5 py-1.5 bg-slate-800 hover:bg-slate-750 text-indigo-300 border border-slate-700/80 rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95"
              title={`${currentProject.name}의 하위 폴더 생성`}
            >
              <FolderPlus size={14} className="text-indigo-400" />
              <span className="hidden sm:inline">하위 폴더</span>
            </button>
          )}

          <button
            onClick={onCreateNote}
            className="flex items-center gap-1 px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95"
          >
            <Plus size={15} />
            <span>새 메모</span>
          </button>
        </div>
      </div>

      {/* 검색창 */}
      <div className="p-3 border-b border-slate-800/60">
        <div className="relative">
          <Search size={14} className="absolute left-3 top-2.5 text-slate-500" />
          <input
            type="text"
            placeholder="제목 또는 내용 검색..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full bg-slate-950 border border-slate-800 rounded-lg pl-8 pr-3 py-1.5 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500 transition-colors"
          />
        </div>
      </div>

      {/* 하위 폴더 섹션: 메인 폴더 접속 시 안의 폴더가 나오도록 표시 */}
      {subprojects.length > 0 && !searchQuery.trim() && (
        <div className="p-2.5 border-b border-slate-800/80 bg-slate-950/40">
          <div className="text-[11px] font-semibold text-slate-400 uppercase tracking-wider mb-2 px-1 flex items-center justify-between">
            <span className="flex items-center gap-1.5">
              <Folder size={12} className="text-indigo-400" />
              하위 폴더 ({subprojects.length})
            </span>
          </div>
          <div className="grid grid-cols-1 gap-1.5">
            {subprojects.map((sub) => {
              const count = noteCounts[sub.id] || 0;
              return (
                <button
                  key={sub.id}
                  onClick={() => onSelectProject?.(sub.id)}
                  className="flex items-center justify-between p-2.5 rounded-xl bg-slate-800/60 hover:bg-slate-800 border border-slate-800 hover:border-indigo-500/50 text-left transition-all group cursor-pointer"
                >
                  <div className="flex items-center gap-2 truncate">
                    {sub.isLocalOnly ? (
                      <Lock size={14} className="text-amber-400 shrink-0" />
                    ) : (
                      <Folder size={14} className="text-indigo-400 group-hover:text-indigo-300 shrink-0" />
                    )}
                    <span className="text-xs font-medium text-slate-200 group-hover:text-white truncate">
                      {sub.name}
                    </span>
                  </div>
                  <div className="flex items-center gap-1.5 text-slate-400 group-hover:text-slate-200 shrink-0">
                    <span className="text-[11px] font-mono text-slate-500 group-hover:text-slate-400">
                      {count}
                    </span>
                    <ChevronRight size={14} className="group-hover:translate-x-0.5 transition-transform" />
                  </div>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* 메모 목록 */}
      <div className="flex-1 overflow-y-auto p-2 space-y-1.5">
        {filteredNotes.length === 0 ? (
          <div className="text-center py-12 px-4 text-slate-500 text-xs">
            {searchQuery ? '검색 결과가 없습니다.' : '메모가 없습니다. 새 메모를 작성해 보세요!'}
          </div>
        ) : (
          filteredNotes.map((note) => {
            const isSelected = activeNoteId === note.id;
            const previewText = note.content
              .replace(/^[#\-*>\s]+/gm, '')
              .replace(/\[[ x]\]/g, '')
              .slice(0, 75);

            return (
              <div
                key={note.id}
                onClick={() => onSelectNote(note.id)}
                className={`group relative p-3 rounded-xl border text-left cursor-pointer transition-all ${
                  isSelected
                    ? 'bg-slate-800/90 border-indigo-500/50 shadow-sm'
                    : 'bg-slate-950/40 border-slate-800/70 hover:bg-slate-800/40 hover:border-slate-700'
                }`}
              >
                <div className="flex items-start justify-between gap-2 mb-1">
                  <h3 className={`text-sm font-semibold truncate ${
                    isSelected ? 'text-indigo-200' : 'text-slate-200'
                  }`}>
                    {note.title || '제목 없음'}
                  </h3>

                  {/* 상태 뱃지 */}
                  <div className="flex items-center gap-1 shrink-0 pt-0.5">
                    {note.isLocalOnly ? (
                      <span className="flex items-center text-[10px] bg-amber-500/10 text-amber-400 px-1.5 py-0.5 rounded border border-amber-500/20" title="이 기기에만 저장됨">
                        <Lock size={10} className="mr-0.5" />
                        로컬
                      </span>
                    ) : note.syncStatus === 'conflict' ? (
                      <span className="flex items-center text-[10px] bg-rose-500/20 text-rose-300 px-1.5 py-0.5 rounded border border-rose-500/30" title="원격 내용과 충돌 발생">
                        <AlertTriangle size={10} className="mr-0.5 text-rose-400" />
                        충돌
                      </span>
                    ) : note.syncStatus === 'synced' ? (
                      <span className="text-emerald-400/80" title="GitHub와 동기화됨">
                        <CloudCheck size={13} />
                      </span>
                    ) : (
                      <span className="text-slate-500" title="동기화 대기 중">
                        <CloudUpload size={13} />
                      </span>
                    )}
                  </div>
                </div>

                {/* 내용 1줄 미리보기 */}
                <p className="text-xs text-slate-400 line-clamp-2 leading-relaxed mb-2 font-normal">
                  {previewText || '(내용 없음)'}
                </p>

                {/* 날짜 & 삭제 버튼 */}
                <div className="flex items-center justify-between text-[11px] text-slate-500 font-mono">
                  <span className="flex items-center gap-1">
                    <Calendar size={11} />
                    {formatDate(note.updatedAt)}
                  </span>

                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      onDeleteNote(note.id);
                    }}
                    className="opacity-0 group-hover:opacity-100 text-slate-500 hover:text-rose-400 p-1 rounded transition-opacity"
                    title="메모 삭제"
                  >
                    <Trash2 size={12} />
                  </button>
                </div>
              </div>
            );
          })
        )}
      </div>
    </div>
  );
};
