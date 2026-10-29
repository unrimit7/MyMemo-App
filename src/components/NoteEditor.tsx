import React, { useState, useRef, useEffect } from 'react';
import type { Note, Project } from '../types';
import { marked } from 'marked';
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
  CheckSquare, 
  List, 
  Code, 
  Quote,
  AlertTriangle,
  Folder
} from 'lucide-react';

interface NoteEditorProps {
  note: Note | null;
  projects: Project[];
  onUpdateNote: (updated: Partial<Note>) => void;
  onBackToList: () => void;
}

export const NoteEditor: React.FC<NoteEditorProps> = ({
  note,
  projects,
  onUpdateNote,
  onBackToList,
}) => {
  const [viewMode, setViewMode] = useState<'edit' | 'preview' | 'split'>('edit');
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // 화면 너비에 따라 기본 뷰 모드 자동 조정 (PC 와이드면 split도 가능)
  useEffect(() => {
    if (window.innerWidth >= 1280) {
      setViewMode('split');
    } else {
      setViewMode('edit');
    }
  }, [note?.id]);

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
    const selected = note.content.substring(start, end);
    const replacement = `${prefix}${selected || '내용'}${suffix}`;

    const newContent = 
      note.content.substring(0, start) + 
      replacement + 
      note.content.substring(end);

    onUpdateNote({ content: newContent });

    // 커서 위치 재조정
    setTimeout(() => {
      textarea.focus();
      textarea.setSelectionRange(
        start + prefix.length,
        start + prefix.length + (selected.length || 2)
      );
    }, 0);
  };

  // 마크다운 파싱 HTML (안전 렌더링)
  const renderMarkdown = (content: string) => {
    try {
      return { __html: marked.parse(content || '') as string };
    } catch {
      return { __html: '<p>마크다운을 렌더링할 수 없습니다.</p>' };
    }
  };

  return (
    <div className="flex-1 flex flex-col h-full bg-slate-950 overflow-hidden">
      {/* 충돌 경고 배너 */}
      {note.syncStatus === 'conflict' && (
        <div className="bg-amber-500/15 border-b border-amber-500/30 p-2.5 px-4 flex items-center justify-between text-xs text-amber-300">
          <div className="flex items-center gap-2">
            <AlertTriangle size={15} className="text-amber-400 shrink-0" />
            <span>다른 기기에서 수정된 내용과 충돌이 감지되어 백업 사본이 별도로 보관되었습니다.</span>
          </div>
          <button
            onClick={() => onUpdateNote({ syncStatus: 'pending' })}
            className="px-2 py-1 bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold rounded text-[11px]"
          >
            충돌 해결 확인
          </button>
        </div>
      )}

      {/* 상단 툴바 1: 내비게이션 & 메타데이터 제어 */}
      <div className="p-3 border-b border-slate-800/80 flex flex-wrap items-center justify-between gap-2 bg-slate-950/70">
        <div className="flex items-center gap-2">
          {/* 모바일 목록 복귀 버튼 */}
          <button
            onClick={onBackToList}
            className="md:hidden p-1.5 -ml-1 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 flex items-center gap-1 text-xs"
          >
            <ArrowLeft size={16} />
            <span>목록</span>
          </button>

          {/* 소속 프로젝트 선택 드롭다운 */}
          <div className="flex items-center gap-1.5 bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-300">
            <Folder size={12} className="text-indigo-400" />
            <select
              value={note.projectId}
              onChange={(e) => onUpdateNote({ projectId: e.target.value })}
              className="bg-transparent text-xs text-slate-200 focus:outline-none cursor-pointer"
            >
              {projects.map((p) => (
                <option key={p.id} value={p.id} className="bg-slate-900 text-slate-200">
                  {p.isLocalOnly ? `🔒 ${p.name}` : p.name}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* 오른쪽: 기기 전용(로컬) 토글 스위치 & 뷰 모드 전환 */}
        <div className="flex items-center gap-2">
          {/* 🔒 기기 전용 토글 버튼 */}
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

          {/* 뷰 모드 전환 버튼들 (편집 / 미리보기 / 분할) */}
          <div className="flex items-center bg-slate-900 border border-slate-800 rounded-lg p-0.5 text-slate-400">
            <button
              onClick={() => setViewMode('edit')}
              className={`p-1.5 rounded text-xs transition-colors ${
                viewMode === 'edit' ? 'bg-indigo-600 text-white' : 'hover:text-slate-200'
              }`}
              title="편집 모드"
            >
              <Edit3 size={13} />
            </button>
            <button
              onClick={() => setViewMode('preview')}
              className={`p-1.5 rounded text-xs transition-colors ${
                viewMode === 'preview' ? 'bg-indigo-600 text-white' : 'hover:text-slate-200'
              }`}
              title="미리보기 모드"
            >
              <Eye size={13} />
            </button>
            <button
              onClick={() => setViewMode('split')}
              className={`hidden lg:block p-1.5 rounded text-xs transition-colors ${
                viewMode === 'split' ? 'bg-indigo-600 text-white' : 'hover:text-slate-200'
              }`}
              title="분할 보기 모드"
            >
              <Columns size={13} />
            </button>
          </div>
        </div>
      </div>

      {/* 상단 툴바 2: 마크다운 빠른 서식 버튼바 (모바일/PC 공용) */}
      {(viewMode === 'edit' || viewMode === 'split') && (
        <div className="px-3 py-1.5 border-b border-slate-800/60 bg-slate-950/40 flex items-center gap-1 overflow-x-auto text-slate-400 text-xs scrollbar-none">
          <button
            onClick={() => insertMarkdown('# ')}
            className="p-1 hover:bg-slate-800 hover:text-slate-200 rounded px-1.5 flex items-center gap-0.5"
            title="큰 제목"
          >
            <Heading1 size={13} />
          </button>
          <button
            onClick={() => insertMarkdown('## ')}
            className="p-1 hover:bg-slate-800 hover:text-slate-200 rounded px-1.5 flex items-center gap-0.5"
            title="중간 제목"
          >
            <Heading2 size={13} />
          </button>
          <div className="w-[1px] h-3.5 bg-slate-800 mx-1" />
          <button
            onClick={() => insertMarkdown('**', '**')}
            className="p-1 hover:bg-slate-800 hover:text-slate-200 rounded px-1.5"
            title="굵게"
          >
            <Bold size={13} />
          </button>
          <button
            onClick={() => insertMarkdown('*', '*')}
            className="p-1 hover:bg-slate-800 hover:text-slate-200 rounded px-1.5"
            title="기울임"
          >
            <Italic size={13} />
          </button>
          <div className="w-[1px] h-3.5 bg-slate-800 mx-1" />
          <button
            onClick={() => insertMarkdown('- [ ] ')}
            className="p-1 hover:bg-slate-800 hover:text-indigo-400 rounded px-2 font-medium flex items-center gap-1 text-[11px] bg-slate-900 border border-slate-800"
            title="할 일(체크리스트)"
          >
            <CheckSquare size={12} className="text-indigo-400" />
            <span>To-Do</span>
          </button>
          <button
            onClick={() => insertMarkdown('- ')}
            className="p-1 hover:bg-slate-800 hover:text-slate-200 rounded px-1.5"
            title="글머리 기호 목록"
          >
            <List size={13} />
          </button>
          <button
            onClick={() => insertMarkdown('> ')}
            className="p-1 hover:bg-slate-800 hover:text-slate-200 rounded px-1.5"
            title="인용구"
          >
            <Quote size={13} />
          </button>
          <button
            onClick={() => insertMarkdown('`', '`')}
            className="p-1 hover:bg-slate-800 hover:text-slate-200 rounded px-1.5"
            title="인라인 코드"
          >
            <Code size={13} />
          </button>
        </div>
      )}

      {/* 제목 입력 영역 */}
      <div className="px-6 pt-4 pb-2">
        <input
          type="text"
          value={note.title}
          onChange={(e) => onUpdateNote({ title: e.target.value })}
          placeholder="메모 제목을 입력하세요..."
          className="w-full bg-transparent text-xl md:text-2xl font-bold text-slate-100 placeholder-slate-600 focus:outline-none"
        />
      </div>

      {/* 본문 영역: 편집창 / 미리보기 / 분할 */}
      <div className="flex-1 flex overflow-hidden p-4 md:px-6 gap-6">
        {/* 편집 에디터 */}
        {(viewMode === 'edit' || viewMode === 'split') && (
          <div className={`flex-1 flex flex-col h-full ${viewMode === 'split' ? 'border-r border-slate-800/80 pr-6' : ''}`}>
            <textarea
              ref={textareaRef}
              value={note.content}
              onChange={(e) => onUpdateNote({ content: e.target.value })}
              placeholder="마크다운 문법으로 자유롭게 작성해 보세요..."
              className="w-full flex-1 bg-transparent text-slate-200 text-sm md:text-base leading-relaxed resize-none focus:outline-none font-mono placeholder-slate-600 scrollbar-thin"
              spellCheck={false}
            />
          </div>
        )}

        {/* 실시간 마크다운 미리보기 */}
        {(viewMode === 'preview' || viewMode === 'split') && (
          <div className="flex-1 h-full overflow-y-auto markdown-preview prose prose-invert max-w-none pr-2">
            <div dangerouslySetInnerHTML={renderMarkdown(note.content)} />
          </div>
        )}
      </div>
    </div>
  );
};
