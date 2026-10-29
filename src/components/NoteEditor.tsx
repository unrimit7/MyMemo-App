import React, { useState, useRef, useEffect, useMemo } from 'react';
import type { Note, Project } from '../types';
import { marked } from 'marked';
import { compressImageToDataUrl } from '../utils/image';
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
  Image as ImageIcon,
  AlertTriangle,
  Folder,
  Loader2
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
  const [isProcessingImage, setIsProcessingImage] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // 화면 너비에 따라 기본 뷰 모드 자동 조정
  useEffect(() => {
    if (window.innerWidth >= 1280) {
      setViewMode('split');
    } else {
      setViewMode('edit');
    }
  }, [note?.id]);

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

    setTimeout(() => {
      textarea.focus();
      textarea.setSelectionRange(
        start + prefix.length,
        start + prefix.length + (selected.length || 2)
      );
    }, 0);
  };

  // 커서 위치에 이미지 Data URL 삽입
  const insertImageAtCursor = (dataUrl: string, altText: string = '이미지') => {
    const textarea = textareaRef.current;
    if (!textarea) {
      onUpdateNote({ content: `${note.content}\n\n![${altText}](${dataUrl})\n\n` });
      return;
    }

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const imageMarkdown = `\n\n![${altText}](${dataUrl})\n\n`;

    const newContent = 
      note.content.substring(0, start) + 
      imageMarkdown + 
      note.content.substring(end);

    onUpdateNote({ content: newContent });

    setTimeout(() => {
      textarea.focus();
      const nextPos = start + imageMarkdown.length;
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
          insertImageAtCursor(compressed, '캡처 이미지');
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

  // 2. 사진 파일 선택 버튼 핸들러
  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessingImage(true);
    try {
      const compressed = await compressImageToDataUrl(file, 1000, 0.75);
      insertImageAtCursor(compressed, file.name.replace(/\.[^/.]+$/, ''));
    } catch (err) {
      console.error('이미지 업로드 실패:', err);
      alert('이미지를 불러오지 못했습니다.');
    } finally {
      setIsProcessingImage(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const renderMarkdown = (content: string) => {
    try {
      return { __html: marked.parse(content || '') as string };
    } catch {
      return { __html: '<p>마크다운을 렌더링할 수 없습니다.</p>' };
    }
  };

  return (
    <div className="flex-1 flex flex-col h-full bg-slate-950 overflow-hidden">
      {/* 숨겨진 파일 선택 인풋 */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileChange}
        accept="image/*"
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

          {/* 계층 경로 빵가루 드롭다운 */}
          <div className="flex items-center gap-1.5 bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-300 max-w-[240px] md:max-w-md truncate" title={folderBreadcrumbs.map(b => b.name).join(' > ')}>
            <Folder size={12} className="text-indigo-400 shrink-0" />
            <select
              value={note.projectId}
              onChange={(e) => onUpdateNote({ projectId: e.target.value })}
              className="bg-transparent text-xs text-slate-200 focus:outline-none cursor-pointer truncate max-w-full"
            >
              {projects.map((p) => {
                const depth = p.parentId ? 1 : 0;
                const prefix = depth === 1 ? '　└ ' : '';
                return (
                  <option key={p.id} value={p.id} className="bg-slate-900 text-slate-200">
                    {prefix}{p.isLocalOnly ? `🔒 ${p.name}` : p.name}
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

      {/* 상단 툴바 2: 마크다운 빠른 서식 및 [사진 첨부] 버튼바 */}
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
          <div className="w-[1px] h-3.5 bg-slate-800 mx-1" />

          {/* 🖼️ 사진 첨부 버튼 */}
          <button
            onClick={() => fileInputRef.current?.click()}
            disabled={isProcessingImage}
            className="p-1 hover:bg-slate-800 text-indigo-400 hover:text-indigo-300 rounded px-2 font-semibold flex items-center gap-1 text-[11px] bg-indigo-500/10 border border-indigo-500/20 active:scale-95"
            title="커서 위치에 사진/이미지 첨부 (또는 복사 후 Ctrl+V)"
          >
            {isProcessingImage ? (
              <Loader2 size={12} className="animate-spin" />
            ) : (
              <ImageIcon size={12} />
            )}
            <span>사진 첨부</span>
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
        {(viewMode === 'edit' || viewMode === 'split') && (
          <div className={`flex-1 flex flex-col h-full ${viewMode === 'split' ? 'border-r border-slate-800/80 pr-6' : ''}`}>
            <textarea
              ref={textareaRef}
              value={note.content}
              onChange={(e) => onUpdateNote({ content: e.target.value })}
              onPaste={handlePaste}
              placeholder="마크다운 문법으로 자유롭게 작성해 보세요... 이미지는 Ctrl + V 로 바로 붙여넣을 수 있습니다!"
              className="w-full flex-1 bg-transparent text-slate-200 text-sm md:text-base leading-relaxed resize-none focus:outline-none font-mono placeholder-slate-600 scrollbar-thin"
              spellCheck={false}
            />
          </div>
        )}

        {(viewMode === 'preview' || viewMode === 'split') && (
          <div className="flex-1 h-full overflow-y-auto markdown-preview prose prose-invert max-w-none pr-2">
            <div dangerouslySetInnerHTML={renderMarkdown(note.content)} />
          </div>
        )}
      </div>
    </div>
  );
};
