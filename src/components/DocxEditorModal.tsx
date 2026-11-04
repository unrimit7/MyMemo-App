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
  Loader2,
  FileText,
  Eye,
  Edit3
} from 'lucide-react';
import { extractDocxTextAndHtml, generateDocxDataUrl } from '../utils/docxViewer';
import { marked } from 'marked';

interface DocxEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: string;
  title: string;
  attId?: string;
  onSave?: (updatedDataUrl: string) => Promise<void> | void;
}

export const DocxEditorModal: React.FC<DocxEditorModalProps> = ({
  isOpen,
  onClose,
  data,
  title,
  onSave,
}) => {
  const [content, setContent] = useState<string>('');
  const [docTitle, setDocTitle] = useState<string>(title.replace(/\.docx?$/i, ''));
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [hasChanges, setHasChanges] = useState<boolean>(false);
  const [saveToast, setSaveToast] = useState<string | null>(null);
  const [errorToast, setErrorToast] = useState<string | null>(null);
  const [viewMode, setViewMode] = useState<'edit' | 'preview' | 'split'>('split');

  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // 워드 내용 추출 및 초기화
  useEffect(() => {
    if (!isOpen || !data) return;

    let isMounted = true;
    setIsLoading(true);
    setHasChanges(false);

    extractDocxTextAndHtml(data)
      .then(({ markdown, text }) => {
        if (!isMounted) return;
        setContent(markdown || text || '');
        setIsLoading(false);
      })
      .catch((err) => {
        console.error('워드 본문 로드 실패:', err);
        if (!isMounted) {
          setErrorToast('워드 문서 내용을 불러오지 못했습니다.');
          setIsLoading(false);
        }
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, data]);

  // 서식 도구 적용 헬퍼
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

  // 표 삽입 헬퍼
  const insertTable = () => {
    const sampleTable = `\n| 항목 | 내용 | 비고 |\n| --- | --- | --- |\n| 항목 1 | 내용 입력 |  |\n| 항목 2 | 내용 입력 |  |\n`;
    applyFormat(sampleTable);
  };

  // 저장 및 메모 반영 핸들러
  const handleSave = async () => {
    setIsSaving(true);
    setSaveToast(null);
    setErrorToast(null);

    try {
      const updatedDataUrl = await generateDocxDataUrl(docTitle, content);
      if (onSave) {
        await onSave(updatedDataUrl);
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
      const exportUrl = await generateDocxDataUrl(docTitle, content);
      const link = document.createElement('a');
      link.href = exportUrl;
      link.download = `${docTitle}.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error('다운로드 생성 실패:', err);
      alert('파일을 생성하지 못했습니다.');
    }
  };

  // 닫기 핸들러
  const handleClose = () => {
    if (hasChanges) {
      if (!confirm('저장되지 않은 수정 내용이 있습니다. 닫으시겠습니까?')) {
        return;
      }
    }
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-black/85 backdrop-blur-sm flex flex-col items-center justify-center p-2 md:p-4 select-none animate-fadeIn">
      {/* 팝업 컨테이너 (완전 불투명) */}
      <div className="w-full max-w-6xl h-[95vh] bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
        {/* 상단 1단: 제목 & 저장 / 다운로드 / 닫기 헤더 */}
        <div className="px-4 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0">
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

          {/* 우측 저장 및 닫기 버튼 */}
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

            {/* 보기 모드 전환 (편집, 분할, 미리보기) */}
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

            {/* 저장 및 반영 버튼 */}
            <button
              type="button"
              onClick={handleSave}
              disabled={isSaving}
              className="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-semibold shadow-md flex items-center gap-1.5 transition-all active:scale-95 disabled:opacity-50"
              title="워드 파일에 저장하고 본문 뷰어에 즉시 반영"
            >
              {isSaving ? <Loader2 size={13} className="animate-spin" /> : <Save size={13} />}
              <span>{isSaving ? '저장 중...' : '💾 저장 및 반영'}</span>
            </button>

            {/* 다운로드 버튼 */}
            <button
              type="button"
              onClick={handleDownload}
              className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors"
              title="수정된 워드(.docx) 파일 다운로드"
            >
              <Download size={15} />
            </button>

            {/* 닫기 버튼 */}
            <button
              type="button"
              onClick={handleClose}
              className="p-1.5 bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 rounded-lg border border-slate-700 transition-colors"
              title="닫기 (Esc)"
            >
              <X size={16} />
            </button>
          </div>
        </div>

        {/* 상단 2단: 서식 툴바 */}
        <div className="px-3 py-1.5 bg-slate-950/70 border-b border-slate-800 flex flex-wrap items-center gap-1.5 shrink-0 text-xs">
          <button
            type="button"
            onClick={() => applyFormat('**', '**')}
            className="p-1.5 hover:bg-slate-800 text-slate-300 hover:text-white rounded"
            title="굵게 (Bold)"
          >
            <Bold size={14} />
          </button>
          <button
            type="button"
            onClick={() => applyFormat('*', '*')}
            className="p-1.5 hover:bg-slate-800 text-slate-300 hover:text-white rounded"
            title="기울임 (Italic)"
          >
            <Italic size={14} />
          </button>

          <span className="w-px h-4 bg-slate-800 mx-1"></span>

          <button
            type="button"
            onClick={() => applyFormat('# ')}
            className="p-1.5 hover:bg-slate-800 text-slate-300 hover:text-white rounded"
            title="제목 1"
          >
            <Heading1 size={14} />
          </button>
          <button
            type="button"
            onClick={() => applyFormat('## ')}
            className="p-1.5 hover:bg-slate-800 text-slate-300 hover:text-white rounded"
            title="제목 2"
          >
            <Heading2 size={14} />
          </button>
          <button
            type="button"
            onClick={() => applyFormat('### ')}
            className="p-1.5 hover:bg-slate-800 text-slate-300 hover:text-white rounded"
            title="제목 3"
          >
            <Heading3 size={14} />
          </button>

          <span className="w-px h-4 bg-slate-800 mx-1"></span>

          <button
            type="button"
            onClick={() => applyFormat('- ')}
            className="p-1.5 hover:bg-slate-800 text-slate-300 hover:text-white rounded"
            title="글머리 기호 목록"
          >
            <List size={14} />
          </button>
          <button
            type="button"
            onClick={() => applyFormat('1. ')}
            className="p-1.5 hover:bg-slate-800 text-slate-300 hover:text-white rounded"
            title="번호 매기기 목록"
          >
            <ListOrdered size={14} />
          </button>
          <button
            type="button"
            onClick={() => applyFormat('> ')}
            className="p-1.5 hover:bg-slate-800 text-slate-300 hover:text-white rounded"
            title="인용구"
          >
            <Quote size={14} />
          </button>
          <button
            type="button"
            onClick={insertTable}
            className="p-1.5 hover:bg-slate-800 text-slate-300 hover:text-white rounded flex items-center gap-1"
            title="표 삽입"
          >
            <Table size={14} />
            <span className="text-[11px]">표</span>
          </button>
        </div>

        {/* 에디터 메인 작업 영역 */}
        <div className="flex-1 overflow-hidden flex flex-col md:flex-row bg-slate-950">
          {isLoading ? (
            <div className="w-full flex items-center justify-center gap-2 py-20 text-slate-400 text-xs">
              <Loader2 size={16} className="animate-spin text-blue-500" />
              <span>워드 문서 텍스트를 불러오는 중입니다...</span>
            </div>
          ) : (
            <>
              {/* 왼쪽: 에디터 텍스트 영역 */}
              {(viewMode === 'edit' || viewMode === 'split') && (
                <div className={`h-full overflow-auto p-4 flex flex-col ${viewMode === 'split' ? 'w-full md:w-1/2 border-r border-slate-800' : 'w-full'}`}>
                  <textarea
                    ref={textareaRef}
                    value={content}
                    onChange={(e) => {
                      setContent(e.target.value);
                      setHasChanges(true);
                    }}
                    placeholder="워드 문서 본문 내용을 입력하거나 수정하세요..."
                    className="w-full h-full bg-slate-900/60 text-slate-200 text-sm font-sans p-4 rounded-xl border border-slate-800 focus:outline-none focus:border-blue-500 resize-none leading-relaxed"
                  />
                </div>
              )}

              {/* 오른쪽: 워드 용지(A4) 스타일 실시간 레이아웃 미리보기 */}
              {(viewMode === 'preview' || viewMode === 'split') && (
                <div className={`h-full overflow-y-auto overflow-x-auto p-4 md:p-8 bg-slate-950 ${viewMode === 'split' ? 'w-full md:w-1/2' : 'w-full'}`}>
                  <div className="w-full max-w-[760px] mx-auto bg-white text-slate-900 shadow-2xl rounded-sm p-8 md:p-12 font-serif text-sm leading-relaxed select-text min-h-fit mb-12">
                    {docTitle && (
                      <h1 className="text-2xl font-bold text-center border-b border-slate-200 pb-4 mb-6 text-slate-900">
                        {docTitle}
                      </h1>
                    )}
                    <div
                      className="prose prose-slate max-w-none text-slate-900 prose-headings:font-serif prose-headings:text-slate-900 prose-p:my-2 [&_table]:w-full [&_table]:border-collapse [&_table]:my-4 [&_th]:border [&_th]:border-slate-300 [&_th]:bg-slate-100 [&_th]:p-2.5 [&_th]:text-xs [&_th]:font-bold [&_th]:text-slate-800 [&_td]:border [&_td]:border-slate-300 [&_td]:p-2 [&_td]:text-xs [&_td]:text-slate-800"
                      dangerouslySetInnerHTML={{ __html: marked.parse(content) as string }}
                    />
                  </div>
                </div>
              )}
            </>
          )}
        </div>

        {/* 하단 상태 바 */}
        <div className="px-4 py-2 bg-slate-900 border-t border-slate-800 flex items-center justify-between text-[11px] text-slate-400">
          <span>문서 글자 수: {content.length}자</span>
          <span className="hidden sm:inline">💾 [저장 및 반영]을 클릭하면 워드(.docx) 규격으로 자동 인코딩되어 메모와 동기화됩니다.</span>
        </div>
      </div>
    </div>
  );
};
