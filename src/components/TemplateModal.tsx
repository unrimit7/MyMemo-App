import React, { useState } from 'react';
import { getNoteTemplates, type NoteTemplate } from '../utils/templates';
import { 
  LayoutTemplate, 
  X, 
  Check, 
  FileText, 
  Calendar, 
  Briefcase, 
  Bug, 
  Lightbulb, 
  Sparkles
} from 'lucide-react';

interface TemplateModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSelectTemplate: (template: NoteTemplate) => void;
}

export const TemplateModal: React.FC<TemplateModalProps> = ({
  isOpen,
  onClose,
  onSelectTemplate,
}) => {
  const templates = getNoteTemplates();
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>(templates[0]?.id || 'meeting');
  const [selectedCategory, setSelectedCategory] = useState<'전체' | '업무' | '기획' | '개인'>('전체');

  if (!isOpen) return null;

  const filteredTemplates = templates.filter((t) => {
    if (selectedCategory === '전체') return true;
    return t.category === selectedCategory;
  });

  const activeTemplate = templates.find((t) => t.id === selectedTemplateId) || templates[0];

  const getTemplateIcon = (iconName: string) => {
    switch (iconName) {
      case 'meeting':
        return <Calendar size={18} className="text-blue-400" />;
      case 'daily':
        return <FileText size={18} className="text-emerald-400" />;
      case 'project':
        return <Briefcase size={18} className="text-purple-400" />;
      case 'bug':
        return <Bug size={18} className="text-rose-400" />;
      case 'weekly':
        return <LayoutTemplate size={18} className="text-amber-400" />;
      case 'idea':
        return <Lightbulb size={18} className="text-yellow-400" />;
      default:
        return <FileText size={18} className="text-indigo-400" />;
    }
  };

  return (
    <div 
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-6 bg-black/80 backdrop-blur-sm animate-in fade-in duration-150"
      onClick={onClose}
    >
      <div 
        className="w-full max-w-4xl h-[85vh] max-h-[650px] bg-slate-900 border border-slate-700/80 rounded-2xl shadow-2xl overflow-hidden flex flex-col text-slate-100 animate-in zoom-in-95 duration-150"
        onClick={(e) => e.stopPropagation()}
      >
        {/* 상단 헤더 */}
        <div className="px-5 py-3.5 border-b border-slate-800 bg-slate-950/70 flex items-center justify-between">
          <div className="flex items-center gap-2.5 font-bold text-slate-100 text-sm sm:text-base">
            <div className="p-1.5 rounded-lg bg-indigo-500/20 text-indigo-400">
              <LayoutTemplate size={18} />
            </div>
            <span>실무 서식 템플릿 선택</span>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-100 hover:bg-slate-800 rounded-lg transition-colors"
          >
            <X size={18} />
          </button>
        </div>

        {/* 메인 본문: 좌측 템플릿 목록 / 우측 미리보기 */}
        <div className="flex-1 flex flex-col md:flex-row min-h-0 overflow-hidden">
          {/* 좌측: 카테고리 탭 & 템플릿 카드 목록 */}
          <div className="w-full md:w-80 border-b md:border-b-0 md:border-r border-slate-800 flex flex-col shrink-0 bg-slate-950/30">
            {/* 카테고리 필터 탭 */}
            <div className="p-2.5 border-b border-slate-800 flex gap-1">
              {(['전체', '업무', '기획', '개인'] as const).map((cat) => (
                <button
                  key={cat}
                  type="button"
                  onClick={() => setSelectedCategory(cat)}
                  className={`flex-1 py-1 rounded-lg text-xs font-semibold transition-colors cursor-pointer ${
                    selectedCategory === cat
                      ? 'bg-indigo-600 text-white'
                      : 'text-slate-400 hover:bg-slate-800 hover:text-slate-200'
                  }`}
                >
                  {cat}
                </button>
              ))}
            </div>

            {/* 템플릿 리스트 */}
            <div className="flex-1 overflow-y-auto p-2 space-y-1.5 scrollbar-thin">
              {filteredTemplates.map((tpl) => {
                const isSelected = tpl.id === selectedTemplateId;
                return (
                  <button
                    key={tpl.id}
                    type="button"
                    onClick={() => setSelectedTemplateId(tpl.id)}
                    className={`w-full p-3 rounded-xl text-left flex items-start gap-3 transition-all cursor-pointer border ${
                      isSelected
                        ? 'bg-indigo-600/25 border-indigo-500/60 shadow-xs'
                        : 'bg-slate-900/60 border-slate-800 hover:border-slate-700 hover:bg-slate-800/60'
                    }`}
                  >
                    <div className="p-2 rounded-lg bg-slate-800 shrink-0 mt-0.5">
                      {getTemplateIcon(tpl.iconName)}
                    </div>
                    <div className="min-w-0 flex-1">
                      <div className="text-xs sm:text-sm font-bold text-slate-200 truncate flex items-center justify-between">
                        <span>{tpl.title}</span>
                        <span className="text-[10px] text-slate-500 font-normal">{tpl.category}</span>
                      </div>
                      <div className="text-[11px] text-slate-400 truncate mt-0.5">
                        {tpl.description}
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>
          </div>

          {/* 우측: 선택된 템플릿 본문 미리보기 */}
          <div className="flex-1 flex flex-col min-h-0 bg-slate-900/40">
            <div className="px-4 py-2.5 border-b border-slate-800 bg-slate-950/40 flex items-center justify-between text-xs text-slate-400 select-none">
              <span className="font-semibold text-slate-300 flex items-center gap-1.5">
                <Sparkles size={13} className="text-amber-400" />
                <span>서식 미리보기</span>
              </span>
              <span className="text-[11px] text-slate-500">선택 시 현재 메모에 즉시 삽입됩니다</span>
            </div>

            <div className="flex-1 overflow-y-auto p-4 md:p-5 scrollbar-thin font-mono text-xs sm:text-sm text-slate-300 leading-relaxed whitespace-pre-wrap select-text bg-slate-950/20">
              {activeTemplate?.content}
            </div>
          </div>
        </div>

        {/* 하단 버튼 바 */}
        <div className="px-5 py-3.5 border-t border-slate-800 bg-slate-950/80 flex items-center justify-between gap-3">
          <div className="text-xs text-slate-400 truncate hidden sm:block">
            선택된 서식: <strong className="text-indigo-300">{activeTemplate?.title}</strong>
          </div>
          <div className="flex items-center gap-2 ml-auto">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors"
            >
              취소
            </button>
            <button
              type="button"
              onClick={() => {
                if (activeTemplate) {
                  onSelectTemplate(activeTemplate);
                  onClose();
                }
              }}
              className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 active:scale-95 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-all shadow-md cursor-pointer"
            >
              <Check size={14} />
              <span>이 서식 적용하기</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
