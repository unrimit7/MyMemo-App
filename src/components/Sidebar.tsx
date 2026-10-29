import React, { useState } from 'react';
import type { Project, GitHubConfig } from '../types';
import { 
  Folder, 
  FolderPlus, 
  Lock, 
  Settings, 
  RefreshCw, 
  CheckCircle2, 
  AlertCircle,
  X,
  FileText
} from 'lucide-react';

interface SidebarProps {
  projects: Project[];
  activeProjectId: string;
  onSelectProject: (id: string) => void;
  onCreateProject: (name: string, isLocalOnly: boolean) => void;
  onDeleteProject: (id: string) => void;
  noteCounts: Record<string, number>;
  gitHubConfig: GitHubConfig | null;
  onOpenSettings: () => void;
  onSync: () => void;
  isSyncing: boolean;
  isOpenMobile: boolean;
  onCloseMobile: () => void;
}

export const Sidebar: React.FC<SidebarProps> = ({
  projects,
  activeProjectId,
  onSelectProject,
  onCreateProject,
  onDeleteProject,
  noteCounts,
  gitHubConfig,
  onOpenSettings,
  onSync,
  isSyncing,
  isOpenMobile,
  onCloseMobile,
}) => {
  const [isAdding, setIsAdding] = useState(false);
  const [newProjectName, setNewProjectName] = useState('');
  const [newIsLocalOnly, setNewIsLocalOnly] = useState(false);

  const handleAddSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newProjectName.trim()) return;
    onCreateProject(newProjectName.trim(), newIsLocalOnly);
    setNewProjectName('');
    setNewIsLocalOnly(false);
    setIsAdding(false);
  };

  const totalNotes = Object.values(noteCounts).reduce((a, b) => a + b, 0);

  return (
    <>
      {/* 모바일 배경 오버레이 */}
      {isOpenMobile && (
        <div 
          className="fixed inset-0 bg-black/60 z-40 md:hidden backdrop-blur-sm"
          onClick={onCloseMobile}
        />
      )}

      <aside className={`
        fixed md:static inset-y-0 left-0 z-50
        w-72 bg-slate-950 border-r border-slate-800
        flex flex-col h-full transition-transform duration-200 ease-in-out
        ${isOpenMobile ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}
      `}>
        {/* 상단 로고 & 타이틀 */}
        <div className="p-4 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-white shadow-lg shadow-indigo-500/20">
              M
            </div>
            <div>
              <h1 className="font-bold text-slate-100 text-base leading-tight">MyMemo</h1>
              <p className="text-xs text-slate-400">GitHub & Private Sync</p>
            </div>
          </div>
          <button 
            onClick={onCloseMobile}
            className="md:hidden text-slate-400 hover:text-white p-1 rounded-lg"
          >
            <X size={20} />
          </button>
        </div>

        {/* 프로젝트 목록 헤더 */}
        <div className="px-4 py-3 flex items-center justify-between text-xs font-semibold text-slate-400 uppercase tracking-wider">
          <span>큰주제 (프로젝트)</span>
          <button
            onClick={() => setIsAdding(!isAdding)}
            className="text-indigo-400 hover:text-indigo-300 flex items-center gap-1 font-normal capitalize hover:bg-slate-900 px-2 py-1 rounded"
          >
            <FolderPlus size={14} />
            <span>추가</span>
          </button>
        </div>

        {/* 새 프로젝트 추가 폼 */}
        {isAdding && (
          <form onSubmit={handleAddSubmit} className="p-3 mx-3 mb-2 bg-slate-900 border border-slate-700/80 rounded-lg space-y-2">
            <input
              type="text"
              placeholder="프로젝트 이름..."
              value={newProjectName}
              onChange={(e) => setNewProjectName(e.target.value)}
              className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-xs text-slate-100 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
              autoFocus
            />
            <label className="flex items-center gap-2 text-xs text-slate-300 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={newIsLocalOnly}
                onChange={(e) => setNewIsLocalOnly(e.target.checked)}
                className="accent-amber-500 rounded"
              />
              <span className="flex items-center gap-1">
                <Lock size={12} className="text-amber-400" />
                기기 전용 프로젝트 (GitHub 미동기화)
              </span>
            </label>
            <div className="flex gap-2 justify-end pt-1">
              <button
                type="button"
                onClick={() => setIsAdding(false)}
                className="px-2 py-1 text-xs text-slate-400 hover:text-slate-200"
              >
                취소
              </button>
              <button
                type="submit"
                className="px-2.5 py-1 text-xs bg-indigo-600 hover:bg-indigo-500 text-white rounded font-medium"
              >
                생성
              </button>
            </div>
          </form>
        )}

        {/* 프로젝트 리스트 */}
        <div className="flex-1 overflow-y-auto px-2 space-y-1">
          {/* 전체 메모 보기 */}
          <button
            onClick={() => {
              onSelectProject('ALL');
              onCloseMobile();
            }}
            className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-sm font-medium transition-colors ${
              activeProjectId === 'ALL'
                ? 'bg-indigo-600/20 text-indigo-300 border border-indigo-500/30'
                : 'text-slate-300 hover:bg-slate-900 hover:text-slate-100'
            }`}
          >
            <div className="flex items-center gap-2.5">
              <FileText size={16} className="text-indigo-400" />
              <span>전체 메모</span>
            </div>
            <span className="text-xs bg-slate-800 text-slate-400 px-1.5 py-0.5 rounded-full font-mono">
              {totalNotes}
            </span>
          </button>

          {projects.map((project) => {
            const count = noteCounts[project.id] || 0;
            const isSelected = activeProjectId === project.id;

            return (
              <div
                key={project.id}
                className="group flex items-center justify-between relative"
              >
                <button
                  onClick={() => {
                    onSelectProject(project.id);
                    onCloseMobile();
                  }}
                  className={`flex-1 flex items-center justify-between px-3 py-2 rounded-lg text-sm transition-colors text-left ${
                    isSelected
                      ? 'bg-slate-800 text-white font-medium'
                      : 'text-slate-300 hover:bg-slate-900 hover:text-slate-100'
                  }`}
                >
                  <div className="flex items-center gap-2 truncate pr-2">
                    {project.isLocalOnly ? (
                      <Lock size={15} className="text-amber-400 shrink-0" />
                    ) : (
                      <Folder size={15} className="text-indigo-400 shrink-0" />
                    )}
                    <span className="truncate">{project.name}</span>
                  </div>
                  <span className="text-xs text-slate-500 font-mono">
                    {count}
                  </span>
                </button>
                {/* 기본 프로젝트가 아닐 때만 삭제 허용 */}
                {project.id !== 'proj-default' && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      if (confirm(`'${project.name}' 프로젝트와 내부 메모를 모두 삭제하시겠습니까?`)) {
                        onDeleteProject(project.id);
                      }
                    }}
                    className="opacity-0 group-hover:opacity-100 text-slate-500 hover:text-rose-400 p-1.5 transition-opacity"
                    title="프로젝트 삭제"
                  >
                    <X size={13} />
                  </button>
                )}
              </div>
            );
          })}
        </div>

        {/* 하단 동기화 및 설정 바 */}
        <div className="p-3 border-t border-slate-800 bg-slate-950/80 space-y-2">
          {/* 동기화 버튼 */}
          <button
            onClick={onSync}
            disabled={isSyncing || !gitHubConfig?.isConnected}
            className={`w-full flex items-center justify-center gap-2 py-2 px-3 rounded-lg text-xs font-medium transition-all ${
              gitHubConfig?.isConnected
                ? 'bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 active:scale-98'
                : 'bg-slate-900 text-slate-500 border border-slate-800 cursor-not-allowed'
            }`}
          >
            <RefreshCw size={14} className={isSyncing ? 'animate-spin text-indigo-400' : ''} />
            <span>
              {isSyncing 
                ? 'GitHub 동기화 중...' 
                : gitHubConfig?.isConnected 
                  ? '지금 GitHub 동기화' 
                  : 'GitHub 미연결'}
            </span>
          </button>

          {/* 설정 버튼 */}
          <div className="flex items-center justify-between pt-1">
            <div className="flex items-center gap-1.5 text-xs text-slate-400 truncate pr-2">
              {gitHubConfig?.isConnected ? (
                <>
                  <CheckCircle2 size={13} className="text-emerald-400 shrink-0" />
                  <span className="truncate">{gitHubConfig.owner}/{gitHubConfig.repo}</span>
                </>
              ) : (
                <>
                  <AlertCircle size={13} className="text-amber-400 shrink-0" />
                  <span>설정 필요</span>
                </>
              )}
            </div>

            <button
              onClick={onOpenSettings}
              className="p-1.5 text-slate-400 hover:text-slate-100 hover:bg-slate-800 rounded-lg transition-colors"
              title="GitHub 연동 설정"
            >
              <Settings size={16} />
            </button>
          </div>
        </div>
      </aside>
    </>
  );
};
