import React, { useState, useMemo } from 'react';
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
  FileText,
  Download,
  ChevronRight,
  ChevronDown
} from 'lucide-react';

interface SidebarProps {
  projects: Project[];
  activeProjectId: string;
  onSelectProject: (id: string) => void;
  onCreateProject: (name: string, isLocalOnly: boolean, parentId?: string | null) => void;
  onDeleteProject: (id: string) => void;
  noteCounts: Record<string, number>;
  gitHubConfig: GitHubConfig | null;
  onOpenSettings: () => void;
  onSync: () => void;
  isSyncing: boolean;
  isOpenMobile: boolean;
  onCloseMobile: () => void;
  onInstallApp?: () => void;
  canInstall?: boolean;
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
  onInstallApp,
  canInstall,
}) => {
  // 새 폴더 모달/인라인 상태: { parentId: string | null | undefined, isOpen: boolean }
  const [addingTarget, setAddingTarget] = useState<{ parentId: string | null; isOpen: boolean }>({
    parentId: null,
    isOpen: false,
  });
  const [newProjectName, setNewProjectName] = useState('');
  const [newIsLocalOnly, setNewIsLocalOnly] = useState(false);

  // 펼쳐진 폴더 ID 세트
  const [expandedFolders, setExpandedFolders] = useState<Set<string>>(new Set());

  // 폴더 접기/펼치기 토글
  const toggleFolder = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    setExpandedFolders((prev) => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  // 각 프로젝트의 depth 계산 맵
  const projectDepthMap = useMemo(() => {
    const depthMap: Record<string, number> = {};
    const getDepth = (id: string): number => {
      if (depthMap[id] !== undefined) return depthMap[id];
      const p = projects.find((proj) => proj.id === id);
      if (!p || !p.parentId) {
        depthMap[id] = 0;
        return 0;
      }
      const d = getDepth(p.parentId) + 1;
      depthMap[id] = d;
      return d;
    };
    projects.forEach((p) => getDepth(p.id));
    return depthMap;
  }, [projects]);

  // 트리 구조 빌드
  const projectTree = useMemo(() => {
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

    return { rootProjects, childrenMap };
  }, [projects]);

  const handleAddSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newProjectName.trim()) return;
    onCreateProject(newProjectName.trim(), newIsLocalOnly, addingTarget.parentId);
    
    // 만약 상위 폴더가 있다면 그 폴더를 자동으로 펼침
    if (addingTarget.parentId) {
      setExpandedFolders((prev) => new Set(prev).add(addingTarget.parentId!));
    }

    setNewProjectName('');
    setNewIsLocalOnly(false);
    setAddingTarget({ parentId: null, isOpen: false });
  };

  const totalNotes = Object.values(noteCounts).reduce((a, b) => a + b, 0);

  // 재귀적 폴더 렌더링 함수
  const renderProjectNode = (project: Project, depth: number) => {
    const children = projectTree.childrenMap[project.id] || [];
    const hasChildren = children.length > 0;
    const isExpanded = expandedFolders.has(project.id);
    const isSelected = activeProjectId === project.id;
    const count = noteCounts[project.id] || 0;
    const canHaveSubfolder = depth < 2; // 최대 2단계(depth=0, depth=1)까지만 하위 폴더 생성 가능

    return (
      <div key={project.id} className="flex flex-col select-none">
        <div
          className={`group flex items-center justify-between rounded-lg text-sm transition-all pr-1.5 ${
            isSelected
              ? 'bg-slate-800 text-white font-medium shadow-sm'
              : 'text-slate-300 hover:bg-slate-900/80 hover:text-slate-100'
          }`}
          style={{ paddingLeft: `${depth * 14 + 8}px` }}
        >
          {/* 폴더 접기/펼치기 아이콘 또는 공백 */}
          <button
            type="button"
            onClick={(e) => toggleFolder(project.id, e)}
            className={`p-1 text-slate-500 hover:text-slate-300 rounded ${
              hasChildren ? 'visible' : 'invisible w-4'
            }`}
          >
            {isExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          </button>

          {/* 폴더 클릭 (메모 목록 선택) */}
          <button
            type="button"
            onClick={() => {
              onSelectProject(project.id);
              onCloseMobile();
            }}
            className="flex-1 flex items-center justify-between py-2 pr-2 truncate text-left"
          >
            <div className="flex items-center gap-1.5 truncate">
              {project.isLocalOnly ? (
                <Lock size={14} className="text-amber-400 shrink-0" />
              ) : (
                <Folder size={14} className="text-indigo-400 shrink-0" />
              )}
              <span className="truncate text-xs md:text-sm">{project.name}</span>
            </div>
            <span className="text-[11px] text-slate-500 font-mono shrink-0 pl-1">
              {count}
            </span>
          </button>

          {/* 호버 액션: 하위 폴더 추가 (depth < 2일 때만) & 삭제 */}
          <div className="flex items-center opacity-0 group-hover:opacity-100 transition-opacity gap-0.5">
            {canHaveSubfolder && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  setAddingTarget({ parentId: project.id, isOpen: true });
                  setNewIsLocalOnly(project.isLocalOnly);
                }}
                className="p-1 text-slate-400 hover:text-indigo-300 hover:bg-slate-800 rounded"
                title="하위 폴더 추가"
              >
                <FolderPlus size={13} />
              </button>
            )}

            {project.id !== 'proj-default' && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  if (confirm(`'${project.name}' 폴더와 하위 메모를 모두 삭제하시겠습니까?`)) {
                    onDeleteProject(project.id);
                  }
                }}
                className="p-1 text-slate-500 hover:text-rose-400 hover:bg-slate-800 rounded"
                title="삭제"
              >
                <X size={13} />
              </button>
            )}
          </div>
        </div>

        {/* 하위 폴더 렌더링 (펼쳐졌을 때) */}
        {hasChildren && isExpanded && (
          <div className="border-l border-slate-800/60 ml-4 space-y-0.5 my-0.5">
            {children.map((child) => renderProjectNode(child, depth + 1))}
          </div>
        )}
      </div>
    );
  };

  const targetParent = projects.find((p) => p.id === addingTarget.parentId);
  const targetParentDepth = addingTarget.parentId ? projectDepthMap[addingTarget.parentId] : -1;

  return (
    <>
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
        {/* 상단 로고 */}
        <div className="p-4 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-white shadow-lg shadow-indigo-500/20">
              M
            </div>
            <div>
              <h1 className="font-bold text-slate-100 text-base leading-tight">MyMemo</h1>
              <p className="text-xs text-slate-400">계층형 폴더 & Private Sync</p>
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
          <span>폴더 및 카테고리</span>
          <button
            onClick={() => setAddingTarget({ parentId: null, isOpen: !addingTarget.isOpen })}
            className="text-indigo-400 hover:text-indigo-300 flex items-center gap-1 font-normal capitalize hover:bg-slate-900 px-2 py-1 rounded"
            title="메인 폴더 추가"
          >
            <FolderPlus size={14} />
            <span>메인 추가</span>
          </button>
        </div>

        {/* 새 폴더 추가 폼 */}
        {addingTarget.isOpen && (
          <form onSubmit={handleAddSubmit} className="p-3 mx-3 mb-2 bg-slate-900 border border-slate-700/80 rounded-lg space-y-2 animate-in fade-in duration-150">
            <div className="text-[11px] text-slate-400 font-medium">
              {targetParent ? (
                <span>
                  📁 <strong className="text-indigo-300">{targetParent.name}</strong> 의 하위 폴더 생성 (
                  {targetParentDepth + 1 === 1 ? '1단계' : '2단계 - 메모 전용'})
                </span>
              ) : (
                <span>📁 최상위 메인 폴더 생성</span>
              )}
            </div>
            <input
              type="text"
              placeholder="폴더 이름 입력..."
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
                기기 전용 (GitHub 미동기화)
              </span>
            </label>
            <div className="flex gap-2 justify-end pt-1">
              <button
                type="button"
                onClick={() => setAddingTarget({ parentId: null, isOpen: false })}
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

        {/* 폴더 리스트 */}
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

          {/* 최상위 폴더부터 재귀 렌더링 */}
          {projectTree.rootProjects.map((proj) => renderProjectNode(proj, 0))}
        </div>

        {/* 하단 동기화 및 설정 바 */}
        <div className="p-3 border-t border-slate-800 bg-slate-950/80 space-y-2">
          {canInstall && onInstallApp && (
            <button
              onClick={onInstallApp}
              className="w-full flex items-center justify-center gap-2 py-2 px-3 rounded-lg text-xs font-semibold bg-indigo-600/20 hover:bg-indigo-600/30 text-indigo-300 border border-indigo-500/40 transition-all shadow-sm active:scale-98"
            >
              <Download size={14} className="text-indigo-400" />
              <span>이 기기에 앱 설치</span>
            </button>
          )}

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
