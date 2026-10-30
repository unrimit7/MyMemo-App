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
  ChevronDown,
  Pencil,
  PanelLeftClose
} from 'lucide-react';

interface SidebarProps {
  projects: Project[];
  activeProjectId: string;
  onSelectProject: (id: string) => void;
  onCreateProject: (name: string, isLocalOnly: boolean, parentId?: string | null) => void;
  onDeleteProject: (id: string) => void;
  onRenameProject: (id: string, newName: string) => void;
  noteCounts: Record<string, number>;
  gitHubConfig: GitHubConfig | null;
  onOpenSettings: () => void;
  onSync: () => void;
  isSyncing: boolean;
  isOpenMobile: boolean;
  onCloseMobile: () => void;
  onInstallApp?: () => void;
  canInstall?: boolean;
  externalAddTarget?: { parentId: string | null; key: number } | null;
  onDropFilesToProject?: (projectId: string, files: FileList) => void;
  isCollapsed?: boolean;
  onToggleCollapse?: () => void;
  pendingCount?: number;
}

export const Sidebar: React.FC<SidebarProps> = ({
  projects,
  activeProjectId,
  onSelectProject,
  onCreateProject,
  onDeleteProject,
  onRenameProject,
  noteCounts,
  gitHubConfig,
  onOpenSettings,
  onSync,
  isSyncing,
  isOpenMobile,
  onCloseMobile,
  onInstallApp,
  canInstall,
  externalAddTarget,
  onDropFilesToProject,
  isCollapsed,
  onToggleCollapse,
  pendingCount,
}) => {
  // 드래그 오버 중인 폴더 ID
  const [dragOverFolderId, setDragOverFolderId] = useState<string | null>(null);

  // 새 폴더 모달/인라인 상태
  const [addingTarget, setAddingTarget] = useState<{ parentId: string | null; isOpen: boolean }>({
    parentId: null,
    isOpen: false,
  });
  const [newProjectName, setNewProjectName] = useState('');
  const [newIsLocalOnly, setNewIsLocalOnly] = useState(false);

  // 폴더 이름 수정 팝업 상태
  const [renamingTarget, setRenamingTarget] = useState<{ id: string; name: string } | null>(null);
  const [renameInputValue, setRenameInputValue] = useState('');

  // 외부(NoteList)에서 하위 폴더 생성 트리거 시 폼 열기
  React.useEffect(() => {
    if (externalAddTarget) {
      setAddingTarget({ parentId: externalAddTarget.parentId, isOpen: true });
      const parent = projects.find((p) => p.id === externalAddTarget.parentId);
      if (parent) setNewIsLocalOnly(parent.isLocalOnly);
    }
  }, [externalAddTarget, projects]);

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
    
    if (addingTarget.parentId) {
      setExpandedFolders((prev) => new Set(prev).add(addingTarget.parentId!));
    }

    setNewProjectName('');
    setNewIsLocalOnly(false);
    setAddingTarget({ parentId: null, isOpen: false });
  };

  const handleRenameSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!renamingTarget || !renameInputValue.trim()) return;
    onRenameProject(renamingTarget.id, renameInputValue.trim());
    setRenamingTarget(null);
  };

  const totalNotes = Object.values(noteCounts).reduce((a, b) => a + b, 0);

  // 재귀적 폴더 렌더링 함수
  const renderProjectNode = (project: Project, depth: number) => {
    const children = projectTree.childrenMap[project.id] || [];
    const hasChildren = children.length > 0;
    const isExpanded = expandedFolders.has(project.id);
    const isSelected = activeProjectId === project.id;
    const count = noteCounts[project.id] || 0;
    const canHaveSubfolder = depth < 2;

    const isDragOver = dragOverFolderId === project.id;

    return (
      <div 
        key={project.id} 
        className="flex flex-col select-none"
        onDragOver={(e) => {
          e.preventDefault();
          e.stopPropagation();
          setDragOverFolderId(project.id);
        }}
        onDragLeave={(e) => {
          e.preventDefault();
          e.stopPropagation();
          if (dragOverFolderId === project.id) setDragOverFolderId(null);
        }}
        onDrop={(e) => {
          e.preventDefault();
          e.stopPropagation();
          setDragOverFolderId(null);
          if (e.dataTransfer.files && e.dataTransfer.files.length > 0 && onDropFilesToProject) {
            onDropFilesToProject(project.id, e.dataTransfer.files);
          }
        }}
      >
        <div
          className={`group flex items-center justify-between rounded-lg text-sm transition-all pr-1.5 ${
            isDragOver
              ? 'bg-indigo-600/30 border-2 border-dashed border-indigo-400 text-white font-bold'
              : isSelected
                ? 'bg-slate-800 text-white font-medium shadow-sm'
                : 'text-slate-300 hover:bg-slate-900/80 hover:text-slate-100'
          }`}
          style={{ paddingLeft: `${depth * 14 + 8}px` }}
        >
          {/* 폴더 접기/펼치기 아이콘 */}
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

          {/* 상시 표시 액션: 하위 폴더 추가 & 이름 수정 & 삭제 */}
          <div className="flex items-center opacity-70 group-hover:opacity-100 transition-opacity gap-1 shrink-0">
            {canHaveSubfolder && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  setAddingTarget({ parentId: project.id, isOpen: true });
                  setNewIsLocalOnly(project.isLocalOnly);
                }}
                className="p-1.5 text-slate-400 hover:text-indigo-300 hover:bg-slate-700/60 rounded-md transition-colors"
                title={`${project.name}의 하위 폴더 추가`}
              >
                <FolderPlus size={14} />
              </button>
            )}

            {/* 폴더 이름 수정 버튼 (기본 폴더 제외) */}
            {project.id !== 'proj-default' && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  setRenamingTarget({ id: project.id, name: project.name });
                  setRenameInputValue(project.name);
                }}
                className="p-1.5 text-slate-400 hover:text-amber-300 hover:bg-slate-700/60 rounded-md transition-colors"
                title="폴더 이름 수정"
              >
                <Pencil size={13} />
              </button>
            )}

            {project.id !== 'proj-default' && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  onDeleteProject(project.id);
                }}
                className="p-1.5 text-slate-500 hover:text-rose-400 hover:bg-slate-700/60 rounded-md transition-colors"
                title="폴더 삭제"
              >
                <X size={14} />
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

      {/* 폴더 이름 수정 모달 팝업 */}
      {renamingTarget && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm animate-in fade-in duration-150">
          <form 
            onSubmit={handleRenameSubmit}
            className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-sm shadow-2xl p-5 text-slate-100 space-y-4 animate-in zoom-in-95 duration-150"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-2.5">
              <div className="p-2 rounded-xl bg-amber-500/10 text-amber-400 border border-amber-500/20">
                <Pencil size={18} />
              </div>
              <div>
                <h3 className="text-sm font-bold text-slate-100">폴더 이름 수정</h3>
                <p className="text-xs text-slate-400">새로운 폴더 이름을 입력하세요.</p>
              </div>
            </div>

            <input
              type="text"
              value={renameInputValue}
              onChange={(e) => setRenameInputValue(e.target.value)}
              className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-indigo-500"
              autoFocus
            />

            <div className="flex justify-end gap-2 pt-1">
              <button
                type="button"
                onClick={() => setRenamingTarget(null)}
                className="px-3 py-1.5 text-xs text-slate-400 hover:text-slate-200"
              >
                취소
              </button>
              <button
                type="submit"
                className="px-4 py-1.5 text-xs bg-indigo-600 hover:bg-indigo-500 text-white font-semibold rounded-xl transition-all"
              >
                수정 완료
              </button>
            </div>
          </form>
        </div>
      )}

      <aside className={`
        fixed md:static inset-y-0 left-0 z-50
        w-72 bg-slate-950 border-r border-slate-800
        flex flex-col h-full transition-all duration-200 ease-in-out
        ${isOpenMobile ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}
        ${isCollapsed ? 'md:hidden' : 'md:flex'}
      `}>
        {/* 상단 타이틀 */}
        <div className="p-4 border-b border-slate-800 flex items-center justify-between">
          <div>
            <h1 className="font-bold text-slate-100 text-base leading-tight">MyMemo</h1>
            <p className="text-xs text-slate-400">계층형 폴더 & Private Sync</p>
          </div>
          <div className="flex items-center gap-1">
            {onToggleCollapse && (
              <button
                onClick={onToggleCollapse}
                className="hidden md:flex p-1.5 text-slate-400 hover:text-slate-100 hover:bg-slate-900 rounded-lg transition-colors"
                title="사이드바(폴더) 접기"
              >
                <PanelLeftClose size={18} />
              </button>
            )}
            <button 
              onClick={onCloseMobile}
              className="md:hidden text-slate-400 hover:text-white p-1 rounded-lg"
            >
              <X size={20} />
            </button>
          </div>
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
            {pendingCount !== undefined && pendingCount > 0 && !isSyncing && (
              <span className="px-1.5 py-0.5 text-[10px] bg-amber-500/20 text-amber-300 border border-amber-500/40 rounded-full font-bold ml-0.5">
                {pendingCount}
              </span>
            )}
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
