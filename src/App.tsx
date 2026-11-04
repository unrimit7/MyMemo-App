import { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import type { Project, Note, GitHubConfig, SyncStatus, AppTheme } from './types';
import { THEME_OPTIONS } from './types';
import { 
  seedInitialDataIfNeeded, 
  getAllProjects, 
  getAllNotes, 
  saveProject, 
  deleteProject, 
  saveNote, 
  deleteNote, 
  getGitHubConfig, 
  saveGitHubConfig,
  trackDeletedNote,
  trackDeletedProject,
  requestPersistentStorage
} from './services/db';
import { syncWithGitHub, getProjectFolderPath } from './services/github';
import { Sidebar } from './components/Sidebar';
import { NoteList } from './components/NoteList';
import { NoteEditor } from './components/NoteEditor';
import { SettingsModal } from './components/SettingsModal';
import { ConfirmModal } from './components/ConfirmModal';
import { CalendarView } from './components/CalendarView';
import { processFileForNote } from './utils/file';
import { getFolderHierarchyPath, updateTitleWithFolderPath, stripFolderPrefixFromTitle } from './utils/folderPath';
import { Calendar as CalendarIcon, X, FolderTree } from 'lucide-react';
import { crossTabSync } from './services/crossTabSync';
import type { MyMemoBackupData } from './utils/backup';

interface DeleteConfirmState {
  isOpen: boolean;
  type: 'note' | 'project';
  id: string;
  title: string;
  message: string;
  subMessage?: string;
}

export function App() {
  const [projects, setProjects] = useState<Project[]>([]);
  const [notes, setNotes] = useState<Note[]>([]);
  const [activeProjectId, setActiveProjectId] = useState<string>(() => {
    return localStorage.getItem('mymemo_active_project_id') || 'ALL';
  });
  const [activeNoteId, setActiveNoteId] = useState<string | null>(() => {
    return localStorage.getItem('mymemo_active_note_id') || null;
  });

  const [gitHubConfig, setGitHubConfig] = useState<GitHubConfig | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isSidebarOpenMobile, setIsSidebarOpenMobile] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncToast, setSyncToast] = useState<string | null>(null);
  const [otherTabsCount, setOtherTabsCount] = useState<number>(0);
  const [isCalendarOpen, setIsCalendarOpen] = useState(false);

  // 모바일 화면 뷰 모드 ('list' 또는 'editor')
  const [mobileScreen, setMobileScreen] = useState<'list' | 'editor'>(() => {
    return (localStorage.getItem('mymemo_mobile_screen') as 'list' | 'editor') || 'list';
  });
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [externalAddTarget, setExternalAddTarget] = useState<{ parentId: string | null; key: number } | null>(null);

  // 데스크톱 좌측 패널(사이드바, 메모목록) 접기/펼치기 상태
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(() => {
    return localStorage.getItem('mymemo_sidebar_collapsed') === 'true';
  });
  const [isNoteListCollapsed, setIsNoteListCollapsed] = useState<boolean>(() => {
    return localStorage.getItem('mymemo_notelist_collapsed') === 'true';
  });

  const toggleSidebar = () => {
    setIsSidebarCollapsed((prev) => {
      const next = !prev;
      localStorage.setItem('mymemo_sidebar_collapsed', String(next));
      return next;
    });
  };

  const toggleNoteList = () => {
    setIsNoteListCollapsed((prev) => {
      const next = !prev;
      localStorage.setItem('mymemo_notelist_collapsed', String(next));
      return next;
    });
  };

  // 좌측 1단(사이드바) 및 2단(메모목록) 마우스 드래그 너비 조절 상태
  const [sidebarWidth, setSidebarWidth] = useState<number>(() => {
    const saved = localStorage.getItem('mymemo_sidebar_width');
    return saved ? parseInt(saved, 10) : 260;
  });
  const [noteListWidth, setNoteListWidth] = useState<number>(() => {
    const saved = localStorage.getItem('mymemo_notelist_width');
    return saved ? parseInt(saved, 10) : 320;
  });
  const [isResizingSidebar, setIsResizingSidebar] = useState(false);
  const [isResizingNoteList, setIsResizingNoteList] = useState(false);

  // 앱 분위기 (테마) 설정 상태
  const [currentTheme, setCurrentTheme] = useState<AppTheme>(() => {
    if (typeof window === 'undefined') return 'navy';
    const saved = localStorage.getItem('mymemo_theme') as AppTheme;
    if (saved && THEME_OPTIONS.some((t) => t.id === saved)) return saved;
    return 'navy';
  });

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', currentTheme);
    localStorage.setItem('mymemo_theme', currentTheme);
  }, [currentTheme]);

  const handleCycleTheme = () => {
    const currentIndex = THEME_OPTIONS.findIndex((t) => t.id === currentTheme);
    const nextIndex = (currentIndex + 1) % THEME_OPTIONS.length;
    setCurrentTheme(THEME_OPTIONS[nextIndex].id);
  };

  // 사이드바 너비 마우스 드래그 리사이징 핸들러
  const handleSidebarResizeStart = (e: React.MouseEvent) => {
    e.preventDefault();
    setIsResizingSidebar(true);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';

    const handleMouseMove = (moveEvent: MouseEvent) => {
      const newWidth = Math.max(180, Math.min(moveEvent.clientX, Math.min(500, window.innerWidth * 0.4)));
      setSidebarWidth(newWidth);
    };

    const handleMouseUp = () => {
      setIsResizingSidebar(false);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
      setSidebarWidth((w) => {
        localStorage.setItem('mymemo_sidebar_width', String(w));
        return w;
      });
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
  };

  // 메모목록 너비 마우스 드래그 리사이징 핸들러
  const handleNoteListResizeStart = (e: React.MouseEvent) => {
    e.preventDefault();
    setIsResizingNoteList(true);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';

    const handleMouseMove = (moveEvent: MouseEvent) => {
      const currentSidebarW = isSidebarCollapsed ? 0 : sidebarWidth;
      const newWidth = Math.max(220, Math.min(moveEvent.clientX - currentSidebarW, Math.min(650, window.innerWidth * 0.5)));
      setNoteListWidth(newWidth);
    };

    const handleMouseUp = () => {
      setIsResizingNoteList(false);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
      setNoteListWidth((w) => {
        localStorage.setItem('mymemo_notelist_width', String(w));
        return w;
      });
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
  };

  // 삭제 확인 모달 상태
  const [deleteConfirm, setDeleteConfirm] = useState<DeleteConfirmState>({
    isOpen: false,
    type: 'note',
    id: '',
    title: '',
    message: '',
  });

  const autoSyncTimerRef = useRef<any>(null);
  const isSyncingRef = useRef(false);

  useEffect(() => {
    const handleBeforeInstall = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };
    window.addEventListener('beforeinstallprompt', handleBeforeInstall);
    return () => window.removeEventListener('beforeinstallprompt', handleBeforeInstall);
  }, []);

  const handleInstallApp = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    if (outcome === 'accepted') {
      setDeferredPrompt(null);
    }
  };

  // GitHub 동기화 공통 실행 로직
  const runSync = useCallback(async (silent: boolean = false, customConfig?: GitHubConfig) => {
    const targetConfig = customConfig || gitHubConfig;
    if (!targetConfig || !targetConfig.isConnected || isSyncingRef.current) {
      if (!silent && (!targetConfig || !targetConfig.isConnected)) {
        setIsSettingsOpen(true);
      }
      return;
    }

    // 다른 창에서 이미 동기화가 진행 중이면 중복 실행 방지
    if (!crossTabSync.acquireSyncLock()) {
      if (!silent) {
        setSyncToast('⚠️ 다른 창에서 이미 GitHub 동기화가 진행 중입니다.');
        setTimeout(() => setSyncToast(null), 3000);
      }
      return;
    }

    isSyncingRef.current = true;
    setIsSyncing(true);
    if (!silent) {
      setSyncToast('GitHub와 동기화 중입니다...');
    }

    try {
      const summary = await syncWithGitHub(targetConfig);
      const updatedNotes = await getAllNotes();
      const updatedProjects = await getAllProjects();
      setNotes(updatedNotes);
      setProjects(updatedProjects);

      const updatedConfig = { ...targetConfig, lastSyncedAt: summary.timestamp };
      await saveGitHubConfig(updatedConfig);
      setGitHubConfig(updatedConfig);

      crossTabSync.broadcastChange('sync_complete');

      if (!silent) {
        setSyncToast(
          `동기화 완료! (업로드: ${summary.uploaded}, 다운로드: ${summary.downloaded}, 로컬보호: ${summary.skippedLocalOnly})`
        );
      } else if (summary.uploaded > 0 || summary.downloaded > 0) {
        setSyncToast(`자동 동기화 완료: 최신 메모를 가져왔습니다.`);
      }
    } catch (err: any) {
      if (!silent) {
        setSyncToast(`동기화 실패: ${err.message || '네트워크 오류'}`);
      }
    } finally {
      crossTabSync.releaseSyncLock();
      isSyncingRef.current = false;
      setIsSyncing(false);
      setTimeout(() => setSyncToast(null), 3500);
    }
  }, [gitHubConfig]);

  // 1. 초기 데이터 로딩 및 앱 켤 때 자동 동기화 1회 실행
  useEffect(() => {
    const initApp = async () => {
      // 브라우저 및 윈도우 스토리지 센스로부터 로컬 데이터 영구 보존 요청
      requestPersistentStorage().catch(() => {});

      await seedInitialDataIfNeeded();
      const loadedProjects = await getAllProjects();
      const loadedNotes = await getAllNotes();
      const config = await getGitHubConfig();

      setProjects(loadedProjects);
      setNotes(loadedNotes);
      setGitHubConfig(config);

      // 마지막으로 선택했던 폴더 및 메모 복원
      const savedProjId = localStorage.getItem('mymemo_active_project_id');
      const savedNoteId = localStorage.getItem('mymemo_active_note_id');

      if (savedProjId && (savedProjId === 'ALL' || loadedProjects.some((p) => p.id === savedProjId))) {
        setActiveProjectId(savedProjId);
      } else {
        setActiveProjectId('ALL');
        localStorage.setItem('mymemo_active_project_id', 'ALL');
      }

      if (savedNoteId && loadedNotes.some((n) => n.id === savedNoteId)) {
        setActiveNoteId(savedNoteId);
      } else if (loadedNotes.length > 0) {
        // 선택된 폴더가 있다면 그 폴더 내의 첫 메모를 우선 선택
        const notesInProj = savedProjId && savedProjId !== 'ALL'
          ? loadedNotes.filter((n) => n.projectId === savedProjId)
          : loadedNotes;
        const targetNote = notesInProj[0] || loadedNotes[0];
        setActiveNoteId(targetNote.id);
        localStorage.setItem('mymemo_active_note_id', targetNote.id);
      }

      // 설정이 연결되어 있고 '앱 시작 시 자동 동기화(syncOnStartup)' 또는 '실시간 자동 동기화(autoSync)' 옵션이 켜져 있을 때만 앱 시작 시 동기화
      if (config && config.isConnected && (config.syncOnStartup || config.autoSync)) {
        runSync(true, config);
      }
    };

    initApp();
  }, []);

  // 멀티 탭/창 실시간 연동 리스너 등록
  useEffect(() => {
    crossTabSync.setOtherTabsCallback((count) => {
      setOtherTabsCount(count);
    });

    const unsubscribe = crossTabSync.subscribe(async (msg) => {
      if (msg.type === 'DATA_CHANGED') {
        const [latestNotes, latestProjects] = await Promise.all([
          getAllNotes(),
          getAllProjects(),
        ]);
        setNotes(latestNotes);
        setProjects(latestProjects);
      }
    });

    return () => {
      unsubscribe();
    };
  }, []);

  // 2. 화면으로 복귀할 때 서비스워커 갱신 확인 및 최신 데이터 동기화 (autoSync 옵션 켜진 경우에만)
  useEffect(() => {
    const handleVisibilityChange = async () => {
      if (document.visibilityState === 'visible') {
        if ('serviceWorker' in navigator) {
          navigator.serviceWorker.getRegistration().then((reg) => reg?.update());
        }
        // 다른 창에서 데이터가 바뀌었을 수 있으므로 로컬 데이터 갱신
        const [latestNotes, latestProjects] = await Promise.all([
          getAllNotes(),
          getAllProjects(),
        ]);
        setNotes(latestNotes);
        setProjects(latestProjects);

        if (gitHubConfig?.isConnected && gitHubConfig.autoSync) {
          runSync(true);
        }
      }
    };
    document.addEventListener('visibilitychange', handleVisibilityChange);
    return () => document.removeEventListener('visibilitychange', handleVisibilityChange);
  }, [gitHubConfig, runSync]);

  // GitHub 업로드 대기 중인 메모 개수 계산
  const pendingCount = useMemo(() => {
    return notes.filter((n) => !n.isLocalOnly && n.syncStatus === 'pending').length;
  }, [notes]);

  // 프로젝트별 메모 개수 계산
  const noteCounts = useMemo(() => {
    const counts: Record<string, number> = {};
    for (const note of notes) {
      counts[note.projectId] = (counts[note.projectId] || 0) + 1;
    }
    return counts;
  }, [notes]);

  // 현재 선택된 메모 객체
  const activeNote = useMemo(() => {
    return notes.find((n) => n.id === activeNoteId) || null;
  }, [notes, activeNoteId]);

  // 프로젝트(폴더) 추가 핸들러
  const handleCreateProject = async (name: string, isLocalOnly: boolean, parentId?: string | null) => {
    const siblings = projects.filter((p) => (p.parentId || null) === (parentId || null));
    const maxOrder = siblings.reduce((max, p) => Math.max(max, p.order ?? 0), 0);
    const newProject: Project = {
      id: `proj-${Date.now()}`,
      name,
      parentId: parentId || null,
      isLocalOnly,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      order: maxOrder + 10,
    };
    await saveProject(newProject);
    setProjects((prev) => [...prev, newProject]);
    setActiveProjectId(newProject.id);
    crossTabSync.broadcastChange('projects');
  };

  // 프로젝트(폴더) 정보(이름 및 기기 전용 여부) 수정 핸들러
  const handleRenameProject = async (id: string, newName: string, isLocalOnly?: boolean) => {
    const target = projects.find((p) => p.id === id);
    if (!target) return;

    const newIsLocalOnly = isLocalOnly !== undefined ? isLocalOnly : target.isLocalOnly;
    const isNameChanged = target.name !== newName;
    const isLocalOnlyChanged = target.isLocalOnly !== newIsLocalOnly;

    if (!isNameChanged && !isLocalOnlyChanged) return;

    // 해당 폴더와 모든 하위(자식, 손자 등) 폴더 ID 목록 추출
    const targetProjIds: string[] = [id];
    const findChildren = (pId: string) => {
      projects.filter((p) => p.parentId === pId).forEach((child) => {
        targetProjIds.push(child.id);
        findChildren(child.id);
      });
    };
    findChildren(id);

    const nowStr = new Date().toISOString();

    // 1. 폴더 목록 갱신 (부모가 기기 전용으로 바뀌면 하위 폴더들도 일괄 기기 전용으로 전환)
    const updatedProjects = projects.map((p) => {
      if (p.id === id) {
        return { ...p, name: newName, isLocalOnly: newIsLocalOnly, updatedAt: nowStr };
      }
      if (newIsLocalOnly && targetProjIds.includes(p.id)) {
        return { ...p, isLocalOnly: true, updatedAt: nowStr };
      }
      if (!newIsLocalOnly && targetProjIds.includes(p.id)) {
        return { ...p, isLocalOnly: false, updatedAt: nowStr };
      }
      return p;
    });

    for (const projId of targetProjIds) {
      const proj = updatedProjects.find((p) => p.id === projId);
      if (proj) await saveProject(proj);
    }
    setProjects(updatedProjects);

    // 2. 해당 폴더 및 하위 폴더에 속한 메모들 갱신
    const affectedNotes = notes.filter((n) => targetProjIds.includes(n.projectId));
    const updatedNotesMap = new Map<string, Note>();

    for (const note of affectedNotes) {
      const newFolderPath = getFolderHierarchyPath(note.projectId, updatedProjects);
      const newTitle = isNameChanged ? updateTitleWithFolderPath(note.title, newFolderPath) : note.title;

      let noteIsLocalOnly = note.isLocalOnly;
      let noteSyncStatus = note.syncStatus;

      // 폴더명이 바뀌었으면 GitHub 원격 파일 경로도 변경되어야 하므로 일반 동기화 대상은 pending으로 전환!
      if (isNameChanged && !noteIsLocalOnly) {
        noteSyncStatus = 'pending';
      }

      if (newIsLocalOnly) {
        // [일반 -> 기기 전용으로 전환]:
        noteIsLocalOnly = true;
        noteSyncStatus = 'local_only';
        // 이전에 GitHub에 업로드되어 있던 파일이라면, 원격 GitHub에서 삭제되도록 톰스톤에 등록!
        if (note.githubPath) {
          await trackDeletedNote(note.id, note.githubPath);
        }
      } else if (isLocalOnlyChanged && !newIsLocalOnly) {
        // [기기 전용 -> 일반 동기화로 전환]:
        noteIsLocalOnly = false;
        noteSyncStatus = 'pending';
      }

      const updatedNote: Note = {
        ...note,
        title: newTitle,
        isLocalOnly: noteIsLocalOnly,
        syncStatus: noteSyncStatus,
        updatedAt: new Date().toISOString(),
      };

      await saveNote(updatedNote);
      updatedNotesMap.set(note.id, updatedNote);
    }

    setNotes((prev) => prev.map((n) => updatedNotesMap.get(n.id) || n));
    crossTabSync.broadcastChange('projects');

    // 3. 토스트 알림 및 원격 동기화 처리
    if (isLocalOnlyChanged) {
      if (newIsLocalOnly) {
        setSyncToast(`🔒 '${newName}' 및 하위 항목이 '기기 전용'으로 전환되었습니다.`);
        // GitHub에 올라가 있던 파일들을 원격에서 안전하게 삭제 동기화
        if (gitHubConfig?.isConnected) {
          setTimeout(() => runSync(true), 1500);
        }
      } else {
        setSyncToast(`📁 '${newName}' 폴더가 'GitHub 동기화 대상'으로 전환되었습니다.`);
        if (gitHubConfig?.isConnected && gitHubConfig.autoSync) {
          setTimeout(() => runSync(true), 1500);
        }
      }
    } else {
      setSyncToast(`✏️ 폴더 이름이 '${newName}'(으)로 변경되었습니다.`);
      if (gitHubConfig?.isConnected && gitHubConfig.autoSync) {
        setTimeout(() => runSync(true), 1500);
      }
    }
    setTimeout(() => setSyncToast(null), 3500);
  };

  // 프로젝트(폴더) 드래그 앤 드롭 이동 및 순서 변경 핸들러
  const handleMoveProject = async (
    draggedId: string,
    targetId: string,
    position: 'before' | 'after' | 'inside'
  ) => {
    if (draggedId === targetId) return;

    const dragged = projects.find((p) => p.id === draggedId);
    if (!dragged) return;

    const isTargetRoot = targetId === 'ROOT';
    const target = isTargetRoot ? null : projects.find((p) => p.id === targetId);
    if (!isTargetRoot && !target) return;

    // 1. 자기 자신의 하위 폴더(자손) 안으로 이동 방지 (순환 참조 방지)
    const isDescendant = (ancestorId: string, checkId: string): boolean => {
      let curr = projects.find((p) => p.id === checkId);
      while (curr && curr.parentId) {
        if (curr.parentId === ancestorId) return true;
        const nextId: string = curr.parentId;
        curr = projects.find((p) => p.id === nextId);
      }
      return false;
    };
    if (!isTargetRoot && isDescendant(draggedId, targetId)) {
      setSyncToast('⚠️ 상위 폴더를 자신의 하위 폴더 안으로 이동할 수 없습니다.');
      setTimeout(() => setSyncToast(null), 3000);
      return;
    }

    // 2. 최대 depth(2단계 하위, 총 3단계 계층) 초과 방지 체크
    const getDepth = (id: string, list: Project[]): number => {
      let depth = 0;
      let curr = list.find((p) => p.id === id);
      while (curr && curr.parentId) {
        depth++;
        const nextId: string = curr.parentId;
        curr = list.find((p) => p.id === nextId);
      }
      return depth;
    };

    const getSubtreeHeight = (rootId: string, list: Project[]): number => {
      const children = list.filter((p) => p.parentId === rootId);
      if (children.length === 0) return 0;
      return 1 + Math.max(...children.map((c) => getSubtreeHeight(c.id, list)));
    };

    const draggedSubtreeHeight = getSubtreeHeight(draggedId, projects);

    let newParentId: string | null = null;
    if (isTargetRoot) {
      newParentId = null;
    } else if (position === 'inside') {
      newParentId = targetId;
      const targetDepth = getDepth(targetId, projects);
      if (targetDepth + 1 + draggedSubtreeHeight > 2) {
        setSyncToast('⚠️ 최대 2단계 하위 폴더까지만 지원됩니다.');
        setTimeout(() => setSyncToast(null), 3000);
        return;
      }
    } else {
      // before 또는 after인 경우: target과 동일한 parentId
      newParentId = target?.parentId || null;
      const parentDepth = newParentId ? getDepth(newParentId, projects) : -1;
      if (parentDepth + 1 + draggedSubtreeHeight > 2) {
        setSyncToast('⚠️ 최대 2단계 하위 폴더까지만 지원됩니다.');
        setTimeout(() => setSyncToast(null), 3000);
        return;
      }
    }

    const isParentChanged = (dragged.parentId || null) !== (newParentId || null);

    // 3. 기기 전용(isLocalOnly) 동기화 보안 규칙: 상위 폴더가 기기 전용이면 이동된 폴더도 기기 전용 전환
    let updatedIsLocalOnly = dragged.isLocalOnly;
    if (newParentId) {
      const parentProj = projects.find((p) => p.id === newParentId);
      if (parentProj?.isLocalOnly) {
        updatedIsLocalOnly = true;
      }
    }

    const nowStr = new Date().toISOString();
    const updatedDragged: Project = {
      ...dragged,
      parentId: newParentId,
      isLocalOnly: updatedIsLocalOnly,
      updatedAt: nowStr,
    };

    let updatedProjectsList = projects.map((p) => (p.id === draggedId ? updatedDragged : p));

    // 하위 폴더들도 부모가 기기 전용으로 바뀌면 연쇄 반영
    if (updatedIsLocalOnly && !dragged.isLocalOnly) {
      const cascadeLocal = (pid: string) => {
        updatedProjectsList = updatedProjectsList.map((p) => {
          if (p.parentId === pid) {
            cascadeLocal(p.id);
            return { ...p, isLocalOnly: true, updatedAt: nowStr };
          }
          return p;
        });
      };
      cascadeLocal(draggedId);
    }

    // 4. newParentId를 공유하는 형제 폴더들 목록 추출 및 순서(order) 재배열
    const siblings = updatedProjectsList
      .filter((p) => (p.parentId || null) === (newParentId || null) && p.id !== draggedId)
      .sort((a, b) => {
        const oA = a.order !== undefined ? a.order : 999999;
        const oB = b.order !== undefined ? b.order : 999999;
        if (oA !== oB) return oA - oB;
        return a.createdAt.localeCompare(b.createdAt);
      });

    if (isTargetRoot || position === 'inside') {
      siblings.push(updatedDragged);
    } else if (position === 'before') {
      const targetIdx = siblings.findIndex((p) => p.id === targetId);
      if (targetIdx !== -1) {
        siblings.splice(targetIdx, 0, updatedDragged);
      } else {
        siblings.unshift(updatedDragged);
      }
    } else {
      // after
      const targetIdx = siblings.findIndex((p) => p.id === targetId);
      if (targetIdx !== -1) {
        siblings.splice(targetIdx + 1, 0, updatedDragged);
      } else {
        siblings.push(updatedDragged);
      }
    }

    // 형제들에게 새로운 순서 번호(order: 10, 20, 30...) 할당
    const orderMap = new Map<string, number>();
    siblings.forEach((s, idx) => {
      orderMap.set(s.id, (idx + 1) * 10);
    });

    const finalProjects = updatedProjectsList.map((p) => {
      if (orderMap.has(p.id)) {
        return { ...p, order: orderMap.get(p.id)!, updatedAt: nowStr };
      }
      return p;
    });

    setProjects(finalProjects);

    // IndexedDB에 수정된 폴더 저장
    for (const p of finalProjects) {
      const orig = projects.find((o) => o.id === p.id);
      if (!orig || orig.parentId !== p.parentId || orig.order !== p.order || orig.isLocalOnly !== p.isLocalOnly) {
        await saveProject(p);
      }
    }

    // 5. 상위 폴더가 변경된 경우, 소속된 메모들의 계층 경로명 및 로컬 전용 상태 갱신
    if (isParentChanged) {
      const targetProjIds: string[] = [draggedId];
      const findChildren = (pId: string) => {
        finalProjects.filter((p) => p.parentId === pId).forEach((child) => {
          targetProjIds.push(child.id);
          findChildren(child.id);
        });
      };
      findChildren(draggedId);

      const affectedNotes = notes.filter((n) => targetProjIds.includes(n.projectId));
      const updatedNotesMap = new Map<string, Note>();

      for (const note of affectedNotes) {
        const newFolderPath = getFolderHierarchyPath(note.projectId, finalProjects);
        const newTitle = updateTitleWithFolderPath(note.title, newFolderPath);

        let noteIsLocalOnly = note.isLocalOnly;
        let noteSyncStatus: SyncStatus = 'pending'; // 폴더 경로가 바뀌었으므로 GitHub 원격 경로 이동 동기화 예약!

        if (updatedIsLocalOnly && !dragged.isLocalOnly) {
          noteIsLocalOnly = true;
          noteSyncStatus = 'local_only';
          if (note.githubPath) {
            await trackDeletedNote(note.id, note.githubPath);
          }
        }

        const updatedNote: Note = {
          ...note,
          title: newTitle,
          isLocalOnly: noteIsLocalOnly,
          syncStatus: noteSyncStatus,
          updatedAt: new Date().toISOString(),
        };

        await saveNote(updatedNote);
        updatedNotesMap.set(note.id, updatedNote);
      }

      if (updatedNotesMap.size > 0) {
        setNotes((prev) => prev.map((n) => updatedNotesMap.get(n.id) || n));
      }

      setSyncToast(
        newParentId
          ? `📁 '${dragged.name}' 폴더가 '${projects.find((p) => p.id === newParentId)?.name}' 하위로 이동되었습니다.`
          : `📁 '${dragged.name}' 폴더가 최상위 메인 폴더로 이동되었습니다.`
      );
      setTimeout(() => setSyncToast(null), 3000);
    }

    crossTabSync.broadcastChange('projects');

    if (gitHubConfig?.isConnected && gitHubConfig.autoSync) {
      setTimeout(() => runSync(true), 1500);
    }
  };

  // 프로젝트(폴더) 일괄 정렬 (가나다순, 최신순 등)
  const handleSortProjects = async (sortType: 'name-asc' | 'name-desc' | 'created-desc' | 'created-asc') => {
    const parentIds = Array.from(new Set(projects.map((p) => p.parentId || null)));
    const updatedProjects = [...projects];

    for (const pid of parentIds) {
      const siblings = updatedProjects.filter((p) => (p.parentId || null) === pid);
      siblings.sort((a, b) => {
        if (sortType === 'name-asc') {
          return a.name.localeCompare(b.name, 'ko', { numeric: true, sensitivity: 'base' });
        }
        if (sortType === 'name-desc') {
          return b.name.localeCompare(a.name, 'ko', { numeric: true, sensitivity: 'base' });
        }
        if (sortType === 'created-desc') {
          return b.createdAt.localeCompare(a.createdAt);
        }
        // created-asc
        return a.createdAt.localeCompare(b.createdAt);
      });

      const nowSortTime = new Date().toISOString();
      siblings.forEach((s, idx) => {
        const item = updatedProjects.find((p) => p.id === s.id);
        if (item) {
          item.order = (idx + 1) * 10;
          item.updatedAt = nowSortTime;
        }
      });
    }

    setProjects(updatedProjects);

    for (const p of updatedProjects) {
      await saveProject(p);
    }

    const sortNames = {
      'name-asc': '가나다순 (A-Z)',
      'name-desc': '가나다 역순 (Z-A)',
      'created-desc': '최신 생성순',
      'created-asc': '과거 생성순',
    };
    setSyncToast(`🔀 폴더가 '${sortNames[sortType]}'으로 정렬되었습니다.`);
    setTimeout(() => setSyncToast(null), 3000);

    crossTabSync.broadcastChange('projects');

    if (gitHubConfig?.isConnected && gitHubConfig.autoSync) {
      setTimeout(() => runSync(true), 1500);
    }
  };

  // 특정 폴더로 파일 드롭 시 새 메모 자동 생성 핸들러
  const handleDropFilesToProject = async (projectId: string, files: FileList) => {
    let targetProjId = projectId;
    if (targetProjId === 'ALL') {
      targetProjId = projects[0]?.id || 'proj-default';
    }

    const currentProj = projects.find((p) => p.id === targetProjId);
    const isLocalOnly = currentProj?.isLocalOnly || false;

    const createdNotes: Note[] = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const processed = await processFileForNote(file);

      const newNote: Note = {
        id: `note-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        projectId: targetProjId,
        title: processed.title,
        content: processed.content,
        attachments: processed.attachments,
        isLocalOnly,
        syncStatus: isLocalOnly ? 'local_only' : 'pending',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };

      await saveNote(newNote);
      createdNotes.push(newNote);
    }

    if (createdNotes.length > 0) {
      setNotes((prev) => [...createdNotes, ...prev]);
      setActiveProjectId(targetProjId);
      setActiveNoteId(createdNotes[0].id);
      setMobileScreen('editor');
      setSyncToast(`${createdNotes.length}개의 파일이 새 메모로 등록되었습니다!`);
      setTimeout(() => setSyncToast(null), 3500);

      if (gitHubConfig?.isConnected && !isLocalOnly && gitHubConfig.autoSync) {
        setTimeout(() => runSync(true), 2000);
      }
    }
  };

  // 폴더 삭제 요청 (팝업 열기)
  const requestDeleteProject = (id: string) => {
    const targetProject = projects.find((p) => p.id === id);
    if (!targetProject) return;

    // 자식 폴더 ID 수집
    const getDescendants = (rootId: string): string[] => {
      const children = projects.filter((p) => p.parentId === rootId);
      let ids = children.map((c) => c.id);
      for (const child of children) {
        ids = ids.concat(getDescendants(child.id));
      }
      return ids;
    };
    const allFamilyIds = [id, ...getDescendants(id)];
    const affectedNoteCount = notes.filter((n) => allFamilyIds.includes(n.projectId)).length;

    setDeleteConfirm({
      isOpen: true,
      type: 'project',
      id,
      title: '폴더 삭제 확인',
      message: `'${targetProject.name}' 폴더를 정말 삭제하시겠습니까?`,
      subMessage: affectedNoteCount > 0 
        ? `포함된 하위 폴더 및 ${affectedNoteCount}개의 메모가 함께 영구 삭제됩니다.` 
        : '비어 있는 폴더입니다.',
    });
  };

  // 메모 삭제 요청 (팝업 열기)
  const requestDeleteNote = (id: string) => {
    const targetNote = notes.find((n) => n.id === id);
    if (!targetNote) return;

    setDeleteConfirm({
      isOpen: true,
      type: 'note',
      id,
      title: '메모 삭제 확인',
      message: `'${targetNote.title || '제목 없는 메모'}'를 정말 삭제하시겠습니까?`,
      subMessage: targetNote.isLocalOnly 
        ? '이 기기에 보관된 메모가 영구 삭제됩니다.'
        : 'GitHub에 동기화된 메모도 함께 삭제됩니다.',
    });
  };

  // 삭제 최종 실행 (ConfirmModal에서 확인 클릭 시)
  const handleConfirmDelete = async () => {
    const { type, id } = deleteConfirm;
    setDeleteConfirm((prev) => ({ ...prev, isOpen: false }));

    if (type === 'note') {
      const targetNote = notes.find((n) => n.id === id);
      if (targetNote && !targetNote.isLocalOnly) {
        const folderPath = getProjectFolderPath(targetNote.projectId, projects);
        const safeTitle = (targetNote.title || '제목없음').replace(/[\/\\:*?"<>|]/g, '_');
        const notesDir = gitHubConfig?.notesDir || 'notes';
        const expectedPath = `${notesDir}/${folderPath}/${safeTitle}_${targetNote.id.slice(-6)}.md`;
        const legacyPath = `${notesDir}/${folderPath}/${safeTitle}.md`;

        await trackDeletedNote(targetNote.id, targetNote.githubPath || expectedPath);
        if (targetNote.githubPath && targetNote.githubPath !== expectedPath) {
          await trackDeletedNote(undefined, expectedPath);
        }
        await trackDeletedNote(undefined, legacyPath);
      }
      await deleteNote(id);
      const remaining = notes.filter((n) => n.id !== id);
      setNotes(remaining);

      // 모바일에서는 메모 삭제 후 즉시 메모 목록('list') 화면으로 복귀
      handleBackToList();
      if (activeNoteId === id) {
        setActiveNoteId(null);
        localStorage.removeItem('mymemo_active_note_id');
      }

      // GitHub 자동 삭제 동기화 예약 (autoSync 설정 시에만)
      if (gitHubConfig?.isConnected && gitHubConfig.autoSync) {
        setTimeout(() => runSync(true), 1500);
      }
      crossTabSync.broadcastChange('notes');
    } else if (type === 'project') {
      const getDescendants = (rootId: string): string[] => {
        const children = projects.filter((p) => p.parentId === rootId);
        let ids = children.map((c) => c.id);
        for (const child of children) {
          ids = ids.concat(getDescendants(child.id));
        }
        return ids;
      };

      const targetIds = [id, ...getDescendants(id)];

      // 속한 메모들의 원격 경로 및 ID 추적 (톰스톤 등록)
      const notesToDelete = notes.filter((n) => targetIds.includes(n.projectId));
      for (const n of notesToDelete) {
        if (!n.isLocalOnly) {
          await trackDeletedNote(n.id, n.githubPath);
        }
      }

      for (const projId of targetIds) {
        const pObj = projects.find((p) => p.id === projId);
        await trackDeletedProject(projId, pObj?.name);
        await deleteProject(projId);
      }

      setProjects((prev) => prev.filter((p) => !targetIds.includes(p.id)));
      setNotes((prev) => prev.filter((n) => !targetIds.includes(n.projectId)));
      if (targetIds.includes(activeProjectId)) {
        handleSelectProject('ALL');
      }

      // GitHub 자동 삭제 동기화 예약 (autoSync 설정 시에만)
      if (gitHubConfig?.isConnected && gitHubConfig.autoSync) {
        setTimeout(() => runSync(true), 1500);
      }
      crossTabSync.broadcastChange('projects');
    }
  };

  // 폴더(프로젝트) 선택 핸들러
  const handleSelectProject = (id: string) => {
    setActiveProjectId(id);
    localStorage.setItem('mymemo_active_project_id', id);
    setIsCalendarOpen(false);
  };

  // 캘린더 메모카드 클릭 시 모달 미리보기 설정 (기본값: true)
  const [calendarModalPreview, setCalendarModalPreview] = useState<boolean>(() => {
    return localStorage.getItem('mymemo_calendar_modal_preview') !== 'false';
  });

  const handleToggleCalendarModalPreview = (enabled: boolean) => {
    setCalendarModalPreview(enabled);
    localStorage.setItem('mymemo_calendar_modal_preview', String(enabled));
  };

  // 캘린더에서 미리보기/편집 팝업 중인 메모 ID
  const [calendarPreviewNoteId, setCalendarPreviewNoteId] = useState<string | null>(null);

  // Esc 키를 누르면 캘린더 미리보기 팝업 창 닫기
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && calendarPreviewNoteId) {
        setCalendarPreviewNoteId(null);
      }
    };
    if (calendarPreviewNoteId) {
      window.addEventListener('keydown', handleKeyDown);
    }
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [calendarPreviewNoteId]);

  // 캘린더 뷰 열기 / 닫기 핸들러
  const handleOpenCalendar = () => {
    setIsCalendarOpen(true);
  };

  const handleCloseCalendar = () => {
    setIsCalendarOpen(false);
    setCalendarPreviewNoteId(null);
  };

  // 캘린더에서 특정 메모를 클릭하여 열기
  const handleSelectNoteFromCalendar = (noteId: string) => {
    setActiveNoteId(noteId);
    if (calendarModalPreview) {
      setCalendarPreviewNoteId(noteId);
    } else {
      setIsCalendarOpen(false);
      handleSelectNote(noteId);
    }
  };

  // 캘린더에서 특정 날짜에 새 메모 작성 (+ 버튼)
  const handleCreateNoteForDate = async (dateStr: string, targetProjectId?: string) => {
    // 1. 선택된 폴더(targetProjectId)가 있으면 해당 폴더에 작성, 없으면 현재 활성 폴더 또는 "일정"/"다이어리" 폴더
    let targetProj: Project | undefined;
    if (targetProjectId && targetProjectId !== 'ALL') {
      targetProj = projects.find((p) => p.id === targetProjectId);
    }
    if (!targetProj && activeProjectId && activeProjectId !== 'ALL') {
      targetProj = projects.find((p) => p.id === activeProjectId);
    }
    if (!targetProj) {
      targetProj = projects.find((p) => p.name === '일정' || p.name === '다이어리');
    }
    if (!targetProj) {
      const newProj: Project = {
        id: `proj-${Date.now()}`,
        name: '일정',
        parentId: null,
        isLocalOnly: false,
        createdAt: new Date().toISOString(),
      };
      await saveProject(newProj);
      targetProj = newProj;
      setProjects((prev) => [newProj, ...prev]);
    }

    const folderPath = getFolderHierarchyPath(targetProj.id, projects);
    const noteTitle = folderPath ? `[${folderPath}] ${dateStr} 업무일지` : `${dateStr} 업무일지`;

    // 캘린더 날짜(dateStr)로 생성 시 createdAt의 날짜도 dateStr 기준(시간은 현재 시각)으로 설정
    const now = new Date();
    const [y, m, d] = dateStr.split('-').map(Number);
    const targetDate = new Date(y, m - 1, d, now.getHours(), now.getMinutes(), now.getSeconds());

    const newNote: Note = {
      id: `note-${Date.now()}`,
      projectId: targetProj.id,
      title: noteTitle,
      content: `# 📋 ${dateStr} (${targetProj.name})\n\n- [ ] `,
      isLocalOnly: targetProj.isLocalOnly,
      syncStatus: targetProj.isLocalOnly ? 'local_only' : 'pending',
      createdAt: !isNaN(targetDate.getTime()) ? targetDate.toISOString() : now.toISOString(),
      updatedAt: now.toISOString(),
    };

    await saveNote(newNote);
    setNotes((prev) => [newNote, ...prev]);
    setIsCalendarOpen(false);
    handleSelectNote(newNote.id);
    crossTabSync.broadcastChange('notes');
  };

  // 캘린더에서 특정 메모를 드래그하여 다른 날짜로 이동하는 핸들러
  const handleMoveNoteToDate = async (noteId: string, targetDateStr: string) => {
    const targetNote = notes.find((n) => n.id === noteId);
    if (!targetNote) return;

    // 1. 제목 내 날짜 패턴이 있으면 변경
    let newTitle = targetNote.title;
    // YYYY-MM-DD or YYYY.MM.DD or YYYY/MM/DD
    const dateMatch = newTitle.match(/(\d{4})([-./])(\d{1,2})\2(\d{1,2})/);
    if (dateMatch) {
      const sep = dateMatch[2];
      const [ty, tm, td] = targetDateStr.split('-');
      const formattedDate = `${ty}${sep}${tm}${sep}${td}`;
      newTitle = newTitle.replace(dateMatch[0], formattedDate);
    } else {
      // 6자리 YYMMDD 패턴 (예: 261102 -> 261029)
      const shortMatch = newTitle.match(/\b(2\d)(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\b/);
      if (shortMatch) {
        const [ty, tm, td] = targetDateStr.split('-');
        const shortDate = `${ty.slice(2)}${tm}${td}`;
        newTitle = newTitle.replace(shortMatch[0], shortDate);
      }
    }

    // 2. createdAt의 날짜 부분을 targetDateStr로 갱신 (시간은 유지)
    const [y, m, d] = targetDateStr.split('-').map(Number);
    const dateObj = new Date(targetNote.createdAt);
    if (!isNaN(dateObj.getTime())) {
      dateObj.setFullYear(y, m - 1, d);
    }
    const newCreatedAt = !isNaN(dateObj.getTime()) ? dateObj.toISOString() : `${targetDateStr}T09:00:00.000Z`;

    // 3. 본문 상단 헤더(# 📋 YYYY-MM-DD 또는 # YYYY-MM-DD 등)가 있다면 새 날짜로 갱신 (없으면 추가)
    let newContent = targetNote.content;
    const headerDateRegex = /(#+\s*(?:📋\s*)?)(\d{4})[-./](\d{1,2})[-./](\d{1,2})/;
    if (headerDateRegex.test(newContent)) {
      newContent = newContent.replace(headerDateRegex, `$1${targetDateStr}`);
    } else {
      newContent = `# 📋 ${targetDateStr}\n\n${newContent}`;
    }

    const updatedNote: Note = {
      ...targetNote,
      title: newTitle,
      content: newContent,
      createdAt: newCreatedAt,
      updatedAt: new Date().toISOString(),
      syncStatus: targetNote.isLocalOnly ? 'local_only' : 'pending',
    };

    await saveNote(updatedNote);
    setNotes((prev) => prev.map((n) => (n.id === noteId ? updatedNote : n)));
    setSyncToast(`📅 '${targetNote.title}' 일정이 ${targetDateStr}로 이동되었습니다.`);
    setTimeout(() => setSyncToast(null), 3000);
    crossTabSync.broadcastChange('notes');
  };

  // 모바일 목록으로 복귀 핸들러
  const handleBackToList = () => {
    setMobileScreen('list');
    localStorage.setItem('mymemo_mobile_screen', 'list');
  };

  // 새 메모 추가 핸들러 (속한 폴더명 + ' - 새 메모' 기본 이름 자동 생성)
  const handleCreateNote = async () => {
    let targetProjId = activeProjectId;
    if (targetProjId === 'ALL') {
      targetProjId = projects[0]?.id || 'proj-default';
    }

    const currentProj = projects.find((p) => p.id === targetProjId);
    const isLocalOnly = currentProj?.isLocalOnly || false;
    const folderPath = targetProjId ? getFolderHierarchyPath(targetProjId, projects) : '';
    const initialTitle = folderPath ? `[${folderPath}] 새 메모` : '새 메모';

    const newNote: Note = {
      id: `note-${Date.now()}`,
      projectId: targetProjId,
      title: initialTitle,
      content: '',
      isLocalOnly,
      syncStatus: isLocalOnly ? 'local_only' : 'pending',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    await saveNote(newNote);
    setNotes((prev) => [newNote, ...prev]);
    handleSelectNote(newNote.id);
    crossTabSync.broadcastChange('notes');
  };

  // 메모 내용 수정 핸들러 (수정 3초 후 자동 동기화 예약)
  const handleUpdateNote = async (updated: Partial<Note>, specificNoteId?: string) => {
    const targetNoteId = specificNoteId || calendarPreviewNoteId || activeNoteId;
    const targetNote = notes.find((n) => n.id === targetNoteId);
    if (!targetNote) return;

    // 소속 폴더 이동 시 알림 토스트 표시 및 활성 폴더 목록 연동
    if (updated.projectId && updated.projectId !== targetNote.projectId) {
      const targetProj = projects.find((p) => p.id === updated.projectId);
      if (targetProj) {
        setSyncToast(`📁 메모가 '${targetProj.name}' 폴더로 이동되었습니다.`);
        setTimeout(() => setSyncToast(null), 3000);
        if (activeProjectId !== 'ALL' && !isCalendarOpen) {
          setActiveProjectId(updated.projectId);
          localStorage.setItem('mymemo_active_project_id', updated.projectId);
        }
      }
    }

    const isLocalOnlyNow = updated.isLocalOnly !== undefined ? updated.isLocalOnly : targetNote.isLocalOnly;

    // 본문에 날짜 헤더(# 📋 YYYY-MM-DD 등)가 입력/수정되면 내부 기준일(createdAt)도 함께 자동 동기화!
    // 이렇게 하면 나중에 사용자가 본문에서 날짜 헤더를 지우더라도, 지정했던 날짜 자리에 그대로 남아있게 됩니다.
    let finalCreatedAt = targetNote.createdAt;
    const contentToCheck = updated.content !== undefined ? updated.content : targetNote.content;
    if (contentToCheck) {
      const headerMatch = contentToCheck.match(/#+\s*.*?\b(\d{4})[-./](\d{1,2})[-./](\d{1,2})\b/);
      if (headerMatch) {
        const [y, m, d] = [Number(headerMatch[1]), Number(headerMatch[2]), Number(headerMatch[3])];
        const dObj = new Date(finalCreatedAt);
        if (!isNaN(dObj.getTime())) {
          dObj.setFullYear(y, m - 1, d);
          finalCreatedAt = dObj.toISOString();
        } else {
          finalCreatedAt = `${headerMatch[1]}-${headerMatch[2].padStart(2, '0')}-${headerMatch[3].padStart(2, '0')}T09:00:00.000Z`;
        }
      }
    }

    const updatedNote: Note = {
      ...targetNote,
      ...updated,
      createdAt: finalCreatedAt,
      updatedAt: new Date().toISOString(),
      syncStatus: isLocalOnlyNow 
        ? 'local_only' 
        : updated.syncStatus || 'pending',
    };

    await saveNote(updatedNote);
    setNotes((prev) => prev.map((n) => (n.id === updatedNote.id ? updatedNote : n)));
    crossTabSync.broadcastChange('notes');

    // autoSync 설정이 true로 켜져 있는 경우에만 3.5초 뒤 자동 업로드
    if (!isLocalOnlyNow && gitHubConfig?.isConnected && gitHubConfig.autoSync) {
      if (autoSyncTimerRef.current) {
        clearTimeout(autoSyncTimerRef.current);
      }
      autoSyncTimerRef.current = setTimeout(() => {
        runSync(true);
      }, 3500);
    }
  };

  // 메모 선택 핸들러
  const handleSelectNote = (id: string) => {
    setActiveNoteId(id);
    localStorage.setItem('mymemo_active_note_id', id);
    setMobileScreen('editor');
    localStorage.setItem('mymemo_mobile_screen', 'editor');
  };

  // 수동 동기화 버튼 클릭
  const handleManualSync = () => {
    runSync(false);
  };

  // 모든 기존 메모의 제목에 현재 소속 폴더 전체 경로 [상위-하위] 일괄 적용
  const handleApplyAllFolderPrefixes = async (): Promise<number> => {
    let updatedCount = 0;
    const nextNotes = [...notes];

    for (let i = 0; i < nextNotes.length; i++) {
      const note = nextNotes[i];
      const folderPath = getFolderHierarchyPath(note.projectId, projects);
      const newTitle = updateTitleWithFolderPath(note.title, folderPath);
      if (newTitle !== note.title) {
        const updated: Note = {
          ...note,
          title: newTitle,
          updatedAt: new Date().toISOString(),
          syncStatus: note.isLocalOnly ? 'local_only' : 'pending',
        };
        await saveNote(updated);
        nextNotes[i] = updated;
        updatedCount++;
      }
    }

    if (updatedCount > 0) {
      setNotes(nextNotes);
      setSyncToast(`총 ${updatedCount}개의 메모에 [폴더경로]가 일괄 적용되었습니다.`);
      setTimeout(() => setSyncToast(null), 3000);
    }
    return updatedCount;
  };

  // 모든 기존 메모의 제목에서 [폴더경로] 대괄호 접두어 일괄 제거
  const handleRemoveAllFolderPrefixes = async (): Promise<number> => {
    let updatedCount = 0;
    const nextNotes = [...notes];

    for (let i = 0; i < nextNotes.length; i++) {
      const note = nextNotes[i];
      const pureTitle = stripFolderPrefixFromTitle(note.title);
      if (pureTitle !== note.title) {
        const updated: Note = {
          ...note,
          title: pureTitle,
          updatedAt: new Date().toISOString(),
          syncStatus: note.isLocalOnly ? 'local_only' : 'pending',
        };
        await saveNote(updated);
        nextNotes[i] = updated;
        updatedCount++;
      }
    }

    if (updatedCount > 0) {
      setNotes(nextNotes);
      setSyncToast(`총 ${updatedCount}개의 메모에서 [폴더경로]가 제거되었습니다.`);
      setTimeout(() => setSyncToast(null), 3000);
    }
    return updatedCount;
  };

  // 설치 및 GitHub 동기화 가이드 메모 열기 또는 생성
  const handleCreateOrOpenGuideNote = async () => {
    const guideContent = `# MyMemo 마크다운 메모장에 오신 것을 환영합니다! 🎉

휴대폰과 PC 어디서나 빠르고 안전하게 동기화되는 마크다운 메모장입니다.
아래 가이드를 따라 하시면 **3분 안에 앱 설치 및 나만의 안전한 클라우드 동기화**를 마칠 수 있습니다!

---

> ⚠️ **[필독] 데이터 안전 및 손실 방지 안내 (백업의 중요성)**
> 
> MyMemo의 데이터는 기본적으로 브라우저 내부(IndexedDB)에 저장됩니다.
> - **브라우저 청소 시 주의**: 브라우저의 '방문 기록'이나 '캐시'를 지우는 것은 안전하지만, **[쿠키 및 사이트 데이터]를 무심코 전체 삭제하거나 기기를 포맷하면 로컬 메모가 모두 삭제**되어 큰 데이터 손실이 발생할 수 있습니다!
> - **1. 안전한 클라우드 영구 보존**: 아래 가이드를 따라 **GitHub 동기화를 꼭 연결해 두세요!** 동기화를 해두시면 컴퓨터가 고장 나거나 브라우저 데이터를 실수로 다 날려도 1초 만에 100% 되살아납니다.
> - **2. '기기 전용(🔒)' 메모는 로컬 백업 필수**: 자물쇠를 걸어둔 '기기 전용' 메모는 보안상 클라우드로 가지 않으므로, **좌측 하단 ⚙️ [설정] ➔ [로컬 전체 데이터 백업(.json)]** 버튼을 통해 PC의 안전한 폴더(D드라이브, 외장하드, USB)에 파일로 백업해 두세요! 백업 없이 기기를 정리하면 복구할 수 없습니다.

---

## 📱 1. 앱 설치 방법 (홈 화면에 추가)

MyMemo는 앱스토어를 거치지 않고 웹 브라우저에서 바로 단독 앱(PWA)으로 설치할 수 있습니다:

- **스마트폰 (갤럭시 / 안드로이드)**:
  - 크롬 또는 삼성인터넷 우측 상단 **메뉴(점 3개)** ➔ **[홈 화면에 추가]** 또는 **[앱 설치]** 클릭
- **스마트폰 (아이폰 / iOS)**:
  - 사파리(Safari) 브라우저 하단 **공유 버튼(네모+화살표)** ➔ **[홈 화면에 추가]** 클릭
- **PC (크롬 / 엣지 / 웨일)**:
  - 브라우저 주소창 오른쪽 끝에 있는 **[모니터/컴퓨터 모양 설치 아이콘]** ➔ **[설치]** 클릭

> 💡 설치하면 바탕화면에 '엔젤기프트 날개' 아이콘이 생성되어 독립된 전용 창으로 편리하게 실행됩니다.

---

## ☁️ 2. GitHub 3분 동기화 설정 가이드

MyMemo의 모든 메모는 **내 개인 GitHub 저장소(Repository)**에 안전하게 백업 및 동기화됩니다.

### 1단계: GitHub 저장소 만들기
1. [GitHub(github.com)](https://github.com)에 로그인합니다.
2. 우측 상단 **\`+\`** 버튼 클릭 ➔ **[New repository]** 선택
3. **Repository name**에 \`my-notes\` 입력
4. 공개 설정은 **Private**(비공개) 권장
5. ⚠️ **중요 꿀팁**: **[Add a README file]**에 체크(V)해 주세요! (체크를 깜빡 잊고 빈 저장소로 만들어도 앱에서 자동 처리됩니다.)
6. 맨 아래 초록색 **[Create repository]** 클릭!

### 2단계: 개인 접근 토큰(PAT) 발급받기
1. GitHub 우측 상단 내 **프로필 사진** 클릭 ➔ **Settings** 메뉴로 들어갑니다.
2. 왼쪽 메뉴 맨 아래 **Developer settings** 클릭
3. **Personal access tokens** ➔ **Tokens (classic)** 클릭
4. **[Generate new token]** ➔ **[Generate new token (classic)]** 클릭
5. **Note**: \`MyMemo\` 입력
6. **Expiration**: \`No expiration\`(만료 없음) 선택 권장
7. **Select scopes(권한)**: 맨 위에 있는 **\`repo\`** (Full control of private repositories) 체크박스 하나만 체크!
8. 맨 아래 초록색 **[Generate token]** 버튼 클릭!
9. 생성된 녹색 상자의 토큰 문자열(\`ghp_...\`로 시작)을 **[복사]**합니다. *(이 창을 닫으면 다시 볼 수 없으니 꼭 복사하세요!)*

### 3단계: MyMemo 앱에서 연결하기
1. MyMemo 앱 화면 좌측 하단(또는 메뉴)의 **⚙️ [설정]** 버튼을 누릅니다.
2. 정보를 입력합니다:
   - **GitHub 아이디 (Owner)**: 내 GitHub 계정 아이디
   - **저장소 이름 (Repository)**: \`my-notes\`
   - **개인 액세스 토큰 (Token)**: 복사한 \`ghp_...\` 토큰 붙여넣기
3. 하단의 **[연결 테스트]** 버튼을 눌러 초록색 성공 메시지를 확인한 뒤 **[저장 및 완료]**를 누르면 끝!

---

## 🚀 3. MyMemo 핵심 기능 사용 팁

- **달력 일정 드래그 이동**: 달력 뷰에서 등록된 일정을 다른 날짜 칸으로 마우스 드래그하면 날짜가 바로 변경됩니다!
- **글자 색상 & 형광펜**: 텍스트를 드래그하고 상단 툴바의 **팔레트 아이콘(🎨)**을 누르면 7가지 색상과 4가지 형광펜을 적용할 수 있습니다.
- **표(Table) & PDF 첨부**: 툴바의 표 아이콘, PDF 첨부 버튼을 통해 문서와 표를 깔끔하게 관리할 수 있습니다.
- **🔒 기기 전용(로컬) 보관 & 백업**: 메모 상단의 자물쇠 아이콘을 켜면 해당 메모는 GitHub로 전송되지 않고 오직 이 기기에만 안전하게 보관됩니다. 좌측 하단 ⚙️ [설정] 창에서 언제든 원클릭으로 전체 백업 파일(.json)을 내려받거나 복원할 수 있습니다.
`;

    const existing = notes.find(
      (n) => n.id === 'note-welcome' || n.title.includes('설치 & GitHub 동기화') || n.title.includes('환영합니다')
    );

    if (existing) {
      const updatedNote: Note = {
        ...existing,
        title: 'MyMemo 설치 & GitHub 동기화 완벽 가이드 📖',
        content: guideContent,
        updatedAt: new Date().toISOString(),
        syncStatus: existing.isLocalOnly ? 'local_only' : 'pending',
      };
      await saveNote(updatedNote);
      setNotes((prev) => prev.map((n) => (n.id === existing.id ? updatedNote : n)));
      setIsCalendarOpen(false);
      handleSelectNote(existing.id);
      setIsSettingsOpen(false);
      setSyncToast('📖 설치 & GitHub 동기화 가이드가 최신 내용으로 갱신되었습니다.');
      setTimeout(() => setSyncToast(null), 3000);
      return;
    }

    const defaultProj = projects.find((p) => !p.isLocalOnly) || projects[0];
    const guideNote: Note = {
      id: `note-guide-${Date.now()}`,
      projectId: defaultProj ? defaultProj.id : 'proj-default',
      title: 'MyMemo 설치 & GitHub 동기화 완벽 가이드 📖',
      content: guideContent,
      isLocalOnly: false,
      syncStatus: 'pending',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    await saveNote(guideNote);
    setNotes((prev) => [guideNote, ...prev]);
    setIsCalendarOpen(false);
    handleSelectNote(guideNote.id);
    setIsSettingsOpen(false);
    setSyncToast('📖 설치 & GitHub 동기화 가이드 메모가 생성되었습니다.');
    setTimeout(() => setSyncToast(null), 3000);
  };

  // 전역 단축키 핸들러 (Ctrl + S: 언제 어디서나 GitHub 동기화 실행)
  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      const isCtrlOrCmd = e.ctrlKey || e.metaKey;
      if (isCtrlOrCmd && e.key.toLowerCase() === 's') {
        e.preventDefault();
        // 메모 에디터(textarea)에 포커스가 있을 때는 NoteEditor 내부 onKeyDown에서 최신 타이핑 버퍼를 즉시 플러시 후 onSync를 부르므로 중복 실행 방지
        const activeEl = document.activeElement;
        const isEditingNote = activeEl && activeEl.tagName === 'TEXTAREA';
        if (!isEditingNote) {
          handleManualSync();
        }
      }
    };

    window.addEventListener('keydown', handleGlobalKeyDown);
    return () => window.removeEventListener('keydown', handleGlobalKeyDown);
  }, [runSync]);

  // 설정 저장 핸들러
  const handleSaveConfig = async (newConfig: GitHubConfig) => {
    await saveGitHubConfig(newConfig);
    setGitHubConfig(newConfig);
    setSyncToast('GitHub 설정이 저장되었습니다. 첫 동기화를 시작합니다.');
    runSync(false, newConfig);
  };

  // 로컬 전체 백업 파일 복원 핸들러
  const handleRestoreBackup = async (
    backup: MyMemoBackupData,
    mode: 'merge' | 'overwrite'
  ): Promise<{ restoredProjects: number; restoredNotes: number }> => {
    let finalProjects = [...projects];
    let finalNotes = [...notes];

    if (mode === 'overwrite') {
      for (const n of notes) {
        await deleteNote(n.id);
      }
      for (const p of projects) {
        await deleteProject(p.id);
      }
      finalProjects = [];
      finalNotes = [];
    }

    // 프로젝트 복원
    for (const bp of backup.projects) {
      const existingIdx = finalProjects.findIndex((p) => p.id === bp.id);
      if (existingIdx !== -1) {
        finalProjects[existingIdx] = bp;
      } else {
        finalProjects.push(bp);
      }
      await saveProject(bp);
    }

    // 메모 복원
    for (const bn of backup.notes) {
      const existingIdx = finalNotes.findIndex((n) => n.id === bn.id);
      if (existingIdx !== -1) {
        finalNotes[existingIdx] = bn;
      } else {
        finalNotes.push(bn);
      }
      await saveNote(bn);
    }

    setProjects(finalProjects);
    setNotes(finalNotes);
    crossTabSync.broadcastChange('sync_complete');

    setSyncToast(`🎉 백업 복원 완료: 폴더 ${backup.projects.length}개, 메모 ${backup.notes.length}개가 안전하게 복원되었습니다.`);
    setTimeout(() => setSyncToast(null), 4000);

    return {
      restoredProjects: backup.projects.length,
      restoredNotes: backup.notes.length,
    };
  };

  return (
    <div className="flex h-screen w-screen overflow-hidden bg-slate-950 text-slate-100 font-sans">
      {/* 토스트 알림 배너 */}
      {syncToast && (
        <div className="fixed top-4 right-4 z-50 bg-slate-800 border border-indigo-500/50 text-slate-100 text-xs py-2 px-4 rounded-xl shadow-xl animate-in slide-in-from-top duration-200">
          {syncToast}
        </div>
      )}

      {/* 1단: 사이드바 (큰주제 / 프로젝트 목록) */}
      <Sidebar
        projects={projects}
        activeProjectId={activeProjectId}
        onSelectProject={handleSelectProject}
        onCreateProject={handleCreateProject}
        onDeleteProject={requestDeleteProject}
        onRenameProject={handleRenameProject}
        onMoveProject={handleMoveProject}
        onSortProjects={handleSortProjects}
        noteCounts={noteCounts}
        gitHubConfig={gitHubConfig}
        onOpenSettings={() => setIsSettingsOpen(true)}
        onSync={handleManualSync}
        isSyncing={isSyncing}
        isOpenMobile={isSidebarOpenMobile}
        onCloseMobile={() => setIsSidebarOpenMobile(false)}
        canInstall={!!deferredPrompt}
        onInstallApp={handleInstallApp}
        externalAddTarget={externalAddTarget}
        onDropFilesToProject={handleDropFilesToProject}
        isCollapsed={isSidebarCollapsed}
        onToggleCollapse={toggleSidebar}
        pendingCount={pendingCount}
        isCalendarActive={isCalendarOpen}
        onOpenCalendar={handleOpenCalendar}
        otherTabsCount={otherTabsCount}
        currentTheme={currentTheme}
        onCycleTheme={handleCycleTheme}
        width={sidebarWidth}
        isResizing={isResizingSidebar}
      />

      {/* 1단 사이드바 너비 조절 바 (Resizer 1) */}
      {!isSidebarCollapsed && (
        <div
          onMouseDown={handleSidebarResizeStart}
          className={`hidden md:block w-1 z-30 cursor-col-resize select-none shrink-0 transition-colors ${
            isResizingSidebar ? 'bg-indigo-500' : 'bg-slate-800 hover:bg-indigo-500'
          }`}
          title="마우스로 드래그하여 폴더 목록 너비 조절"
        />
      )}

      {/* 2단 & 3단 영역: 캘린더 모드일 때는 대형 캘린더 뷰, 일반 모드일 때는 메모 목록 & 에디터 */}
      {isCalendarOpen ? (
        <CalendarView
          notes={notes}
          projects={projects}
          activeProjectId={activeProjectId}
          onSelectNote={handleSelectNoteFromCalendar}
          onCreateNoteForDate={handleCreateNoteForDate}
          onClose={handleCloseCalendar}
          onMoveNoteDate={handleMoveNoteToDate}
          onOpenSidebar={() => setIsSidebarOpenMobile(true)}
        />
      ) : (
        <>
          {/* 2단: 메모 목록 (소주제 목록) */}
          <div 
            style={typeof window !== 'undefined' && window.innerWidth >= 768 && !isNoteListCollapsed ? { width: `${noteListWidth}px` } : undefined}
            className={`
              ${mobileScreen === 'list' ? 'flex' : 'hidden'} 
              ${isNoteListCollapsed ? 'md:hidden' : 'md:flex'}
              ${isResizingNoteList ? 'select-none transition-none' : ''}
              w-full h-full shrink-0 overflow-hidden
            `}
          >
            <NoteList
              notes={notes}
              projects={projects}
              activeProjectId={activeProjectId}
              activeNoteId={activeNoteId}
              noteCounts={noteCounts}
              onSelectProject={handleSelectProject}
              onSelectNote={handleSelectNote}
              onCreateNote={handleCreateNote}
              onDeleteNote={requestDeleteNote}
              onOpenSidebar={() => setIsSidebarOpenMobile(true)}
              onRequestCreateSubfolder={(parentId) => {
                setExternalAddTarget({ parentId, key: Date.now() });
                setIsSidebarOpenMobile(true);
              }}
              onDropFiles={(files) => handleDropFilesToProject(activeProjectId, files)}
              isSidebarCollapsed={isSidebarCollapsed}
              onToggleSidebar={toggleSidebar}
              onToggleCollapse={toggleNoteList}
              pendingCount={pendingCount}
              onSync={handleManualSync}
              isSyncing={isSyncing}
            />
          </div>

          {/* 2단 메모 목록 너비 조절 바 (Resizer 2) */}
          {!isNoteListCollapsed && (
            <div
              onMouseDown={handleNoteListResizeStart}
              className={`hidden md:block w-1 z-30 cursor-col-resize select-none shrink-0 transition-colors ${
                isResizingNoteList ? 'bg-indigo-500' : 'bg-slate-800 hover:bg-indigo-500'
              }`}
              title="마우스로 드래그하여 메모 목록 너비 조절"
            />
          )}

          {/* 3단: 메모 에디터 & 본문 내용 */}
          <div className={`
            ${mobileScreen === 'editor' ? 'flex' : 'hidden'} md:flex
            flex-1 min-w-0 h-full overflow-hidden
          `}>
            <NoteEditor
              note={activeNote}
              projects={projects}
              onUpdateNote={handleUpdateNote}
              onBackToList={handleBackToList}
              onDeleteNote={requestDeleteNote}
              isSidebarCollapsed={isSidebarCollapsed}
              onToggleSidebar={toggleSidebar}
              isNoteListCollapsed={isNoteListCollapsed}
              onToggleNoteList={toggleNoteList}
              onSync={handleManualSync}
              isSyncing={isSyncing}
            />
          </div>
        </>
      )}

      {/* 설정 모달 */}
      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        config={gitHubConfig}
        onSaveConfig={handleSaveConfig}
        noteCount={notes.length}
        projects={projects}
        notes={notes}
        currentTheme={currentTheme}
        onChangeTheme={setCurrentTheme}
        onRestoreBackup={handleRestoreBackup}
        onApplyFolderPrefixes={handleApplyAllFolderPrefixes}
        onRemoveFolderPrefixes={handleRemoveAllFolderPrefixes}
        onCreateGuideNote={handleCreateOrOpenGuideNote}
        calendarModalPreview={calendarModalPreview}
        onToggleCalendarModalPreview={handleToggleCalendarModalPreview}
      />

      {/* 캘린더 메모 미리보기/편집 전용 팝업 창 (모달) */}
      {calendarPreviewNoteId && (() => {
        const previewNote = notes.find((n) => n.id === calendarPreviewNoteId) || null;
        if (!previewNote) return null;
        return (
          <div 
            className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-4 md:p-6 bg-black/80 backdrop-blur-xs animate-in fade-in duration-150"
            onClick={() => setCalendarPreviewNoteId(null)}
          >
            <div 
              className="w-full h-full max-w-6xl bg-slate-900 border border-slate-700/90 rounded-2xl shadow-2xl flex flex-col overflow-hidden animate-in zoom-in-95 duration-150"
              onClick={(e) => e.stopPropagation()}
            >
              {/* 팝업 상단 헤더 바 */}
              <div className="px-4 py-2.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between shrink-0">
                <div className="flex items-center gap-2">
                  <CalendarIcon size={16} className="text-indigo-400" />
                  <span className="text-xs md:text-sm font-bold text-slate-200">
                    캘린더 메모 미리보기 & 편집
                  </span>
                  <span className="hidden sm:inline-block text-[11px] text-slate-500 font-mono">
                    (단축키 및 실시간 자동저장 지원)
                  </span>
                </div>

                <div className="flex items-center gap-2">
                  {/* 해당 폴더로 직접 이동하기 버튼 */}
                  <button
                    type="button"
                    onClick={() => {
                      const targetId = calendarPreviewNoteId;
                      setCalendarPreviewNoteId(null);
                      setIsCalendarOpen(false);
                      handleSelectNote(targetId);
                    }}
                    className="flex items-center gap-1 px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg text-xs font-medium transition-colors border border-slate-700/70 active:scale-95"
                    title="캘린더를 닫고 이 메모가 속한 폴더로 직접 이동"
                  >
                    <FolderTree size={13} />
                    <span>폴더로 이동</span>
                  </button>

                  {/* 닫기 버튼 */}
                  <button
                    type="button"
                    onClick={() => setCalendarPreviewNoteId(null)}
                    className="flex items-center gap-1 px-2.5 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95"
                    title="닫기 (Esc)"
                  >
                    <X size={14} />
                    <span>닫기</span>
                  </button>
                </div>
              </div>

              {/* 에디터 본체: 이미지와 동일하게 보기 및 편집, 단축키 100% 동작 */}
              <div className="flex-1 min-h-0 overflow-hidden">
                <NoteEditor
                  note={previewNote}
                  projects={projects}
                  onUpdateNote={(updated) => handleUpdateNote(updated, previewNote.id)}
                  onBackToList={() => setCalendarPreviewNoteId(null)}
                  onDeleteNote={(id) => {
                    setCalendarPreviewNoteId(null);
                    requestDeleteNote(id);
                  }}
                  isSidebarCollapsed={true}
                  onToggleSidebar={() => {}}
                  isNoteListCollapsed={true}
                  onToggleNoteList={() => {}}
                  onSync={handleManualSync}
                  isSyncing={isSyncing}
                />
              </div>
            </div>
          </div>
        );
      })()}

      {/* 삭제 확인 전용 모달 */}
      <ConfirmModal
        isOpen={deleteConfirm.isOpen}
        title={deleteConfirm.title}
        message={deleteConfirm.message}
        subMessage={deleteConfirm.subMessage}
        onConfirm={handleConfirmDelete}
        onCancel={() => setDeleteConfirm((prev) => ({ ...prev, isOpen: false }))}
      />
    </div>
  );
}

export default App;
