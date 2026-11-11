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
import { QuickSwitcherModal } from './components/QuickSwitcherModal';
import { FloatingNoteWindow } from './components/FloatingNoteWindow';
import { processFileForNote } from './utils/file';
import { getFolderHierarchyPath, updateTitleWithFolderPath, stripFolderPrefixFromTitle } from './utils/folderPath';
import { Calendar as CalendarIcon, X, FolderTree } from 'lucide-react';
import { crossTabSync } from './services/crossTabSync';
import type { MyMemoBackupData } from './utils/backup';
import { getGuideNotes } from './utils/guideNotes';

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
  const [noteSearchQuery, setNoteSearchQuery] = useState<string>('');
  const [isQuickSwitcherOpen, setIsQuickSwitcherOpen] = useState<boolean>(false);
  const [floatingNotes, setFloatingNotes] = useState<{ id: string; isPopout: boolean }[]>([]);

  const handleOpenFloatingNote = (id: string, isPopout: boolean = false) => {
    setFloatingNotes((prev) => {
      const existing = prev.find((n) => n.id === id);
      if (existing) {
        return prev.map((n) => (n.id === id ? { ...n, isPopout } : n));
      }
      return [...prev, { id, isPopout }];
    });
  };

  const handleCloseFloatingNote = (id: string) => {
    setFloatingNotes((prev) => prev.filter((n) => n.id !== id));
  };

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
      try {
        // 브라우저 및 윈도우 스토리지 센스로부터 로컬 데이터 영구 보존 요청
        requestPersistentStorage().catch(() => {});

        await seedInitialDataIfNeeded();
        const loadedProjects = await getAllProjects();
        const loadedNotes = await getAllNotes();
        const config = await getGitHubConfig();

        // 📖 사용가이드 관련 메모가 엉뚱한 폴더에 있으면 '사용가이드' 폴더로 자동 복귀 및 정리
        const guideFolder = loadedProjects.find((p) => p.name.trim() === '사용가이드');
        if (guideFolder) {
          let hasMovedGuide = false;
          for (let i = 0; i < loadedNotes.length; i++) {
            const n = loadedNotes[i];
            const isGuideNote =
              n.id === 'note-welcome' ||
              n.title.includes('설치 & GitHub 동기화') ||
              n.title.includes('MyMemo 설치') ||
              n.title.startsWith('[사용가이드]');

            if (isGuideNote && n.projectId !== guideFolder.id) {
              const moved: Note = {
                ...n,
                projectId: guideFolder.id,
                updatedAt: new Date().toISOString(),
              };
              await saveNote(moved);
              loadedNotes[i] = moved;
              hasMovedGuide = true;
            }
          }
          if (hasMovedGuide) {
            console.log('[MyMemo] 엉뚱한 폴더에 위치했던 가이드 메모를 사용가이드 폴더로 자동 복귀시켰습니다.');
          }
        }

        setProjects(loadedProjects);
        setNotes(loadedNotes);
        setGitHubConfig(config);

        // 마지막으로 선택했던 폴더 및 메모 복원
        let savedProjId: string | null = null;
        let savedNoteId: string | null = null;
        try {
          savedProjId = localStorage.getItem('mymemo_active_project_id');
          savedNoteId = localStorage.getItem('mymemo_active_note_id');
        } catch {}

        if (savedProjId && (savedProjId === 'ALL' || loadedProjects.some((p) => p.id === savedProjId))) {
          setActiveProjectId(savedProjId);
          if (savedProjId !== 'ALL') {
            try { localStorage.setItem('mymemo_last_folder_id', savedProjId); } catch {}
          }
        } else {
          setActiveProjectId('ALL');
          try { localStorage.setItem('mymemo_active_project_id', 'ALL'); } catch {}
        }

        const currentFolderId = (savedProjId && savedProjId !== 'ALL' && loadedProjects.some((p) => p.id === savedProjId))
          ? savedProjId
          : 'ALL';

        const folderNotes = loadedNotes
          .filter((n) => currentFolderId === 'ALL' || n.projectId === currentFolderId)
          .sort((a, b) => new Date(b.updatedAt || b.createdAt).getTime() - new Date(a.updatedAt || a.createdAt).getTime());

        if (savedNoteId && folderNotes.some((n) => n.id === savedNoteId)) {
          setActiveNoteId(savedNoteId);
        } else if (folderNotes.length > 0) {
          setActiveNoteId(folderNotes[0].id);
          try { localStorage.setItem('mymemo_active_note_id', folderNotes[0].id); } catch {}
        } else {
          setActiveNoteId(null);
          try { localStorage.removeItem('mymemo_active_note_id'); } catch {}
        }

        // 설정이 연결되어 있고 '앱 시작 시 자동 동기화(syncOnStartup)' 또는 '실시간 자동 동기화(autoSync)' 옵션이 켜져 있을 때만 앱 시작 시 동기화
        if (config && config.isConnected && (config.syncOnStartup || config.autoSync)) {
          runSync(true, config);
        }
      } catch (err) {
        console.warn('MyMemo 초기 데이터 로드 중 경고(인앱 브라우저 등):', err);
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
    try { localStorage.setItem('mymemo_last_folder_id', newProject.id); } catch {}
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
    try {
      localStorage.setItem('mymemo_active_project_id', id);
      if (id !== 'ALL') {
        localStorage.setItem('mymemo_last_folder_id', id);
      }
    } catch {}
    setIsCalendarOpen(false);

    // 해당 폴더의 메모 목록 조회 (최신순 정렬)
    const folderNotes = notes
      .filter((n) => (id === 'ALL' ? true : n.projectId === id))
      .sort((a, b) => new Date(b.updatedAt || b.createdAt).getTime() - new Date(a.updatedAt || a.createdAt).getTime());

    if (folderNotes.length > 0) {
      const latestNoteId = folderNotes[0].id;
      setActiveNoteId(latestNoteId);
      try {
        localStorage.setItem('mymemo_active_note_id', latestNoteId);
      } catch {}
    } else {
      // 해당 폴더에 메모가 없으면 에디터의 이전 메모 표시를 비움
      setActiveNoteId(null);
      try {
        localStorage.removeItem('mymemo_active_note_id');
      } catch {}
    }
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

  // 위키링크([[문서명]]) 클릭 시 특정 제목으로 새 메모 바로 생성
  const handleCreateNoteWithTitle = async (title: string) => {
    let targetProjId = activeProjectId;
    if (targetProjId === 'ALL') {
      targetProjId = projects[0]?.id || 'proj-default';
    }
    const currentProj = projects.find((p) => p.id === targetProjId);
    const isLocalOnly = currentProj?.isLocalOnly || false;

    const newNote: Note = {
      id: `note-${Date.now()}`,
      projectId: targetProjId,
      title,
      content: `# ${title}\n\n`,
      isLocalOnly,
      syncStatus: isLocalOnly ? 'local_only' : 'pending',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    await saveNote(newNote);
    setNotes((prev) => [newNote, ...prev]);
    handleSelectNote(newNote.id);
    crossTabSync.broadcastChange('notes');
    return newNote;
  };

  // 본문 내 #태그 클릭 시 메모 목록에서 해당 태그 즉시 검색 필터링
  const handleSearchTag = (tag: string) => {
    setNoteSearchQuery(tag);
    setIsNoteListCollapsed(false);
    setMobileScreen('list');
    localStorage.setItem('mymemo_mobile_screen', 'list');
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
    try {
      localStorage.setItem('mymemo_active_note_id', id);
      const target = notes.find((n) => n.id === id);
      if (target && target.projectId) {
        localStorage.setItem('mymemo_last_folder_id', target.projectId);
      }
    } catch {}
    setMobileScreen('editor');
    try { localStorage.setItem('mymemo_mobile_screen', 'editor'); } catch {}
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

  // 사용가이드 폴더 및 최신 가이드 메모 일괄 생성/갱신
  const handleCreateOrOpenGuideNote = async () => {
    // 1. '사용가이드' 폴더가 있는지 확인, 없으면 자동 생성
    let guideFolder = projects.find((p) => p.name.trim() === '사용가이드');
    let nextProjects = [...projects];

    if (!guideFolder) {
      guideFolder = {
        id: `proj-guide-${Date.now()}`,
        name: '사용가이드',
        isLocalOnly: false,
        color: '#6366f1',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        order: 9999,
      };
      await saveProject(guideFolder);
      nextProjects = [...projects, guideFolder];
      setProjects(nextProjects);
    }

    // 2. 가이드 메모 목록 가져오기
    const guideDefs = getGuideNotes();
    const updatedNotesList = [...notes];
    let firstGuideId: string | null = null;

    for (let i = 0; i < guideDefs.length; i++) {
      const def = guideDefs[i];
      // 이미 같은 제목이나 관련 가이드 메모가 있는지 탐색
      const existing = updatedNotesList.find(
        (n) =>
          n.title.trim() === def.title.trim() ||
          (i === 0 && n.title.includes('MyMemo 시작하기')) ||
          (i === 1 && (n.id === 'note-welcome' || n.title.includes('설치 & GitHub 동기화') || n.title.includes('앱 설치(PWA) & GitHub')))
      );

      if (existing) {
        const updated: Note = {
          ...existing,
          projectId: guideFolder.id, // 항상 사용가이드 폴더에 배치
          title: def.title,
          content: def.content,
          updatedAt: new Date().toISOString(),
        };
        await saveNote(updated);
        const idx = updatedNotesList.findIndex((n) => n.id === existing.id);
        if (idx !== -1) updatedNotesList[idx] = updated;
        if (i === 0) firstGuideId = existing.id;
      } else {
        const newNote: Note = {
          id: `note-guide-${Date.now()}-${i}`,
          projectId: guideFolder.id,
          title: def.title,
          content: def.content,
          isLocalOnly: false,
          syncStatus: 'pending',
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        };
        await saveNote(newNote);
        updatedNotesList.unshift(newNote);
        if (i === 0) firstGuideId = newNote.id;
      }
    }

    // 3. 다른 폴더에 남아있는 기존 가이드/환영 메모가 있다면 사용가이드 폴더로 모두 복귀
    for (let j = 0; j < updatedNotesList.length; j++) {
      const n = updatedNotesList[j];
      const isGuide =
        n.id === 'note-welcome' ||
        n.title.includes('설치 & GitHub 동기화') ||
        n.title.includes('MyMemo 설치') ||
        n.title.startsWith('[사용가이드]');

      if (isGuide && n.projectId !== guideFolder.id) {
        const moved: Note = {
          ...n,
          projectId: guideFolder.id,
          updatedAt: new Date().toISOString(),
        };
        await saveNote(moved);
        updatedNotesList[j] = moved;
      }
    }

    setNotes(updatedNotesList);
    setActiveProjectId(guideFolder.id);
    localStorage.setItem('mymemo_active_project_id', guideFolder.id);

    if (firstGuideId) {
      handleSelectNote(firstGuideId);
    }

    setIsCalendarOpen(false);
    setIsSettingsOpen(false);
    crossTabSync.broadcastChange('notes');
    crossTabSync.broadcastChange('projects');

    setSyncToast('📖 [사용가이드] 폴더에 최신 사용 설명서가 준비되었습니다.');
    setTimeout(() => setSyncToast(null), 3500);
  };

  // 전역 단축키 핸들러 (Ctrl + S: 동기화, Ctrl + K / Ctrl + P: 퀵 스위처)
  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      const isCtrlOrCmd = e.ctrlKey || e.metaKey;

      // Ctrl + P / Cmd + P: 초고속 퀵 스위처 열기
      if (isCtrlOrCmd && e.key.toLowerCase() === 'p') {
        e.preventDefault();
        setIsQuickSwitcherOpen(true);
        return;
      }

      // Ctrl + K / Cmd + K: 텍스트 미선택 상태 시 퀵 스위처 열기
      if (isCtrlOrCmd && e.key.toLowerCase() === 'k') {
        const activeEl = document.activeElement;
        const isEditingNote = activeEl && activeEl.tagName === 'TEXTAREA';
        if (!isEditingNote) {
          e.preventDefault();
          setIsQuickSwitcherOpen(true);
          return;
        } else {
          const ta = activeEl as HTMLTextAreaElement;
          if (ta.selectionStart === ta.selectionEnd) {
            e.preventDefault();
            setIsQuickSwitcherOpen(true);
            return;
          }
        }
      }

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
        onOpenGuide={handleCreateOrOpenGuideNote}
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
              searchQuery={noteSearchQuery}
              onSearchQueryChange={setNoteSearchQuery}
              onOpenQuickSwitcher={() => setIsQuickSwitcherOpen(true)}
              onSelectProject={handleSelectProject}
              onSelectNote={handleSelectNote}
              onCreateNote={handleCreateNote}
              onDeleteNote={requestDeleteNote}
              onOpenFloatingNote={handleOpenFloatingNote}
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
              allNotes={notes}
              onSelectNote={handleSelectNote}
              onCreateNote={handleCreateNote}
              onCreateNoteWithTitle={handleCreateNoteWithTitle}
              onSearchTag={handleSearchTag}
              onOpenQuickSwitcher={() => setIsQuickSwitcherOpen(true)}
              onOpenFloatingNote={handleOpenFloatingNote}
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
                  allNotes={notes}
                  onSelectNote={(id) => {
                    setCalendarPreviewNoteId(null);
                    handleSelectNote(id);
                  }}
                  onCreateNoteWithTitle={handleCreateNoteWithTitle}
                  onSearchTag={(tag) => {
                    setCalendarPreviewNoteId(null);
                    handleSearchTag(tag);
                  }}
                  onOpenQuickSwitcher={() => setIsQuickSwitcherOpen(true)}
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

      {/* 초고속 퀵 스위처 모달 (Ctrl + K / Ctrl + P) */}
      <QuickSwitcherModal
        isOpen={isQuickSwitcherOpen}
        onClose={() => setIsQuickSwitcherOpen(false)}
        notes={notes}
        projects={projects}
        onSelectNote={handleSelectNote}
        onCreateNoteWithTitle={handleCreateNoteWithTitle}
      />

      {/* 삭제 확인 전용 모달 */}
      <ConfirmModal
        isOpen={deleteConfirm.isOpen}
        title={deleteConfirm.title}
        message={deleteConfirm.message}
        subMessage={deleteConfirm.subMessage}
        onConfirm={handleConfirmDelete}
        onCancel={() => setDeleteConfirm((prev) => ({ ...prev, isOpen: false }))}
      />

      {/* 플로팅 독립 메모 창 (다중 메모 분리/비교/동시 편집 및 OS 새 창 팝업) */}
      {floatingNotes.map((item, index) => {
        const fNote = notes.find((n) => n.id === item.id);
        if (!fNote) return null;
        return (
          <FloatingNoteWindow
            key={fNote.id}
            note={fNote}
            allNotes={notes}
            projects={projects}
            initialPopout={item.isPopout}
            initialX={typeof window !== 'undefined' ? Math.min(Math.max(20, window.innerWidth - 680), 80 + (index % 6) * 36) : 100}
            initialY={typeof window !== 'undefined' ? Math.min(Math.max(20, window.innerHeight - 560), 60 + (index % 6) * 36) : 80}
            zIndex={70 + index}
            onClose={() => handleCloseFloatingNote(fNote.id)}
            onUpdateNote={handleUpdateNote}
          />
        );
      })}
    </div>
  );
}

export default App;
