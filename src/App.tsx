import { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import type { Project, Note, GitHubConfig } from './types';
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
  trackDeletedNote
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
      isSyncingRef.current = false;
      setIsSyncing(false);
      setTimeout(() => setSyncToast(null), 3500);
    }
  }, [gitHubConfig]);

  // 1. 초기 데이터 로딩 및 앱 켤 때 자동 동기화 1회 실행
  useEffect(() => {
    const initApp = async () => {
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

  // 2. 화면으로 복귀할 때 서비스워커 갱신 확인 및 자동 동기화 (autoSync 옵션 켜진 경우에만)
  useEffect(() => {
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        if ('serviceWorker' in navigator) {
          navigator.serviceWorker.getRegistration().then((reg) => reg?.update());
        }
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
    const newProject: Project = {
      id: `proj-${Date.now()}`,
      name,
      parentId: parentId || null,
      isLocalOnly,
      createdAt: new Date().toISOString(),
    };
    await saveProject(newProject);
    setProjects((prev) => [...prev, newProject]);
    setActiveProjectId(newProject.id);
  };

  // 프로젝트(폴더) 이름 수정 핸들러
  const handleRenameProject = async (id: string, newName: string) => {
    const target = projects.find((p) => p.id === id);
    if (!target || target.name === newName) return;

    const updatedProject: Project = { ...target, name: newName };
    const updatedProjects = projects.map((p) => (p.id === id ? updatedProject : p));
    await saveProject(updatedProject);
    setProjects(updatedProjects);

    // 하위 폴더 및 해당 폴더에 속한 메모들의 제목 접두어([상위-하위]) 자동 갱신
    const affectedNotes = notes.filter((n) => {
      let curr: Project | undefined = projects.find((p) => p.id === n.projectId);
      while (curr) {
        if (curr.id === id) return true;
        if (!curr.parentId) break;
        curr = projects.find((p) => p.id === curr?.parentId);
      }
      return false;
    });

    const updatedNotesMap = new Map<string, Note>();
    for (const note of affectedNotes) {
      const newFolderPath = getFolderHierarchyPath(note.projectId, updatedProjects);
      const newTitle = updateTitleWithFolderPath(note.title, newFolderPath);
      const updatedNote: Note = {
        ...note,
        title: newTitle,
        syncStatus: note.isLocalOnly ? 'local_only' : 'pending',
        updatedAt: new Date().toISOString(),
      };
      await saveNote(updatedNote);
      updatedNotesMap.set(note.id, updatedNote);
    }

    setNotes((prev) =>
      prev.map((n) => updatedNotesMap.get(n.id) || (n.projectId === id && !n.isLocalOnly ? { ...n, syncStatus: 'pending' } : n))
    );

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
    }
  };

  // 폴더(프로젝트) 선택 핸들러
  const handleSelectProject = (id: string) => {
    setActiveProjectId(id);
    localStorage.setItem('mymemo_active_project_id', id);
    setIsCalendarOpen(false);
  };

  // 캘린더 뷰 열기 / 닫기 핸들러
  const handleOpenCalendar = () => {
    setIsCalendarOpen(true);
  };

  const handleCloseCalendar = () => {
    setIsCalendarOpen(false);
  };

  // 캘린더에서 특정 메모를 클릭하여 열기
  const handleSelectNoteFromCalendar = (noteId: string) => {
    setIsCalendarOpen(false);
    handleSelectNote(noteId);
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
    const newNote: Note = {
      id: `note-${Date.now()}`,
      projectId: targetProj.id,
      title: noteTitle,
      content: `# 📋 ${dateStr} (${targetProj.name})\n\n- [ ] `,
      isLocalOnly: targetProj.isLocalOnly,
      syncStatus: targetProj.isLocalOnly ? 'local_only' : 'pending',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    await saveNote(newNote);
    setNotes((prev) => [newNote, ...prev]);
    setIsCalendarOpen(false);
    handleSelectNote(newNote.id);
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
  };

  // 메모 내용 수정 핸들러 (수정 3초 후 자동 동기화 예약)
  const handleUpdateNote = async (updated: Partial<Note>) => {
    if (!activeNote) return;

    // 소속 폴더 이동 시 알림 토스트 표시 및 활성 폴더 목록 연동
    if (updated.projectId && updated.projectId !== activeNote.projectId) {
      const targetProj = projects.find((p) => p.id === updated.projectId);
      if (targetProj) {
        setSyncToast(`📁 메모가 '${targetProj.name}' 폴더로 이동되었습니다.`);
        setTimeout(() => setSyncToast(null), 3000);
        if (activeProjectId !== 'ALL') {
          setActiveProjectId(updated.projectId);
          localStorage.setItem('mymemo_active_project_id', updated.projectId);
        }
      }
    }

    const isLocalOnlyNow = updated.isLocalOnly !== undefined ? updated.isLocalOnly : activeNote.isLocalOnly;

    const updatedNote: Note = {
      ...activeNote,
      ...updated,
      updatedAt: new Date().toISOString(),
      syncStatus: isLocalOnlyNow 
        ? 'local_only' 
        : updated.syncStatus || 'pending',
    };

    await saveNote(updatedNote);
    setNotes((prev) => prev.map((n) => (n.id === updatedNote.id ? updatedNote : n)));

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
        onApplyFolderPrefixes={handleApplyAllFolderPrefixes}
        onRemoveFolderPrefixes={handleRemoveAllFolderPrefixes}
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
    </div>
  );
}

export default App;
