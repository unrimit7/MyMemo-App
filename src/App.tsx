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
  saveGitHubConfig 
} from './services/db';
import { syncWithGitHub } from './services/github';
import { Sidebar } from './components/Sidebar';
import { NoteList } from './components/NoteList';
import { NoteEditor } from './components/NoteEditor';
import { SettingsModal } from './components/SettingsModal';

export function App() {
  const [projects, setProjects] = useState<Project[]>([]);
  const [notes, setNotes] = useState<Note[]>([]);
  const [activeProjectId, setActiveProjectId] = useState<string>('ALL');
  const [activeNoteId, setActiveNoteId] = useState<string | null>(null);

  const [gitHubConfig, setGitHubConfig] = useState<GitHubConfig | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isSidebarOpenMobile, setIsSidebarOpenMobile] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncToast, setSyncToast] = useState<string | null>(null);

  // 모바일 화면 뷰 모드 ('list' 또는 'editor')
  const [mobileScreen, setMobileScreen] = useState<'list' | 'editor'>('list');
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);

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

      if (loadedNotes.length > 0 && !activeNoteId) {
        setActiveNoteId(loadedNotes[0].id);
      }

      // 설정이 연결되어 있다면 앱 켜질 때 자동 동기화
      if (config && config.isConnected) {
        runSync(true, config);
      }
    };

    initApp();
  }, []);

  // 2. 화면으로 복귀할 때 (모바일에서 화면 켜거나 다른 앱 다녀왔을 때) 자동 동기화
  useEffect(() => {
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible' && gitHubConfig?.isConnected) {
        runSync(true);
      }
    };
    document.addEventListener('visibilitychange', handleVisibilityChange);
    return () => document.removeEventListener('visibilitychange', handleVisibilityChange);
  }, [gitHubConfig, runSync]);

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

  // 프로젝트 추가 핸들러
  const handleCreateProject = async (name: string, isLocalOnly: boolean) => {
    const newProject: Project = {
      id: `proj-${Date.now()}`,
      name,
      isLocalOnly,
      createdAt: new Date().toISOString(),
    };
    await saveProject(newProject);
    setProjects((prev) => [...prev, newProject]);
    setActiveProjectId(newProject.id);
  };

  // 프로젝트 삭제 핸들러
  const handleDeleteProject = async (id: string) => {
    await deleteProject(id);
    setProjects((prev) => prev.filter((p) => p.id !== id));
    setNotes((prev) => prev.filter((n) => n.projectId !== id));
    if (activeProjectId === id) {
      setActiveProjectId('ALL');
    }
  };

  // 새 메모 추가 핸들러
  const handleCreateNote = async () => {
    let targetProjId = activeProjectId;
    if (targetProjId === 'ALL') {
      targetProjId = projects[0]?.id || 'proj-default';
    }

    const currentProj = projects.find((p) => p.id === targetProjId);
    const isLocalOnly = currentProj?.isLocalOnly || false;

    const newNote: Note = {
      id: `note-${Date.now()}`,
      projectId: targetProjId,
      title: '새 메모',
      content: '',
      isLocalOnly,
      syncStatus: isLocalOnly ? 'local_only' : 'pending',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    await saveNote(newNote);
    setNotes((prev) => [newNote, ...prev]);
    setActiveNoteId(newNote.id);
    setMobileScreen('editor');
  };

  // 메모 내용 수정 핸들러 (수정 3초 후 자동 동기화 예약)
  const handleUpdateNote = async (updated: Partial<Note>) => {
    if (!activeNote) return;

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

    // 로컬 전용 메모가 아니고 GitHub 연결 상태라면 3.5초 뒤 자동 동기화
    if (!isLocalOnlyNow && gitHubConfig?.isConnected) {
      if (autoSyncTimerRef.current) {
        clearTimeout(autoSyncTimerRef.current);
      }
      autoSyncTimerRef.current = setTimeout(() => {
        runSync(true);
      }, 3500);
    }
  };

  // 메모 삭제 핸들러
  const handleDeleteNote = async (id: string) => {
    await deleteNote(id);
    const remaining = notes.filter((n) => n.id !== id);
    setNotes(remaining);
    if (activeNoteId === id) {
      setActiveNoteId(remaining[0]?.id || null);
      if (remaining.length === 0) {
        setMobileScreen('list');
      }
    }
  };

  // 메모 선택 핸들러
  const handleSelectNote = (id: string) => {
    setActiveNoteId(id);
    setMobileScreen('editor');
  };

  // 수동 동기화 버튼 클릭
  const handleManualSync = () => {
    runSync(false);
  };

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
        onSelectProject={setActiveProjectId}
        onCreateProject={handleCreateProject}
        onDeleteProject={handleDeleteProject}
        noteCounts={noteCounts}
        gitHubConfig={gitHubConfig}
        onOpenSettings={() => setIsSettingsOpen(true)}
        onSync={handleManualSync}
        isSyncing={isSyncing}
        isOpenMobile={isSidebarOpenMobile}
        onCloseMobile={() => setIsSidebarOpenMobile(false)}
        canInstall={!!deferredPrompt}
        onInstallApp={handleInstallApp}
      />

      {/* 2단: 메모 목록 (소주제 목록) */}
      <div className={`
        ${mobileScreen === 'list' ? 'flex' : 'hidden'} md:flex
        w-full md:w-80 lg:w-96 h-full shrink-0
      `}>
        <NoteList
          notes={notes}
          projects={projects}
          activeProjectId={activeProjectId}
          activeNoteId={activeNoteId}
          onSelectNote={handleSelectNote}
          onCreateNote={handleCreateNote}
          onDeleteNote={handleDeleteNote}
          onOpenSidebar={() => setIsSidebarOpenMobile(true)}
        />
      </div>

      {/* 3단: 메모 에디터 & 본문 내용 */}
      <div className={`
        ${mobileScreen === 'editor' ? 'flex' : 'hidden'} md:flex
        flex-1 h-full overflow-hidden
      `}>
        <NoteEditor
          note={activeNote}
          projects={projects}
          onUpdateNote={handleUpdateNote}
          onBackToList={() => setMobileScreen('list')}
        />
      </div>

      {/* 설정 모달 */}
      <SettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        config={gitHubConfig}
        onSaveConfig={handleSaveConfig}
      />
    </div>
  );
}

export default App;
