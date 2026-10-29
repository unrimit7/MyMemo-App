import { useState, useEffect, useMemo } from 'react';
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

  // 1. 초기 데이터 로딩
  const loadData = async () => {
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
  };

  useEffect(() => {
    loadData();
  }, []);

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
    // 현재 선택된 프로젝트 확인 (ALL이면 첫 번째 일반 프로젝트 또는 기본)
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

  // 메모 내용 수정 핸들러
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

  // GitHub 전체 동기화 실행
  const handleSync = async () => {
    if (!gitHubConfig || !gitHubConfig.isConnected) {
      setIsSettingsOpen(true);
      return;
    }

    setIsSyncing(true);
    setSyncToast('GitHub와 동기화 중입니다...');

    try {
      const summary = await syncWithGitHub(gitHubConfig);
      // 동기화 후 최신 데이터 리로드
      const updatedNotes = await getAllNotes();
      const updatedProjects = await getAllProjects();
      setNotes(updatedNotes);
      setProjects(updatedProjects);

      // 설정에 마지막 동기화 시간 갱신
      const updatedConfig = { ...gitHubConfig, lastSyncedAt: summary.timestamp };
      await saveGitHubConfig(updatedConfig);
      setGitHubConfig(updatedConfig);

      setSyncToast(
        `동기화 완료! (업로드: ${summary.uploaded}, 다운로드: ${summary.downloaded}, 로컬보호: ${summary.skippedLocalOnly})`
      );
    } catch (err: any) {
      setSyncToast(`동기화 실패: ${err.message || '네트워크 오류'}`);
    } finally {
      setIsSyncing(false);
      setTimeout(() => setSyncToast(null), 4500);
    }
  };

  // 설정 저장 핸들러
  const handleSaveConfig = async (newConfig: GitHubConfig) => {
    await saveGitHubConfig(newConfig);
    setGitHubConfig(newConfig);
    setSyncToast('GitHub 설정이 저장되었습니다.');
    setTimeout(() => setSyncToast(null), 3000);
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
        onSync={handleSync}
        isSyncing={isSyncing}
        isOpenMobile={isSidebarOpenMobile}
        onCloseMobile={() => setIsSidebarOpenMobile(false)}
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
