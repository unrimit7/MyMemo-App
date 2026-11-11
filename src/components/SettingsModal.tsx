import React, { useState, useEffect, useRef } from 'react';
import type { GitHubConfig, Project, Note } from '../types';
import { testGitHubConnection } from '../services/github';
import { exportBackupData, downloadBackupBlob, parseBackupFile, type MyMemoBackupData } from '../utils/backup';
import { 
  X, 
  Database, 
  ShieldCheck, 
  ExternalLink,
  CheckCircle,
  AlertCircle,
  Loader2,
  FolderTree,
  Tag,
  RotateCcw,
  BookOpen,
  Calendar as CalendarIcon,
  HardDrive,
  Download,
  Upload,
  FileArchive,
  Palette,
  Settings,
  User,
  Mail,
  Copy
} from 'lucide-react';
import { type AppTheme, THEME_OPTIONS } from '../types';

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  config: GitHubConfig | null;
  onSaveConfig: (config: GitHubConfig) => void;
  noteCount?: number;
  projects?: Project[];
  notes?: Note[];
  currentTheme?: AppTheme;
  onChangeTheme?: (theme: AppTheme) => void;
  onRestoreBackup?: (backup: MyMemoBackupData, mode: 'merge' | 'overwrite') => Promise<{ restoredProjects: number; restoredNotes: number }>;
  onApplyFolderPrefixes?: () => Promise<number>;
  onRemoveFolderPrefixes?: () => Promise<number>;
  onCreateGuideNote?: () => void;
  calendarModalPreview?: boolean;
  onToggleCalendarModalPreview?: (enabled: boolean) => void;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({
  isOpen,
  onClose,
  config,
  onSaveConfig,
  noteCount = 0,
  projects = [],
  notes = [],
  currentTheme = 'navy',
  onChangeTheme,
  onRestoreBackup,
  onApplyFolderPrefixes,
  onRemoveFolderPrefixes,
  onCreateGuideNote,
  calendarModalPreview = true,
  onToggleCalendarModalPreview,
}) => {
  const [owner, setOwner] = useState('');
  const [repo, setRepo] = useState('');
  const [token, setToken] = useState('');
  const [branch, setBranch] = useState('main');
  const [notesDir, setNotesDir] = useState('notes');
  const [autoSync, setAutoSync] = useState(false);
  const [syncOnStartup, setSyncOnStartup] = useState(false);

  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);

  const [isBatchProcessing, setIsBatchProcessing] = useState(false);
  const [batchResult, setBatchResult] = useState<{ success: boolean; message: string } | null>(null);

  // 로컬 전체 백업 및 복원 상태
  const [includeAttachments, setIncludeAttachments] = useState(true);
  const [backupResult, setBackupResult] = useState<{ success: boolean; message: string } | null>(null);
  const [pendingRestore, setPendingRestore] = useState<MyMemoBackupData | null>(null);
  const [isRestoring, setIsRestoring] = useState(false);
  const restoreFileInputRef = useRef<HTMLInputElement>(null);

  // 전체 백업 생성 및 다운로드
  const handleExportBackup = () => {
    try {
      setBackupResult(null);
      const { filename, blob, sizeKb, attachmentsCount } = exportBackupData(projects, notes, includeAttachments);
      downloadBackupBlob(blob, filename);

      const sizeStr = sizeKb > 1024 ? `${(sizeKb / 1024).toFixed(1)} MB` : `${sizeKb} KB`;
      setBackupResult({
        success: true,
        message: `백업 파일(${filename}, ${sizeStr})이 다운로드되었습니다. (메모 ${notes.length}개, 폴더 ${projects.length}개, 첨부 ${attachmentsCount}개)`,
      });
    } catch (err: any) {
      setBackupResult({
        success: false,
        message: `백업 파일 생성 실패: ${err.message || '알 수 없는 오류'}`,
      });
    }
  };

  // 백업 파일 선택 시 검증
  const handleFileSelected = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const parsed = await parseBackupFile(file);
      setPendingRestore(parsed);
      setBackupResult(null);
    } catch (err: any) {
      setBackupResult({
        success: false,
        message: err.message || '백업 파일을 읽을 수 없습니다.',
      });
    } finally {
      if (restoreFileInputRef.current) restoreFileInputRef.current.value = '';
    }
  };

  // 백업 복원 실행
  const handleExecuteRestore = async (mode: 'merge' | 'overwrite') => {
    if (!pendingRestore || !onRestoreBackup) return;

    setIsRestoring(true);
    try {
      const res = await onRestoreBackup(pendingRestore, mode);
      setBackupResult({
        success: true,
        message: `복원 성공! (폴더 ${res.restoredProjects}개, 메모 ${res.restoredNotes}개 반영 완료)`,
      });
      setPendingRestore(null);
    } catch (err: any) {
      setBackupResult({
        success: false,
        message: `복원 중 오류 발생: ${err.message || '오류'}`,
      });
    } finally {
      setIsRestoring(false);
    }
  };

  const handleApplyPrefixes = async () => {
    if (!onApplyFolderPrefixes) return;
    if (!window.confirm('모든 기존 메모의 제목에 현재 소속 폴더의 전체 경로 [상위-하위]를 일괄 적용하시겠습니까?')) return;
    setIsBatchProcessing(true);
    setBatchResult(null);
    try {
      const count = await onApplyFolderPrefixes();
      setBatchResult({ success: true, message: `총 ${count}개의 메모 제목에 [폴더경로]가 일괄 적용되었습니다.` });
    } catch {
      setBatchResult({ success: false, message: '일괄 적용 중 오류가 발생했습니다.' });
    } finally {
      setIsBatchProcessing(false);
    }
  };

  const handleRemovePrefixes = async () => {
    if (!onRemoveFolderPrefixes) return;
    if (!window.confirm('모든 메모의 제목에서 [폴더경로] 대괄호 접두어를 제거하고 순수 제목만 남기시겠습니까?')) return;
    setIsBatchProcessing(true);
    setBatchResult(null);
    try {
      const count = await onRemoveFolderPrefixes();
      setBatchResult({ success: true, message: `총 ${count}개의 메모 제목에서 [폴더경로]를 제거했습니다.` });
    } catch {
      setBatchResult({ success: false, message: '일괄 제거 중 오류가 발생했습니다.' });
    } finally {
      setIsBatchProcessing(false);
    }
  };

  useEffect(() => {
    if (config) {
      setOwner(config.owner || '');
      setRepo(config.repo || '');
      setToken(config.token || '');
      setBranch(config.branch || 'main');
      setNotesDir(config.notesDir || 'notes');
      setAutoSync(config.autoSync ?? false);
      setSyncOnStartup(config.syncOnStartup ?? false);
    }
  }, [config, isOpen]);

  if (!isOpen) return null;

  const handleTestConnection = async () => {
    if (!owner.trim() || !repo.trim() || !token.trim()) {
      setTestResult({ success: false, message: '아이디, 저장소 이름, 토큰을 모두 입력해 주세요.' });
      return;
    }

    setTesting(true);
    setTestResult(null);

    const result = await testGitHubConnection({
      owner: owner.trim(),
      repo: repo.trim(),
      token: token.trim(),
      branch: branch.trim() || 'main',
      notesDir: notesDir.trim() || 'notes',
      isConnected: false,
    });

    setTesting(false);
    setTestResult(result);
  };

  const handleSave = () => {
    if (!owner.trim() || !repo.trim() || !token.trim()) {
      alert('아이디, 저장소 이름, 토큰을 모두 입력해 주세요.');
      return;
    }

    onSaveConfig({
      owner: owner.trim(),
      repo: repo.trim(),
      token: token.trim(),
      branch: branch.trim() || 'main',
      notesDir: notesDir.trim() || 'notes',
      isConnected: true,
      autoSync: autoSync,
      syncOnStartup: syncOnStartup,
      lastSyncedAt: config?.lastSyncedAt,
    });

    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm animate-in fade-in duration-150">
      <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-lg shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* 모달 헤더 */}
        <div className="p-4 px-6 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-indigo-500/10 border border-indigo-500/20 flex items-center justify-center text-indigo-400">
              <Settings size={18} />
            </div>
            <div>
              <h2 className="font-bold text-slate-100 text-base">설정</h2>
              <p className="text-xs text-slate-400">앱 환경, 테마 및 GitHub 동기화를 설정합니다.</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800"
          >
            <X size={18} />
          </button>
        </div>

        {/* 모달 본문 (스크롤) */}
        <div className="p-6 overflow-y-auto space-y-4 text-xs">
          {/* 앱 분위기 (테마) 설정 */}
          <div className="p-4 rounded-xl border border-slate-800 bg-slate-950/60 space-y-3">
            <div className="flex items-center justify-between">
              <h4 className="text-xs font-bold text-slate-200 flex items-center gap-2">
                <Palette size={15} className="text-indigo-400" />
                <span>앱 분위기 (테마) 설정</span>
              </h4>
              <span className="text-[11px] text-slate-400">
                스크롤바 & 배경 실시간 적용
              </span>
            </div>

            <p className="text-[11px] text-slate-400 leading-relaxed">
              원하는 분위기를 선택하면 스크롤바와 배경 화면의 분위기가 즉시 조화롭게 변경됩니다.
            </p>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-1">
              {THEME_OPTIONS.map((t) => {
                const isActive = currentTheme === t.id;
                return (
                  <button
                    key={t.id}
                    type="button"
                    onClick={() => onChangeTheme?.(t.id)}
                    className={`flex items-center justify-between p-2.5 rounded-xl border text-left transition-all cursor-pointer ${
                      isActive
                        ? 'border-indigo-500 bg-indigo-600/20 text-white shadow-md ring-1 ring-indigo-500/50'
                        : 'border-slate-800 bg-slate-900/70 hover:border-slate-700 hover:bg-slate-800 text-slate-300'
                    }`}
                  >
                    <div className="flex items-center gap-2.5 min-w-0">
                      {/* 색상 미리보기 원형 인디케이터 */}
                      <div 
                        className="w-5 h-5 rounded-full border border-white/20 shrink-0 flex items-center justify-center shadow-inner"
                        style={{ backgroundColor: t.bgHex }}
                      >
                        <div 
                          className="w-2 h-2 rounded-full"
                          style={{ backgroundColor: t.accentHex }}
                        />
                      </div>
                      <div className="truncate">
                        <div className="font-semibold text-xs leading-tight truncate">
                          {t.name}
                        </div>
                        <div className="text-[10px] text-slate-400 leading-tight truncate">
                          {t.desc}
                        </div>
                      </div>
                    </div>

                    {isActive && (
                      <span className="text-indigo-400 shrink-0 text-xs font-bold pl-1">
                        ✓
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
          </div>

          {/* 제작자 정보 & 피드백 안내 카드 */}
          <div className="p-4 rounded-xl border border-indigo-500/30 bg-gradient-to-br from-indigo-950/30 via-slate-900/80 to-slate-950/60 shadow-sm space-y-2.5">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <div className="w-6 h-6 rounded-lg bg-indigo-500/20 border border-indigo-500/30 flex items-center justify-center text-indigo-400 font-bold text-xs">
                  ✍️
                </div>
                <h4 className="text-xs font-bold text-slate-100 flex items-center gap-1.5">
                  <span>제작자 정보 및 개선 문의</span>
                  <span className="text-[10px] px-1.5 py-0.2 rounded-full bg-indigo-500/10 text-indigo-300 border border-indigo-500/20 font-normal">
                    MyMemo
                  </span>
                </h4>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs pt-0.5">
              <div className="flex items-center gap-2 p-2 rounded-lg bg-slate-900/80 border border-slate-800">
                <User size={13} className="text-indigo-400 shrink-0" />
                <span className="text-slate-400">제작자 :</span>
                <span className="font-semibold text-slate-100">DHJung</span>
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-900/80 border border-slate-800 min-w-0">
                <div className="flex items-center gap-2 min-w-0">
                  <Mail size={13} className="text-indigo-400 shrink-0" />
                  <span className="text-slate-400 shrink-0">E-mail :</span>
                  <a
                    href="mailto:unrimit7@gmail.com?subject=[MyMemo%20%EA%B0%9C%EC%84%A0%20%EB%B0%8F%20%EB%AC%B8%EC%9D%98]"
                    className="font-semibold text-indigo-400 hover:text-indigo-300 truncate underline"
                    title="메일 앱으로 바로 보내기"
                  >
                    unrimit7@gmail.com
                  </a>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    navigator.clipboard.writeText('unrimit7@gmail.com');
                    alert('이메일 주소(unrimit7@gmail.com)가 클립보드에 복사되었습니다.');
                  }}
                  className="p-1 hover:bg-slate-800 text-slate-400 hover:text-slate-200 rounded transition-colors ml-1 shrink-0"
                  title="이메일 주소 복사"
                >
                  <Copy size={12} />
                </button>
              </div>
            </div>

            <p className="text-[11px] text-slate-400 leading-relaxed bg-slate-950/60 p-2.5 rounded-lg border border-slate-800/80 flex items-start gap-1.5">
              <span className="text-indigo-400 shrink-0 mt-0.5">💬</span>
              <span>
                버그나 개선 요청 사항을 메일로 알려 주시면 개선에 참조하도록 하겠습니다. 소중한 의견 언제든 환영합니다!
              </span>
            </p>
          </div>

          {/* GitHub 클라우드 동기화 섹션 헤더 */}
          <div className="pt-1">
            <div className="flex items-center gap-2 pb-2 border-b border-slate-800/80">
              <div className="w-5 h-5 rounded-md bg-slate-800 flex items-center justify-center text-slate-200">
                <svg className="w-3.5 h-3.5 fill-current" viewBox="0 0 24 24">
                  <path fillRule="evenodd" clipRule="evenodd" d="M12 2C6.477 2 2 6.484 2 12.017c0 4.425 2.865 8.18 6.839 9.504.5.092.682-.217.682-.483 0-.237-.008-.868-.013-1.703-2.782.605-3.369-1.343-3.369-1.343-.454-1.158-1.11-1.466-1.11-1.466-.908-.62.069-.608.069-.608 1.003.07 1.53 1.032 1.53 1.032.892 1.53 2.341 1.088 2.91.832.092-.647.35-1.088.636-1.338-2.22-.253-4.555-1.113-4.555-4.951 0-1.093.39-1.988 1.029-2.688-.103-.253-.446-1.272.098-2.65 0 0 .84-.27 2.75 1.026A9.564 9.564 0 0112 6.844c.85.004 1.705.115 2.504.337 1.909-1.296 2.747-1.027 2.747-1.027.546 1.379.202 2.398.1 2.651.64.7 1.028 1.595 1.028 2.688 0 3.848-2.339 4.695-4.566 4.943.359.309.678.92.678 1.855 0 1.338-.012 2.419-.012 2.747 0 .268.18.58.688.482A10.019 10.019 0 0022 12.017C22 6.484 17.522 2 12 2z" />
                </svg>
              </div>
              <h4 className="text-xs font-bold text-slate-200">GitHub 클라우드 동기화</h4>
            </div>
          </div>

          {/* 보안 알림 배너 */}
          <div className="p-3 bg-emerald-500/10 border border-emerald-500/20 rounded-xl flex items-start gap-2.5 text-emerald-300">
            <ShieldCheck size={18} className="shrink-0 mt-0.5 text-emerald-400" />
            <div className="leading-relaxed">
              <strong>안심 보안 안내:</strong> 입력하신 토큰은 어떤 외부 서버로도 전송되지 않으며, 오직 본인 기기(브라우저 Local)에만 안전하게 저장되어 GitHub와 직접 통신합니다.
            </div>
          </div>

          {/* 설치 & 동기화 가이드 열기 배너 */}
          {onCreateGuideNote && (
            <div className="p-3 bg-indigo-950/40 border border-indigo-500/30 rounded-xl flex items-center justify-between gap-3 text-indigo-300">
              <div className="flex items-center gap-2 min-w-0">
                <BookOpen size={16} className="text-indigo-400 shrink-0" />
                <span className="text-[11px] text-indigo-200 truncate">처음이라 설정이나 기능 사용법이 궁금하신가요?</span>
              </div>
              <button
                type="button"
                onClick={onCreateGuideNote}
                className="px-2.5 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shrink-0 transition-colors shadow-sm"
              >
                가이드 메모 보기
              </button>
            </div>
          )}

          {/* GitHub 사용자명 */}
          <div>
            <label className="block text-slate-300 font-semibold mb-1">
              GitHub 아이디 (Owner) <span className="text-rose-400">*</span>
            </label>
            <input
              type="text"
              placeholder="예: octocat"
              value={owner}
              onChange={(e) => setOwner(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-slate-100 placeholder-slate-600 focus:outline-none focus:border-indigo-500 text-sm"
            />
          </div>

          {/* 저장소 이름 */}
          <div>
            <label className="block text-slate-300 font-semibold mb-1">
              저장소 이름 (Repository) <span className="text-rose-400">*</span>
            </label>
            <input
              type="text"
              placeholder="예: my-notes"
              value={repo}
              onChange={(e) => setRepo(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-slate-100 placeholder-slate-600 focus:outline-none focus:border-indigo-500 text-sm"
            />
          </div>

          {/* 토큰 (PAT) */}
          <div>
            <div className="flex items-center justify-between mb-1">
              <label className="text-slate-300 font-semibold">
                개인 접근 토큰 (PAT) <span className="text-rose-400">*</span>
              </label>
              <a
                href="https://github.com/settings/tokens"
                target="_blank"
                rel="noreferrer"
                className="text-indigo-400 hover:text-indigo-300 flex items-center gap-1 font-normal"
              >
                <span>토큰 발급 바로가기</span>
                <ExternalLink size={11} />
              </a>
            </div>
            <input
              type="password"
              placeholder="ghp_xxxxxxxxxxxx..."
              value={token}
              onChange={(e) => setToken(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-slate-100 placeholder-slate-600 focus:outline-none focus:border-indigo-500 text-sm font-mono"
            />
            <p className="text-slate-500 mt-1 text-[11px]">
              * 권한 설정: Classic 토큰 기준 <strong>'repo'</strong> 체크박스만 선택하시면 됩니다.
            </p>
          </div>

          {/* 추가 설정 (접이식 형태 또는 기본) */}
          <div className="grid grid-cols-2 gap-3 pt-1">
            <div>
              <label className="block text-slate-400 font-medium mb-1">브랜치</label>
              <input
                type="text"
                value={branch}
                onChange={(e) => setBranch(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-slate-200 text-xs focus:outline-none focus:border-indigo-500"
              />
            </div>
            <div>
              <label className="block text-slate-400 font-medium mb-1">저장 폴더 경로</label>
              <input
                type="text"
                value={notesDir}
                onChange={(e) => setNotesDir(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-slate-200 text-xs focus:outline-none focus:border-indigo-500"
              />
            </div>
          </div>

          {/* 동기화 동작 모드 설정 (2가지 독립 옵션) */}
          <div className="p-4 rounded-xl border border-slate-800 bg-slate-950/60 space-y-3.5">
            <h4 className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
              <span>동기화 방식 설정 (기기별 독립 적용)</span>
            </h4>

            {/* 1. 앱 기동 시 자동 동기화 */}
            <div className="space-y-1">
              <label className="flex items-center gap-2 text-xs font-bold text-slate-200 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={syncOnStartup}
                  onChange={(e) => setSyncOnStartup(e.target.checked)}
                  className="accent-indigo-500 rounded w-4 h-4 cursor-pointer"
                />
                <span>1. 앱 시작(기동) 시 최신 메모 자동 동기화</span>
              </label>
              <p className="text-[11px] text-slate-400 leading-relaxed pl-6">
                {syncOnStartup ? (
                  <span className="text-emerald-400 font-medium">✅ 앱을 켤 때 다른 기기나 동료가 올린 최신 메모를 자동으로 1회 가져옵니다.</span>
                ) : (
                  <span>앱을 켜도 동기화하지 않고, 오직 수동으로 [동기화] 버튼을 누를 때만 동기화합니다.</span>
                )}
              </p>
            </div>

            <div className="h-[1px] bg-slate-800/80" />

            {/* 2. 메모 작성 중 실시간 자동 동기화 */}
            <div className="space-y-1">
              <label className="flex items-center gap-2 text-xs font-bold text-slate-200 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={autoSync}
                  onChange={(e) => setAutoSync(e.target.checked)}
                  className="accent-indigo-500 rounded w-4 h-4 cursor-pointer"
                />
                <span>2. 메모 작성 중 실시간 자동 동기화</span>
              </label>
              <p className="text-[11px] text-slate-400 leading-relaxed pl-6">
                {autoSync ? (
                  <span className="text-amber-400 font-medium">⚡ 작성 중 3.5초 후 자동으로 GitHub에 커밋 및 업로드됩니다.</span>
                ) : (
                  <span className="text-indigo-300 font-medium">🛡️ <strong>수동 동기화 모드 (여러 사람 협업 시 강력 권장)</strong>: 평소에는 로컬에 저장되고, <strong>[동기화]</strong> 버튼이나 Ctrl+S를 누를 때만 GitHub에 안전하게 반영합니다.</span>
                )}
              </p>
            </div>
          </div>

          {/* 캘린더 동작 설정 */}
          <div className="p-4 rounded-xl border border-slate-800 bg-slate-950/60 space-y-3">
            <h4 className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
              <CalendarIcon size={14} className="text-indigo-400" />
              <span>캘린더 동작 설정</span>
            </h4>

            <div className="space-y-1">
              <label className="flex items-center gap-2 text-xs font-bold text-slate-200 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={calendarModalPreview}
                  onChange={(e) => onToggleCalendarModalPreview?.(e.target.checked)}
                  className="accent-indigo-500 rounded w-4 h-4 cursor-pointer"
                />
                <span>캘린더 메모카드 클릭시 미리보기</span>
              </label>
              <p className="text-[11px] text-slate-400 leading-relaxed pl-6">
                {calendarModalPreview ? (
                  <span className="text-emerald-400 font-medium">
                    ✅ 캘린더에서 메모를 클릭하면 캘린더를 닫지 않고 <strong>새창(팝업)에서 바로 보기 및 편집</strong>합니다. (단축키 지원)
                  </span>
                ) : (
                  <span>
                    체크 해제 시 캘린더가 닫히며 해당 폴더와 메모 편집 화면으로 직접 이동합니다.
                  </span>
                )}
              </p>
            </div>
          </div>

          {/* 메모 제목 및 폴더 전체경로 일괄 관리 */}
          <div className="p-4 rounded-xl border border-slate-800 bg-slate-950/60 space-y-3">
            <div className="flex items-center justify-between">
              <h4 className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
                <FolderTree size={14} className="text-indigo-400" />
                <span>메모 제목의 폴더 전체경로 관리</span>
              </h4>
              <span className="text-[11px] text-slate-500 font-mono">총 {noteCount}개 메모</span>
            </div>

            <p className="text-[11px] text-slate-400 leading-relaxed">
              기존 메모들의 제목에 속한 폴더의 전체 경로(예: <code className="text-indigo-300 bg-indigo-950/60 px-1 py-0.5 rounded border border-indigo-500/20 font-mono">[일정-개인] 결혼식</code>)를 일괄 적용하거나 깔끔하게 정리할 수 있습니다.
            </p>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-1">
              <button
                type="button"
                onClick={handleApplyPrefixes}
                disabled={isBatchProcessing}
                className="flex items-center justify-center gap-1.5 py-2 px-3 bg-indigo-600/20 hover:bg-indigo-600/30 text-indigo-300 border border-indigo-500/30 rounded-xl text-xs font-semibold transition-all active:scale-95 disabled:opacity-50"
              >
                {isBatchProcessing ? <Loader2 size={13} className="animate-spin" /> : <Tag size={13} />}
                <span>모든 메모에 [폴더경로] 일괄 적용</span>
              </button>

              <button
                type="button"
                onClick={handleRemovePrefixes}
                disabled={isBatchProcessing}
                className="flex items-center justify-center gap-1.5 py-2 px-3 bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 rounded-xl text-xs font-semibold transition-all active:scale-95 disabled:opacity-50"
              >
                {isBatchProcessing ? <Loader2 size={13} className="animate-spin" /> : <RotateCcw size={13} />}
                <span>제목에서 [폴더경로] 일괄 제거</span>
              </button>
            </div>

            {batchResult && (
              <div className={`p-2.5 rounded-lg border text-[11px] flex items-center gap-1.5 animate-in fade-in ${
                batchResult.success 
                  ? 'bg-emerald-500/10 border-emerald-500/30 text-emerald-300' 
                  : 'bg-rose-500/10 border-rose-500/30 text-rose-300'
              }`}>
                {batchResult.success ? <CheckCircle size={14} className="shrink-0" /> : <AlertCircle size={14} className="shrink-0" />}
                <span>{batchResult.message}</span>
              </div>
            )}
          </div>

          {/* 로컬 전체 데이터 백업 및 복원 (기기 전용 데이터 포함) */}
          <div className="p-4 rounded-xl border border-slate-800 bg-slate-950/60 space-y-3">
            <div className="flex items-center justify-between">
              <h4 className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
                <HardDrive size={14} className="text-emerald-400" />
                <span>로컬 전체 데이터 백업 및 복원</span>
              </h4>
              <span className="text-[11px] text-emerald-400 font-semibold px-2 py-0.5 rounded-full bg-emerald-950/60 border border-emerald-500/30">
                기기 전용(🔒) 포함
              </span>
            </div>

            <p className="text-[11px] text-slate-400 leading-relaxed">
              GitHub에 업로드되지 않는 <strong>'기기 전용' 메모를 포함한 모든 메모, 폴더 구조, 첨부파일(PDF/이미지)</strong>을 내 PC에 파일 1개(.json)로 안전하게 다운로드하거나 언제든 복원할 수 있습니다.
            </p>

            {/* 백업 옵션 선택 */}
            <div className="flex items-center gap-2 pt-1">
              <label className="flex items-center gap-2 text-xs text-slate-300 cursor-pointer">
                <input
                  type="checkbox"
                  checked={includeAttachments}
                  onChange={(e) => setIncludeAttachments(e.target.checked)}
                  className="accent-indigo-500 rounded w-4 h-4 cursor-pointer"
                />
                <span>이미지 및 PDF 첨부파일 포함</span>
              </label>
              <span className="text-[11px] text-slate-500">
                {includeAttachments ? '(완전 백업)' : '(텍스트만 초경량 백업)'}
              </span>
            </div>

            {/* 백업 및 복원 버튼 */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-1">
              <button
                type="button"
                onClick={handleExportBackup}
                className="flex items-center justify-center gap-1.5 py-2 px-3 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/30 rounded-xl text-xs font-semibold transition-all active:scale-95"
                title="모든 메모와 폴더를 단일 JSON 백업 파일로 내보내기"
              >
                <Download size={13} />
                <span>전체 백업 파일(.json) 다운로드</span>
              </button>

              <button
                type="button"
                onClick={() => restoreFileInputRef.current?.click()}
                className="flex items-center justify-center gap-1.5 py-2 px-3 bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 rounded-xl text-xs font-semibold transition-all active:scale-95"
                title="저장해 둔 백업 파일(.json)을 불러와 복원"
              >
                <Upload size={13} />
                <span>백업 파일로 복원하기...</span>
              </button>

              <input
                ref={restoreFileInputRef}
                type="file"
                accept=".json,application/json"
                className="hidden"
                onChange={handleFileSelected}
              />
            </div>

            {/* 복원 대기 확인 패널 */}
            {pendingRestore && (
              <div className="p-3 rounded-xl border border-indigo-500/40 bg-indigo-950/40 space-y-2.5 animate-in fade-in">
                <div className="flex items-center justify-between text-xs font-bold text-indigo-200">
                  <div className="flex items-center gap-1.5">
                    <FileArchive size={14} className="text-indigo-400" />
                    <span>백업 파일 정보 확인</span>
                  </div>
                  <button
                    type="button"
                    onClick={() => setPendingRestore(null)}
                    className="text-slate-400 hover:text-white text-xs"
                  >
                    취소
                  </button>
                </div>

                <div className="text-[11px] text-slate-300 bg-slate-900/80 p-2 rounded-lg border border-slate-800 space-y-1">
                  <div>📅 백업 일시: {new Date(pendingRestore.createdAt).toLocaleString()}</div>
                  <div className="flex gap-3 text-indigo-300 font-semibold">
                    <span>📁 폴더: {pendingRestore.stats.projectsCount}개</span>
                    <span>📝 메모: {pendingRestore.stats.notesCount}개</span>
                    <span>📎 첨부파일: {pendingRestore.stats.attachmentsCount}개</span>
                  </div>
                </div>

                <div className="flex items-center justify-end gap-2 pt-1">
                  <button
                    type="button"
                    onClick={() => handleExecuteRestore('merge')}
                    disabled={isRestoring}
                    className="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95 disabled:opacity-50"
                    title="기존 메모를 유지하면서 백업본을 합칩니다."
                  >
                    기존 데이터에 병합 복원
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      if (window.confirm('기존 데이터를 모두 삭제하고 백업본으로 완전히 덮어쓰시겠습니까?')) {
                        handleExecuteRestore('overwrite');
                      }
                    }}
                    disabled={isRestoring}
                    className="px-3 py-1.5 bg-rose-600/80 hover:bg-rose-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all active:scale-95 disabled:opacity-50"
                    title="기존 데이터를 지우고 백업본 상태로 복원합니다."
                  >
                    전체 덮어쓰기 복원
                  </button>
                </div>
              </div>
            )}

            {/* 결과 메시지 */}
            {backupResult && (
              <div className={`p-2.5 rounded-lg border text-[11px] flex items-center gap-1.5 animate-in fade-in ${
                backupResult.success 
                  ? 'bg-emerald-500/10 border-emerald-500/30 text-emerald-300' 
                  : 'bg-rose-500/10 border-rose-500/30 text-rose-300'
              }`}>
                {backupResult.success ? <CheckCircle size={14} className="shrink-0" /> : <AlertCircle size={14} className="shrink-0" />}
                <span className="leading-relaxed">{backupResult.message}</span>
              </div>
            )}
          </div>

          {/* 연결 테스트 결과 메시지 */}
          {testResult && (
            <div className={`p-3 rounded-xl border flex items-center gap-2 ${
              testResult.success
                ? 'bg-emerald-500/10 border-emerald-500/30 text-emerald-300'
                : 'bg-rose-500/10 border-rose-500/30 text-rose-300'
            }`}>
              {testResult.success ? <CheckCircle size={16} /> : <AlertCircle size={16} />}
              <span className="font-medium">{testResult.message}</span>
            </div>
          )}
        </div>

        {/* 모달 푸터 */}
        <div className="p-4 px-6 border-t border-slate-800 bg-slate-950/50 flex items-center justify-between">
          <button
            type="button"
            onClick={handleTestConnection}
            disabled={testing}
            className="flex items-center gap-1.5 px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-semibold transition-all border border-slate-700 disabled:opacity-50"
          >
            {testing ? <Loader2 size={13} className="animate-spin" /> : <Database size={13} />}
            <span>연결 테스트</span>
          </button>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 text-slate-400 hover:text-slate-200 text-xs font-semibold"
            >
              닫기
            </button>
            <button
              type="button"
              onClick={handleSave}
              className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-semibold shadow-lg shadow-indigo-600/20 transition-all active:scale-95"
            >
              저장 및 완료
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
