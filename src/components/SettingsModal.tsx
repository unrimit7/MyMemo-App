import React, { useState, useEffect } from 'react';
import type { GitHubConfig } from '../types';
import { testGitHubConnection } from '../services/github';
import { 
  X, 
  Database, 
  ShieldCheck, 
  ExternalLink,
  CheckCircle,
  AlertCircle,
  Loader2
} from 'lucide-react';

interface SettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  config: GitHubConfig | null;
  onSaveConfig: (config: GitHubConfig) => void;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({
  isOpen,
  onClose,
  config,
  onSaveConfig,
}) => {
  const [owner, setOwner] = useState('');
  const [repo, setRepo] = useState('');
  const [token, setToken] = useState('');
  const [branch, setBranch] = useState('main');
  const [notesDir, setNotesDir] = useState('notes');

  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);

  useEffect(() => {
    if (config) {
      setOwner(config.owner || '');
      setRepo(config.repo || '');
      setToken(config.token || '');
      setBranch(config.branch || 'main');
      setNotesDir(config.notesDir || 'notes');
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
            <div className="w-8 h-8 rounded-lg bg-slate-800 flex items-center justify-center text-indigo-400">
              <svg className="w-5 h-5 fill-current" viewBox="0 0 24 24">
                <path fillRule="evenodd" clipRule="evenodd" d="M12 2C6.477 2 2 6.484 2 12.017c0 4.425 2.865 8.18 6.839 9.504.5.092.682-.217.682-.483 0-.237-.008-.868-.013-1.703-2.782.605-3.369-1.343-3.369-1.343-.454-1.158-1.11-1.466-1.11-1.466-.908-.62.069-.608.069-.608 1.003.07 1.53 1.032 1.53 1.032.892 1.53 2.341 1.088 2.91.832.092-.647.35-1.088.636-1.338-2.22-.253-4.555-1.113-4.555-4.951 0-1.093.39-1.988 1.029-2.688-.103-.253-.446-1.272.098-2.65 0 0 .84-.27 2.75 1.026A9.564 9.564 0 0112 6.844c.85.004 1.705.115 2.504.337 1.909-1.296 2.747-1.027 2.747-1.027.546 1.379.202 2.398.1 2.651.64.7 1.028 1.595 1.028 2.688 0 3.848-2.339 4.695-4.566 4.943.359.309.678.92.678 1.855 0 1.338-.012 2.419-.012 2.747 0 .268.18.58.688.482A10.019 10.019 0 0022 12.017C22 6.484 17.522 2 12 2z" />
              </svg>
            </div>
            <div>
              <h2 className="font-bold text-slate-100 text-base">GitHub 연동 설정</h2>
              <p className="text-xs text-slate-400">개인 저장소에 안전하게 메모를 동기화합니다.</p>
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
          {/* 보안 알림 배너 */}
          <div className="p-3 bg-emerald-500/10 border border-emerald-500/20 rounded-xl flex items-start gap-2.5 text-emerald-300">
            <ShieldCheck size={18} className="shrink-0 mt-0.5 text-emerald-400" />
            <div className="leading-relaxed">
              <strong>안심 보안 안내:</strong> 입력하신 토큰은 어떤 외부 서버로도 전송되지 않으며, 오직 본인 기기(브라우저 Local)에만 안전하게 저장되어 GitHub와 직접 통신합니다.
            </div>
          </div>

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
