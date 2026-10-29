export interface Project {
  id: string;
  name: string;
  isLocalOnly: boolean; // 기기 전용 여부 (GitHub 동기화에서 배제)
  color?: string;
  createdAt: string;
}

export type SyncStatus = 'synced' | 'pending' | 'conflict' | 'local_only' | 'syncing';

export interface Note {
  id: string;
  projectId: string;
  title: string;
  content: string;
  isLocalOnly: boolean; // 기기 전용 여부 (true면 GitHub 동기화 제외)
  githubSha?: string; // GitHub 원격 파일 SHA
  githubPath?: string; // GitHub 저장소 내 파일 경로
  syncStatus: SyncStatus;
  conflictContent?: string; // 충돌 발생 시 원격 내용 임시 보관
  createdAt: string;
  updatedAt: string;
}

export interface GitHubConfig {
  owner: string;
  repo: string;
  token: string;
  branch: string;
  notesDir: string;
  isConnected: boolean;
  lastSyncedAt?: string;
}

export interface SyncSummary {
  uploaded: number;
  downloaded: number;
  conflicts: number;
  skippedLocalOnly: number;
  timestamp: string;
}
