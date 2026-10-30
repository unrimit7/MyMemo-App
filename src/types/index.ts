export interface Project {
  id: string;
  name: string;
  parentId?: string | null; // 상위 폴더 ID (null이면 메인 프로젝트)
  isLocalOnly: boolean; // 기기 전용 여부 (GitHub 동기화에서 배제)
  color?: string;
  createdAt: string;
}

export type SyncStatus = 'synced' | 'pending' | 'conflict' | 'local_only' | 'syncing';

export interface AttachmentItem {
  id: string; // 첨부파일 고유 키 (예: 'att-12345')
  name: string; // 파일 이름 (예: '문서.pdf', '캡처.png')
  type: string; // MIME 타입 (예: 'application/pdf', 'image/webp')
  data: string; // Base64 data URL
}

export interface Note {
  id: string;
  projectId: string; // 소속된 폴더(프로젝트) ID
  title: string;
  content: string;
  attachments?: Record<string, AttachmentItem>; // 첨부파일 사전
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
  autoSync?: boolean; // false면 수동 동기화 모드(권장), true면 변경 3.5초 후 자동 업로드
  lastSyncedAt?: string;
}

export interface SyncSummary {
  uploaded: number;
  downloaded: number;
  conflicts: number;
  skippedLocalOnly: number;
  timestamp: string;
}
