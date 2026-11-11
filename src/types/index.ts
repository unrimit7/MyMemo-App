export interface Project {
  id: string;
  name: string;
  parentId?: string | null; // 상위 폴더 ID (null이면 메인 프로젝트)
  isLocalOnly: boolean; // 기기 전용 여부 (GitHub 동기화에서 배제)
  color?: string;
  createdAt: string;
  updatedAt?: string; // 폴더 이동/수정 일시 (GitHub 동기화 시 덮어쓰기 방지)
  order?: number; // 정렬 순서 인덱스 (드래그 앤 드롭 및 사용자 정의 순서)
}

export type SyncStatus = 'synced' | 'pending' | 'conflict' | 'local_only' | 'syncing';

export interface AttachmentItem {
  id: string; // 첨부파일 고유 키 (예: 'att-12345')
  name: string; // 파일 이름 (예: '문서.pdf', '캡처.png')
  type: string; // MIME 타입 (예: 'application/pdf', 'image/webp')
  data: string; // Base64 data URL
  originalData?: string; // 최초 웹 수정 전 원본 백업 데이터 URL
  originalDate?: string; // 최초 백업 일시 (ISO 문자열)
  editorEngine?: 'exceljs' | 'sheetjs' | 'superdoc' | 'legacy'; // 선호 편집/저장 엔진 (엑셀: exceljs | 워드: superdoc)
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
  syncOnStartup?: boolean; // true면 앱 시작(기동) 시 최신 메모 자동 1회 동기화
  syncOriginalBackup?: boolean; // true면 첨부파일의 원본 백업 버전도 GitHub에 함께 동기화 (기본값: false, 기기 내부 보존)
  lastSyncedAt?: string;
}

export interface SyncSummary {
  uploaded: number;
  downloaded: number;
  conflicts: number;
  skippedLocalOnly: number;
  timestamp: string;
  warnings?: string[];
  blocked?: string[];
}

export type AppTheme = 'navy' | 'black' | 'charcoal' | 'forest' | 'wine' | 'sepia' | 'light';

export interface ThemeOption {
  id: AppTheme;
  name: string;
  desc: string;
  bgHex: string;
  accentHex: string;
  isDark: boolean;
}

export const THEME_OPTIONS: ThemeOption[] = [
  { id: 'navy', name: '미드나잇 네이비', desc: '깊고 세련된 다크 블루 (기본)', bgHex: '#0b1120', accentHex: '#6366f1', isDark: true },
  { id: 'black', name: '딥 블랙 (OLED)', desc: '배터리 절약 & 칠흑의 올블랙', bgHex: '#000000', accentHex: '#38bdf8', isDark: true },
  { id: 'charcoal', name: '모던 차콜', desc: '차분하고 단정한 슬레이트 그레이', bgHex: '#121316', accentHex: '#818cf8', isDark: true },
  { id: 'forest', name: '포레스트 나이트', desc: '눈이 편안한 딥 에메랄드 그린', bgHex: '#06120d', accentHex: '#10b981', isDark: true },
  { id: 'wine', name: '로즈 버건디', desc: '우아하고 포근한 와인 버건디', bgHex: '#12080f', accentHex: '#f43f5e', isDark: true },
  { id: 'sepia', name: '웜 세피아', desc: '아늑한 서재 감성의 다크 브라운', bgHex: '#130f0c', accentHex: '#f59e0b', isDark: true },
  { id: 'light', name: '클래식 라이트', desc: '화사하고 깔끔한 소프트 화이트', bgHex: '#f8fafc', accentHex: '#4f46e5', isDark: false },
];
