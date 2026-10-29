import { openDB, type DBSchema, type IDBPDatabase } from 'idb';
import type { Project, Note, GitHubConfig } from '../types';

interface MyMemoDB extends DBSchema {
  projects: {
    key: string;
    value: Project;
  };
  notes: {
    key: string;
    value: Note;
    indexes: {
      'by-project': string;
      'by-updated': string;
    };
  };
  settings: {
    key: string;
    value: any;
  };
}

const DB_NAME = 'MyMemo_DB';
const DB_VERSION = 1;

let dbPromise: Promise<IDBPDatabase<MyMemoDB>> | null = null;

export async function getDB(): Promise<IDBPDatabase<MyMemoDB>> {
  if (!dbPromise) {
    dbPromise = openDB<MyMemoDB>(DB_NAME, DB_VERSION, {
      upgrade(db) {
        if (!db.objectStoreNames.contains('projects')) {
          db.createObjectStore('projects', { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains('notes')) {
          const noteStore = db.createObjectStore('notes', { keyPath: 'id' });
          noteStore.createIndex('by-project', 'projectId');
          noteStore.createIndex('by-updated', 'updatedAt');
        }
        if (!db.objectStoreNames.contains('settings')) {
          db.createObjectStore('settings');
        }
      },
    });
  }
  return dbPromise;
}

// 초기 기본 데이터 시딩
export async function seedInitialDataIfNeeded(): Promise<void> {
  const db = await getDB();
  const projectCount = await db.count('projects');

  if (projectCount === 0) {
    const defaultProject: Project = {
      id: 'proj-default',
      name: '빠른 메모',
      isLocalOnly: false,
      color: '#6366f1',
      createdAt: new Date().toISOString(),
    };

    const secureProject: Project = {
      id: 'proj-secure',
      name: '보안 메모 (기기 전용)',
      isLocalOnly: true,
      color: '#ef4444',
      createdAt: new Date().toISOString(),
    };

    await db.put('projects', defaultProject);
    await db.put('projects', secureProject);

    const welcomeNote: Note = {
      id: 'note-welcome',
      projectId: 'proj-default',
      title: 'MyMemo에 오신 것을 환영합니다! 🎉',
      isLocalOnly: false,
      syncStatus: 'pending',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      content: `# MyMemo 마크다운 메모장

휴대폰과 PC 어디서나 빠르고 안전하게 동기화되는 마크다운 메모장입니다.

## 🚀 주요 기능
- **오프라인 우선**: 인터넷이 없어도 언제나 작성 및 저장 가능
- **GitHub 동기화**: 마크다운(.md) 파일 형태로 내 레포지토리에 보관
- **🔒 기기 전용(로컬) 보관**: 민감한 정보는 클라우드로 가지 않고 오직 이 기기에만 보관!
- **충돌 방지 분기**: 동시 수정 시 덮어쓰지 않고 안전하게 사본 분기

## ✍️ 체크리스트 예시
- [x] MyMemo 열어보기
- [ ] 우측 상단 ⚙️ 설정에서 GitHub 토큰 연결하기
- [ ] 새 메모 작성해 보기
- [ ] 기기 전용 🔒 토글 켜보기
`,
    };

    const secretNote: Note = {
      id: 'note-secret',
      projectId: 'proj-secure',
      title: '비밀번호 및 보안 정보 예시 🔒',
      isLocalOnly: true,
      syncStatus: 'local_only',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      content: `# 기기 전용 보안 메모

이 메모는 **[🔒 이 기기에만 저장]**이 활성화되어 있어 **GitHub로 절대 전송되지 않습니다.**

- 비밀번호, 개인 식별 번호 등을 안전하게 이 기기 내부에만 보관할 수 있습니다.
- 상단 자물쇠 아이콘이 활성화되어 있는지 확인해 보세요!
`,
    };

    await db.put('notes', welcomeNote);
    await db.put('notes', secretNote);
  }
}

// Project CRUD
export async function getAllProjects(): Promise<Project[]> {
  const db = await getDB();
  return db.getAll('projects');
}

export async function saveProject(project: Project): Promise<void> {
  const db = await getDB();
  await db.put('projects', project);
}

export async function deleteProject(id: string): Promise<void> {
  const db = await getDB();
  // 프로젝트에 속한 메모들도 함께 삭제하거나 정리
  const notes = await db.getAllFromIndex('notes', 'by-project', id);
  for (const note of notes) {
    await db.delete('notes', note.id);
  }
  await db.delete('projects', id);
}

// Note CRUD
export async function getAllNotes(): Promise<Note[]> {
  const db = await getDB();
  return db.getAllFromIndex('notes', 'by-updated');
}

export async function getNotesByProject(projectId: string): Promise<Note[]> {
  const db = await getDB();
  return db.getAllFromIndex('notes', 'by-project', projectId);
}

export async function getNoteById(id: string): Promise<Note | undefined> {
  const db = await getDB();
  return db.get('notes', id);
}

export async function saveNote(note: Note): Promise<void> {
  const db = await getDB();
  await db.put('notes', note);
}

export async function deleteNote(id: string): Promise<void> {
  const db = await getDB();
  await db.delete('notes', id);
}

// Settings CRUD
export async function getGitHubConfig(): Promise<GitHubConfig | null> {
  const db = await getDB();
  return (await db.get('settings', 'github_config')) || null;
}

export async function saveGitHubConfig(config: GitHubConfig): Promise<void> {
  const db = await getDB();
  await db.put('settings', config, 'github_config');
}
