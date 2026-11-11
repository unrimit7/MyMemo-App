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

/**
 * 브라우저 및 윈도우 스토리지 센스(디스크 정리)로부터 MyMemo 데이터를 영구 보존(Persistent Storage)하도록 요청합니다.
 */
export async function requestPersistentStorage(): Promise<boolean> {
  if (typeof navigator !== 'undefined' && navigator.storage && navigator.storage.persist) {
    try {
      const isPersisted = await navigator.storage.persisted();
      if (!isPersisted) {
        return await navigator.storage.persist();
      }
      return isPersisted;
    } catch {
      return false;
    }
  }
  return false;
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
      updatedAt: new Date().toISOString(),
    };

    const secureProject: Project = {
      id: 'proj-secure',
      name: '보안 메모 (기기 전용)',
      isLocalOnly: true,
      color: '#ef4444',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    const guideProject: Project = {
      id: 'proj-guide',
      name: '사용가이드',
      isLocalOnly: false,
      color: '#6366f1',
      order: 9999,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    await db.put('projects', defaultProject);
    await db.put('projects', secureProject);
    await db.put('projects', guideProject);

    const welcomeNote: Note = {
      id: 'note-welcome',
      projectId: 'proj-guide',
      title: 'MyMemo 설치 & GitHub 동기화 완벽 가이드 📖',
      isLocalOnly: false,
      syncStatus: 'pending',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      content: `# MyMemo 마크다운 메모장에 오신 것을 환영합니다! 🎉

휴대폰과 PC 어디서나 빠르고 안전하게 동기화되는 마크다운 메모장입니다.
아래 가이드를 따라 하시면 **3분 안에 앱 설치 및 나만의 안전한 클라우드 동기화**를 마칠 수 있습니다!

---

> ⚠️ **[필독] 처음 사용하는 분들을 위한 데이터 안전 & 백업 필수 안내**
> 
> MyMemo는 별도의 회원가입 없이 브라우저/기기 내부 저장소(IndexedDB)에 데이터를 즉시 보관합니다.
> **사용자가 백업을 하지 않고 브라우저 기록(쿠키/사이트 데이터)을 청소하거나 앱을 삭제하면, 작성한 메모와 첨부파일이 모두 사라져 큰 데이터 손실이 발생할 수 있습니다!**
> 
> 🛡️ **소중한 데이터를 안전하게 지키는 2가지 방법:**
> 1. **GitHub 동기화 연결 (가장 권장)**: 아래 가이드를 따라 GitHub를 연결해 두시면, 기기를 교체하거나 브라우저를 초기화해도 언제든 원클릭으로 모든 메모가 100% 복원됩니다.
> 2. **주기적인 로컬 전체 백업 (.json 저장)**: 특히 자물쇠가 달린 **'기기 전용(🔒)'** 메모는 클라우드로 가지 않으므로, 좌측 하단 **⚙️ [설정] ➔ [로컬 전체 데이터 백업(.json)]** 버튼을 통해 주기적으로 내 컴퓨터(D드라이브, USB 등)에 파일 하나로 백업해 두세요! 백업 없이 기기를 정리하면 복구할 수 없습니다.

---

## 📱 1. 앱 설치 방법 (홈 화면에 추가)

MyMemo는 앱스토어를 거치지 않고 웹 브라우저에서 바로 단독 앱(PWA)으로 설치할 수 있습니다:

- **스마트폰 (갤럭시 / 안드로이드)**:
  - 크롬 또는 삼성인터넷 우측 상단 **메뉴(점 3개)** ➔ **[홈 화면에 추가]** 또는 **[앱 설치]** 클릭
- **스마트폰 (아이폰 / iOS)**:
  - 사파리(Safari) 브라우저 하단 **공유 버튼(네모+화살표)** ➔ **[홈 화면에 추가]** 클릭
- **PC (크롬 / 엣지 / 웨일)**:
  - 브라우저 주소창 오른쪽 끝에 있는 **[모니터/컴퓨터 모양 설치 아이콘]** ➔ **[설치]** 클릭

> 💡 설치하면 바탕화면에 '엔젤기프트 날개' 아이콘이 생성되어 독립된 전용 창으로 편리하게 실행됩니다.

---

## ☁️ 2. GitHub 3분 동기화 설정 가이드

MyMemo의 모든 메모는 **내 개인 GitHub 저장소(Repository)**에 안전하게 백업 및 동기화됩니다.

### 1단계: GitHub 저장소 만들기
1. [GitHub(github.com)](https://github.com)에 로그인합니다.
2. 우측 상단 **\`+\`** 버튼 클릭 ➔ **[New repository]** 선택
3. **Repository name**에 \`my-notes\` 입력
4. 공개 설정은 **Private**(비공개) 권장
5. ⚠️ **중요 꿀팁**: **[Add a README file]**에 체크(V)해 주세요! (체크를 깜빡 잊고 빈 저장소로 만들어도 앱에서 자동 처리됩니다.)
6. 맨 아래 초록색 **[Create repository]** 클릭!

### 2단계: 개인 접근 토큰(PAT) 발급받기
1. GitHub 우측 상단 내 **프로필 사진** 클릭 ➔ **Settings** 메뉴로 들어갑니다.
2. 왼쪽 메뉴 맨 아래 **Developer settings** 클릭
3. **Personal access tokens** ➔ **Tokens (classic)** 클릭
4. **[Generate new token]** ➔ **[Generate new token (classic)]** 클릭
5. **Note**: \`MyMemo\` 입력
6. **Expiration**: \`No expiration\`(만료 없음) 선택 권장
7. **Select scopes(권한)**: 맨 위에 있는 **\`repo\`** (Full control of private repositories) 체크박스 하나만 체크!
8. 맨 아래 초록색 **[Generate token]** 버튼 클릭!
9. 생성된 녹색 상자의 토큰 문자열(\`ghp_...\`로 시작)을 **[복사]**합니다. *(이 창을 닫으면 다시 볼 수 없으니 꼭 복사하세요!)*

### 3단계: MyMemo 앱에서 연결하기
1. MyMemo 앱 화면 좌측 하단(또는 메뉴)의 **⚙️ [설정]** 버튼을 누릅니다.
2. 정보를 입력합니다:
   - **GitHub 아이디 (Owner)**: 내 GitHub 계정 아이디
   - **저장소 이름 (Repository)**: \`my-notes\`
   - **개인 액세스 토큰 (Token)**: 복사한 \`ghp_...\` 토큰 붙여넣기
3. 하단의 **[연결 테스트]** 버튼을 눌러 초록색 성공 메시지를 확인한 뒤 **[저장 및 완료]**를 누르면 끝!

---

## 🚀 3. MyMemo 핵심 기능 사용 팁

- **달력 일정 드래그 이동**: 달력 뷰에서 등록된 일정을 다른 날짜 칸으로 마우스 드래그하면 날짜가 바로 변경됩니다!
- **글자 색상 & 형광펜**: 텍스트를 드래그하고 상단 툴바의 **팔레트 아이콘(🎨)**을 누르면 7가지 색상과 4가지 형광펜을 적용할 수 있습니다.
- **표(Table) & PDF 첨부**: 툴바의 표 아이콘, PDF 첨부 버튼을 통해 문서와 표를 깔끔하게 관리할 수 있습니다.
- **🔒 기기 전용(로컬) 보관 & 백업**: 메모 상단의 자물쇠 아이콘을 켜면 해당 메모는 GitHub로 전송되지 않고 오직 이 기기에만 안전하게 보관됩니다. 좌측 하단 ⚙️ [설정] 창에서 언제든 원클릭으로 전체 백업 파일(.json)을 내려받거나 복원할 수 있습니다.
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
  } else {
    // 기존 사용자의 환영 메모에도 최신 백업 안내 문구가 반영되도록 자동 보강
    try {
      const existingWelcome = await db.get('notes', 'note-welcome');
      if (existingWelcome && !existingWelcome.content.includes('데이터 안전 & 백업 필수 안내')) {
        const warningBlock = `> ⚠️ **[필독] 처음 사용하는 분들을 위한 데이터 안전 & 백업 필수 안내**\n> \n> MyMemo는 별도의 회원가입 없이 브라우저/기기 내부 저장소(IndexedDB)에 데이터를 즉시 보관합니다.\n> **사용자가 백업을 하지 않고 브라우저 기록(쿠키/사이트 데이터)을 청소하거나 앱을 삭제하면, 작성한 메모와 첨부파일이 모두 사라져 큰 데이터 손실이 발생할 수 있습니다!**\n> \n> 🛡️ **소중한 데이터를 안전하게 지키는 2가지 방법:**\n> 1. **GitHub 동기화 연결 (가장 권장)**: 아래 가이드를 따라 GitHub를 연결해 두시면, 기기를 교체하거나 브라우저를 초기화해도 언제든 원클릭으로 모든 메모가 100% 복원됩니다.\n> 2. **주기적인 로컬 전체 백업 (.json 저장)**: 특히 자물쇠가 달린 **'기기 전용(🔒)'** 메모는 클라우드로 가지 않으므로, 좌측 하단 **⚙️ [설정] ➔ [로컬 전체 데이터 백업(.json)]** 버튼을 통해 주기적으로 내 컴퓨터(D드라이브, USB 등)에 파일 하나로 백업해 두세요! 백업 없이 기기를 정리하면 복구할 수 없습니다.\n\n---\n\n`;

        if (existingWelcome.content.includes('---')) {
          const parts = existingWelcome.content.split('---');
          existingWelcome.content = `${parts[0]}---\n\n${warningBlock}${parts.slice(1).join('---').replace(/^\s*>\s*⚠️[\s\S]*?---\n\n/, '')}`;
        } else {
          existingWelcome.content = `${existingWelcome.content}\n\n${warningBlock}`;
        }
        existingWelcome.updatedAt = new Date().toISOString();
        await db.put('notes', existingWelcome);
      }
    } catch {
      // ignore
    }
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

// 원격 삭제 추적 큐 (톰스톤 시스템)
export interface DeletedTombstone {
  id?: string;
  path?: string;
  deletedAt: string;
}

export async function trackDeletedNote(id?: string, path?: string): Promise<void> {
  if (!id && !path) return;
  const db = await getDB();
  const current: DeletedTombstone[] = (await db.get('settings', 'deleted_tombstones')) || [];
  const exists = current.some((t) => (id && t.id === id) || (path && t.path === path));
  if (!exists) {
    current.push({ id, path, deletedAt: new Date().toISOString() });
    await db.put('settings', current, 'deleted_tombstones');
  }

  // 하위 호환성 유지
  if (path) {
    await trackDeletedRemotePath(path);
  }
}

export async function getDeletedTombstones(): Promise<DeletedTombstone[]> {
  const db = await getDB();
  return (await db.get('settings', 'deleted_tombstones')) || [];
}

export async function clearDeletedTombstones(itemsToRemove: { id?: string; path?: string }[]): Promise<void> {
  const db = await getDB();
  const current: DeletedTombstone[] = (await db.get('settings', 'deleted_tombstones')) || [];
  const remaining = current.filter((t) => {
    return !itemsToRemove.some((r) => (r.id && r.id === t.id) || (r.path && r.path === t.path));
  });
  await db.put('settings', remaining, 'deleted_tombstones');
}

// 기존 함수 하위 호환성
export async function trackDeletedRemotePath(path: string): Promise<void> {
  if (!path) return;
  const db = await getDB();
  const current: string[] = (await db.get('settings', 'deleted_remote_paths')) || [];
  if (!current.includes(path)) {
    current.push(path);
    await db.put('settings', current, 'deleted_remote_paths');
  }
}

export async function getDeletedRemotePaths(): Promise<string[]> {
  const db = await getDB();
  return (await db.get('settings', 'deleted_remote_paths')) || [];
}

export async function clearDeletedRemotePaths(pathsToRemove: string[]): Promise<void> {
  const db = await getDB();
  const current: string[] = (await db.get('settings', 'deleted_remote_paths')) || [];
  const remaining = current.filter((p) => !pathsToRemove.includes(p));
  await db.put('settings', remaining, 'deleted_remote_paths');
}

// 폴더(프로젝트) 삭제 추적 큐 (폴더 톰스톤 시스템)
export interface DeletedProjectTombstone {
  id: string;
  name?: string;
  deletedAt: string;
}

export async function trackDeletedProject(id: string, name?: string): Promise<void> {
  if (!id) return;
  const db = await getDB();
  const current: DeletedProjectTombstone[] = (await db.get('settings', 'deleted_project_tombstones')) || [];
  const exists = current.some((t) => t.id === id);
  if (!exists) {
    current.push({ id, name, deletedAt: new Date().toISOString() });
    await db.put('settings', current, 'deleted_project_tombstones');
  }
}

export async function getDeletedProjects(): Promise<DeletedProjectTombstone[]> {
  const db = await getDB();
  return (await db.get('settings', 'deleted_project_tombstones')) || [];
}

export async function clearDeletedProjects(idsToRemove: string[]): Promise<void> {
  const db = await getDB();
  const current: DeletedProjectTombstone[] = (await db.get('settings', 'deleted_project_tombstones')) || [];
  const remaining = current.filter((t) => !idsToRemove.includes(t.id));
  await db.put('settings', remaining, 'deleted_project_tombstones');
}
