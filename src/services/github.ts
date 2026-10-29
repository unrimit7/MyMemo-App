import type { GitHubConfig, Note, Project, SyncSummary } from '../types';
import { saveNote, getAllNotes, getAllProjects, saveProject, getDeletedRemotePaths, clearDeletedRemotePaths } from './db';

// UTF-8 안전한 Base64 인코딩/디코딩 (한글 및 이모지 깨짐 방지)
export function utf8ToBase64(str: string): string {
  return btoa(
    encodeURIComponent(str).replace(/%([0-9A-F]{2})/g, (_, p1) =>
      String.fromCharCode(parseInt(p1, 16))
    )
  );
}

export function base64ToUtf8(str: string): string {
  return decodeURIComponent(
    atob(str)
      .split('')
      .map((c) => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
      .join('')
  );
}

// GitHub API 헤더 생성
function getHeaders(token: string) {
  return {
    Authorization: `token ${token}`,
    Accept: 'application/vnd.github.v3+json',
    'Content-Type': 'application/json',
  };
}

// 1. 연결 테스트
export async function testGitHubConnection(config: GitHubConfig): Promise<{ success: boolean; message: string }> {
  try {
    const res = await fetch(`https://api.github.com/repos/${config.owner}/${config.repo}`, {
      headers: getHeaders(config.token),
    });

    if (res.status === 200) {
      const data = await res.json();
      return { success: true, message: `성공! 저장소 '${data.full_name}'에 정상 연결되었습니다.` };
    } else if (res.status === 401) {
      return { success: false, message: '인증 실패: 토큰(PAT)이 잘못되었거나 만료되었습니다.' };
    } else if (res.status === 404) {
      return { success: false, message: '저장소를 찾을 수 없습니다. 아이디 또는 레포지토리 이름을 확인해 주세요.' };
    } else {
      const err = await res.json();
      return { success: false, message: `연결 오류 (${res.status}): ${err.message || '알 수 없는 오류'}` };
    }
  } catch (error: any) {
    return { success: false, message: `네트워크 오류: ${error.message}` };
  }
}

// 폴더 ID로부터 전체 계층 경로 문자열 구하기 (예: "업무/기획/디자인")
export function getProjectFolderPath(projectId: string, projects: Project[]): string {
  const parts: string[] = [];
  let curr = projects.find((p) => p.id === projectId);

  while (curr) {
    parts.unshift(curr.name.replace(/[\/\\:*?"<>|]/g, '_'));
    if (!curr.parentId) break;
    curr = projects.find((p) => p.id === curr?.parentId);
  }

  return parts.length > 0 ? parts.join('/') : '기타';
}

// 마크다운 Frontmatter 파서/시리얼라이저
export function serializeNoteToMarkdown(note: Note, folderPath: string): string {
  const frontmatter = `---
id: "${note.id}"
title: "${note.title.replace(/"/g, '\\"')}"
folder: "${folderPath.replace(/"/g, '\\"')}"
createdAt: "${note.createdAt}"
updatedAt: "${note.updatedAt}"
---

`;
  return frontmatter + note.content;
}

export function parseMarkdownToNote(raw: string): {
  id?: string;
  title?: string;
  folder?: string;
  content: string;
  createdAt?: string;
  updatedAt?: string;
} {
  const fmRegex = /^---\r?\n([\s\S]*?)\r?\n---\r?\n([\s\S]*)$/;
  const match = raw.match(fmRegex);

  if (!match) {
    return { content: raw };
  }

  const fmText = match[1];
  const content = match[2];
  const result: any = { content };

  fmText.split('\n').forEach((line) => {
    const colonIdx = line.indexOf(':');
    if (colonIdx !== -1) {
      const key = line.slice(0, colonIdx).trim();
      let val = line.slice(colonIdx + 1).trim();
      if (val.startsWith('"') && val.endsWith('"')) {
        val = val.slice(1, -1);
      }
      result[key] = val;
    }
  });

  return result;
}

// 2. 단일 파일 업로드/업데이트
export async function uploadFileToGitHub(
  config: GitHubConfig,
  path: string,
  content: string,
  commitMessage: string,
  sha?: string
): Promise<{ success: boolean; sha?: string; error?: string }> {
  try {
    const url = `https://api.github.com/repos/${config.owner}/${config.repo}/contents/${path}`;
    const body: any = {
      message: commitMessage,
      content: utf8ToBase64(content),
      branch: config.branch || 'main',
    };
    if (sha) {
      body.sha = sha;
    }

    const res = await fetch(url, {
      method: 'PUT',
      headers: getHeaders(config.token),
      body: JSON.stringify(body),
    });

    if (res.ok) {
      const data = await res.json();
      return { success: true, sha: data.content?.sha };
    } else {
      const err = await res.json();
      return { success: false, error: err.message };
    }
  } catch (err: any) {
    return { success: false, error: err.message };
  }
}

// 2-2. 파일 삭제
export async function deleteFileFromGitHub(
  config: GitHubConfig,
  path: string,
  sha: string,
  commitMessage: string
): Promise<{ success: boolean; error?: string }> {
  try {
    const url = `https://api.github.com/repos/${config.owner}/${config.repo}/contents/${path}`;
    const res = await fetch(url, {
      method: 'DELETE',
      headers: getHeaders(config.token),
      body: JSON.stringify({
        message: commitMessage,
        sha,
        branch: config.branch || 'main',
      }),
    });
    return { success: res.ok };
  } catch (err: any) {
    return { success: false, error: err.message };
  }
}

// 3. GitHub 전체 동기화 실행
export async function syncWithGitHub(config: GitHubConfig): Promise<SyncSummary> {
  const summary: SyncSummary = {
    uploaded: 0,
    downloaded: 0,
    conflicts: 0,
    skippedLocalOnly: 0,
    timestamp: new Date().toISOString(),
  };

  if (!config.isConnected || !config.token) {
    throw new Error('GitHub 설정이 완료되지 않았습니다.');
  }

  const projects = await getAllProjects();
  const notes = await getAllNotes();

  // 상위 조상 중 하나라도 isLocalOnly이면 true를 반환하는 함수
  const isChainLocalOnly = (projectId: string): boolean => {
    let curr = projects.find((p) => p.id === projectId);
    while (curr) {
      if (curr.isLocalOnly) return true;
      if (!curr.parentId) break;
      curr = projects.find((p) => p.id === curr?.parentId);
    }
    return false;
  };

  // 1) 원격 저장소 파일 트리 가져오기
  const notesDir = (config.notesDir || 'notes').replace(/^\/|\/$/g, '');
  const treeUrl = `https://api.github.com/repos/${config.owner}/${config.repo}/git/trees/${config.branch || 'main'}?recursive=1`;
  
  const treeRes = await fetch(treeUrl, { headers: getHeaders(config.token) });
  if (!treeRes.ok) {
    throw new Error('원격 저장소 브랜치 트리 조회 실패: ' + (await treeRes.text()));
  }
  const treeData = await treeRes.json();
  const remoteFiles: Array<{ path: string; sha: string; url: string }> = (treeData.tree || []).filter(
    (item: any) => item.type === 'blob' && item.path.startsWith(notesDir) && item.path.endsWith('.md')
  );

  const remoteFileMap = new Map(remoteFiles.map((f) => [f.path, f]));

  // 1-2) 사용자가 로컬에서 삭제한 파일들을 GitHub 원격에서도 삭제
  const pendingDeletes = await getDeletedRemotePaths();
  const deletedSuccessPaths: string[] = [];
  for (const delPath of pendingDeletes) {
    const remoteTarget = remoteFileMap.get(delPath);
    if (remoteTarget) {
      const delRes = await deleteFileFromGitHub(config, delPath, remoteTarget.sha, `Delete note: ${delPath}`);
      if (delRes.success) {
        deletedSuccessPaths.push(delPath);
        remoteFileMap.delete(delPath);
      }
    } else {
      deletedSuccessPaths.push(delPath);
    }
  }
  if (deletedSuccessPaths.length > 0) {
    await clearDeletedRemotePaths(deletedSuccessPaths);
  }

  // 2) 로컬 메모들을 GitHub로 업로드 (단, isLocalOnly 또는 상위 폴더가 로컬 전용이면 제외!)
  for (const note of notes) {
    if (note.isLocalOnly || isChainLocalOnly(note.projectId)) {
      summary.skippedLocalOnly++;
      continue;
    }

    const folderPath = getProjectFolderPath(note.projectId, projects);
    const safeTitle = (note.title || '제목없음').replace(/[\/\\:*?"<>|]/g, '_');
    const expectedPath = `${notesDir}/${folderPath}/${safeTitle}_${note.id.slice(-6)}.md`;

    const remoteFile = remoteFileMap.get(expectedPath);
    const serializedContent = serializeNoteToMarkdown(note, folderPath);

    if (remoteFile) {
      if (note.githubSha && note.githubSha !== remoteFile.sha && note.syncStatus === 'pending') {
        summary.conflicts++;
        
        const remoteContentRes = await fetch(remoteFile.url, { headers: getHeaders(config.token) });
        const remoteContentData = await remoteContentRes.json();
        const remoteDecoded = base64ToUtf8(remoteContentData.content);

        const conflictNote: Note = {
          id: `note-${Date.now()}-conflict`,
          projectId: note.projectId,
          title: `[충돌 복사본] ${note.title}`,
          content: remoteDecoded,
          isLocalOnly: false,
          syncStatus: 'conflict',
          conflictContent: remoteDecoded,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        };
        await saveNote(conflictNote);

        const upRes = await uploadFileToGitHub(
          config,
          expectedPath,
          serializedContent,
          `Update note: ${note.title} (resolved)`,
          remoteFile.sha
        );
        if (upRes.success) {
          note.githubSha = upRes.sha;
          note.githubPath = expectedPath;
          note.syncStatus = 'synced';
          await saveNote(note);
          summary.uploaded++;
        }
      } else if (note.syncStatus === 'pending' || !note.githubSha) {
        const upRes = await uploadFileToGitHub(
          config,
          expectedPath,
          serializedContent,
          `Update note: ${note.title}`,
          remoteFile.sha
        );
        if (upRes.success) {
          note.githubSha = upRes.sha;
          note.githubPath = expectedPath;
          note.syncStatus = 'synced';
          await saveNote(note);
          summary.uploaded++;
        }
      } else {
        note.syncStatus = 'synced';
        await saveNote(note);
      }
      remoteFileMap.delete(expectedPath);
    } else {
      const upRes = await uploadFileToGitHub(
        config,
        expectedPath,
        serializedContent,
        `Create note: ${note.title}`
      );
      if (upRes.success) {
        note.githubSha = upRes.sha;
        note.githubPath = expectedPath;
        note.syncStatus = 'synced';
        await saveNote(note);
        summary.uploaded++;
      }
    }
  }

  // 3) 원격에만 있고 로컬에 없는 파일들을 로컬로 다운로드 및 폴더 트리 동적 생성
  for (const [path, remoteFile] of remoteFileMap.entries()) {
    const fileRes = await fetch(remoteFile.url, { headers: getHeaders(config.token) });
    if (!fileRes.ok) continue;

    const fileData = await fileRes.json();
    const rawContent = base64ToUtf8(fileData.content);
    const parsed = parseMarkdownToNote(rawContent);

    // 경로 파싱: notes/메인폴더/하위폴더/파일명.md
    const relativePath = path.slice(notesDir.length).replace(/^\//, '');
    const segments = relativePath.split('/');
    segments.pop(); // 파일명 제외한 폴더 세그먼트들

    // 폴더 체인 찾기 또는 순차 생성 (최대 3단계: root -> sub1 -> sub2)
    let parentId: string | null = null;
    if (segments.length === 0) {
      // 루트 바로 아래 파일
      segments.push('빠른 메모');
    }

    for (let i = 0; i < Math.min(segments.length, 3); i++) {
      const segName = segments[i];
      let found = projects.find((p) => p.name === segName && (p.parentId || null) === parentId);
      if (!found) {
        found = {
          id: `proj-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
          name: segName,
          parentId: parentId,
          isLocalOnly: false,
          createdAt: new Date().toISOString(),
        };
        await saveProject(found);
        projects.push(found);
      }
      parentId = found.id;
    }

    const targetProjectId = parentId || 'proj-default';

    const newNote: Note = {
      id: parsed.id || `note-${Date.now()}`,
      projectId: targetProjectId,
      title: parsed.title || path.split('/').pop()?.replace('.md', '') || '제목 없음',
      content: parsed.content,
      isLocalOnly: false,
      githubSha: remoteFile.sha,
      githubPath: path,
      syncStatus: 'synced',
      createdAt: parsed.createdAt || new Date().toISOString(),
      updatedAt: parsed.updatedAt || new Date().toISOString(),
    };

    await saveNote(newNote);
    summary.downloaded++;
  }

  return summary;
}
