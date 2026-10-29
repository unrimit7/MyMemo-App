import type { GitHubConfig, Note, SyncSummary } from '../types';
import { saveNote, getAllNotes, getAllProjects, saveProject } from './db';

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

// 마크다운 Frontmatter 파서/시리얼라이저
export function serializeNoteToMarkdown(note: Note, projectName: string): string {
  const frontmatter = `---
id: "${note.id}"
title: "${note.title.replace(/"/g, '\\"')}"
project: "${projectName.replace(/"/g, '\\"')}"
createdAt: "${note.createdAt}"
updatedAt: "${note.updatedAt}"
---

`;
  return frontmatter + note.content;
}

export function parseMarkdownToNote(raw: string): {
  id?: string;
  title?: string;
  projectName?: string;
  content: string;
  createdAt?: string;
  updatedAt?: string;
} {
  const fmRegex = /^---\r?\n([\s\S]*?)\r?\n---\r?\n([\s\S]*)$/;
  const match = raw.match(fmRegex);

  if (!match) {
    // Frontmatter가 없는 일반 마크다운 파일
    return {
      content: raw,
    };
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

  // 2) 로컬 메모들을 GitHub로 업로드 (단, isLocalOnly === true는 엄격히 건너뜀!)
  for (const note of notes) {
    if (note.isLocalOnly) {
      summary.skippedLocalOnly++;
      continue;
    }

    const project = projects.find((p) => p.id === note.projectId);
    // 만약 상위 프로젝트 자체가 기기 전용이라면 이 메모도 건너뜀
    if (project && project.isLocalOnly) {
      summary.skippedLocalOnly++;
      continue;
    }

    const safeProjName = (project?.name || '기타').replace(/[\/\\:*?"<>|]/g, '_');
    const safeTitle = (note.title || '제목없음').replace(/[\/\\:*?"<>|]/g, '_');
    const expectedPath = `${notesDir}/${safeProjName}/${safeTitle}_${note.id.slice(-6)}.md`;

    const remoteFile = remoteFileMap.get(expectedPath);
    const serializedContent = serializeNoteToMarkdown(note, project?.name || '기타');

    if (remoteFile) {
      // 이미 원격에 파일이 있음
      if (note.githubSha && note.githubSha !== remoteFile.sha && note.syncStatus === 'pending') {
        // 충돌 감지! (원격도 바뀌었고 로컬도 수정됨)
        summary.conflicts++;
        
        // 원격 내용 다운로드
        const remoteContentRes = await fetch(remoteFile.url, { headers: getHeaders(config.token) });
        const remoteContentData = await remoteContentRes.json();
        const remoteDecoded = base64ToUtf8(remoteContentData.content);

        // 충돌 사본 메모 생성 (데이터 유실 방지)
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

        // 로컬 원본은 내 변경사항으로 새 sha 반영
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
        // 로컬 변경사항 업로드
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
        // 이미 최신 동기화 상태
        note.syncStatus = 'synced';
        await saveNote(note);
      }
      remoteFileMap.delete(expectedPath);
    } else {
      // 원격에 없는 새 메모 -> 업로드 생성
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

  // 3) 원격에만 있고 로컬에 없는 파일들을 로컬로 다운로드
  for (const [path, remoteFile] of remoteFileMap.entries()) {
    const fileRes = await fetch(remoteFile.url, { headers: getHeaders(config.token) });
    if (!fileRes.ok) continue;

    const fileData = await fileRes.json();
    const rawContent = base64ToUtf8(fileData.content);
    const parsed = parseMarkdownToNote(rawContent);

    // 프로젝트 찾기 또는 자동 생성
    const projName = parsed.projectName || path.split('/')[1] || '원격 메모';
    let targetProj = projects.find((p) => p.name === projName);
    if (!targetProj) {
      targetProj = {
        id: `proj-${Date.now()}`,
        name: projName,
        isLocalOnly: false,
        createdAt: new Date().toISOString(),
      };
      await saveProject(targetProj);
      projects.push(targetProj);
    }

    const newNote: Note = {
      id: parsed.id || `note-${Date.now()}`,
      projectId: targetProj.id,
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
