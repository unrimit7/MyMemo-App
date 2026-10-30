import type { AttachmentItem, GitHubConfig, Note, Project, SyncSummary } from '../types';
import { cleanLegacyAttachments } from '../utils/attachment';
import { 
  saveNote, 
  deleteNote,
  getAllNotes, 
  getAllProjects, 
  saveProject, 
  getDeletedTombstones, 
  clearDeletedTombstones 
} from './db';

// UTF-8 안전한 Base64 인코딩/디코딩
export function utf8ToBase64(str: string): string {
  return btoa(
    encodeURIComponent(str).replace(/%([0-9A-F]{2})/g, (_, p1) =>
      String.fromCharCode(parseInt(p1, 16))
    )
  );
}

export function base64ToUtf8(str: string): string {
  try {
    return decodeURIComponent(
      atob(str.replace(/\s/g, ''))
        .split('')
        .map((c) => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
        .join('')
    );
  } catch {
    return atob(str.replace(/\s/g, ''));
  }
}

function getHeaders(token: string) {
  return {
    Authorization: `token ${token}`,
    Accept: 'application/vnd.github.v3+json',
    'Content-Type': 'application/json',
  };
}

// GitHub Contents API URL 안전한 인코딩 생성
export function getContentsUrl(config: GitHubConfig, path: string): string {
  const encodedPath = path
    .split('/')
    .map((seg) => encodeURIComponent(seg))
    .join('/');
  return `https://api.github.com/repos/${config.owner}/${config.repo}/contents/${encodedPath}`;
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

// 폴더 ID로부터 전체 계층 경로 구하기
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
  let body = note.content;
  if (note.attachments && Object.keys(note.attachments).length > 0) {
    const serializedAtts = JSON.stringify(note.attachments);
    body += `\n\n<!-- MYMEMO_ATTACHMENTS:\n${serializedAtts}\n:MYMEMO_ATTACHMENTS -->\n`;
  }
  return frontmatter + body;
}

export function parseMarkdownToNote(raw: string): {
  id?: string;
  title?: string;
  folder?: string;
  content: string;
  attachments?: Record<string, AttachmentItem>;
  createdAt?: string;
  updatedAt?: string;
} {
  const fmRegex = /^---\r?\n([\s\S]*?)\r?\n---\r?\n([\s\S]*)$/;
  const match = raw.match(fmRegex);

  let rawContent = raw;
  const result: any = {};

  if (match) {
    const fmText = match[1];
    rawContent = match[2];

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
  }

  // 첨부파일 딕셔너리 파싱
  let parsedAttachments: Record<string, AttachmentItem> = {};
  const attMatch = rawContent.match(/<!--\s*MYMEMO_ATTACHMENTS:\r?\n([\s\S]*?)\r?\n:MYMEMO_ATTACHMENTS\s*-->/);
  if (attMatch) {
    try {
      parsedAttachments = JSON.parse(attMatch[1]);
      rawContent = rawContent.replace(attMatch[0], '').trim();
    } catch (e) {
      console.warn('첨부파일 파싱 실패:', e);
    }
  }

  // 본문 내 레거시 Base64 자동 정리
  const { content: cleanedContent, attachments: finalAttachments } = cleanLegacyAttachments(
    rawContent,
    parsedAttachments
  );

  result.content = cleanedContent;
  result.attachments = finalAttachments;
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
    const url = getContentsUrl(config, path);
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
      // 409 Conflict 발생 시 최신 SHA를 가져와 재시도(Last-Write-Wins)
      if (res.status === 409) {
        try {
          const getRes = await fetch(url, { headers: getHeaders(config.token) });
          if (getRes.ok) {
            const getData = await getRes.json();
            if (getData.sha) {
              body.sha = getData.sha;
              const retryRes = await fetch(url, {
                method: 'PUT',
                headers: getHeaders(config.token),
                body: JSON.stringify(body),
              });
              if (retryRes.ok) {
                const retryData = await retryRes.json();
                return { success: true, sha: retryData.content?.sha };
              }
            }
          }
        } catch {
          // retry fallback
        }
      }
      const err = await res.json().catch(() => ({}));
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
    const url = getContentsUrl(config, path);
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

// 3. GitHub 전체 동기화 실행 (톰스톤 삭제 보장 & 삭제된 원격 메모 로컬 정리)
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

  // 1-1) 원격 폴더 구조(.folders.json) 동기화: 메모가 0개인 빈 폴더도 기기 간에 사라지지 않도록 복원
  const foldersJsonFile = (treeData.tree || []).find(
    (item: any) => item.type === 'blob' && item.path === `${notesDir}/.folders.json`
  );

  if (foldersJsonFile) {
    try {
      const foldersRes = await fetch(foldersJsonFile.url, { headers: getHeaders(config.token) });
      if (foldersRes.ok) {
        const foldersData = await foldersRes.json();
        const rawFolders = base64ToUtf8(foldersData.content);
        const remoteProjects: Project[] = JSON.parse(rawFolders);

        for (const rp of remoteProjects) {
          const localExists = projects.find((p) => p.id === rp.id || (p.name === rp.name && p.parentId === rp.parentId));
          if (!localExists) {
            await saveProject(rp);
            projects.push(rp);
          }
        }
      }
    } catch (e) {
      console.warn('원격 폴더 구조 로드 실패:', e);
    }
  }

  // 1-2) 톰스톤(삭제 큐) 처리: 로컬에서 삭제된 파일들을 원격에서도 확실히 제거
  const tombstones = await getDeletedTombstones();
  const deletedSuccessTombstones: { id?: string; path?: string }[] = [];

  for (const t of tombstones) {
    let targetPath = t.path;
    let targetFile = t.path ? remoteFileMap.get(t.path) : undefined;

    // 경로로 바로 못 찾았다면 idSuffix로 찾기
    if (!targetFile && t.id) {
      const idSuffix = `_${t.id.slice(-6)}.md`;
      for (const [p, f] of remoteFileMap.entries()) {
        if (p.endsWith(idSuffix)) {
          targetFile = f;
          targetPath = p;
          break;
        }
      }
    }

    if (targetFile && targetPath) {
      const delRes = await deleteFileFromGitHub(config, targetPath, targetFile.sha, `Delete note: ${targetPath}`);
      if (delRes.success) {
        remoteFileMap.delete(targetPath);
        deletedSuccessTombstones.push(t);
      }
    } else {
      // 원격에도 이미 존재하지 않으면 처리 완료로 간주
      deletedSuccessTombstones.push(t);
    }
  }

  // 2) 로컬 메모와 원격 메모 간 양방향 동기화 (업로드 및 원격 변경사항 다운로드)
  for (const note of notes) {
    if (note.isLocalOnly || isChainLocalOnly(note.projectId)) {
      summary.skippedLocalOnly++;
      continue;
    }

    const folderPath = getProjectFolderPath(note.projectId, projects);
    const safeTitle = (note.title || '제목없음').replace(/[\/\\:*?"<>|]/g, '_');
    const expectedPath = `${notesDir}/${folderPath}/${safeTitle}_${note.id.slice(-6)}.md`;
    const idSuffix = `_${note.id.slice(-6)}.md`;
    const legacyPath = `${notesDir}/${folderPath}/${safeTitle}.md`;

    // 1. 해당 로컬 메모와 일치하는 원격 파일 찾기 (경로, ID 접미사, 또는 레거시 경로 기준)
    let remoteFile = remoteFileMap.get(expectedPath);
    let remotePath = expectedPath;

    if (!remoteFile && note.githubPath && remoteFileMap.has(note.githubPath)) {
      remoteFile = remoteFileMap.get(note.githubPath);
      remotePath = note.githubPath;
    }

    if (!remoteFile) {
      for (const [p, f] of remoteFileMap.entries()) {
        if (p.endsWith(idSuffix)) {
          remoteFile = f;
          remotePath = p;
          break;
        }
      }
    }

    if (!remoteFile && remoteFileMap.has(legacyPath)) {
      remoteFile = remoteFileMap.get(legacyPath);
      remotePath = legacyPath;
    }

    const serializedContent = serializeNoteToMarkdown(note, folderPath);

    if (remoteFile) {
      // 3단계의 '신규 파일 다운로드'에서 중복 처리되지 않도록 맵에서 제거
      remoteFileMap.delete(remotePath);
      if (note.githubPath && note.githubPath !== remotePath) {
        remoteFileMap.delete(note.githubPath);
      }
      if (remoteFileMap.has(legacyPath)) {
        // 구버전 파일이 원격에 따로 남아있다면 원격에서 즉시 정리하여 좀비 방지!
        const legacyTarget = remoteFileMap.get(legacyPath)!;
        await deleteFileFromGitHub(config, legacyPath, legacyTarget.sha, `Clean legacy duplicate: ${legacyPath}`);
        remoteFileMap.delete(legacyPath);
      }

      const isRemoteChanged = note.githubSha !== remoteFile.sha;
      const isLocalPending = note.syncStatus === 'pending';

      if (!isLocalPending && isRemoteChanged) {
        // [경우 A] 로컬은 수정되지 않았으나, 다른 기기(PC 등)에서 수정하여 원격 SHA가 변경된 경우:
        // -> 원격의 최신 내용을 다운로드하여 로컬 메모에 덮어쓰기 반영!
        const fileRes = await fetch(remoteFile.url, { headers: getHeaders(config.token) });
        if (fileRes.ok) {
          const fileData = await fileRes.json();
          const rawContent = base64ToUtf8(fileData.content);
          const parsed = parseMarkdownToNote(rawContent);

          note.title = parsed.title || note.title;
          note.content = parsed.content;
          note.attachments = parsed.attachments;
          note.githubSha = remoteFile.sha;
          note.githubPath = remotePath;
          note.updatedAt = parsed.updatedAt || new Date().toISOString();
          note.syncStatus = 'synced';
          await saveNote(note);
          summary.downloaded++;
        }
      } else if (isLocalPending && !isRemoteChanged) {
        // [경우 B] 로컬에서만 수정되었고, 원격은 변경되지 않은 경우:
        // -> 로컬 내용을 원격에 업로드!
        // 만약 파일 경로가 바뀌었다면 이전 파일 정리
        if (remotePath !== expectedPath) {
          await deleteFileFromGitHub(config, remotePath, remoteFile.sha, `Rename note: ${note.title}`);
        }

        const upRes = await uploadFileToGitHub(
          config,
          expectedPath,
          serializedContent,
          `Update note: ${note.title}`,
          remotePath === expectedPath ? remoteFile.sha : undefined
        );
        if (upRes.success) {
          note.githubSha = upRes.sha;
          note.githubPath = expectedPath;
          note.syncStatus = 'synced';
          await saveNote(note);
          summary.uploaded++;
        }
      } else if (isLocalPending && isRemoteChanged) {
        // [경우 C] 양쪽 기기에서 동시에 수정된 경우 (충돌):
        // 더 최신 수정 일시를 가진 쪽을 우선 반영하거나 분기
        const fileRes = await fetch(remoteFile.url, { headers: getHeaders(config.token) });
        if (fileRes.ok) {
          const fileData = await fileRes.json();
          const rawContent = base64ToUtf8(fileData.content);
          const parsed = parseMarkdownToNote(rawContent);

          const localTime = new Date(note.updatedAt).getTime();
          const remoteTime = new Date(parsed.updatedAt || 0).getTime();

          if (localTime >= remoteTime) {
            // 로컬이 더 최신이면 원격에 덮어쓰기
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
            // 원격이 더 최신이면 원격 내용 반영
            note.title = parsed.title || note.title;
            note.content = parsed.content;
            note.attachments = parsed.attachments;
            note.githubSha = remoteFile.sha;
            note.githubPath = remotePath;
            note.updatedAt = parsed.updatedAt || new Date().toISOString();
            note.syncStatus = 'synced';
            await saveNote(note);
            summary.downloaded++;
          }
        }
      } else {
        // [경우 D] 양쪽 모두 이미 동일한 상태
        note.syncStatus = 'synced';
        note.githubPath = remotePath;
        await saveNote(note);
      }
    } else {
      // [경우 E] 원격에 해당 파일이 없는 경우:
      const isRecentlyModified = Date.now() - new Date(note.updatedAt).getTime() < 5 * 60 * 1000;

      if (note.syncStatus === 'pending' || !note.githubPath || isRecentlyModified) {
        // 새로 작성되었거나 최근 5분 내 수정된 로컬 메모는 원격에 생성 업로드!
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
      } else {
        // 다른 기기에서 삭제되어 원격에서 사라진 경우 로컬에서도 삭제 반영
        await deleteNote(note.id);
      }
    }
  }

  // 3) 원격에만 있고 로컬에 없는 신규 파일들을 다운로드
  for (const [path, remoteFile] of remoteFileMap.entries()) {
    const fileRes = await fetch(remoteFile.url, { headers: getHeaders(config.token) });
    if (!fileRes.ok) continue;

    const fileData = await fileRes.json();
    const rawContent = base64ToUtf8(fileData.content);
    const parsed = parseMarkdownToNote(rawContent);

    // ★ 핵심: 로컬 톰스톤(삭제 기록)에 있는 메모인지 확인!
    const isDeletedLocally = tombstones.some(
      (t) => (parsed.id && t.id === parsed.id) || 
             (t.path && t.path === path) ||
             (t.id && path.endsWith(`_${t.id.slice(-6)}.md`))
    );
    if (isDeletedLocally) {
      // 로컬에서 이미 삭제된 메모가 원격에 남아있는 경우 원격 파일도 삭제하고 다운로드하지 않음!
      const delRes = await deleteFileFromGitHub(config, path, remoteFile.sha, `Delete tombstone note: ${path}`);
      if (delRes.success) {
        deletedSuccessTombstones.push({ id: parsed.id, path });
      }
      continue;
    }

    // ★ 1. 만약 로컬에 이미 같은 ID를 가진 메모가 있다면 타임스탬프 비교 갱신
    const existingLocalNote = parsed.id ? notes.find((n) => n.id === parsed.id) : null;
    if (existingLocalNote) {
      const localTime = new Date(existingLocalNote.updatedAt || 0).getTime();
      const remoteTime = new Date(parsed.updatedAt || 0).getTime();

      if (remoteTime > localTime) {
        existingLocalNote.title = parsed.title || existingLocalNote.title;
        existingLocalNote.content = parsed.content;
        existingLocalNote.attachments = parsed.attachments;
        existingLocalNote.githubSha = remoteFile.sha;
        existingLocalNote.githubPath = path;
        existingLocalNote.updatedAt = parsed.updatedAt || new Date().toISOString();
        existingLocalNote.syncStatus = 'synced';
        await saveNote(existingLocalNote);
        summary.downloaded++;
      } else {
        // 로컬이 더 최신이면 이전 경로의 원격 잔여 파일 정리
        if (existingLocalNote.githubPath && existingLocalNote.githubPath !== path) {
          await deleteFileFromGitHub(config, path, remoteFile.sha, `Remove obsolete duplicate: ${path}`);
        }
      }
      continue;
    }

    // ★ 2. 제목 기반 매칭: 동일 제목 메모가 로컬에 이미 존재하는 경우 중복 신규 생성 차단!
    const relativePath = path.slice(notesDir.length).replace(/^\//, '');
    const segments = relativePath.split('/');
    const remoteFileName = segments.pop()?.replace('.md', '') || '';
    const cleanRemoteTitle = (parsed.title || remoteFileName).replace(/_\w{6}$/, '');

    const existingByTitle = notes.find((n) => {
      const nTitleClean = n.title.replace(/_\w{6}$/, '');
      return nTitleClean === cleanRemoteTitle;
    });

    if (existingByTitle) {
      const localTime = new Date(existingByTitle.updatedAt || 0).getTime();
      const remoteTime = new Date(parsed.updatedAt || 0).getTime();

      if (remoteTime > localTime) {
        existingByTitle.content = parsed.content;
        existingByTitle.attachments = parsed.attachments;
        existingByTitle.updatedAt = parsed.updatedAt || new Date().toISOString();
        existingByTitle.githubSha = remoteFile.sha;
        existingByTitle.githubPath = path;
        existingByTitle.syncStatus = 'synced';
        await saveNote(existingByTitle);
        summary.downloaded++;
      } else {
        // 로컬이 최신 상태이므로 원격의 중복 잔여 파일 삭제!
        await deleteFileFromGitHub(config, path, remoteFile.sha, `Remove legacy duplicate: ${path}`);
      }
      continue;
    }

    const folderSegments = path.slice(notesDir.length).replace(/^\//, '').split('/');
    folderSegments.pop();

    let parentId: string | null = null;
    if (folderSegments.length === 0) {
      folderSegments.push('빠른 메모');
    }

    for (let i = 0; i < Math.min(folderSegments.length, 3); i++) {
      const segName = folderSegments[i];
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
      attachments: parsed.attachments,
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

  // 성공적으로 처리된 톰스톤 정리
  if (deletedSuccessTombstones.length > 0) {
    await clearDeletedTombstones(deletedSuccessTombstones);
  }

  // 4) 원격에 전체 폴더 구조(.folders.json) 동기화 보관 (메모가 하나도 없는 빈 폴더도 기기 간 완벽 유지)
  try {
    const syncableProjects = projects.filter((p) => !p.isLocalOnly);
    const serializedProjects = JSON.stringify(syncableProjects, null, 2);
    const foldersPath = `${notesDir}/.folders.json`;

    await uploadFileToGitHub(
      config,
      foldersPath,
      serializedProjects,
      'Sync folder structure',
      foldersJsonFile?.sha
    );
  } catch (e) {
    console.warn('원격 폴더 목록 업로드 실패:', e);
  }

  return summary;
}
