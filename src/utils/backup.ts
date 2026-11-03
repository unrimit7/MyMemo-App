import type { Note, Project } from '../types';

export interface MyMemoBackupData {
  version: number;
  appName: 'MyMemo';
  createdAt: string;
  stats: {
    projectsCount: number;
    notesCount: number;
    attachmentsCount: number;
    totalSizeBytes?: number;
  };
  projects: Project[];
  notes: Note[];
}

/**
 * 전체 메모 및 프로젝트를 단일 JSON 백업 파일로 내보내기합니다.
 * @param projects 프로젝트 목록
 * @param notes 메모 목록
 * @param includeAttachments 첨부파일(PDF/이미지 등) 포함 여부 (기본 true)
 */
export function exportBackupData(
  projects: Project[],
  notes: Note[],
  includeAttachments: boolean = true
): { filename: string; blob: Blob; sizeKb: number; attachmentsCount: number } {
  let totalAttachmentsCount = 0;

  // 첨부파일 포함 여부에 따라 데이터 정돈
  const processedNotes: Note[] = notes.map((note) => {
    if (!note.attachments || Object.keys(note.attachments).length === 0) {
      return note;
    }

    totalAttachmentsCount += Object.keys(note.attachments).length;

    if (includeAttachments) {
      return note;
    }

    // 첨부파일 미포함 옵션인 경우 용량 절약을 위해 첨부 데이터 제거
    return {
      ...note,
      attachments: {},
    };
  });

  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
  const suffix = includeAttachments ? 'full' : 'text_only';
  const filename = `mymemo_backup_${dateStr}_${suffix}.json`;

  const backupData: MyMemoBackupData = {
    version: 2,
    appName: 'MyMemo',
    createdAt: now.toISOString(),
    stats: {
      projectsCount: projects.length,
      notesCount: processedNotes.length,
      attachmentsCount: includeAttachments ? totalAttachmentsCount : 0,
    },
    projects,
    notes: processedNotes,
  };

  const jsonString = JSON.stringify(backupData, null, 2);
  const blob = new Blob([jsonString], { type: 'application/json;charset=utf-8' });
  const sizeKb = Math.round(blob.size / 1024);

  return {
    filename,
    blob,
    sizeKb,
    attachmentsCount: backupData.stats.attachmentsCount,
  };
}

/**
 * 브라우저에서 Blob 데이터를 다운로드하도록 트리거합니다.
 */
export function downloadBackupBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

/**
 * 사용자가 선택한 JSON 백업 파일의 유효성을 검사하고 데이터를 파싱합니다.
 */
export async function parseBackupFile(file: File): Promise<MyMemoBackupData> {
  const text = await file.text();
  let parsed: any;
  try {
    parsed = JSON.parse(text);
  } catch {
    throw new Error('유효한 JSON 파일 형식이 아닙니다.');
  }

  if (!parsed || (parsed.appName !== 'MyMemo' && !Array.isArray(parsed.notes))) {
    throw new Error('MyMemo 백업 파일 형식이 아닙니다.');
  }

  const projects: Project[] = Array.isArray(parsed.projects) ? parsed.projects : [];
  const notes: Note[] = Array.isArray(parsed.notes) ? parsed.notes : [];

  let attachmentsCount = 0;
  for (const n of notes) {
    if (n.attachments) {
      attachmentsCount += Object.keys(n.attachments).length;
    }
  }

  return {
    version: parsed.version || 1,
    appName: 'MyMemo',
    createdAt: parsed.createdAt || new Date().toISOString(),
    stats: {
      projectsCount: projects.length,
      notesCount: notes.length,
      attachmentsCount,
    },
    projects,
    notes,
  };
}
