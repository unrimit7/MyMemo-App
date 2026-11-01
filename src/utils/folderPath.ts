import type { Project } from '../types';

/**
 * 프로젝트 ID에 대한 전체 계층 경로를 하이픈(-)으로 연결하여 반환합니다.
 * 예: "일정 > 개인" -> "일정-개인"
 */
export function getFolderHierarchyPath(projectId: string, projects: Project[]): string {
  const parts: string[] = [];
  let curr = projects.find((p) => p.id === projectId);

  while (curr) {
    const cleanName = curr.name.trim().replace(/[\/\\:*?"<>|\[\]]/g, '');
    if (cleanName) {
      parts.unshift(cleanName);
    }
    if (!curr.parentId) break;
    curr = projects.find((p) => p.id === curr?.parentId);
  }

  return parts.length > 0 ? parts.join('-') : '';
}

/**
 * 기존 제목에서 앞부분의 [폴더경로] 대괄호 접두어를 새 폴더 경로로 교체하거나 추가합니다.
 * 예: "[업무-회의] 결혼식" + "일정-개인" -> "[일정-개인] 결혼식"
 *     "결혼식" + "일정-개인" -> "[일정-개인] 결혼식"
 *     "[일정-개인] [중요] 결혼식" + "일정-업무" -> "[일정-업무] [중요] 결혼식"
 */
export function updateTitleWithFolderPath(title: string, newFolderPath: string): string {
  const trimmed = (title || '').trim();
  const match = trimmed.match(/^\s*\[([^\]]+)\]\s*(.*)$/);
  
  let pureTitle = trimmed;
  if (match) {
    pureTitle = match[2].trim();
  }
  if (!pureTitle) {
    pureTitle = '제목 없음';
  }

  if (!newFolderPath) {
    return pureTitle;
  }
  return `[${newFolderPath}] ${pureTitle}`;
}

/**
 * 제목에서 앞부분의 [ ... ] 접두어를 제거하고 순수 제목만 반환합니다.
 */
export function stripFolderPrefixFromTitle(title: string): string {
  const trimmed = (title || '').trim();
  const match = trimmed.match(/^\s*\[([^\]]+)\]\s*(.*)$/);
  if (match && match[2].trim()) {
    return match[2].trim();
  }
  return trimmed || '제목 없음';
}
