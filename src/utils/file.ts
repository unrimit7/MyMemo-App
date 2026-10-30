import { compressImageToDataUrl } from './image';
import type { AttachmentItem } from '../types';

export interface ProcessedDroppedFile {
  title: string;
  content: string;
  attachments?: Record<string, AttachmentItem>;
  type: 'text' | 'image' | 'pdf' | 'other';
}

// 드롭된 파일을 분석하여 마크다운 메모용 제목과 본문으로 변환 (거대한 Base64를 본문에서 숨기고 단일 줄 태그로 삽입)
export async function processFileForNote(file: File): Promise<ProcessedDroppedFile> {
  const fileName = file.name;
  const safeName = fileName.replace(/`/g, "'").replace(/[\[\]]/g, '').trim();
  const baseTitle = fileName.replace(/\.[^/.]+$/, '');
  const safeBaseTitle = safeName.replace(/\.[^/.]+$/, '');
  const mimeType = file.type;

  // 1. 텍스트 또는 마크다운 파일
  if (
    mimeType.startsWith('text/') ||
    fileName.endsWith('.txt') ||
    fileName.endsWith('.md') ||
    fileName.endsWith('.json') ||
    fileName.endsWith('.csv')
  ) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = (e.target?.result as string) || '';
        resolve({
          title: baseTitle,
          content: text,
          type: 'text',
        });
      };
      reader.onerror = () => {
        resolve({
          title: baseTitle,
          content: '(파일을 읽지 못했습니다)',
          type: 'text',
        });
      };
      reader.readAsText(file);
    });
  }

  // 2. 이미지 파일
  if (mimeType.startsWith('image/')) {
    try {
      const dataUrl = await compressImageToDataUrl(file, 1000, 0.75);
      const attId = `att-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
      return {
        title: safeBaseTitle,
        content: `# 🖼️ ${safeBaseTitle}\n\n![${safeName}](attachment:${attId})\n\n`,
        attachments: {
          [attId]: {
            id: attId,
            name: fileName,
            type: mimeType || 'image/webp',
            data: dataUrl,
          },
        },
        type: 'image',
      };
    } catch {
      return {
        title: safeBaseTitle,
        content: `이미지 파일: ${fileName}`,
        type: 'image',
      };
    }
  }

  // 3. PDF 파일
  if (mimeType === 'application/pdf' || fileName.endsWith('.pdf')) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const dataUrl = e.target?.result as string;
        const attId = `att-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
        resolve({
          title: `[PDF] ${safeBaseTitle}`,
          content: `# 📄 ${safeName}\n\n[📄 PDF 문서: ${safeName}](attachment:${attId})\n\n`,
          attachments: {
            [attId]: {
              id: attId,
              name: fileName,
              type: 'application/pdf',
              data: dataUrl,
            },
          },
          type: 'pdf',
        });
      };
      reader.onerror = () => {
        resolve({
          title: `[PDF] ${safeBaseTitle}`,
          content: `PDF 파일: ${fileName}`,
          type: 'pdf',
        });
      };
      reader.readAsDataURL(file);
    });
  }

  // 4. 기타 파일 (링크 형태)
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const dataUrl = e.target?.result as string;
      const attId = `att-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
      resolve({
        title: safeBaseTitle,
        content: `# 📎 ${safeName}\n\n[📎 파일: ${safeName}](attachment:${attId})\n\n`,
        attachments: {
          [attId]: {
            id: attId,
            name: fileName,
            type: mimeType || 'application/octet-stream',
            data: dataUrl,
          },
        },
        type: 'other',
      });
    };
    reader.onerror = () => {
      resolve({
        title: baseTitle,
        content: `파일: ${fileName}`,
        type: 'other',
      });
    };
    reader.readAsDataURL(file);
  });
}
