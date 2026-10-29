import { compressImageToDataUrl } from './image';

export interface ProcessedDroppedFile {
  title: string;
  content: string;
  type: 'text' | 'image' | 'pdf' | 'other';
}

// 드롭된 파일을 분석하여 마크다운 메모용 제목과 본문으로 변환
export async function processFileForNote(file: File): Promise<ProcessedDroppedFile> {
  const fileName = file.name;
  const baseTitle = fileName.replace(/\.[^/.]+$/, '');
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
      return {
        title: baseTitle,
        content: `# ${baseTitle}\n\n![${baseTitle}](${dataUrl})\n\n`,
        type: 'image',
      };
    } catch {
      return {
        title: baseTitle,
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
        // HTML object/iframe 태그로 브라우저 내장 PDF 뷰어 임베드
        const pdfContent = `# 📄 ${fileName}\n\n<object data="${dataUrl}" type="application/pdf" width="100%" height="600px">\n  <p>PDF를 볼 수 없습니다. <a href="${dataUrl}" download="${fileName}">다운로드</a></p>\n</object>\n\n`;
        resolve({
          title: `[PDF] ${baseTitle}`,
          content: pdfContent,
          type: 'pdf',
        });
      };
      reader.onerror = () => {
        resolve({
          title: `[PDF] ${baseTitle}`,
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
      resolve({
        title: baseTitle,
        content: `# 📎 ${fileName}\n\n[다운로드: ${fileName}](${dataUrl})\n`,
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
