import type { Note } from '../types';
import { marked } from 'marked';
import { renderAttachmentsInHtml } from './attachment';

/**
 * 메모 내용을 깔끔하고 표준적인 A4 PDF 문서 양식으로 렌더링하여 브라우저의 PDF 저장 / 인쇄 창을 실행합니다.
 */
export function exportNoteToPdf(
  note: Note,
  folderBreadcrumbs: string = '',
  customContent?: string
): void {
  const contentToRender = customContent !== undefined ? customContent : note.content;
  const rawHtml = marked.parse(contentToRender || '') as string;
  const renderedHtml = renderAttachmentsInHtml(rawHtml, note.attachments);

  const createdDate = new Date(note.createdAt).toLocaleDateString('ko-KR', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });
  const updatedDate = new Date(note.updatedAt).toLocaleDateString('ko-KR', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  // 숨김 iframe 생성
  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.right = '0';
  iframe.style.bottom = '0';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = '0';
  iframe.style.visibility = 'hidden';
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow?.document;
  if (!doc) {
    alert('PDF 생성 창을 초기화하지 못했습니다.');
    document.body.removeChild(iframe);
    return;
  }

  const printDocumentHtml = `
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <title>${(note.title || '메모').replace(/[<>:"/\\|?*]/g, '_')}</title>
  <style>
    @page {
      size: A4 portrait;
      margin: 18mm 15mm 20mm 15mm;
    }
    *, *::before, *::after {
      box-sizing: border-box;
    }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Pretendard", "Segoe UI", Roboto, sans-serif;
      color: #0f172a;
      background: #ffffff;
      line-height: 1.65;
      font-size: 13.5px;
      margin: 0;
      padding: 0;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    
    /* 문서 상단 헤더 */
    .doc-header {
      border-bottom: 2px solid #e2e8f0;
      padding-bottom: 14px;
      margin-bottom: 22px;
    }
    .doc-folder {
      font-size: 11.5px;
      font-weight: 600;
      color: #6366f1;
      margin-bottom: 4px;
      display: flex;
      align-items: center;
      gap: 4px;
    }
    .doc-title {
      font-size: 24px;
      font-weight: 800;
      color: #0f172a;
      line-height: 1.3;
      margin: 4px 0 8px 0;
      word-break: break-all;
    }
    .doc-meta {
      font-size: 11px;
      color: #64748b;
      display: flex;
      gap: 16px;
    }

    /* 본문 마크다운 서식 */
    .doc-body {
      color: #1e293b;
    }
    h1, h2, h3, h4, h5, h6 {
      color: #0f172a;
      font-weight: 700;
      margin-top: 24px;
      margin-bottom: 10px;
      break-after: avoid;
    }
    h1 {
      font-size: 18px;
      border-bottom: 1.5px solid #cbd5e1;
      padding-bottom: 6px;
    }
    h2 {
      font-size: 16px;
      border-bottom: 1px solid #e2e8f0;
      padding-bottom: 4px;
    }
    h3 { font-size: 14.5px; }
    p {
      margin: 8px 0 12px 0;
      word-break: break-word;
    }
    strong { font-weight: 700; color: #0f172a; }
    em { font-style: italic; }
    
    /* 표 */
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 16px 0;
      font-size: 12.5px;
      break-inside: avoid;
    }
    th, td {
      border: 1px solid #cbd5e1;
      padding: 7px 10px;
      text-align: left;
    }
    th {
      background-color: #f1f5f9;
      font-weight: 700;
      color: #1e293b;
    }
    tr:nth-child(even) td {
      background-color: #f8fafc;
    }

    /* 인용구 */
    blockquote {
      margin: 12px 0;
      padding: 8px 14px;
      background-color: #f8fafc;
      border-left: 4px solid #6366f1;
      color: #334155;
      font-style: normal;
      border-radius: 0 6px 6px 0;
      break-inside: avoid;
    }

    /* 코드 */
    code {
      font-family: ui-monospace, SFMono-Regular, Consolas, monospace;
      font-size: 11.5px;
      background-color: #f1f5f9;
      color: #475569;
      padding: 2px 5px;
      border-radius: 4px;
      border: 1px solid #e2e8f0;
    }
    pre {
      background-color: #f8fafc;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      padding: 12px;
      overflow-x: auto;
      margin: 12px 0;
      break-inside: avoid;
    }
    pre code {
      border: none;
      padding: 0;
      background: transparent;
      color: #1e293b;
      font-size: 12px;
      line-height: 1.5;
    }

    /* 목록 */
    ul, ol {
      padding-left: 22px;
      margin: 8px 0;
    }
    li {
      margin: 4px 0;
    }

    /* 이미지 */
    img {
      max-width: 100%;
      height: auto;
      border-radius: 6px;
      margin: 12px 0;
      border: 1px solid #e2e8f0;
      break-inside: avoid;
    }

    /* 링크 */
    a {
      color: #4f46e5;
      text-decoration: none;
    }

    /* 인라인 PDF / 첨부 뷰어 출력 시 프린트 양식 */
    object {
      display: none !important;
    }
    .print-att-card {
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      background: #f8fafc;
      padding: 10px 14px;
      margin: 12px 0;
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 12px;
      font-weight: 600;
      color: #334155;
      break-inside: avoid;
    }
  </style>
</head>
<body>
  <div class="doc-header">
    ${folderBreadcrumbs ? `<div class="doc-folder">📁 ${folderBreadcrumbs}</div>` : ''}
    <h1 class="doc-title">${note.title || '제목 없는 메모'}</h1>
    <div class="doc-meta">
      <span>작성일: ${createdDate}</span>
      <span>최종 수정: ${updatedDate}</span>
    </div>
  </div>
  <div class="doc-body">
    ${renderedHtml}
  </div>
</body>
</html>
  `.trim();

  doc.open();
  doc.write(printDocumentHtml);
  doc.close();

  // 이미지가 있을 수 있으므로 약간의 로딩 대기 후 프린트 실행
  setTimeout(() => {
    try {
      iframe.contentWindow?.focus();
      iframe.contentWindow?.print();
    } catch (e) {
      console.error('프린트 실행 실패:', e);
    } finally {
      setTimeout(() => {
        if (document.body.contains(iframe)) {
          document.body.removeChild(iframe);
        }
      }, 2000);
    }
  }, 350);
}
