import type { AttachmentItem } from '../types';

/**
 * 본문에 포함된 거대한 Base64 Data URL(<object>, <img>, <a>)을 깔끔한 단일 줄 첨부 태그로 추출 및 변환합니다.
 * 에디터에는 [📄 PDF: 파일명.pdf](attachment:att-xxx) 또는 ![이미지: 이름](attachment:att-xxx)만 남게 됩니다.
 */
export function cleanLegacyAttachments(
  content: string,
  existingAttachments: Record<string, AttachmentItem> = {}
): {
  content: string;
  attachments: Record<string, AttachmentItem>;
  hasChanges: boolean;
} {
  const attachments: Record<string, AttachmentItem> = { ...existingAttachments };
  let hasChanges = false;
  let nextIdCounter = Date.now();

  const getNewId = () => `att-${nextIdCounter++}-${Math.random().toString(36).slice(2, 6)}`;

  // 1. 거대한 <object data="data:application/pdf;base64,..."> 태그 추출
  let cleaned = content.replace(
    /<object\s+[^>]*data="(data:application\/pdf;base64,[^"]+)"[^>]*>([\s\S]*?)<\/object>/gi,
    (_match, dataUrl, inner) => {
      hasChanges = true;
      let name = '첨부_문서.pdf';

      const downloadMatch = inner.match(/download="([^"]+)"/i);
      if (downloadMatch && downloadMatch[1]) {
        name = downloadMatch[1];
      } else {
        const titleMatch = content.match(/#\s+📄\s+([^\n\r]+)/);
        if (titleMatch && titleMatch[1]) {
          name = titleMatch[1].trim();
        }
      }

      // 이미 같은 데이터가 첨부 사전에 있는지 확인
      let existingId = Object.keys(attachments).find((k) => attachments[k].data === dataUrl);
      if (!existingId) {
        existingId = getNewId();
        attachments[existingId] = {
          id: existingId,
          name,
          type: 'application/pdf',
          data: dataUrl,
        };
      }

      return `\n\n[📄 PDF: ${name}](attachment:${existingId})\n\n`;
    }
  );

  // 2. 본문에 직접 들어간 거대한 이미지 ![alt](data:image/...) 추출
  cleaned = cleaned.replace(
    /!\[([^\]]*)\]\((data:image\/[^;]+;base64,[^\)]+)\)/gi,
    (_match, altText, dataUrl) => {
      hasChanges = true;
      const name = (altText || '이미지').trim();

      let existingId = Object.keys(attachments).find((k) => attachments[k].data === dataUrl);
      if (!existingId) {
        existingId = getNewId();
        attachments[existingId] = {
          id: existingId,
          name,
          type: 'image/webp',
          data: dataUrl,
        };
      }

      return `\n\n![${name}](attachment:${existingId})\n\n`;
    }
  );

  // 3. 본문에 들어간 일반 data URL 링크 [text](data:...) 추출
  cleaned = cleaned.replace(
    /\[([^\]]*)\]\((data:[^;]+;base64,[^\)]+)\)/gi,
    (match, text, dataUrl) => {
      if (match.startsWith('!')) return match; // 이미지는 2번에서 처리
      hasChanges = true;
      const name = (text || '파일').trim();
      const mimeMatch = dataUrl.match(/^data:([^;]+);/);
      const mimeType = mimeMatch ? mimeMatch[1] : 'application/octet-stream';

      let existingId = Object.keys(attachments).find((k) => attachments[k].data === dataUrl);
      if (!existingId) {
        existingId = getNewId();
        attachments[existingId] = {
          id: existingId,
          name,
          type: mimeType,
          data: dataUrl,
        };
      }

      if (mimeType === 'application/pdf') {
        return `\n\n[📄 PDF: ${name}](attachment:${existingId})\n\n`;
      }
      return `\n\n[📎 파일: ${name}](attachment:${existingId})\n\n`;
    }
  );

  return {
    content: cleaned.trim(),
    attachments,
    hasChanges,
  };
}

/**
 * 마크다운 미리보기 렌더링 시 attachment:id 링크를 실제 브라우저 뷰어/이미지로 치환합니다.
 */
export function renderAttachmentsInHtml(
  html: string,
  attachments: Record<string, AttachmentItem> = {}
): string {
  if (!html || Object.keys(attachments).length === 0) {
    return html;
  }

  let result = html;

  // 1. <img src="attachment:att-xxx"> 치환
  result = result.replace(
    /<img\s+([^>]*?)src=["']attachment:([^"']+)["']([^>]*?)>/gi,
    (_match, pre, attId, post) => {
      const att = attachments[attId];
      if (!att) {
        return `<span class="inline-block p-2 rounded bg-rose-950/40 text-rose-400 text-xs border border-rose-800">⚠️ 첨부 이미지(${attId})를 찾을 수 없습니다.</span>`;
      }
      return `<img ${pre}src="${att.data}" alt="${att.name}" class="rounded-xl max-w-full my-3 border border-slate-800 shadow-md cursor-pointer hover:opacity-95 transition-opacity"${post} />`;
    }
  );

  // 2. <a href="attachment:att-xxx"> 치환 (PDF 뷰어 또는 다운로드 카드 렌더링)
  result = result.replace(
    /<a\s+[^>]*href=["']attachment:([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi,
    (_match, attId, labelText) => {
      const att = attachments[attId];
      if (!att) {
        return `<span class="inline-block p-2 rounded bg-rose-950/40 text-rose-400 text-xs border border-rose-800">⚠️ 첨부 파일(${attId})을 찾을 수 없습니다.</span>`;
      }

      const isPdf = att.type === 'application/pdf' || att.name.endsWith('.pdf') || att.data.startsWith('data:application/pdf');

      if (isPdf) {
        return `
<div class="my-4 rounded-2xl border border-slate-800 bg-slate-900/80 overflow-hidden shadow-2xl">
  <div class="px-4 py-3 bg-slate-900 border-b border-slate-800 flex items-center justify-between gap-3">
    <div class="flex items-center gap-2 text-xs font-bold text-slate-100 truncate">
      <span class="text-base">📄</span>
      <span class="truncate">${att.name}</span>
    </div>
    <a href="${att.data}" download="${att.name}" class="shrink-0 px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-semibold no-underline flex items-center gap-1.5 shadow-md shadow-indigo-600/20 active:scale-95 transition-all">
      <span>⬇️ 다운로드</span>
    </a>
  </div>
  <object data="${att.data}" type="application/pdf" width="100%" height="650px" class="block bg-slate-950">
    <div class="p-8 text-center text-xs text-slate-400 space-y-3">
      <p class="text-slate-300 font-semibold">브라우저에서 직접 PDF를 미리볼 수 없는 환경입니다.</p>
      <a href="${att.data}" download="${att.name}" class="inline-block px-4 py-2 bg-slate-800 hover:bg-slate-700 text-indigo-400 rounded-xl font-medium border border-slate-700">
        📄 ${att.name} 파일 내려받기
      </a>
    </div>
  </object>
</div>
        `.trim();
      }

      // PDF가 아닌 일반 첨부파일 카드
      return `
<div class="my-3 inline-flex items-center justify-between gap-4 p-3 rounded-xl border border-slate-800 bg-slate-900/80 hover:border-slate-700 transition-all max-w-md">
  <div class="flex items-center gap-2 truncate">
    <span class="text-lg">📎</span>
    <span class="text-xs font-medium text-slate-200 truncate">${att.name || labelText}</span>
  </div>
  <a href="${att.data}" download="${att.name}" class="shrink-0 px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-indigo-300 rounded-lg text-xs font-semibold no-underline">
    다운로드
  </a>
</div>
      `.trim();
    }
  );

  return result;
}
