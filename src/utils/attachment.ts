import type { AttachmentItem } from '../types';

// Blob URL 캐시 (대용량 Base64 문자열 반복 파싱 부하 및 메모리 낭비 방지)
const blobUrlCache = new Map<string, string>();

export function getAttachmentBlobUrl(att: AttachmentItem): string {
  if (!att.data) return '';
  if (att.data.startsWith('blob:') || att.data.startsWith('http')) return att.data;
  
  if (blobUrlCache.has(att.id)) {
    return blobUrlCache.get(att.id)!;
  }

  try {
    const commaIdx = att.data.indexOf(',');
    if (commaIdx === -1) return att.data;
    const header = att.data.substring(0, commaIdx);
    const base64 = att.data.substring(commaIdx + 1);
    const mimeMatch = header.match(/:(.*?);/);
    const mime = mimeMatch ? mimeMatch[1] : att.type || 'application/octet-stream';
    const binary = atob(base64);
    const array = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      array[i] = binary.charCodeAt(i);
    }
    const blob = new Blob([array], { type: mime });
    const url = URL.createObjectURL(blob);
    blobUrlCache.set(att.id, url);
    return url;
  } catch {
    return att.data;
  }
}

/**
 * 첨부파일 객체를 HTML 마크업으로 변환합니다.
 * PDF인 경우: 헤더(제목 + 열기/다운로드 버튼) + 반응형 PDF 뷰어(데스크톱 인라인 뷰어 / 모바일 다이렉트 뷰어)
 * 이미지인 경우: <img> 태그
 * 일반 파일인 경우: 다운로드 카드
 */
export function createAttachmentCardHtml(att: AttachmentItem, labelText: string = ''): string {
  const isPdf =
    att.type === 'application/pdf' ||
    att.name.toLowerCase().endsWith('.pdf') ||
    att.data.startsWith('data:application/pdf');

  // 사용자가 마크다운 [대괄호] 안에 수정한 표시 이름이 있으면 우선 사용
  const cleanLabel = (labelText || '')
    .replace(/^[\s📄📎🖼️]*/, '')
    .replace(/^(?:PDF\s*문서:\s*|파일:\s*)/i, '')
    .trim();
  const displayName = cleanLabel || att.name;
  const downloadName = isPdf && !displayName.toLowerCase().endsWith('.pdf') ? `${displayName}.pdf` : displayName;
  const targetUrl = isPdf ? getAttachmentBlobUrl(att) : att.data;

  if (isPdf) {
    return `
<div class="my-5 rounded-2xl border border-slate-800 bg-slate-900/90 overflow-hidden shadow-2xl">
  <!-- 상단 헤더: 문서명과 액션 버튼 -->
  <div class="px-4 py-3 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3">
    <div class="flex items-center gap-2 text-xs md:text-sm font-bold text-slate-100 min-w-0">
      <span class="text-base md:text-lg">📄</span>
      <span class="truncate">${displayName}</span>
    </div>
    <div class="flex items-center gap-2 shrink-0">
      <a href="${targetUrl}" download="${downloadName}" class="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-semibold no-underline flex items-center gap-1.5 shadow-md shadow-indigo-600/20 active:scale-95 transition-all">
        <span>⬇️ PDF 열기 / 다운로드</span>
      </a>
    </div>
  </div>
  <!-- 본문 인라인 PDF 뷰어 (PC 데스크톱 브라우저는 바로 임베드 표시, 모바일은 기기 뷰어 열기 카드 표시) -->
  <div class="w-full bg-slate-950">
    <object data="${targetUrl}" type="application/pdf" width="100%" class="w-full h-[650px] md:h-[750px] block border-0 bg-slate-950">
      <div class="p-8 text-center text-xs md:text-sm text-slate-400 space-y-3 bg-slate-950">
        <p class="text-slate-200 font-semibold text-sm">모바일 환경에서는 기기 PDF 앱으로 문서를 바로 열람할 수 있습니다.</p>
        <p class="text-xs text-slate-400">아래 버튼을 누르면 기기에서 즉시 문서를 열람하거나 다운로드합니다.</p>
        <a href="${targetUrl}" download="${downloadName}" class="inline-flex items-center gap-2 px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-semibold shadow-lg shadow-indigo-600/30">
          <span>📄 ${displayName} 바로 열기 / 다운로드</span>
        </a>
      </div>
    </object>
  </div>
</div>
    `.trim();
  }

  const isImage =
    att.type?.startsWith('image/') ||
    att.data?.startsWith('data:image/');

  if (isImage) {
    return `<img src="${att.data}" alt="${cleanLabel || att.name}" class="rounded-xl max-w-full my-3 border border-slate-800 shadow-md cursor-pointer hover:opacity-95 transition-opacity" />`;
  }

  // 일반 파일 카드
  return `
<div class="my-3 inline-flex items-center justify-between gap-4 p-3 rounded-xl border border-slate-800 bg-slate-900/80 hover:border-slate-700 transition-all max-w-md">
  <div class="flex items-center gap-2 truncate">
    <span class="text-lg">📎</span>
    <span class="text-xs font-medium text-slate-200 truncate">${displayName}</span>
  </div>
  <a href="${att.data}" download="${displayName}" class="shrink-0 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-indigo-300 rounded-lg text-xs font-semibold no-underline">
    다운로드
  </a>
</div>
  `.trim();
}

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

      const safeName = name.replace(/`/g, "'").replace(/[\[\]]/g, '').trim();
      return `\n\n[📄 PDF: ${safeName}](attachment:${existingId})\n\n`;
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

      const safeName = name.replace(/`/g, "'").replace(/[\[\]]/g, '').trim();
      return `\n\n![${safeName}](attachment:${existingId})\n\n`;
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

      const safeName = name.replace(/`/g, "'").replace(/[\[\]]/g, '').trim();
      if (mimeType === 'application/pdf') {
        return `\n\n[📄 PDF: ${safeName}](attachment:${existingId})\n\n`;
      }
      return `\n\n[📎 파일: ${safeName}](attachment:${existingId})\n\n`;
    }
  );

  // 4. 이미 생성된 attachment 태그 중에서 백틱(`)이 포함되어 마크다운 파싱이 깨지는 태그 정돈
  cleaned = cleaned.replace(
    /\[([^\]]*`[^\]]*)\]\((attachment:[a-zA-Z0-9_-]+)\)/g,
    (_match, label, attRef) => {
      hasChanges = true;
      return `[${label.replace(/`/g, "'")}](${attRef})`;
    }
  );

  return {
    content: cleaned,
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

      // alt 속성에 지정된 크기 옵션 파싱 (예: ![사진|작게], ![사진|보통], ![사진|300], ![사진|50%])
      let widthStyle = '';
      const fullTag = `${pre} ${post}`;
      const altMatch = fullTag.match(/alt=["']([^"']*)["']/i);
      let cleanAlt = att.name;

      if (altMatch) {
        const rawAlt = altMatch[1];
        const pipeIdx = rawAlt.indexOf('|');
        if (pipeIdx !== -1) {
          cleanAlt = rawAlt.slice(0, pipeIdx).trim() || att.name;
          const sizeSpec = rawAlt.slice(pipeIdx + 1).trim().toLowerCase();
          if (sizeSpec === '작게' || sizeSpec === 'small' || sizeSpec === 's') {
            widthStyle = 'max-width: 180px; max-height: 260px; width: auto; height: auto; display: block;';
          } else if (sizeSpec === '보통' || sizeSpec === '중간' || sizeSpec === 'medium' || sizeSpec === 'm') {
            widthStyle = 'max-width: 340px; max-height: 480px; width: auto; height: auto; display: block;';
          } else if (sizeSpec === '크게' || sizeSpec === 'large' || sizeSpec === 'l') {
            widthStyle = 'max-width: 560px; max-height: 750px; width: auto; height: auto; display: block;';
          } else if (sizeSpec === '전체' || sizeSpec === 'full' || sizeSpec === '100%') {
            widthStyle = 'max-width: 100%; max-height: none; width: auto; height: auto; display: block;';
          } else if (/^\d+(px)?$/.test(sizeSpec)) {
            const pxVal = parseInt(sizeSpec.replace('px', ''), 10);
            widthStyle = `max-width: ${pxVal}px; max-height: ${Math.round(pxVal * 1.6)}px; width: auto; height: auto; display: block;`;
          } else if (/^\d+%$/.test(sizeSpec)) {
            widthStyle = `max-width: ${sizeSpec}; max-height: none; width: auto; height: auto; display: block;`;
          }
        }
      }

      return `<img ${pre}src="${att.data}" alt="${cleanAlt}" style="${widthStyle}" class="rounded-xl my-3 border border-slate-800 shadow-md cursor-zoom-in hover:opacity-95 transition-opacity note-image" data-att-id="${attId}"${post} />`;
    }
  );

  // 2. <a href="attachment:att-xxx"> 치환
  result = result.replace(
    /<a\s+[^>]*href=["']attachment:([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi,
    (_match, attId, labelText) => {
      const att = attachments[attId];
      if (!att) {
        return `<span class="inline-block p-2 rounded bg-rose-950/40 text-rose-400 text-xs border border-rose-800">⚠️ 첨부 파일(${attId})을 찾을 수 없습니다.</span>`;
      }
      return createAttachmentCardHtml(att, labelText);
    }
  );

  // 3. marked가 파일명의 백틱(`), 괄호 등으로 인해 <a> 태그로 파싱하지 못하고 텍스트로 남긴 경우
  // 예: <p>[📄 PDF: ...](attachment:att-xxx)</p>
  result = result.replace(
    /(?:<p>)?\s*!?\[([\s\S]*?)\]\((?:attachment:)?(att-[a-zA-Z0-9_-]+)\)\s*(?:<\/p>)?/gi,
    (_match, labelText, attId) => {
      const att = attachments[attId];
      if (!att) {
        return `<span class="inline-block p-2 rounded bg-rose-950/40 text-rose-400 text-xs border border-rose-800">⚠️ 첨부 파일(${attId})을 찾을 수 없습니다.</span>`;
      }
      return createAttachmentCardHtml(att, labelText);
    }
  );

  return result;
}
