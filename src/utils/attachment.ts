import type { AttachmentItem } from '../types';
import { parseExcelWorkbook } from './excelViewer';

function getColumnLabel(index: number): string {
  let label = '';
  let num = index;
  while (num >= 0) {
    label = String.fromCharCode((num % 26) + 65) + label;
    num = Math.floor(num / 26) - 1;
  }
  return label;
}

// Blob URL 캐시 (대용량 Base64 문자열 반복 파싱 부하 및 메모리 낭비 방지)
interface BlobCacheEntry {
  url: string;
  signature: string;
}
const blobUrlCache = new Map<string, BlobCacheEntry>();

export function invalidateAttachmentBlobUrl(attId?: string): void {
  if (attId) {
    const entry = blobUrlCache.get(attId);
    if (entry) {
      try {
        URL.revokeObjectURL(entry.url);
      } catch {
        // ignore
      }
      blobUrlCache.delete(attId);
    }
  } else {
    blobUrlCache.forEach((entry) => {
      try {
        URL.revokeObjectURL(entry.url);
      } catch {
        // ignore
      }
    });
    blobUrlCache.clear();
  }
}

export function getAttachmentBlobUrl(att: AttachmentItem): string {
  if (!att.data) return '';
  if (att.data.startsWith('blob:') || att.data.startsWith('http')) return att.data;
  
  const currentSig = `${att.data.length}_${att.data.slice(-32)}`;
  const cached = blobUrlCache.get(att.id);
  if (cached && cached.signature === currentSig) {
    return cached.url;
  }
  if (cached) {
    try {
      URL.revokeObjectURL(cached.url);
    } catch {
      // ignore
    }
    blobUrlCache.delete(att.id);
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
    blobUrlCache.set(att.id, { url, signature: currentSig });
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

  const isHwp =
    att.name.toLowerCase().endsWith('.hwp') ||
    att.name.toLowerCase().endsWith('.hwpx') ||
    att.type === 'application/x-hwp' ||
    att.type === 'application/haansofthwp' ||
    att.type === 'application/vnd.hancom.hwp' ||
    att.type === 'application/vnd.hancom.hwpx';

  const isExcel =
    att.name.toLowerCase().endsWith('.xlsx') ||
    att.name.toLowerCase().endsWith('.xls') ||
    att.name.toLowerCase().endsWith('.csv') ||
    att.type?.includes('spreadsheet') ||
    att.type === 'text/csv' ||
    att.data?.startsWith('data:application/vnd.openxmlformats') ||
    att.data?.startsWith('data:text/csv');

  // 사용자가 마크다운 [대괄호] 안에 수정한 표시 이름이 있으면 우선 사용
  const cleanLabel = (labelText || '')
    .replace(/^[\s📄📎🖼️📝📊]*/, '')
    .replace(/^(?:PDF\s*문서:\s*|한글\s*문서:\s*|한글:\s*|엑셀\s*문서:\s*|엑셀:\s*|파일:\s*)/i, '')
    .trim();
  const displayName = cleanLabel || att.name;
  const downloadName = isPdf && !displayName.toLowerCase().endsWith('.pdf')
    ? `${displayName}.pdf`
    : isHwp && !displayName.toLowerCase().endsWith('.hwp') && !displayName.toLowerCase().endsWith('.hwpx')
    ? `${displayName}.hwp`
    : displayName;
  const targetUrl = (isPdf || isHwp) ? getAttachmentBlobUrl(att) : att.data;

  if (isExcel) {
    const isCsv = att.name.toLowerCase().endsWith('.csv') || att.type === 'text/csv';
    const parsed = parseExcelWorkbook(att.data, att.id);
    const sheetNames = parsed.sheetNames || ['Sheet1'];
    const firstSheet = sheetNames[0] || 'Sheet1';
    const rows = parsed.sheets[firstSheet] || [];
    const totalRows = rows.length;
    let maxCols = 0;
    rows.forEach((r) => {
      maxCols = Math.max(maxCols, r.length);
    });
    const colCount = Math.max(maxCols, 1);
    const previewRows = rows.slice(0, 12);

    const sheetTabsHtml = sheetNames.length > 1
      ? `<div class="flex items-center gap-1 overflow-x-auto max-w-xs py-0.5">
          ${sheetNames.map((sName, sIdx) => `
            <button type="button" class="excel-sheet-pill px-2 py-0.5 text-[11px] font-semibold rounded ${sIdx === 0 ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40' : 'bg-slate-800 text-slate-400 hover:text-slate-200 border border-slate-700/60'}" data-att-id="${att.id}" data-sheet-name="${sName}">${sName}</button>
          `).join('')}
        </div>`
      : '';

    const tableHeaderHtml = `
      <thead class="bg-slate-900 sticky top-0 z-10 border-b border-slate-800">
        <tr>
          <th class="w-10 px-2 py-1.5 text-center text-[10px] font-bold text-slate-500 bg-slate-950 border-r border-slate-800 sticky left-0 z-20">#</th>
          ${Array.from({ length: colCount }).map((_, cIdx) => `
            <th class="px-2.5 py-1.5 text-[11px] font-mono font-semibold text-slate-400 border-r border-slate-800/80 min-w-[75px]">${getColumnLabel(cIdx)}</th>
          `).join('')}
        </tr>
      </thead>
    `;

    const tableBodyHtml = previewRows.map((row, rIdx) => `
      <tr class="hover:bg-indigo-950/20 transition-colors">
        <td class="px-2 py-1 text-center text-[10px] font-mono font-bold text-slate-500 bg-slate-950/90 border-r border-slate-800 sticky left-0 z-10">${rIdx + 1}</td>
        ${Array.from({ length: colCount }).map((_, cIdx) => {
          const val = row[cIdx] !== undefined && row[cIdx] !== null ? String(row[cIdx]) : '';
          return `<td class="px-2.5 py-1 text-xs text-slate-200 border-r border-slate-800/50 break-words whitespace-pre-wrap max-w-xs">${val}</td>`;
        }).join('')}
      </tr>
    `).join('');

    return `
<div class="my-5 rounded-2xl border border-slate-800 bg-slate-900/90 overflow-hidden shadow-2xl excel-inline-viewer" data-att-id="${att.id}" data-file-name="${displayName}">
  <!-- 상단 헤더: 문서명, 시트 탭, 직접 편집, 메모로 가져오기, 전면 뷰어, 다운로드 버튼 -->
  <div class="px-3.5 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-2.5">
    <div class="flex items-center gap-2 text-xs md:text-sm font-bold text-slate-100 min-w-0">
      <span class="text-base md:text-lg">📊</span>
      <span class="truncate" title="${displayName}">${displayName}</span>
      <span class="text-[10px] px-1.5 py-0.5 rounded bg-emerald-500/10 text-emerald-300 border border-emerald-500/20 font-bold shrink-0 uppercase">
        ${isCsv ? 'CSV' : 'XLSX'}
      </span>
      ${sheetTabsHtml}
    </div>

    <div class="flex items-center gap-1.5 shrink-0">
      <!-- 스프레드시트 직접 편집 버튼 -->
      <button type="button" class="excel-btn-edit px-2.5 py-1 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 rounded-lg text-xs font-semibold border border-emerald-500/30 flex items-center gap-1 active:scale-95 transition-all cursor-pointer" data-att-id="${att.id}" data-file-name="${displayName}" title="웹에서 표 데이터 직접 편집 및 행/열 추가·삭제">
        <span>✏️ 직접 편집</span>
      </button>

      <!-- 본문 메모로 가져오기 버튼 -->
      <button type="button" class="excel-btn-import px-2.5 py-1 bg-indigo-600/30 hover:bg-indigo-600 text-indigo-200 hover:text-white rounded-lg text-xs font-semibold border border-indigo-500/40 flex items-center gap-1 active:scale-95 transition-all cursor-pointer" data-att-id="${att.id}" data-file-name="${displayName}" title="표 데이터를 마크다운 표로 메모 본문에 가져오기">
        <span>📝 표 가져오기</span>
      </button>

      <!-- 전면 뷰어로 크게 보기 버튼 -->
      <button type="button" class="excel-btn-fullscreen px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-indigo-300 rounded-lg text-xs font-semibold border border-indigo-500/30 flex items-center gap-1 active:scale-95 transition-all cursor-pointer" data-att-id="${att.id}" data-file-name="${displayName}" title="전체 행/열 전면 뷰어로 크게 보기">
        <span>🔍 크게 보기</span>
      </button>

      <!-- 다운로드 버튼 -->
      <a href="${att.data}" download="${displayName}" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg text-xs font-semibold border border-slate-700 flex items-center gap-1 active:scale-95 transition-all no-underline" title="기기에 파일로 다운로드">
        <span>⬇️ 저장</span>
      </a>
    </div>
  </div>

  <!-- 인라인 표 그리드 뷰어 -->
  <div class="w-full bg-slate-950 p-2 md:p-3 overflow-x-auto max-h-[420px]">
    <div class="inline-block min-w-full align-middle border border-slate-800 rounded-xl overflow-hidden bg-slate-900/60 shadow-inner">
      <table class="min-w-full border-collapse text-left" id="excel-table-${att.id}">
        ${tableHeaderHtml}
        <tbody class="divide-y divide-slate-800/60 font-sans">
          ${tableBodyHtml}
        </tbody>
      </table>
    </div>
  </div>

  <!-- 하단 행 카운트 안내 바 -->
  <div class="px-3.5 py-1.5 bg-slate-900 border-t border-slate-800 flex items-center justify-between text-[11px] text-slate-400">
    <span>전체 ${totalRows}개 행 • ${colCount}개 열 ${totalRows > 12 ? `(상위 12개 행 미리보기)` : ''}</span>
    <span class="text-slate-500 hidden sm:inline">전체 확인 및 셀 수정은 [🔍 크게 보기] 또는 [✏️ 직접 편집]을 클릭하세요.</span>
  </div>
</div>
    `.trim();
  }

  if (isHwp) {
    return `
<div class="my-5 rounded-2xl border border-slate-800 bg-slate-900/90 overflow-hidden shadow-2xl hwp-inline-viewer" data-att-id="${att.id}" data-file-name="${displayName}">
  <!-- 상단 헤더: 문서명, 페이지 이동, 메모로 가져오기, 전면 뷰어, 다운로드 버튼 -->
  <div class="px-3.5 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-2.5">
    <div class="flex items-center gap-2 text-xs md:text-sm font-bold text-slate-100 min-w-0">
      <span class="text-base md:text-lg">📝</span>
      <span class="truncate" title="${displayName}">${displayName}</span>
      <span class="text-[10px] px-1.5 py-0.5 rounded bg-blue-500/10 text-blue-300 border border-blue-500/20 font-bold shrink-0">
        HWP
      </span>
    </div>

    <div class="flex items-center gap-1.5 shrink-0">
      <!-- 페이지 이동 컨트롤 -->
      <div class="flex items-center gap-1 bg-slate-950 border border-slate-800 rounded-lg px-2 py-0.5 text-xs text-slate-300">
        <button type="button" class="hwp-btn-prev px-1.5 py-0.5 hover:text-white hover:bg-slate-800 rounded transition-colors disabled:opacity-30 disabled:hover:bg-transparent" data-att-id="${att.id}" title="이전 페이지">◀</button>
        <span class="hwp-page-indicator font-semibold text-xs px-1" id="hwp-page-indicator-${att.id}">1 / ?</span>
        <button type="button" class="hwp-btn-next px-1.5 py-0.5 hover:text-white hover:bg-slate-800 rounded transition-colors disabled:opacity-30 disabled:hover:bg-transparent" data-att-id="${att.id}" title="다음 페이지">▶</button>
      </div>

      <!-- 한글 직접 편집 버튼 -->
      <button type="button" class="hwp-btn-edit px-2.5 py-1 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 rounded-lg text-xs font-semibold border border-emerald-500/30 flex items-center gap-1 active:scale-95 transition-all" data-att-id="${att.id}" data-file-name="${displayName}" title="한글 문서 원본을 웹에서 직접 수정/편집">
        <span>✏️ 직접 편집</span>
      </button>

      <!-- 본문 메모로 가져오기 버튼 -->
      <button type="button" class="hwp-btn-import px-2.5 py-1 bg-indigo-600/30 hover:bg-indigo-600 text-indigo-200 hover:text-white rounded-lg text-xs font-semibold border border-indigo-500/40 flex items-center gap-1 active:scale-95 transition-all" data-att-id="${att.id}" data-file-name="${displayName}" title="한글 문서 내용을 마크다운 본문으로 가져오기">
        <span>📝 가져오기</span>
      </button>

      <!-- 전면 뷰어로 크게 보기 버튼 -->
      <button type="button" class="hwp-btn-fullscreen px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-indigo-300 rounded-lg text-xs font-semibold border border-indigo-500/30 flex items-center gap-1 active:scale-95 transition-all" data-att-id="${att.id}" data-file-name="${displayName}" title="전면 뷰어로 크게 보기">
        <span>🔍 크게 보기</span>
      </button>

      <!-- 다운로드 버튼 -->
      <a href="${targetUrl}" download="${downloadName}" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg text-xs font-semibold border border-slate-700 flex items-center gap-1 active:scale-95 transition-all no-underline" title="기기에 파일로 다운로드">
        <span>⬇️ 저장</span>
      </a>
    </div>
  </div>

  <!-- 본문 인라인 HWP SVG 뷰어 (스마트폰 및 PC 브라우저에서 화면에 즉시 렌더링!) -->
  <div class="w-full bg-slate-950 p-2 md:p-4 flex flex-col items-center justify-center overflow-x-auto min-h-[300px]">
    <div class="hwp-loading-box flex items-center justify-center gap-2 py-10 text-xs text-slate-400" id="hwp-loading-${att.id}">
      <span class="animate-spin inline-block w-4 h-4 border-2 border-indigo-500 border-t-transparent rounded-full"></span>
      <span>한글(HWP) 문서 페이지를 렌더링하는 중입니다...</span>
    </div>
    <div id="hwp-svg-container-${att.id}" class="hwp-svg-container rounded-lg shadow-xl bg-white p-2 md:p-4 text-slate-900 max-w-full cursor-zoom-in hidden" data-att-id="${att.id}" data-file-name="${displayName}" title="클릭하면 전면 뷰어로 크게 볼 수 있습니다."></div>
  </div>
</div>
    `.trim();
  }

  if (isPdf) {
    return `
<div class="my-5 rounded-2xl border border-slate-800 bg-slate-900/90 overflow-hidden shadow-2xl pdf-inline-viewer" data-att-id="${att.id}" data-file-name="${displayName}">
  <!-- 상단 헤더: 문서명, 페이지 이동, 전면 뷰어, 다운로드 버튼 -->
  <div class="px-3.5 py-2.5 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-2.5">
    <div class="flex items-center gap-2 text-xs md:text-sm font-bold text-slate-100 min-w-0">
      <span class="text-base md:text-lg">📄</span>
      <span class="truncate" title="${displayName}">${displayName}</span>
    </div>

    <div class="flex items-center gap-1.5 shrink-0">
      <!-- 페이지 이동 컨트롤 -->
      <div class="flex items-center gap-1 bg-slate-950 border border-slate-800 rounded-lg px-2 py-0.5 text-xs text-slate-300">
        <button type="button" class="pdf-btn-prev px-1.5 py-0.5 hover:text-white hover:bg-slate-800 rounded transition-colors disabled:opacity-30 disabled:hover:bg-transparent" data-att-id="${att.id}" title="이전 페이지">◀</button>
        <span class="pdf-page-indicator font-semibold text-xs px-1" id="pdf-page-indicator-${att.id}">1 / ?</span>
        <button type="button" class="pdf-btn-next px-1.5 py-0.5 hover:text-white hover:bg-slate-800 rounded transition-colors disabled:opacity-30 disabled:hover:bg-transparent" data-att-id="${att.id}" title="다음 페이지">▶</button>
      </div>

      <!-- PDF 직접 주석/서명 편집 버튼 -->
      <button type="button" class="pdf-btn-edit px-2.5 py-1 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 rounded-lg text-xs font-semibold border border-emerald-500/30 flex items-center gap-1 active:scale-95 transition-all" data-att-id="${att.id}" data-file-name="${displayName}" title="PDF 위에 텍스트 입력, 형광펜, 펜 드로잉, 서명/도장 편집">
        <span>✏️ 주석/서명</span>
      </button>

      <!-- 본문 메모로 가져오기 버튼 -->
      <button type="button" class="pdf-btn-import px-2.5 py-1 bg-indigo-600/30 hover:bg-indigo-600 text-indigo-200 hover:text-white rounded-lg text-xs font-semibold border border-indigo-500/40 flex items-center gap-1 active:scale-95 transition-all" data-att-id="${att.id}" data-file-name="${displayName}" title="PDF 문서의 텍스트를 마크다운 본문으로 가져오기">
        <span>📝 메모로 가져오기</span>
      </button>

      <!-- 전면 뷰어로 크게 보기 버튼 -->
      <button type="button" class="pdf-btn-fullscreen px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-indigo-300 rounded-lg text-xs font-semibold border border-indigo-500/30 flex items-center gap-1 active:scale-95 transition-all" data-att-id="${att.id}" data-file-name="${displayName}" title="전면 뷰어로 크게 보기">
        <span>🔍 크게 보기</span>
      </button>

      <!-- 다운로드 버튼 (명시적으로 저장을 원하는 경우에만) -->
      <a href="${targetUrl}" download="${downloadName}" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white rounded-lg text-xs font-semibold border border-slate-700 flex items-center gap-1 active:scale-95 transition-all no-underline" title="기기에 파일로 다운로드">
        <span>⬇️ 저장</span>
      </a>
    </div>
  </div>

  <!-- 본문 인라인 PDF 캔버스 뷰어 (모바일 스마트폰 및 모든 PC 브라우저에서 다운로드 없이 화면에 즉시 렌더링!) -->
  <div class="w-full bg-slate-950 p-2 md:p-4 flex flex-col items-center justify-center overflow-x-auto min-h-[350px]">
    <div class="pdf-loading-box flex items-center justify-center gap-2 py-10 text-xs text-slate-400" id="pdf-loading-${att.id}">
      <span class="animate-spin inline-block w-4 h-4 border-2 border-indigo-500 border-t-transparent rounded-full"></span>
      <span>PDF 문서 페이지를 렌더링하는 중입니다...</span>
    </div>
    <canvas id="pdf-canvas-${att.id}" class="pdf-canvas rounded-lg shadow-xl bg-white max-w-full cursor-zoom-in hidden" data-att-id="${att.id}" data-file-name="${displayName}" title="클릭하면 전면 뷰어로 크게 볼 수 있습니다."></canvas>
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
      if (mimeType === 'application/pdf' || name.toLowerCase().endsWith('.pdf')) {
        return `\n\n[📄 PDF: ${safeName}](attachment:${existingId})\n\n`;
      }
      if (name.toLowerCase().endsWith('.hwp') || name.toLowerCase().endsWith('.hwpx') || mimeType.includes('hwp')) {
        return `\n\n[📝 한글: ${safeName}](attachment:${existingId})\n\n`;
      }
      if (
        name.toLowerCase().endsWith('.xlsx') ||
        name.toLowerCase().endsWith('.xls') ||
        name.toLowerCase().endsWith('.csv') ||
        mimeType.includes('spreadsheet') ||
        mimeType === 'text/csv'
      ) {
        return `\n\n[📊 엑셀: ${safeName}](attachment:${existingId})\n\n`;
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

      return `<span class="relative inline-block group my-3 note-image-wrapper align-middle" style="${widthStyle || 'max-width: 100%;'}">
        <img ${pre}src="${att.data}" alt="${cleanAlt}" style="width: 100%; height: auto; display: block;" class="rounded-xl border border-slate-800 shadow-md cursor-zoom-in hover:opacity-95 transition-opacity note-image" data-att-id="${attId}"${post} />
        <button type="button" class="img-btn-edit absolute top-2 right-2 bg-slate-900/90 hover:bg-indigo-600 text-slate-200 hover:text-white px-2 py-1 rounded-lg text-xs font-bold shadow-lg border border-slate-700/80 flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity backdrop-blur-xs cursor-pointer" data-att-id="${attId}" data-file-name="${cleanAlt}" title="이미지 편집 (그리기, 자르기, 영역 이동, 모자이크, 텍스트)">
          ✏️ 편집
        </button>
      </span>`;
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
