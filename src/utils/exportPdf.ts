import type { Note, AttachmentItem } from '../types';
import { marked } from 'marked';
import { loadPdfDocument } from './pdfViewer';
import { loadHwpDocument, renderHwpPageToSvg } from './hwpViewer';
import { parseExcelWorkbookWithStyles, renderExcelTableHtml } from './excelViewer';
import { renderDocxToContainer } from './docxViewer';
import { extractPptxTextAndMarkdown } from './pptxViewer';

/**
 * PDF 내보내기 시 첨부 문서(PDF/한글/엑셀/워드/PPT) 내용 포함 방식
 * - expanded: 펼친(접기 해제) 첨부만 내용 포함, 접힌 첨부는 이름 한 줄 (기본)
 * - all: 접힘 상태와 관계없이 모두 내용 포함
 * - none: 모든 첨부 문서를 이름 한 줄로만 표시 (이미지는 본문의 일부이므로 항상 표시)
 */
export type PdfAttachmentMode = 'expanded' | 'all' | 'none';

const EXCEL_PRINT_MAX_ROWS = 200;

const HTML_ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
function escapeHtml(s: string): string {
  return (s || '').replace(/[&<>"']/g, (c) => HTML_ESCAPES[c] || c);
}

type AttKind = 'pdf' | 'hwp' | 'excel' | 'word' | 'ppt' | 'other';

function getAttachmentKind(att: AttachmentItem): AttKind {
  const name = (att.name || '').toLowerCase();
  const type = att.type || '';
  if (type === 'application/pdf' || name.endsWith('.pdf') || (att.data || '').startsWith('data:application/pdf')) return 'pdf';
  if (
    name.endsWith('.hwp') || name.endsWith('.hwpx') ||
    ['application/x-hwp', 'application/haansofthwp', 'application/vnd.hancom.hwp', 'application/vnd.hancom.hwpx'].includes(type)
  ) return 'hwp';
  if (name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv') || /spreadsheetml|ms-excel|text\/csv/.test(type)) return 'excel';
  if (name.endsWith('.docx') || /wordprocessingml/.test(type)) return 'word';
  if (name.endsWith('.pptx') || /presentationml/.test(type)) return 'ppt';
  return 'other';
}

/** 내용을 PDF에 넣을 수 있는 문서형 첨부인지 (내보내기 선택 창을 띄울지 판단에 사용) */
export function isPrintableDocumentAttachment(att: AttachmentItem | undefined): boolean {
  if (!att || !att.data) return false;
  return getAttachmentKind(att) !== 'other';
}

/** 본문에서 참조하는 문서형 첨부가 있는지 */
export function contentHasDocumentAttachments(content: string, attachments: Record<string, AttachmentItem> = {}): boolean {
  const ids = new Set<string>();
  (content || '').replace(/(att-[a-zA-Z0-9_-]+)/g, (m) => {
    ids.add(m);
    return m;
  });
  for (const id of ids) {
    if (isPrintableDocumentAttachment(attachments[id])) return true;
  }
  return false;
}

const KIND_ICON: Record<AttKind, string> = { pdf: '📄', hwp: '📝', excel: '📊', word: '📘', ppt: '📽️', other: '📎' };
const KIND_LABEL: Record<AttKind, string> = { pdf: 'PDF 문서', hwp: '한글 문서', excel: '엑셀', word: '워드 문서', ppt: 'PPT', other: '첨부 파일' };

function attachmentCardHtml(att: AttachmentItem, displayName: string, note?: string): string {
  const kind = getAttachmentKind(att);
  return `
    <div class="print-att-card">
      <span>${KIND_ICON[kind]}</span>
      <span>${escapeHtml(displayName)} (${KIND_LABEL[kind]}${note ? ` · ${note}` : ''})</span>
    </div>
  `;
}

/** 미리보기 화면의 토글(<details>) 펼침 상태. 제목(summary 글자)과 순서로 마크다운 블록과 짝을 맞춘다. */
export interface PdfToggleOpenState {
  summary: string;
  open: boolean;
}

function normalizeSummaryText(s: string): string {
  return (s || '')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/[*_`~]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * 마크다운의 토글(<details>) 블록을 내보내기 방식과 화면 펼침 상태에 맞게 변환한다.
 * 펼침 판정: all → 항상 펼침 / expanded·none → 미리보기 화면에서 펼쳐 둔 것만 펼침
 *   (미리보기가 없는 편집 전용 화면이면 모두 접힌 것으로 본다)
 * - 문서형 첨부가 든 블록('📁 접기'): 펼침이면 토글을 풀어 첨부 내용을 넣고, 접힘이면 '접힌 첨부' 한 줄 카드
 *   (none 모드에서는 풀어도 첨부가 이름 한 줄로 표시됨)
 * - 글만 있는 토글: 펼침이면 <details open> 으로 내용까지, 접힘이면 제목 한 줄만
 */
function applyDetailsMode(
  markdown: string,
  attachments: Record<string, AttachmentItem>,
  mode: PdfAttachmentMode,
  openStates?: PdfToggleOpenState[]
): string {
  const used = new Set<number>();
  const isOpenOnScreen = (summaryRaw: string): boolean => {
    if (!openStates || openStates.length === 0) return false;
    const key = normalizeSummaryText(summaryRaw);
    for (let i = 0; i < openStates.length; i++) {
      if (used.has(i)) continue;
      if (normalizeSummaryText(openStates[i].summary) === key) {
        used.add(i);
        return openStates[i].open;
      }
    }
    return false;
  };

  return (markdown || '').replace(
    /<details([^>]*)>(\s*)<summary>([\s\S]*?)<\/summary>([\s\S]*?)<\/details>/gi,
    (block: string, attrs: string, gap: string, summary: string, inner: string) => {
      const open = mode === 'all' ? true : isOpenOnScreen(summary);

      const ids: string[] = [];
      inner.replace(/(?:\(|attachment:)(att-[a-zA-Z0-9_-]+)/g, (m: string, id: string) => {
        if (!ids.includes(id)) ids.push(id);
        return m;
      });
      const docAtts = ids.map((id) => attachments[id]).filter((a): a is AttachmentItem => isPrintableDocumentAttachment(a));

      if (docAtts.length > 0) {
        if (open) {
          return `\n\n${inner.trim()}\n\n`;
        }
        const cards = docAtts.map((a) => attachmentCardHtml(a, a.name, '접힌 첨부')).join('');
        return `\n\n<div class="print-collapsed-group">${cards}</div>\n\n`;
      }

      // 글 토글: 열림 여부만 바꾼다 (내부 마크다운 해석은 미리보기와 동일하게 유지)
      const cleanAttrs = attrs.replace(/\s+open(?:=(?:"[^"]*"|'[^']*'|\S+))?/gi, '');
      if (open) {
        return block.replace(/^<details[^>]*>/i, `<details${cleanAttrs} open>`);
      }
      return `<details${cleanAttrs}>${gap}<summary>${summary}</summary>${inner}</details>`;
    }
  );
}

/**
 * SVG 문자열에서 viewBox 또는 width/height를 파싱하여 가로(Landscape) 여부를 판별합니다.
 */
function isSvgLandscape(svgStr: string): boolean {
  if (!svgStr) return false;
  const vbMatch = svgStr.match(/viewBox=["']\s*([-\d.]+)\s+([-\d.]+)\s+([\d.]+)\s+([\d.]+)\s*["']/i);
  if (vbMatch) {
    const w = parseFloat(vbMatch[3]);
    const h = parseFloat(vbMatch[4]);
    return w > h;
  }
  const wMatch = svgStr.match(/width=["']([\d.]+)p?[xt]?["']/i);
  const hMatch = svgStr.match(/height=["']([\d.]+)p?[xt]?["']/i);
  if (wMatch && hMatch) {
    return parseFloat(wMatch[1]) > parseFloat(hMatch[1]);
  }
  return false;
}

/**
 * 가로(Landscape, width > height)로 작성된 SVG 문서를 90도 시계방향으로 눕혀
 * 세로 A4 규격(viewBox: 0 0 H W)에 맞추고 꽉 차게 표시하도록 변환합니다.
 */
function rotateSvg90Degrees(svgStr: string): string {
  const vbMatch = svgStr.match(/viewBox=["']\s*([-\d.]+)\s+([-\d.]+)\s+([\d.]+)\s+([\d.]+)\s*["']/i);
  if (!vbMatch) return svgStr;

  const w = parseFloat(vbMatch[3]);
  const h = parseFloat(vbMatch[4]);
  if (w <= h) return svgStr; // 이미 세로 문서이면 회전하지 않음

  const openTagEnd = svgStr.indexOf('>');
  if (openTagEnd === -1) return svgStr;

  const closeTagStart = svgStr.lastIndexOf('</svg>');
  if (closeTagStart === -1) return svgStr;

  let openTag = svgStr.slice(0, openTagEnd + 1);
  const innerContent = svgStr.slice(openTagEnd + 1, closeTagStart);

  // viewBox를 세로 비율 (0 0 H W)로 치환하고 고정 너비/높이는 제거하여 반응형으로 맞춤
  openTag = openTag.replace(/viewBox=["'][^"']*["']/i, `viewBox="0 0 ${h} ${w}"`);
  openTag = openTag.replace(/\s+(?:width|height)=["'][^"']*["']/gi, '');

  return `${openTag}<g transform="rotate(90) translate(0, -${h})">${innerContent}</g></svg>`;
}

/**
 * 인쇄/PDF 내보내기용 첨부파일 HTML 비동기 변환기
 * PDF 및 한글(HWP) 첨부 문서를 실제 고화질 페이지 이미지/SVG로 렌더링하여 인쇄 양식에 온전히 포함시킵니다.
 * 가로(Landscape) 설정 문서는 자동으로 가로 방향 페이지(@page landscape)로 눕혀 꽉 차게 렌더링합니다.
 */
async function renderAttachmentsForPrint(
  html: string,
  attachments: Record<string, AttachmentItem> = {},
  includeContent: boolean = true
): Promise<string> {
  if (!html || Object.keys(attachments).length === 0) return html;

  let result = html;

  // 1. 일반 이미지 (<img src="attachment:att-xxx">)
  result = result.replace(
    /<img\s+([^>]*?)src=["']attachment:([^"']+)["']([^>]*?)>/gi,
    (_match, pre, attId, post) => {
      const att = attachments[attId];
      if (!att || !att.data) return '';
      return `<img ${pre}src="${att.data}" style="max-width: 100%; height: auto; display: block; margin: 12px auto; border-radius: 6px; border: 1px solid #e2e8f0;" ${post} />`;
    }
  );

  // 2. 문서 및 링크: <a href="attachment:att-xxx"> 또는 [📄 PDF: ...](attachment:att-xxx)
  const attLinks: Array<{ fullMatch: string; attId: string; label: string }> = [];

  // 패턴 A: <a href="attachment:att-xxx">라벨</a>
  result.replace(
    /<a\s+[^>]*href=["']attachment:([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi,
    (fullMatch, attId, label) => {
      attLinks.push({ fullMatch, attId, label });
      return fullMatch;
    }
  );

  // 패턴 B: [라벨](attachment:att-xxx) (p 태그 여부 무관)
  result.replace(
    /(?:<p>)?\s*!?\[([\s\S]*?)\]\((?:attachment:)?(att-[a-zA-Z0-9_-]+)\)\s*(?:<\/p>)?/gi,
    (fullMatch, label, attId) => {
      attLinks.push({ fullMatch, attId, label });
      return fullMatch;
    }
  );

  for (const { fullMatch, attId, label } of attLinks) {
    const att = attachments[attId];
    if (!att || !att.data) continue;

    // 아이콘(이모지)과 '엑셀:' 같은 접두어 제거. u 플래그로 이모지를 통째로 지워 '�' 깨짐 방지
    const cleanLabel = (label || '')
      .replace(/<[^>]+>/g, '')
      .replace(/^[\s\p{Extended_Pictographic}\uFE0F\u200D]*/u, '')
      .replace(/^(?:PDF\s*문서|PDF|한글\s*문서|한글|엑셀|워드|PPT|파워포인트|파일)\s*:\s*/i, '')
      .trim();
    const displayName = cleanLabel || att.name;

    const kind = getAttachmentKind(att);
    const isPdf = kind === 'pdf';
    const isHwp = kind === 'hwp';

    // '첨부 내용 빼기' 모드: 문서형 첨부는 이름 한 줄로만
    if (!includeContent && kind !== 'other') {
      result = result.split(fullMatch).join(attachmentCardHtml(att, displayName));
      continue;
    }

    if (kind === 'excel') {
      try {
        const parsed = await parseExcelWorkbookWithStyles(att.data, att.id);
        let sheetsHtml = '';
        for (const sName of parsed.sheetNames || []) {
          const { tableHeaderHtml, tableBodyHtml, totalRows } = renderExcelTableHtml(parsed, sName, EXCEL_PRINT_MAX_ROWS);
          if (!totalRows) continue;
          const cutNote = totalRows > EXCEL_PRINT_MAX_ROWS ? ` (전체 ${totalRows}행 중 앞 ${EXCEL_PRINT_MAX_ROWS}행)` : '';
          sheetsHtml += `
            <div class="print-excel-sheet">
              <div class="print-excel-sheet-name">시트: ${escapeHtml(sName)}${cutNote}</div>
              <table class="print-excel-table">${tableHeaderHtml}<tbody>${tableBodyHtml}</tbody></table>
            </div>
          `;
        }
        const replacement = `
          <div class="print-doc-container print-doc-flow">
            <div class="print-doc-header">
              <span class="print-doc-badge print-doc-badge-excel">📊 엑셀</span>
              <span class="print-doc-title">${escapeHtml(displayName)} (시트 ${(parsed.sheetNames || []).length}개)</span>
            </div>
            ${sheetsHtml || '<div class="print-empty">(내용이 없는 통합 문서)</div>'}
          </div>
        `;
        result = result.split(fullMatch).join(replacement);
      } catch (err) {
        console.error('엑셀 인쇄용 렌더링 실패:', err);
        result = result.split(fullMatch).join(attachmentCardHtml(att, displayName, '내용 표시 실패'));
      }
      continue;
    }

    if (kind === 'word') {
      // docx-preview로 화면 밖 임시 영역에 렌더링한 뒤 HTML(스타일 포함)을 인쇄 문서에 옮긴다.
      // 이미지는 base64로 넣어 인쇄용 iframe에서도 보이게 한다.
      const holder = document.createElement('div');
      holder.style.cssText = 'position:fixed;left:-20000px;top:0;width:900px;visibility:hidden;';
      document.body.appendChild(holder);
      try {
        await renderDocxToContainer(att.data, holder, { useBase64URL: true } as any);
        const replacement = `
          <div class="print-doc-container print-doc-flow">
            <div class="print-doc-header">
              <span class="print-doc-badge print-doc-badge-word">📘 워드</span>
              <span class="print-doc-title">${escapeHtml(displayName)}</span>
            </div>
            <div class="print-docx">${holder.innerHTML}</div>
          </div>
        `;
        result = result.split(fullMatch).join(replacement);
      } catch (err) {
        console.error('워드 인쇄용 렌더링 실패:', err);
        result = result.split(fullMatch).join(attachmentCardHtml(att, displayName, '내용 표시 실패'));
      } finally {
        holder.remove();
      }
      continue;
    }

    if (kind === 'ppt') {
      // 슬라이드 그림 렌더러는 뷰어 컴포넌트 안에만 있어, 슬라이드별 제목·본문·표를 글로 넣는다.
      try {
        const { markdown } = await extractPptxTextAndMarkdown(att.data);
        const body = marked.parse(markdown.replace(/^# .*\n/, ''), { breaks: true, gfm: true }) as string;
        const replacement = `
          <div class="print-doc-container print-doc-flow">
            <div class="print-doc-header">
              <span class="print-doc-badge print-doc-badge-ppt">📽️ PPT</span>
              <span class="print-doc-title">${escapeHtml(displayName)} (슬라이드 내용)</span>
            </div>
            <div class="print-ppt-text">${body}</div>
          </div>
        `;
        result = result.split(fullMatch).join(replacement);
      } catch (err) {
        console.error('PPT 인쇄용 렌더링 실패:', err);
        result = result.split(fullMatch).join(attachmentCardHtml(att, displayName, '내용 표시 실패'));
      }
      continue;
    }

    if (isPdf) {
      try {
        const doc = await loadPdfDocument(att.data, att.id);
        const total = doc.numPages;
        let pagesHtml = '';
        for (let p = 1; p <= total; p++) {
          const page = await doc.getPage(p);
          const unscaledVp = page.getViewport({ scale: 1.0 });
          const isLandscape = unscaledVp.width > unscaledVp.height;
          // 가로 문서는 90도 시계방향 회전하여 세로 A4 페이지에 꽉 차게 눕혀 렌더링
          const viewport = page.getViewport({
            scale: 1.6,
            rotation: isLandscape ? 90 : 0,
          });
          const canvas = document.createElement('canvas');
          canvas.width = viewport.width;
          canvas.height = viewport.height;
          const ctx = canvas.getContext('2d');
          if (ctx) {
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            await (page.render({ canvasContext: ctx, viewport } as any)).promise;
            const dataUrl = canvas.toDataURL('image/png');
            pagesHtml += `
              <div class="print-doc-page ${isLandscape ? 'is-landscape-rotated' : 'is-portrait'}">
                <img src="${dataUrl}" alt="페이지 ${p}" />
                <div class="print-page-number">- ${p} / ${total} 페이지 ${isLandscape ? '(가로 90° 눕힘)' : ''} -</div>
              </div>
            `;
          }
        }

        const replacement = `
          <div class="print-doc-container">
            <div class="print-doc-header">
              <span class="print-doc-badge">📄 PDF 문서</span>
              <span class="print-doc-title">${displayName} (총 ${total}페이지)</span>
            </div>
            <div class="print-doc-pages">
              ${pagesHtml}
            </div>
          </div>
        `;
        result = result.split(fullMatch).join(replacement);
      } catch (err) {
        console.error('PDF 인쇄용 렌더링 실패:', err);
        const fallback = `
          <div class="print-att-card">
            <span>📄</span>
            <span>${displayName} (PDF 문서 첨부)</span>
          </div>
        `;
        result = result.split(fullMatch).join(fallback);
      }
    } else if (isHwp) {
      try {
        const hwpDoc = await loadHwpDocument(att.data, att.id);
        const total = hwpDoc.pageCount();
        let pagesHtml = '';
        for (let p = 0; p < total; p++) {
          let svg = renderHwpPageToSvg(hwpDoc, p);
          const isLandscape = isSvgLandscape(svg);
          // 가로 설정 문서는 90도 회전(눕힘)하여 세로 A4 페이지에 꽉 차게 변환
          if (isLandscape) {
            svg = rotateSvg90Degrees(svg);
          }
          pagesHtml += `
            <div class="print-doc-page ${isLandscape ? 'is-landscape-rotated' : 'is-portrait'}">
              <div class="print-hwp-svg-page">${svg}</div>
              <div class="print-page-number">- ${p + 1} / ${total} 페이지 ${isLandscape ? '(가로 90° 눕힘)' : ''} -</div>
            </div>
          `;
        }

        const replacement = `
          <div class="print-doc-container">
            <div class="print-doc-header">
              <span class="print-doc-badge print-doc-badge-hwp">📝 한글(HWP) 문서</span>
              <span class="print-doc-title">${displayName} (총 ${total}페이지)</span>
            </div>
            <div class="print-doc-pages">
              ${pagesHtml}
            </div>
          </div>
        `;
        result = result.split(fullMatch).join(replacement);
      } catch (err) {
        console.error('HWP 인쇄용 렌더링 실패:', err);
        const fallback = `
          <div class="print-att-card">
            <span>📝</span>
            <span>${displayName} (한글 문서 첨부)</span>
          </div>
        `;
        result = result.split(fullMatch).join(fallback);
      }
    } else {
      result = result.split(fullMatch).join(attachmentCardHtml(att, displayName));
    }
  }

  return result;
}

/**
 * 메모 내용을 깔끔하고 표준적인 A4 PDF 문서 양식으로 렌더링하여 브라우저의 PDF 저장 / 인쇄 창을 실행합니다.
 * 본문에 포함된 PDF 및 한글(HWP) 문서도 실제 페이지 그래픽으로 함께 인쇄물에 포함됩니다.
 */
export async function exportNoteToPdf(
  note: Note,
  folderBreadcrumbs: string = '',
  customContent?: string,
  attachmentMode: PdfAttachmentMode = 'expanded',
  toggleOpenStates?: PdfToggleOpenState[]
): Promise<void> {
  const contentToRender = customContent !== undefined ? customContent : note.content;
  let normalizedContent = applyDetailsMode(contentToRender || '', note.attachments || {}, attachmentMode, toggleOpenStates);
  normalizedContent = normalizedContent.replace(/^(\s*-\s*\[[ xX]\])([^\s\n])/gm, '$1 $2');
  normalizedContent = normalizedContent.replace(/==([^=\n]+)==/g, '<mark style="background-color: #fef08a; color: #0f172a; padding: 2px 5px; border-radius: 4px;">$1</mark>');
  const rawHtml = marked.parse(normalizedContent, { breaks: true, gfm: true }) as string;
  const renderedHtml = await renderAttachmentsForPrint(rawHtml, note.attachments, attachmentMode !== 'none');

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
      margin: 14mm 12mm 14mm 12mm;
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
    mark {
      background-color: #fef08a;
      color: #0f172a;
      padding: 1px 4px;
      border-radius: 3px;
    }
    
    /* 표 */
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 16px 0;
      font-size: 12px;
      line-height: 1.5;
      break-inside: avoid;
    }
    th, td {
      border: 1px solid #cbd5e1;
      padding: 7px 10px;
      text-align: left;
      vertical-align: middle;
      word-break: keep-all;
      overflow-wrap: break-word;
    }
    th {
      background-color: #f1f5f9;
      font-weight: 700;
      color: #1e293b;
      white-space: nowrap;
    }
    th[align="center"], td[align="center"] {
      text-align: center;
    }
    th[align="right"], td[align="right"] {
      text-align: right;
    }
    th[align="left"], td[align="left"] {
      text-align: left;
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

    /* 인쇄용 PDF 및 한글(HWP) 문서 뷰어 컨테이너 */
    .print-doc-container {
      margin: 20px 0;
      border: 1.5px solid #cbd5e1;
      border-radius: 8px;
      overflow: hidden;
      background: #ffffff;
      break-inside: avoid-page;
      page-break-inside: avoid;
    }
    .print-doc-header {
      background: #f8fafc;
      border-bottom: 1.5px solid #cbd5e1;
      padding: 8px 14px;
      font-size: 13px;
      font-weight: 700;
      color: #0f172a;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .print-doc-badge {
      background: #4f46e5;
      color: #ffffff;
      padding: 2px 7px;
      border-radius: 4px;
      font-size: 10.5px;
      font-weight: 700;
    }
    .print-doc-badge-hwp {
      background: #2563eb;
    }
    .print-doc-title {
      font-size: 12.5px;
      color: #1e293b;
    }
    /* 문서 각 페이지 (세로 A4 용지 1페이지에 온전히 맞춤) */
    .print-doc-page {
      padding: 10px 0;
      text-align: center;
      border-bottom: 1px dashed #cbd5e1;
      box-sizing: border-box;
      break-inside: avoid;
      page-break-inside: avoid;
      break-before: page;
      page-break-before: always;
      width: 100%;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
    }
    .print-doc-page:first-child {
      break-before: auto;
      page-break-before: auto;
    }
    .print-doc-page:last-child {
      border-bottom: none;
    }
    .print-doc-page img {
      width: 100% !important;
      max-width: 100% !important;
      max-height: 250mm !important;
      height: auto !important;
      border: 1px solid #cbd5e1;
      border-radius: 4px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.08);
      margin: 0 auto;
      display: block;
      object-fit: contain;
    }
    .print-hwp-svg-page {
      width: 100%;
      display: flex;
      justify-content: center;
      align-items: center;
    }
    .print-hwp-svg-page svg {
      width: 100% !important;
      max-width: 100% !important;
      max-height: 250mm !important;
      height: auto !important;
      border: 1px solid #cbd5e1;
      border-radius: 4px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.08);
      margin: 0 auto;
      display: block;
    }
    .print-page-number {
      font-size: 11px;
      color: #64748b;
      font-weight: 600;
      margin-top: 6px;
    }

    /* 엑셀·워드·PPT처럼 길어질 수 있는 첨부는 페이지를 넘겨 이어지게 */
    .print-doc-container.print-doc-flow {
      break-inside: auto !important;
      page-break-inside: auto !important;
    }
    .print-doc-badge-excel { background: #059669; }
    .print-doc-badge-word { background: #2563eb; }
    .print-doc-badge-ppt { background: #ea580c; }
    .print-empty { padding: 10px 14px; font-size: 12px; color: #64748b; }
    .print-excel-sheet { padding: 8px 10px 12px; }
    .print-excel-sheet + .print-excel-sheet { border-top: 1px dashed #cbd5e1; }
    .print-excel-sheet-name { font-size: 12px; font-weight: 700; color: #065f46; margin: 2px 0 6px; }
    table.print-excel-table {
      width: 100% !important;
      table-layout: auto;
      margin: 0 !important;
      font-size: 9.5px !important;
      break-inside: auto !important;
    }
    table.print-excel-table tr { break-inside: avoid; }
    table.print-excel-table td, table.print-excel-table th {
      padding: 3px 5px !important;
      font-size: 9.5px !important;
      max-width: none !important;
      word-break: break-all;
      white-space: pre-wrap !important;
    }
    table.print-excel-table td.excel-td-num { position: static !important; }
    /* 워드: A4 페이지 폭(210mm)을 인쇄 영역(186mm)에 맞추고 페이지마다 나눠 인쇄 */
    .print-docx { padding: 6px 0; }
    .print-docx section {
      zoom: 0.88;
      box-shadow: none !important;
      margin: 0 auto !important;
      break-after: page;
      page-break-after: always;
    }
    .print-docx section:last-of-type { break-after: auto; page-break-after: auto; }
    .print-ppt-text { padding: 4px 14px 10px; font-size: 12.5px; }
    .print-ppt-text h2 { font-size: 13.5px; margin-top: 14px; }
    .print-collapsed-group { margin: 10px 0; }
    /* 글 토글(<details>): 접힌 것은 제목 한 줄, 펼친 것은 테두리 상자 안에 내용까지 */
    details {
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      padding: 8px 14px;
      margin: 10px 0;
      background: #ffffff;
    }
    details > summary {
      font-weight: 700;
      font-size: 13.5px;
      color: #0f172a;
      cursor: default;
    }
    details[open] > summary {
      border-bottom: 1px solid #e2e8f0;
      padding-bottom: 6px;
      margin-bottom: 8px;
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
      }, 3000);
    }
  }, 450);
}
