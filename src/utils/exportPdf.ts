import type { Note, AttachmentItem } from '../types';
import { marked } from 'marked';
import { loadPdfDocument } from './pdfViewer';
import { loadHwpDocument, renderHwpPageToSvg } from './hwpViewer';

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
  attachments: Record<string, AttachmentItem> = {}
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

    const cleanLabel = (label || '')
      .replace(/^[\s📄📎🖼️📝]*/, '')
      .replace(/^(?:PDF\s*문서:\s*|한글\s*문서:\s*|한글:\s*|파일:\s*)/i, '')
      .trim();
    const displayName = cleanLabel || att.name;

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
      const fallback = `
        <div class="print-att-card">
          <span>📎</span>
          <span>${displayName} (첨부 파일)</span>
        </div>
      `;
      result = result.split(fullMatch).join(fallback);
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
  customContent?: string
): Promise<void> {
  const contentToRender = customContent !== undefined ? customContent : note.content;
  let normalizedContent = (contentToRender || '').replace(/^(\s*-\s*\[[ xX]\])([^\s\n])/gm, '$1 $2');
  normalizedContent = normalizedContent.replace(/==([^=\n]+)==/g, '<mark style="background-color: #fef08a; color: #0f172a; padding: 2px 5px; border-radius: 4px;">$1</mark>');
  const rawHtml = marked.parse(normalizedContent, { breaks: true, gfm: true }) as string;
  const renderedHtml = await renderAttachmentsForPrint(rawHtml, note.attachments);

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
