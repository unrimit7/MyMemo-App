import { renderAsync, type Options } from 'docx-preview';
import * as mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';

// Data URL 또는 Base64 문자열을 ArrayBuffer로 변환
export function dataUrlToArrayBuffer(dataOrUrl: string): ArrayBuffer {
  const commaIdx = dataOrUrl.indexOf(',');
  const base64 = commaIdx !== -1 ? dataOrUrl.slice(commaIdx + 1) : dataOrUrl;
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

// DOCX 캐시 관리
const docxCache = new Map<string, ArrayBuffer>();

export function invalidateDocxCache(attId?: string) {
  if (attId) {
    docxCache.delete(attId);
  } else {
    docxCache.clear();
  }
}

/**
 * Word (.docx) 문서를 특정 DOM 컨테이너에 고해상도로 렌더링합니다.
 */
export async function renderDocxToContainer(
  dataOrUrl: string,
  container: HTMLElement,
  options?: Partial<Options>
): Promise<void> {
  const arrayBuffer = dataUrlToArrayBuffer(dataOrUrl);
  container.innerHTML = '';

  const defaultOpts: Partial<Options> = {
    inWrapper: true,
    ignoreWidth: false,
    ignoreHeight: false,
    ignoreFonts: false,
    breakPages: true,
    className: 'docx-preview-rendered',
    trimXmlDeclaration: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    ...options,
  };

  await renderAsync(arrayBuffer, container, undefined, defaultOpts);
}

/**
 * HTML 문자열을 깔끔한 마크다운으로 변환
 */
function htmlToMarkdown(html: string): string {
  const div = document.createElement('div');
  div.innerHTML = html;

  // 헤딩 변환
  for (let i = 1; i <= 6; i++) {
    div.querySelectorAll(`h${i}`).forEach((el) => {
      const prefix = '#'.repeat(i);
      el.outerHTML = `\n\n${prefix} ${el.textContent?.trim()}\n\n`;
    });
  }

  // 문단 변환
  div.querySelectorAll('p').forEach((el) => {
    el.outerHTML = `\n\n${el.textContent?.trim()}\n\n`;
  });

  // 강조 태그
  div.querySelectorAll('strong, b').forEach((el) => {
    el.outerHTML = `**${el.textContent}**`;
  });
  div.querySelectorAll('em, i').forEach((el) => {
    el.outerHTML = `*${el.textContent}*`;
  });

  // 목록 변환
  div.querySelectorAll('li').forEach((el) => {
    el.outerHTML = `\n- ${el.textContent?.trim()}`;
  });

  // 표(Table) 변환
  div.querySelectorAll('table').forEach((table) => {
    const rows = Array.from(table.querySelectorAll('tr'));
    if (rows.length === 0) return;

    const tableLines: string[] = [];
    let maxCols = 0;

    rows.forEach((tr, rIdx) => {
      const cells = Array.from(tr.querySelectorAll('th, td')).map((c) =>
        (c.textContent || '').replace(/\|/g, '\\|').replace(/\n/g, ' ').trim()
      );
      maxCols = Math.max(maxCols, cells.length);
      tableLines.push(`| ${cells.join(' | ')} |`);

      // 1행 아래 구분선 삽입
      if (rIdx === 0) {
        tableLines.push(`| ${cells.map(() => '---').join(' | ')} |`);
      }
    });

    table.outerHTML = `\n\n${tableLines.join('\n')}\n\n`;
  });

  // 불필요한 연속 공백 및 줄바꿈 정리
  return (div.textContent || '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Word (.docx) 문서에서 텍스트, HTML, 그리고 마크다운을 추출합니다.
 */
export async function extractDocxTextAndHtml(
  dataOrUrl: string
): Promise<{ text: string; html: string; markdown: string }> {
  try {
    const arrayBuffer = dataUrlToArrayBuffer(dataOrUrl);

    // mammoth를 사용하여 깔끔한 HTML과 텍스트 추출
    const [htmlResult, textResult] = await Promise.all([
      mammoth.convertToHtml({ arrayBuffer }),
      mammoth.extractRawText({ arrayBuffer }),
    ]);

    const html = htmlResult.value || '';
    const text = textResult.value || '';
    const markdown = htmlToMarkdown(html) || text;

    return { text, html, markdown };
  } catch (err) {
    console.error('워드 문서 텍스트 추출 실패:', err);
    return { text: '', html: '', markdown: '' };
  }
}

/**
 * 텍스트나 문단을 새로운 Word (.docx) Data URL로 생성합니다.
 */
export async function generateDocxDataUrl(title: string, content: string): Promise<string> {
  const lines = content.split('\n');
  const paragraphs: Paragraph[] = [];

  // 문서 제목 문단
  if (title) {
    paragraphs.push(
      new Paragraph({
        text: title,
        heading: HeadingLevel.TITLE,
        spacing: { after: 240 },
      })
    );
  }

  // 본문 라인들을 문단으로 변환
  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      paragraphs.push(new Paragraph({ spacing: { after: 120 } }));
      return;
    }

    if (trimmed.startsWith('# ')) {
      paragraphs.push(
        new Paragraph({
          text: trimmed.slice(2).trim(),
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 240, after: 120 },
        })
      );
    } else if (trimmed.startsWith('## ')) {
      paragraphs.push(
        new Paragraph({
          text: trimmed.slice(3).trim(),
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 200, after: 100 },
        })
      );
    } else if (trimmed.startsWith('### ')) {
      paragraphs.push(
        new Paragraph({
          text: trimmed.slice(4).trim(),
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 160, after: 80 },
        })
      );
    } else if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun(trimmed.slice(2).trim())],
          bullet: { level: 0 },
          spacing: { after: 80 },
        })
      );
    } else {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun(line)],
          spacing: { after: 100 },
        })
      );
    }
  });

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: paragraphs,
      },
    ],
  });

  const base64Out = await Packer.toBase64String(doc);
  return `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${base64Out}`;
}
