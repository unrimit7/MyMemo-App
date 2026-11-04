import { renderAsync, type Options } from 'docx-preview';
import * as mammoth from 'mammoth';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  HeadingLevel, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType,
  BorderStyle,
  AlignmentType
} from 'docx';

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
    className: 'docx-doc-page',
    trimXmlDeclaration: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    ...options,
  };

  await renderAsync(arrayBuffer, container, undefined, defaultOpts);

  // docx-preview가 생성한 wrapper의 gray 배경을 투명으로 설정하고 A4 페이지들이 깔끔하게 중앙 정렬되도록 보정
  const wrapper = container.querySelector<HTMLElement>('.docx-doc-page-wrapper') || container.querySelector<HTMLElement>('[class*="-wrapper"]');
  if (wrapper) {
    wrapper.style.setProperty('background', 'transparent', 'important');
    wrapper.style.setProperty('padding', '8px 0', 'important');
    wrapper.style.setProperty('display', 'flex');
    wrapper.style.setProperty('flex-direction', 'column');
    wrapper.style.setProperty('align-items', 'center');
    wrapper.style.setProperty('width', '100%');
    wrapper.style.setProperty('box-sizing', 'border-box');
  }

  const sections = container.querySelectorAll<HTMLElement>('section.docx-doc-page, section[class*="docx-doc"]');
  sections.forEach((sec) => {
    sec.style.setProperty('background', '#ffffff', 'important');
    sec.style.setProperty('color', '#0f172a', 'important');
    sec.style.setProperty('box-shadow', '0 10px 30px -5px rgba(0, 0, 0, 0.45)', 'important');
    sec.style.setProperty('margin', '12px auto 24px auto', 'important');
    sec.style.setProperty('border-radius', '4px', 'important');
    sec.style.setProperty('box-sizing', 'border-box');
    sec.style.setProperty('max-width', '100%');
  });
}

/**
 * DOM 노드를 재귀적으로 깔끔한 마크다운으로 변환 (서식 및 표 완벽 보존)
 */
function nodeToMarkdown(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent || '';
  }
  if (node.nodeType !== Node.ELEMENT_NODE) {
    return '';
  }

  const el = node as HTMLElement;
  const tag = el.tagName.toLowerCase();
  const childrenMd = Array.from(el.childNodes).map(nodeToMarkdown).join('');

  switch (tag) {
    case 'h1':
      return `\n\n# ${childrenMd.trim()}\n\n`;
    case 'h2':
      return `\n\n## ${childrenMd.trim()}\n\n`;
    case 'h3':
      return `\n\n### ${childrenMd.trim()}\n\n`;
    case 'h4':
    case 'h5':
    case 'h6':
      return `\n\n#### ${childrenMd.trim()}\n\n`;
    case 'p':
      return `\n\n${childrenMd.trim()}\n\n`;
    case 'strong':
    case 'b':
      return `**${childrenMd}**`;
    case 'em':
    case 'i':
      return `*${childrenMd}*`;
    case 'u':
      return `<u>${childrenMd}</u>`;
    case 's':
    case 'del':
      return `~~${childrenMd}~~`;
    case 'code':
      return `\`${childrenMd}\``;
    case 'blockquote':
      return `\n\n> ${childrenMd.trim().replace(/\n/g, '\n> ')}\n\n`;
    case 'ul':
      return `\n\n${childrenMd}\n\n`;
    case 'ol':
      return `\n\n${childrenMd}\n\n`;
    case 'li': {
      const isOrdered = el.parentElement?.tagName.toLowerCase() === 'ol';
      const index = Array.from(el.parentElement?.children || []).indexOf(el) + 1;
      const prefix = isOrdered ? `${index}. ` : '- ';
      return `\n${prefix}${childrenMd.trim()}`;
    }
    case 'table': {
      const rows = Array.from(el.querySelectorAll('tr'));
      if (rows.length === 0) return '';
      const tableLines: string[] = [];

      rows.forEach((tr, rIdx) => {
        const cells = Array.from(tr.querySelectorAll('th, td')).map((c) =>
          Array.from(c.childNodes).map(nodeToMarkdown).join('').replace(/\|/g, '\\|').replace(/\n/g, ' ').trim()
        );
        tableLines.push(`| ${cells.join(' | ')} |`);
        if (rIdx === 0) {
          tableLines.push(`| ${cells.map(() => '---').join(' | ')} |`);
        }
      });
      return `\n\n${tableLines.join('\n')}\n\n`;
    }
    case 'a': {
      const href = el.getAttribute('href') || '';
      return `[${childrenMd}](${href})`;
    }
    case 'br':
      return '\n';
    case 'hr':
      return '\n\n---\n\n';
    default:
      return childrenMd;
  }
}

/**
 * HTML 문자열을 깔끔한 마크다운으로 변환
 */
export function htmlToMarkdown(html: string): string {
  const container = document.createElement('div');
  container.innerHTML = html;
  return nodeToMarkdown(container)
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

    // mammoth를 사용하여 구조화된 HTML과 원문 텍스트 추출
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
 * 마크다운/텍스트 내용을 표준 Word (.docx) Data URL로 생성합니다.
 * 표(Table), 헤딩(Heading), 목록(Bullet List), 문단을 docx 정식 객체로 변환합니다.
 */
export async function generateDocxDataUrl(title: string, content: string): Promise<string> {
  const lines = content.split('\n');
  const docChildren: (Paragraph | Table)[] = [];

  // 문서 제목 문단
  if (title) {
    docChildren.push(
      new Paragraph({
        text: title,
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { after: 360 },
      })
    );
  }

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const trimmed = line.trim();

    // 1. 빈 줄
    if (!trimmed) {
      i++;
      continue;
    }

    // 2. 마크다운 표 감지 (| 열1 | 열2 | ...)
    if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
      const tableRowsData: string[][] = [];
      while (i < lines.length && lines[i].trim().startsWith('|') && lines[i].trim().endsWith('|')) {
        const rowStr = lines[i].trim();
        // 구분선(| --- | --- |) 건너뛰기
        if (!/^\|(?:\s*:?-+:?\s*\|)+$/.test(rowStr)) {
          const cells = rowStr
            .slice(1, -1)
            .split('|')
            .map((c) => c.trim().replace(/\\\|/g, '|'));
          tableRowsData.push(cells);
        }
        i++;
      }

      if (tableRowsData.length > 0) {
        const maxCols = Math.max(...tableRowsData.map((r) => r.length));
        const tableRows = tableRowsData.map((rowCells, rIdx) => {
          return new TableRow({
            tableHeader: rIdx === 0,
            children: Array.from({ length: maxCols }).map((_, cIdx) => {
              const cellText = rowCells[cIdx] || '';
              return new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: cellText,
                        bold: rIdx === 0,
                      }),
                    ],
                    spacing: { before: 80, after: 80 },
                  }),
                ],
                width: {
                  size: Math.floor(100 / maxCols),
                  type: WidthType.PERCENTAGE,
                },
                shading: rIdx === 0 ? { fill: 'F1F5F9' } : undefined,
              });
            }),
          });
        });

        docChildren.push(
          new Table({
            rows: tableRows,
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: 'CBD5E1' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CBD5E1' },
              left: { style: BorderStyle.SINGLE, size: 1, color: 'CBD5E1' },
              right: { style: BorderStyle.SINGLE, size: 1, color: 'CBD5E1' },
              insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
              insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
            },
          })
        );
        // 표 뒤 여백용 문단
        docChildren.push(new Paragraph({ spacing: { after: 200 } }));
      }
      continue;
    }

    // 3. 헤딩
    if (trimmed.startsWith('# ')) {
      docChildren.push(
        new Paragraph({
          text: trimmed.slice(2).trim(),
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 280, after: 120 },
        })
      );
    } else if (trimmed.startsWith('## ')) {
      docChildren.push(
        new Paragraph({
          text: trimmed.slice(3).trim(),
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 240, after: 100 },
        })
      );
    } else if (trimmed.startsWith('### ')) {
      docChildren.push(
        new Paragraph({
          text: trimmed.slice(4).trim(),
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 200, after: 80 },
        })
      );
    } else if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
      // 4. 글머리 기호 목록
      docChildren.push(
        new Paragraph({
          children: [new TextRun(trimmed.slice(2).trim())],
          bullet: { level: 0 },
          spacing: { after: 80 },
        })
      );
    } else {
      // 5. 일반 문단
      docChildren.push(
        new Paragraph({
          children: [new TextRun(line)],
          spacing: { after: 120 },
        })
      );
    }

    i++;
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: docChildren,
      },
    ],
  });

  const base64Out = await Packer.toBase64String(doc);
  return `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${base64Out}`;
}
