// 메모 본문 첨부 카드(엑셀·워드·한글·PDF·PPT)의 버튼 판별, 본문 가져오기, 저장/원본 복원을 메인 편집기와 창 분리가 함께 쓰도록 모은 모듈
import type { AttachmentItem } from '../types';
import { PPTX_EDITING_ENABLED } from './featureFlags';
import { invalidateAttachmentBlobUrl } from './attachment';
import { parseExcelWorkbook, sheetRowsToMarkdownTable, invalidateExcelCache } from './excelViewer';
import { extractDocxTextAndHtml, invalidateDocxCache } from './docxViewer';
import { loadHwpDocument, extractHwpText, invalidateHwpCache } from './hwpViewer';
import { loadPdfDocument, extractPdfText, invalidatePdfCache } from './pdfViewer';
import { extractPptxTextAndMarkdown, invalidatePptxCache } from './pptxViewer';

export type CardKind = 'excel' | 'docx' | 'hwp' | 'pdf' | 'pptx';
export type CardActionType = 'fullscreen' | 'edit' | 'import';

export interface CardAction {
  kind: CardKind;
  action: CardActionType;
  attId: string;
  fileName: string;
  el: HTMLElement;
}

// 클릭 대상 → 카드 동작. 앞에 있는 선택자가 먼저 판정된다 (버튼이 카드 본문보다 우선).
const CARD_SELECTORS: { sel: string; kind: CardKind; action: CardActionType }[] = [
  { sel: '.excel-btn-edit', kind: 'excel', action: 'edit' },
  { sel: '.excel-btn-import', kind: 'excel', action: 'import' },
  { sel: '.excel-btn-fullscreen', kind: 'excel', action: 'fullscreen' },
  { sel: '.docx-btn-edit', kind: 'docx', action: 'edit' },
  { sel: '.docx-btn-import', kind: 'docx', action: 'import' },
  { sel: '.docx-btn-fullscreen', kind: 'docx', action: 'fullscreen' },
  { sel: '.docx-container', kind: 'docx', action: 'fullscreen' },
  { sel: '.hwp-btn-edit', kind: 'hwp', action: 'edit' },
  { sel: '.hwp-btn-import', kind: 'hwp', action: 'import' },
  { sel: '.hwp-btn-fullscreen', kind: 'hwp', action: 'fullscreen' },
  { sel: '.hwp-svg-container', kind: 'hwp', action: 'fullscreen' },
  { sel: '.pdf-btn-edit', kind: 'pdf', action: 'edit' },
  { sel: '.pdf-btn-import', kind: 'pdf', action: 'import' },
  { sel: '.pdf-btn-fullscreen', kind: 'pdf', action: 'fullscreen' },
  { sel: 'canvas.pdf-canvas', kind: 'pdf', action: 'fullscreen' },
  { sel: '.pptx-btn-edit', kind: 'pptx', action: 'edit' },
  { sel: '.pptx-btn-import', kind: 'pptx', action: 'import' },
  { sel: '.pptx-btn-fullscreen', kind: 'pptx', action: 'fullscreen' },
  { sel: '.pptx-container', kind: 'pptx', action: 'fullscreen' },
];

export function resolveCardAction(target: HTMLElement | null): CardAction | null {
  if (!target) return null;
  for (const { sel, kind, action } of CARD_SELECTORS) {
    const el = target.closest<HTMLElement>(sel);
    if (!el) continue;
    if (kind === 'pptx' && action === 'edit' && !PPTX_EDITING_ENABLED) return null;
    return {
      kind,
      action,
      attId: el.getAttribute('data-att-id') || '',
      fileName: el.getAttribute('data-file-name') || '',
      el,
    };
  }
  return null;
}

// 추출은 됐지만 가져올 내용이 없을 때 (메시지를 그대로 사용자에게 보여 준다)
export class ImportEmptyError extends Error {}

const IMPORT_FAIL_MESSAGE: Record<CardKind, string> = {
  excel: '엑셀 표 데이터를 읽지 못했습니다.',
  docx: '워드 문서 내용을 추출하지 못했습니다.',
  hwp: '한글 문서 텍스트를 추출하지 못했습니다.',
  pdf: 'PDF 문서 텍스트를 추출하지 못했습니다.',
  pptx: '파워포인트 슬라이드 내용을 추출하지 못했습니다.',
};

export function importErrorMessage(kind: CardKind, err: unknown): string {
  return err instanceof ImportEmptyError ? err.message : IMPORT_FAIL_MESSAGE[kind];
}

// 첨부 내용을 메모 본문 끝에 덧붙일 마크다운 블록으로 만든다. 내용이 없으면 ImportEmptyError를 던진다.
export async function buildImportBlock(kind: CardKind, att: AttachmentItem): Promise<{ block: string; message: string }> {
  const name = att.name;
  if (kind === 'excel') {
    const parsed = parseExcelWorkbook(att.data, att.id);
    const firstSheet = parsed.sheetNames[0] || 'Sheet1';
    const rows = parsed.sheets[firstSheet] || [];
    const md = sheetRowsToMarkdownTable(rows);
    if (!md.trim()) throw new ImportEmptyError('가져올 수 있는 표 데이터가 없습니다.');
    return {
      block: `\n\n### 📊 ${name} (${firstSheet})\n\n${md}\n\n`,
      message: `'${name}' 표 데이터(${rows.length}행)를 마크다운 표로 메모에 성공적으로 가져왔습니다!`,
    };
  }
  if (kind === 'docx') {
    const { markdown, text } = await extractDocxTextAndHtml(att.data);
    const body = markdown || text;
    if (!body.trim()) throw new ImportEmptyError('추출할 수 있는 문서 내용이 없습니다.');
    return {
      block: `\n\n### 📄 ${name} 내용\n\n${body}\n\n`,
      message: `'${name}' 문서 내용(${body.length}자)을 메모에 성공적으로 가져왔습니다!`,
    };
  }
  if (kind === 'hwp') {
    const doc = await loadHwpDocument(att.data, att.id);
    const text = extractHwpText(doc);
    if (!text.trim()) throw new ImportEmptyError('추출할 수 있는 본문 텍스트가 없습니다.');
    return {
      block: `\n\n### 📝 ${name} 본문\n\n${text}\n\n`,
      message: `'${name}' 문서의 본문 텍스트(${text.length}자)를 메모에 성공적으로 가져왔습니다!`,
    };
  }
  if (kind === 'pdf') {
    const doc = await loadPdfDocument(att.data, att.id);
    const text = await extractPdfText(doc);
    if (!text.trim()) throw new ImportEmptyError('추출할 수 있는 텍스트가 없습니다. (스캔 이미지로만 구성된 PDF는 텍스트 레이어가 포함되어 있지 않습니다.)');
    return {
      block: `\n\n### 📄 ${name} 본문\n\n${text}\n\n`,
      message: `'${name}' 문서의 본문 텍스트(${text.length}자)를 메모에 성공적으로 가져왔습니다!`,
    };
  }
  const { markdown, text } = await extractPptxTextAndMarkdown(att.data);
  const body = markdown || text;
  if (!body.trim()) throw new ImportEmptyError('추출할 수 있는 슬라이드 내용이 없습니다.');
  return {
    block: `\n\n${body}\n\n`,
    message: `'${name}' 발표 자료 내용을 메모에 성공적으로 가져왔습니다!`,
  };
}

// 편집 저장본을 첨부 목록에 반영한다. 처음 저장할 때 원래 data를 originalData로 보관한다(원본으로 되돌리기용).
export function withSavedAttachment(
  atts: Record<string, AttachmentItem>,
  attId: string,
  updatedDataUrl: string,
  extra: Partial<AttachmentItem> = {},
): Record<string, AttachmentItem> {
  const existing = atts[attId];
  const originalData = existing?.originalData || existing?.data || updatedDataUrl;
  const originalDate = existing?.originalDate || new Date().toISOString();
  return {
    ...atts,
    [attId]: { ...(existing || ({} as AttachmentItem)), data: updatedDataUrl, originalData, originalDate, ...extra },
  };
}

// 보관된 원본으로 되돌린다. 원본 보관본이 없으면 null.
export function withRestoredOriginal(
  atts: Record<string, AttachmentItem>,
  attId: string,
): Record<string, AttachmentItem> | null {
  const existing = atts[attId];
  if (!existing?.originalData) return null;
  return {
    ...atts,
    [attId]: { ...existing, data: existing.originalData, originalData: undefined, originalDate: undefined },
  };
}

// 저장·복원 뒤 미리보기가 옛 내용을 보여 주지 않도록 형식별 캐시를 비운다.
export function invalidateCardCaches(kind: CardKind, attId: string): void {
  invalidateAttachmentBlobUrl(attId);
  if (kind === 'excel') { invalidateExcelCache(attId); invalidateExcelCache(); }
  else if (kind === 'docx') { invalidateDocxCache(attId); invalidateDocxCache(); }
  else if (kind === 'hwp') { invalidateHwpCache(attId); invalidateHwpCache(); }
  else if (kind === 'pdf') invalidatePdfCache(attId);
  else invalidatePptxCache(attId);
}
