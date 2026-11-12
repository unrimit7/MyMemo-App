import { describe, it, expect } from 'vitest';
import { serializeNoteToMarkdown, parseMarkdownToNote } from '../src/services/github';
import type { NoteItem, AttachmentItem } from '../src/types';

describe('syncOriginalBackup GitHub 동기화 제외 및 기기 내부 보존 검증', () => {
  const dummyNote: NoteItem = {
    id: 'note-123',
    title: '테스트 문서',
    content: '본문 내용입니다.',
    folder: '',
    createdAt: '2026-11-12T00:00:00Z',
    updatedAt: '2026-11-12T00:00:00Z',
    attachments: {
      'att-1': {
        id: 'att-1',
        name: '보고서.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        size: 15000,
        data: 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,MODIFIED_DATA',
        originalData: 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,ORIGINAL_RAW_BACKUP_BYTES',
        originalDate: '2026-11-10T12:00:00Z',
      } as AttachmentItem,
    },
  };

  it('기본 설정(syncOriginalBackup: false)일 때 직렬화 결과에서 originalData와 originalDate가 엄격히 제외되어야 함', () => {
    const md = serializeNoteToMarkdown(dummyNote, '', false);

    // originalData 내용이 마크다운 텍스트 어디에도 포함되지 않아야 함
    expect(md).not.toContain('ORIGINAL_RAW_BACKUP_BYTES');
    expect(md).toContain('MODIFIED_DATA');

    // 역직렬화(parseMarkdownToNote) 시에도 originalData가 없어야 함
    const parsed = parseMarkdownToNote(md);
    expect(parsed.attachments).toBeDefined();
    expect(parsed.attachments!['att-1'].data).toBe('data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,MODIFIED_DATA');
    expect(parsed.attachments!['att-1'].originalData).toBeUndefined();
    expect(parsed.attachments!['att-1'].originalDate).toBeUndefined();
  });

  it('syncOriginalBackup: true로 명시할 때만 originalData가 마크다운에 보존되어야 함', () => {
    const md = serializeNoteToMarkdown(dummyNote, '', true);

    expect(md).toContain('ORIGINAL_RAW_BACKUP_BYTES');
    const parsed = parseMarkdownToNote(md);
    expect(parsed.attachments!['att-1'].originalData).toBe('data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,ORIGINAL_RAW_BACKUP_BYTES');
  });
});
