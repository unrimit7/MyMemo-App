// @vitest-environment jsdom
// 첨부 카드 공통 동작(버튼 판별, 저장 시 원본 보관, 원본 복원)의 회귀 테스트
import { describe, it, expect } from 'vitest';
import { resolveCardAction, withSavedAttachment, withRestoredOriginal } from '../src/utils/attachmentCardActions';
import type { AttachmentItem } from '../src/types';

function el(html: string, sel: string): HTMLElement {
  const host = document.createElement('div');
  host.innerHTML = html;
  return host.querySelector<HTMLElement>(sel)!;
}

describe('resolveCardAction', () => {
  it('버튼 안쪽 span을 눌러도 버튼 동작으로 판별한다', () => {
    const t = el('<button class="pdf-btn-import" data-att-id="a1" data-file-name="x.pdf"><span>가져오기</span></button>', 'span');
    expect(resolveCardAction(t)).toMatchObject({ kind: 'pdf', action: 'import', attId: 'a1', fileName: 'x.pdf' });
  });

  it('워드 용지와 PPT 카드 본문 클릭은 크게 보기다', () => {
    expect(resolveCardAction(el('<div class="docx-container" data-att-id="d"><p>본문</p></div>', 'p'))).toMatchObject({ kind: 'docx', action: 'fullscreen' });
    expect(resolveCardAction(el('<div class="pptx-container" data-att-id="p"><h3>제목</h3></div>', 'h3'))).toMatchObject({ kind: 'pptx', action: 'fullscreen' });
  });

  it('PPT 편집 기능이 꺼져 있으면 편집 버튼을 동작으로 보지 않는다', () => {
    expect(resolveCardAction(el('<button class="pptx-btn-edit" data-att-id="p">편집</button>', 'button'))).toBeNull();
  });

  it('카드와 관계없는 요소는 null', () => {
    expect(resolveCardAction(el('<p>일반 글</p>', 'p'))).toBeNull();
  });
});

describe('withSavedAttachment / withRestoredOriginal', () => {
  const base: Record<string, AttachmentItem> = {
    a1: { id: 'a1', name: 'x.pdf', type: 'application/pdf', data: 'data:original' } as AttachmentItem,
  };

  it('처음 저장할 때 원래 내용을 originalData로 보관한다', () => {
    const once = withSavedAttachment(base, 'a1', 'data:v1');
    expect(once.a1.data).toBe('data:v1');
    expect(once.a1.originalData).toBe('data:original');
    expect(once.a1.originalDate).toBeTruthy();
  });

  it('두 번째 저장에서도 처음 원본을 그대로 유지한다', () => {
    const twice = withSavedAttachment(withSavedAttachment(base, 'a1', 'data:v1'), 'a1', 'data:v2');
    expect(twice.a1.data).toBe('data:v2');
    expect(twice.a1.originalData).toBe('data:original');
  });

  it('원본으로 되돌리면 data가 원본이 되고 보관본은 비워진다', () => {
    const restored = withRestoredOriginal(withSavedAttachment(base, 'a1', 'data:v1'), 'a1')!;
    expect(restored.a1.data).toBe('data:original');
    expect(restored.a1.originalData).toBeUndefined();
  });

  it('보관본이 없으면 복원하지 않는다', () => {
    expect(withRestoredOriginal(base, 'a1')).toBeNull();
  });
});
