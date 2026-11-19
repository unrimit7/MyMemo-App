import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { decodeBytesSmart, dataUrlToUint8Array } from '../src/utils/hexViewer';
import { processFileForNote } from '../src/utils/file';
import { createAttachmentCardHtml } from '../src/utils/attachment';
import type { AttachmentItem } from '../src/types';

describe('텍스트 파일 인코딩(UTF-8 및 ANSI/EUC-KR) 자동 감지 디코딩 및 첨부 카드 테스트', () => {
  it('1. EUC-KR(CP949) 바이트 배열을 자동으로 감지하여 한글 깨짐 없이 정상 디코딩해야 한다', () => {
    // '산기협 : unrimit@naver.com'의 EUC-KR 바이트 (BB EA B1 E2 C7 F9 ...)
    const euckrBytes = new Uint8Array([
      0xbb, 0xea, 0xb1, 0xe2, 0xc7, 0xf9, 0x20, 0x3a, 0x20,
      0x75, 0x6e, 0x72, 0x69, 0x6d, 0x69, 0x74, 0x40, 0x6e,
      0x61, 0x76, 0x65, 0x72, 0x2e, 0x63, 0x6f, 0x6d,
    ]);

    const result = decodeBytesSmart(euckrBytes);
    expect(result.encoding).toBe('euc-kr');
    expect(result.text).toBe('산기협 : unrimit@naver.com');
  });

  it('2. 표준 UTF-8 인코딩 텍스트는 UTF-8로 우선 감지하여 정상 디코딩해야 한다', () => {
    const originalText = '안녕하세요! MyMemo UTF-8 텍스트 테스트입니다.';
    const utf8Bytes = new TextEncoder().encode(originalText);

    const result = decodeBytesSmart(utf8Bytes);
    expect(result.encoding).toBe('utf-8');
    expect(result.text).toBe(originalText);
  });

  it('3. 순수 영문/숫자/기호(ASCII) 텍스트는 UTF-8로 정상 디코딩되어야 한다', () => {
    const originalText = 'Hello World 12345! @naver.com';
    const asciiBytes = new TextEncoder().encode(originalText);

    const result = decodeBytesSmart(asciiBytes);
    expect(result.encoding).toBe('utf-8');
    expect(result.text).toBe(originalText);
  });

  it('4. processFileForNote 함수에서 EUC-KR .txt 파일을 드롭했을 때 .bin과 동일하게 첨부 카드로 래핑되고 한글이 정상 보존되어야 한다', async () => {
    const euckrBytes = new Uint8Array([
      0xbb, 0xea, 0xb1, 0xe2, 0xc7, 0xf9, 0x20, 0x3a, 0x20,
      0x75, 0x6e, 0x72, 0x69, 0x6d, 0x69, 0x74, 0x40, 0x6e,
      0x61, 0x76, 0x65, 0x72, 0x2e, 0x63, 0x6f, 0x6d,
    ]);

    const mockFile = new File([euckrBytes], '시크릿.txt', { type: 'text/plain' });
    const processed = await processFileForNote(mockFile);

    expect(processed.type).toBe('text');
    expect(processed.title).toBe('[텍스트] 시크릿');
    expect(processed.content).toContain('[📄 텍스트: 시크릿.txt]');
    expect(processed.attachments).toBeDefined();

    const attKey = Object.keys(processed.attachments!)[0];
    const att = processed.attachments![attKey];
    expect(att.name).toBe('시크릿.txt');

    // 첨부 데이터 역디코딩 검증
    const bytes = dataUrlToUint8Array(att.data);
    const decoded = decodeBytesSmart(bytes);
    expect(decoded.encoding).toBe('euc-kr');
    expect(decoded.text).toContain('산기협 : unrimit@naver.com');
    expect(decoded.text).not.toContain('\uFFFD'); //  유니코드 대체 문자 없음
  });

  it('5. createAttachmentCardHtml에서 .txt 첨부파일에 대해 text-inline-viewer 카드와 텍스트 편집기 열기 버튼을 렌더링해야 한다', () => {
    const euckrBytes = new Uint8Array([
      0xbb, 0xea, 0xb1, 0xe2, 0xc7, 0xf9, 0x20, 0x3a, 0x20,
      0x75, 0x6e, 0x72, 0x69, 0x6d, 0x69, 0x74, 0x40, 0x6e,
      0x61, 0x76, 0x65, 0x72, 0x2e, 0x63, 0x6f, 0x6d,
    ]);
    const base64 = Buffer.from(euckrBytes).toString('base64');
    const att: AttachmentItem = {
      id: 'att-test-txt',
      name: '시크릿.txt',
      type: 'text/plain',
      data: `data:text/plain;base64,${base64}`,
    };

    const cardHtml = createAttachmentCardHtml(att, '[📄 텍스트: 시크릿.txt]');
    expect(cardHtml).toContain('text-inline-viewer');
    expect(cardHtml).toContain('TEXT / EUC-KR');
    expect(cardHtml).toContain('📝 텍스트 편집기 열기');
    expect(cardHtml).toContain('산기협 : unrimit@naver.com');
  });

  it('6. 마크다운(.md) 파일은 첨부 카드가 아니라 마크다운 본문으로 직접 불러와야 한다', async () => {
    const mdContent = '# 제목\n\n- 항목 1\n- 항목 2';
    const utf8Bytes = new TextEncoder().encode(mdContent);
    const mockMdFile = new File([utf8Bytes], 'README.md', { type: 'text/markdown' });
    const processed = await processFileForNote(mockMdFile);

    expect(processed.type).toBe('text');
    expect(processed.title).toBe('README');
    expect(processed.content).toBe(mdContent);
    expect(processed.attachments).toBeUndefined();
  });

  it('7. 실제 다운로드 폴더의 시크릿 (1).txt 파일이 존재하는 경우 검증', () => {
    const samplePath = 'c:/Users/JUNG/Downloads/시크릿 (1).txt';
    if (fs.existsSync(samplePath)) {
      const buf = fs.readFileSync(samplePath);
      const result = decodeBytesSmart(new Uint8Array(buf));
      expect(result.encoding).toBe('euc-kr');
      expect(result.text).toContain('산기협');
      expect(result.text).toContain('unrimit@naver.com');
      expect(result.text).not.toContain('\uFFFD');
    }
  });
});
