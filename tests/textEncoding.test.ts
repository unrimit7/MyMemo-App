import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { decodeBytesSmart } from '../src/utils/hexViewer';
import { processFileForNote } from '../src/utils/file';

describe('텍스트 파일 인코딩(UTF-8 및 ANSI/EUC-KR) 자동 감지 디코딩 테스트', () => {
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

  it('4. processFileForNote 함수에서 EUC-KR 텍스트 파일을 드롭했을 때 깨짐() 없이 한글이 복원되어야 한다', async () => {
    const euckrBytes = new Uint8Array([
      0xbb, 0xea, 0xb1, 0xe2, 0xc7, 0xf9, 0x20, 0x3a, 0x20,
      0x75, 0x6e, 0x72, 0x69, 0x6d, 0x69, 0x74, 0x40, 0x6e,
      0x61, 0x76, 0x65, 0x72, 0x2e, 0x63, 0x6f, 0x6d,
    ]);

    const mockFile = new File([euckrBytes], '시크릿.txt', { type: 'text/plain' });
    const processed = await processFileForNote(mockFile);

    expect(processed.type).toBe('text');
    expect(processed.title).toBe('시크릿');
    expect(processed.content).toContain('산기협 : unrimit@naver.com');
    expect(processed.content).not.toContain('\uFFFD'); //  유니코드 대체 문자 없음
  });

  it('5. 실제 다운로드 폴더의 시크릿 (1).txt 파일이 존재하는 경우 검증', () => {
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
