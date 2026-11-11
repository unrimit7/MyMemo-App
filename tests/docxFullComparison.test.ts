import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { JSDOM } from 'jsdom';
import { detectDocxUnsupportedFeatures } from '../src/utils/docxLossDetector';

const samplesDir = path.resolve(__dirname, '../docs/office-upgrade/samples/real');
const outDir = path.resolve(__dirname, '../docs/office-upgrade/out');

const dom = new JSDOM();
const parser = new dom.window.DOMParser();

async function parseXml(zip: JSZip, entryPath: string): Promise<Document> {
  const content = await zip.file(entryPath)!.async('text');
  return parser.parseFromString(content, 'text/xml');
}

describe('DOCX 무손실 업그레이드 전체 검증 (Step 2 - SuperDoc 실제 편집 회귀 방지 검증)', () => {
  describe('1. 워드A_사업보고서_원본: 실제 수정 및 구조 보존 검증', () => {
    const origPath = path.join(samplesDir, '워드A_사업보고서_원본.docx');
    const editedPath = path.join(outDir, '워드A_사업보고서_원본_수정저장결과.docx');

    it('원본 및 실제 SuperDoc 수정저장 결과 파일이 모두 존재해야 함', () => {
      expect(fs.existsSync(origPath)).toBe(true);
      expect(fs.existsSync(editedPath)).toBe(true);
    });

    it('수정한 본문 제목과 표 셀 텍스트가 정확히 반영되어야 함', async () => {
      const zip = await JSZip.loadAsync(fs.readFileSync(editedPath));
      const docXml = await zip.file('word/document.xml')!.async('text');

      expect(docXml.includes('2026년 하반기 사업 점검 보고서 (수정완료)')).toBe(true);
      expect(docXml.includes('개발/디자인 (개편)')).toBe(true);
    });

    it('수정한 텍스트 외 모든 XML 구조 요소 수(표, 변경추적, 구역, 각주, 메모, 서식 등)가 원본과 100% 동일해야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const editedZip = await JSZip.loadAsync(fs.readFileSync(editedPath));

      const origXml = await parseXml(origZip, 'word/document.xml');
      const editedXml = await parseXml(editedZip, 'word/document.xml');

      // 표 (1개)
      expect(editedXml.getElementsByTagName('w:tbl').length).toBe(origXml.getElementsByTagName('w:tbl').length);
      expect(editedXml.getElementsByTagName('w:tbl').length).toBe(1);
      // 세로 병합 (2개)
      expect(editedXml.getElementsByTagName('w:vMerge').length).toBe(origXml.getElementsByTagName('w:vMerge').length);
      expect(editedXml.getElementsByTagName('w:vMerge').length).toBe(2);
      // 변경 내용 추적 (삽입 3개, 삭제 1개)
      expect(editedXml.getElementsByTagName('w:ins').length).toBe(origXml.getElementsByTagName('w:ins').length);
      expect(editedXml.getElementsByTagName('w:ins').length).toBe(3);
      expect(editedXml.getElementsByTagName('w:del').length).toBe(origXml.getElementsByTagName('w:del').length);
      expect(editedXml.getElementsByTagName('w:del').length).toBe(1);
      // 구역 (2개 - 1구역 세로, 2구역 가로 2단)
      expect(editedXml.getElementsByTagName('w:sectPr').length).toBe(origXml.getElementsByTagName('w:sectPr').length);
      expect(editedXml.getElementsByTagName('w:sectPr').length).toBe(2);
      // 각주 및 메모 참조
      expect(editedXml.getElementsByTagName('w:footnoteReference').length).toBe(origXml.getElementsByTagName('w:footnoteReference').length);
      expect(editedXml.getElementsByTagName('w:footnoteReference').length).toBe(1);
      expect(editedXml.getElementsByTagName('w:commentReference').length).toBe(origXml.getElementsByTagName('w:commentReference').length);
      expect(editedXml.getElementsByTagName('w:commentReference').length).toBe(1);
      // 그림/드로잉
      expect(editedXml.getElementsByTagName('w:drawing').length).toBe(origXml.getElementsByTagName('w:drawing').length);
      expect(editedXml.getElementsByTagName('w:drawing').length).toBe(1);
      // 런 서식 요소
      expect(editedXml.getElementsByTagName('w:b').length).toBe(origXml.getElementsByTagName('w:b').length);
      expect(editedXml.getElementsByTagName('w:b').length).toBe(9);
      expect(editedXml.getElementsByTagName('w:i').length).toBe(origXml.getElementsByTagName('w:i').length);
      expect(editedXml.getElementsByTagName('w:i').length).toBe(1);
      expect(editedXml.getElementsByTagName('w:u').length).toBe(origXml.getElementsByTagName('w:u').length);
      expect(editedXml.getElementsByTagName('w:u').length).toBe(2);
      expect(editedXml.getElementsByTagName('w:strike').length).toBe(origXml.getElementsByTagName('w:strike').length);
      expect(editedXml.getElementsByTagName('w:strike').length).toBe(1);
      expect(editedXml.getElementsByTagName('w:color').length).toBe(origXml.getElementsByTagName('w:color').length);
      expect(editedXml.getElementsByTagName('w:color').length).toBe(2);
      expect(editedXml.getElementsByTagName('w:highlight').length).toBe(origXml.getElementsByTagName('w:highlight').length);
      expect(editedXml.getElementsByTagName('w:highlight').length).toBe(1);
      expect(editedXml.getElementsByTagName('w:vertAlign').length).toBe(origXml.getElementsByTagName('w:vertAlign').length);
      expect(editedXml.getElementsByTagName('w:vertAlign').length).toBe(3);
    });

    it('외부 파트(머리글, 바닥글 PAGE 필드, 각주, 메모, 미디어, 스타일)가 온전히 유지되어야 함', async () => {
      const zip = await JSZip.loadAsync(fs.readFileSync(editedPath));
      expect(zip.file('word/header1.xml')).not.toBeNull();
      expect(zip.file('word/footer1.xml')).not.toBeNull();
      expect(zip.file('word/footnotes.xml')).not.toBeNull();
      expect(zip.file('word/comments.xml')).not.toBeNull();
      expect(zip.file('word/media/image1.png')).not.toBeNull();
      expect(zip.file('word/styles.xml')).not.toBeNull();

      const footerXml = await zip.file('word/footer1.xml')!.async('text');
      expect(footerXml.includes('PAGE') || footerXml.includes('w:fldSimple')).toBe(true);
    });
  });

  describe('2. 워드A_사업보고서_한글: 실제 수정 및 한글 2018 호환 검증', () => {
    const origPath = path.join(samplesDir, '워드A_사업보고서_한글.docx');
    const editedPath = path.join(outDir, '워드A_사업보고서_한글_수정저장결과.docx');

    it('수정한 본문 제목과 표 셀 텍스트가 정확히 반영되어야 함', async () => {
      const zip = await JSZip.loadAsync(fs.readFileSync(editedPath));
      const docXml = await zip.file('word/document.xml')!.async('text');

      expect(docXml.includes('2026년 하반기 사업 점검 보고서 (수정완료)')).toBe(true);
      expect(docXml.includes('개발/디자인 (개편)')).toBe(true);
    });

    it('수정한 텍스트 외 모든 XML 구조 요소 수와 한글 고유 HYPERLINK 필드가 원본과 동일해야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const editedZip = await JSZip.loadAsync(fs.readFileSync(editedPath));

      const origXml = await parseXml(origZip, 'word/document.xml');
      const editedXml = await parseXml(editedZip, 'word/document.xml');

      expect(editedXml.getElementsByTagName('w:tbl').length).toBe(1);
      expect(editedXml.getElementsByTagName('w:vMerge').length).toBe(2);
      expect(editedXml.getElementsByTagName('w:ins').length).toBe(3);
      expect(editedXml.getElementsByTagName('w:del').length).toBe(1);
      expect(editedXml.getElementsByTagName('w:sectPr').length).toBe(2);
      expect(editedXml.getElementsByTagName('w:footnoteReference').length).toBe(1);
      expect(editedXml.getElementsByTagName('w:commentReference').length).toBe(1);
      expect(editedXml.getElementsByTagName('w:drawing').length).toBe(1);

      const editedDocXml = await editedZip.file('word/document.xml')!.async('text');
      expect(editedDocXml.includes('HYPERLINK')).toBe(true);
    });
  });

  describe('3. 워드B_신청서양식_원본: 실제 수정 및 양식 컨트롤 보존 검증', () => {
    const origPath = path.join(samplesDir, '워드B_신청서양식_원본.docx');
    const editedPath = path.join(outDir, '워드B_신청서양식_원본_수정저장결과.docx');

    it('수정한 성명 칸 본문과 표 셀(사용 목적) 텍스트가 정확히 반영되어야 함', async () => {
      const zip = await JSZip.loadAsync(fs.readFileSync(editedPath));
      const docXml = await zip.file('word/document.xml')!.async('text');

      expect(docXml.includes('김가상 (수정완료)')).toBe(true);
      expect(docXml.includes('외부 고객 시연회 진행 (개편)')).toBe(true);
    });

    it('수정한 텍스트 외 모든 컨트롤 요소 수(sdt 5개, 체크박스 3개, 드롭다운 1개, 표 1개, 글상자 1개)가 원본과 동일해야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const editedZip = await JSZip.loadAsync(fs.readFileSync(editedPath));

      const origXml = await parseXml(origZip, 'word/document.xml');
      const editedXml = await parseXml(editedZip, 'word/document.xml');

      // 표 (1개)
      expect(editedXml.getElementsByTagName('w:tbl').length).toBe(1);
      // 열 병합 (4개)
      expect(editedXml.getElementsByTagName('w:gridSpan').length).toBe(4);
      // SDT 내용 컨트롤 (5개)
      expect(editedXml.getElementsByTagName('w:sdt').length).toBe(5);
      // 체크박스 (3개)
      expect(editedXml.getElementsByTagName('w14:checkbox').length).toBe(3);
      // 드롭다운 목록 (1개)
      expect(editedXml.getElementsByTagName('w:dropDownList').length).toBe(1);
      // 글상자 (1개)
      expect(editedXml.getElementsByTagName('w:txbxContent').length).toBe(1);
      // 책갈피 (1개)
      const origBookmarks = Array.from(origXml.getElementsByTagName('w:bookmarkStart')).filter(b => b.getAttribute('w:name') !== '_GoBack');
      const editedBookmarks = Array.from(editedXml.getElementsByTagName('w:bookmarkStart')).filter(b => b.getAttribute('w:name') !== '_GoBack');
      expect(editedBookmarks.length).toBe(origBookmarks.length);
      expect(editedBookmarks.length).toBe(1);
      // 구역 (1개)
      expect(editedXml.getElementsByTagName('w:sectPr').length).toBe(1);
    });

    it('첫 페이지 다른 머리글(header1/header2) 및 바닥글이 온전히 유지되어야 함', async () => {
      const zip = await JSZip.loadAsync(fs.readFileSync(editedPath));
      expect(zip.file('word/header1.xml')).not.toBeNull();
      expect(zip.file('word/header2.xml')).not.toBeNull();
      expect(zip.file('word/footer1.xml')).not.toBeNull();
    });
  });

  describe('4. 워드B_신청서양식_한글: 실제 수정 및 한글 DrawingML/VML 호환 검증', () => {
    const origPath = path.join(samplesDir, '워드B_신청서양식_한글.docx');
    const editedPath = path.join(outDir, '워드B_신청서양식_한글_수정저장결과.docx');

    it('수정한 성명 칸 본문과 표 셀 텍스트가 정확히 반영되어야 함', async () => {
      const zip = await JSZip.loadAsync(fs.readFileSync(editedPath));
      const docXml = await zip.file('word/document.xml')!.async('text');

      expect(docXml.includes('김가상 (수정완료)')).toBe(true);
      expect(docXml.includes('외부 고객 시연회 진행 (개편)')).toBe(true);
    });

    it('한글 특화 AlternateContent 글상자(2개) 및 체크박스 기호(☒, ☐)가 온전히 유지되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const editedZip = await JSZip.loadAsync(fs.readFileSync(editedPath));

      const origXml = await parseXml(origZip, 'word/document.xml');
      const editedXml = await parseXml(editedZip, 'word/document.xml');

      expect(editedXml.getElementsByTagName('w:tbl').length).toBe(1);
      expect(editedXml.getElementsByTagName('w:sdt').length).toBe(5);
      expect(editedXml.getElementsByTagName('w:txbxContent').length).toBe(2);
      expect(editedXml.getElementsByTagName('w:gridSpan').length).toBe(4);

      const editedDocXml = await editedZip.file('word/document.xml')!.async('text');
      expect(editedDocXml.includes('☒') && editedDocXml.includes('☐')).toBe(true);
    });
  });

  describe('5. 손실 감지기 (docxLossDetector) 검증', () => {
    it('일반 문서는 빈 경고 목록을 반환해야 함', async () => {
      const buf = fs.readFileSync(path.join(samplesDir, '워드A_사업보고서_원본.docx'));
      const dataUrl = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${buf.toString('base64')}`;
      const report = await detectDocxUnsupportedFeatures(dataUrl);

      expect(report.hasVbaMacro).toBe(false);
      expect(report.hasOleEmbeddings).toBe(false);
      expect(report.unsupportedList).toEqual([]);
    });

    it('VBA 매크로 또는 OLE 개체가 포함된 경우 올바르게 감지하여 경고해야 함', async () => {
      const zip = new JSZip();
      zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types/>');
      zip.file('word/vbaProject.bin', 'fake macro binary data');
      zip.file('word/embeddings/oleObject1.bin', 'fake ole binary');
      const zipBuf = await zip.generateAsync({ type: 'nodebuffer' });
      const dataUrl = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${zipBuf.toString('base64')}`;

      const report = await detectDocxUnsupportedFeatures(dataUrl);
      expect(report.hasVbaMacro).toBe(true);
      expect(report.hasOleEmbeddings).toBe(true);
      expect(report.unsupportedList.length).toBe(2);
      expect(report.unsupportedList).toContain('VBA 매크로 코드');
      expect(report.unsupportedList).toContain('OLE 삽입 개체 (임베디드 오피스 파일)');
    });
  });
});
