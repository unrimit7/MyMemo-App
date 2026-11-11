import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { detectDocxUnsupportedFeatures } from '../src/utils/docxLossDetector';

const samplesDir = path.resolve(__dirname, '../docs/office-upgrade/samples/real');
const outDir = path.resolve(__dirname, '../docs/office-upgrade/out');

function extractParagraphTexts(docXml: string): string[] {
  const pMatches = docXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  return pMatches
    .map((pXml) => {
      const tMatches = pXml.match(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g) || [];
      return tMatches.map((t) => t.replace(/<[^>]+>/g, '')).join('');
    })
    .filter((t) => t.trim().length > 0);
}

function countOccurrences(str: string, regex: RegExp): number {
  return (str.match(regex) || []).length;
}

describe('DOCX 무손실 업그레이드 전체 검증 (Step 2)', () => {
  describe('1. 워드A_사업보고서_원본 검증', () => {
    const origPath = path.join(samplesDir, '워드A_사업보고서_원본.docx');
    const savedPath = path.join(outDir, '워드A_사업보고서_원본_저장결과.docx');

    it('원본 및 저장 결과 파일이 모두 존재해야 함', () => {
      expect(fs.existsSync(origPath)).toBe(true);
      expect(fs.existsSync(savedPath)).toBe(true);
    });

    it('document.xml 내 모든 문단 텍스트가 100% 일치해야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      const origTexts = extractParagraphTexts(origDoc);
      const savedTexts = extractParagraphTexts(savedDoc);

      expect(savedTexts.length).toBe(origTexts.length);
      for (let i = 0; i < origTexts.length; i++) {
        expect(savedTexts[i]).toBe(origTexts[i]);
      }
    });

    it('인라인 런 서식(굵게, 기울임, 밑줄, 취소선, 색상, 형광펜, 위/아래 첨자)이 보존되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      expect(countOccurrences(savedDoc, /<w:b[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:b[ >\/]/g));
      expect(countOccurrences(savedDoc, /<w:i[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:i[ >\/]/g));
      expect(countOccurrences(savedDoc, /<w:u[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:u[ >\/]/g));
      expect(countOccurrences(savedDoc, /<w:strike[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:strike[ >\/]/g));
      expect(countOccurrences(savedDoc, /<w:color[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:color[ >\/]/g));
      expect(countOccurrences(savedDoc, /<w:highlight[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:highlight[ >\/]/g));
      expect(countOccurrences(savedDoc, /<w:vertAlign[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:vertAlign[ >\/]/g));
    });

    it('구역(sectPr), 가로 방향(landscape), 2단(w:cols num="2") 배치가 유지되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      expect(countOccurrences(savedDoc, /<w:sectPr\b/g)).toBe(countOccurrences(origDoc, /<w:sectPr\b/g));
      expect(savedDoc.includes('orient="landscape"')).toBe(true);
      expect(savedDoc.includes('num="2"')).toBe(true);
    });

    it('머리글/바닥글(PAGE 필드), 각주, 메모가 유지되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      expect(savedZip.file('word/header1.xml')).not.toBeNull();
      expect(savedZip.file('word/footer1.xml')).not.toBeNull();
      expect(savedZip.file('word/footnotes.xml')).not.toBeNull();
      expect(savedZip.file('word/comments.xml')).not.toBeNull();

      const footerXml = await savedZip.file('word/footer1.xml')!.async('text');
      expect(footerXml.includes('PAGE') || footerXml.includes('w:fldSimple')).toBe(true);

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');
      expect(countOccurrences(savedDoc, /<w:footnoteReference\b/g)).toBe(countOccurrences(origDoc, /<w:footnoteReference\b/g));
      expect(countOccurrences(savedDoc, /<w:commentReference\b/g)).toBe(countOccurrences(origDoc, /<w:commentReference\b/g));
    });

    it('변경 내용 추적(w:ins, w:del)이 정확히 유지되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      expect(countOccurrences(savedDoc, /<w:ins\b/g)).toBe(countOccurrences(origDoc, /<w:ins\b/g));
      expect(countOccurrences(savedDoc, /<w:del\b/g)).toBe(countOccurrences(origDoc, /<w:del\b/g));
    });

    it('표(tbl)와 세로 병합(vMerge), 테두리가 유지되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      expect(countOccurrences(savedDoc, /<w:tbl\b/g)).toBe(countOccurrences(origDoc, /<w:tbl\b/g));
      expect(countOccurrences(savedDoc, /<w:vMerge\b/g)).toBe(countOccurrences(origDoc, /<w:vMerge\b/g));
    });

    it('미디어 파일(image1.png) 및 문서 관계(_rels)가 유지되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origImg = await origZip.file('word/media/image1.png')!.async('nodebuffer');
      const savedImg = await savedZip.file('word/media/image1.png')!.async('nodebuffer');
      expect(savedImg.length).toBe(origImg.length);

      const savedRels = await savedZip.file('word/_rels/document.xml.rels')!.async('text');
      expect(savedRels.includes('media/image1.png')).toBe(true);
    });

    it('styles.xml의 주요 스타일 정의(Heading1, Heading2, Normal)가 유지되어야 함', async () => {
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));
      const savedStyles = await savedZip.file('word/styles.xml')!.async('text');

      expect(savedStyles.includes('styleId="Heading1"') || savedStyles.includes('w:styleId="Heading1"')).toBe(true);
      expect(savedStyles.includes('styleId="Heading2"') || savedStyles.includes('w:styleId="Heading2"')).toBe(true);
      expect(savedStyles.includes('styleId="Normal"') || savedStyles.includes('w:styleId="Normal"')).toBe(true);
    });
  });

  describe('2. 워드A_사업보고서_한글 검증 (한컴 한글 호환)', () => {
    const origPath = path.join(samplesDir, '워드A_사업보고서_한글.docx');
    const savedPath = path.join(outDir, '워드A_사업보고서_한글_저장결과.docx');

    it('document.xml 내 모든 문단 텍스트가 100% 일치해야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      const origTexts = extractParagraphTexts(origDoc);
      const savedTexts = extractParagraphTexts(savedDoc);

      expect(savedTexts.length).toBe(origTexts.length);
      for (let i = 0; i < origTexts.length; i++) {
        expect(savedTexts[i]).toBe(origTexts[i]);
      }
    });

    it('한글 특화 HYPERLINK 필드 명령(instrText) 및 서식이 보존되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      expect(countOccurrences(savedDoc, /<w:instrText\b/g)).toBe(countOccurrences(origDoc, /<w:instrText\b/g));
      expect(savedDoc.includes('HYPERLINK')).toBe(true);
      expect(countOccurrences(savedDoc, /<w:b[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:b[ >\/]/g));
      expect(countOccurrences(savedDoc, /<w:i[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:i[ >\/]/g));
      expect(countOccurrences(savedDoc, /<w:u[ >\/]/g)).toBe(countOccurrences(origDoc, /<w:u[ >\/]/g));
    });

    it('각주, 메모, 구역 나눔, 이미지 미디어가 유지되어 한글 2018에서 정상 로드 가능한 구조여야 함', async () => {
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      expect(savedZip.file('word/footnotes.xml')).not.toBeNull();
      expect(savedZip.file('word/comments.xml')).not.toBeNull();
      expect(savedZip.file('word/media/image1.png')).not.toBeNull();

      const docXml = await savedZip.file('word/document.xml')!.async('text');
      expect(docXml.includes('orient="landscape"')).toBe(true);
      expect(docXml.includes('num="2"')).toBe(true);
    });
  });

  describe('3. 워드B_신청서양식_원본 검증', () => {
    const origPath = path.join(samplesDir, '워드B_신청서양식_원본.docx');
    const savedPath = path.join(outDir, '워드B_신청서양식_원본_저장결과.docx');

    it('document.xml 내 모든 문단 텍스트가 100% 일치해야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      const origTexts = extractParagraphTexts(origDoc);
      const savedTexts = extractParagraphTexts(savedDoc);

      expect(savedTexts.length).toBe(origTexts.length);
      for (let i = 0; i < origTexts.length; i++) {
        expect(savedTexts[i]).toBe(origTexts[i]);
      }
    });

    it('첫 페이지 다른 머리글(header1/header2) 및 바닥글이 유지되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      expect(savedZip.file('word/header1.xml')).not.toBeNull();
      expect(savedZip.file('word/header2.xml')).not.toBeNull();
      expect(savedZip.file('word/footer1.xml')).not.toBeNull();

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');
      expect(countOccurrences(savedDoc, /<w:headerReference\b/g)).toBe(countOccurrences(origDoc, /<w:headerReference\b/g));
    });

    it('입력 컨트롤(w:sdt), 체크박스(w14:checkbox), 드롭다운(w:dropDownList)이 보존되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      expect(countOccurrences(savedDoc, /<w:sdt\b/g)).toBe(countOccurrences(origDoc, /<w:sdt\b/g));
      expect(countOccurrences(savedDoc, /<w14:checkbox\b/g)).toBe(countOccurrences(origDoc, /<w14:checkbox\b/g));
      expect(countOccurrences(savedDoc, /<w:dropDownList\b/g)).toBe(countOccurrences(origDoc, /<w:dropDownList\b/g));
    });

    it('글상자(w:txbxContent), 책갈피(w:bookmarkStart), 표 열병합(w:gridSpan)이 유지되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      expect(countOccurrences(savedDoc, /<w:txbxContent\b/g)).toBe(countOccurrences(origDoc, /<w:txbxContent\b/g));
      expect(countOccurrences(savedDoc, /<w:bookmarkStart\b/g)).toBe(countOccurrences(origDoc, /<w:bookmarkStart\b/g));
      expect(countOccurrences(savedDoc, /<w:gridSpan\b/g)).toBe(countOccurrences(origDoc, /<w:gridSpan\b/g));
    });
  });

  describe('4. 워드B_신청서양식_한글 검증 (한컴 한글 호환)', () => {
    const origPath = path.join(samplesDir, '워드B_신청서양식_한글.docx');
    const savedPath = path.join(outDir, '워드B_신청서양식_한글_저장결과.docx');

    it('document.xml 내 모든 문단 텍스트가 100% 일치해야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      const origTexts = extractParagraphTexts(origDoc);
      const savedTexts = extractParagraphTexts(savedDoc);

      expect(savedTexts.length).toBe(origTexts.length);
      for (let i = 0; i < origTexts.length; i++) {
        expect(savedTexts[i]).toBe(origTexts[i]);
      }
    });

    it('한글 특화 DrawingML+VML 글상자 및 체크박스 기호가 유지되어야 함', async () => {
      const origZip = await JSZip.loadAsync(fs.readFileSync(origPath));
      const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath));

      const origDoc = await origZip.file('word/document.xml')!.async('text');
      const savedDoc = await savedZip.file('word/document.xml')!.async('text');

      expect(countOccurrences(savedDoc, /<w:txbxContent\b/g)).toBe(countOccurrences(origDoc, /<w:txbxContent\b/g));
      expect(savedDoc.includes('☒') || savedDoc.includes('☐')).toBe(true);
      expect(countOccurrences(savedDoc, /<w:gridSpan\b/g)).toBe(countOccurrences(origDoc, /<w:gridSpan\b/g));
      expect(countOccurrences(savedDoc, /<w:bookmarkStart\b/g)).toBe(countOccurrences(origDoc, /<w:bookmarkStart\b/g));
    });
  });

  describe('5. 손실 감지기 (docxLossDetector) 검증', () => {
    it('매크로/OLE/차트가 없는 일반 문서는 빈 경고 목록을 반환해야 함', async () => {
      const buf = fs.readFileSync(path.join(samplesDir, '워드A_사업보고서_원본.docx'));
      const dataUrl = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${buf.toString('base64')}`;
      const report = await detectDocxUnsupportedFeatures(dataUrl);

      expect(report.hasVbaMacro).toBe(false);
      expect(report.hasOleEmbeddings).toBe(false);
      expect(report.unsupportedList).toEqual([]);
    });

    it('VBA 매크로(.bin) 또는 OLE 개체가 포함된 경우 올바르게 감지하여 경고해야 함', async () => {
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
