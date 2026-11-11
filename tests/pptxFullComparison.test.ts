import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { JSDOM } from 'jsdom';
import {
  replaceRunTextInSlideXml,
  deleteSlideFromPptx,
  reorderSlidesInPptx,
} from '../src/utils/pptxAdapter';
import { detectPptxUnsupportedFeatures } from '../src/utils/pptxLossDetector';

function parseXmlDoc(xmlStr: string): Document {
  const dom = new JSDOM(xmlStr, { contentType: 'text/xml' });
  return dom.window.document;
}

describe('3단계 PPTX OpenXML In-Place 검증 테스트 스위트', () => {
  const samplesDir = path.resolve('docs/office-upgrade/samples/real');
  const outDir = path.resolve('docs/office-upgrade/out');

  it('1. PPT_A 원본 슬라이드 2의 본문 run 치환 시 형광펜, 하이퍼링크, 첨자 등 형제 run 서식이 100% 보존되어야 한다', async () => {
    const rawBuf = fs.readFileSync(path.join(samplesDir, 'PPT_A_사업보고_원본.pptx'));
    const zip = await JSZip.loadAsync(rawBuf);
    const s2Xml = await zip.file('ppt/slides/slide2.xml')!.async('text');

    const { updatedXml, replaced } = replaceRunTextInSlideXml(s2Xml, '빨강', '파랑');
    expect(replaced).toBe(true);
    expect(updatedXml).toContain('파랑');
    expect(updatedXml).not.toContain('빨강');

    const doc = parseXmlDoc(updatedXml);
    // 16개 run 검사
    const runs = doc.getElementsByTagName('a:r');
    expect(runs.length).toBeGreaterThanOrEqual(15);

    // 하이퍼링크 (hlinkClick) 보존
    const hlinks = doc.getElementsByTagName('a:hlinkClick');
    expect(hlinks.length).toBeGreaterThanOrEqual(1);

    // 형광펜 (highlight) 보존
    const highlights = doc.getElementsByTagName('a:highlight');
    expect(highlights.length).toBeGreaterThanOrEqual(1);

    // 슬라이드 번호 필드 (fld) 보존
    const flds = doc.getElementsByTagName('a:fld');
    expect(flds.length).toBeGreaterThanOrEqual(1);

    // 위/아래 첨자 (baseline) 보존
    expect(updatedXml).toContain('baseline=');
    // 취소선 (strike) 보존
    expect(updatedXml).toContain('strike=');
  });

  it('2. PPT_A 슬라이드 7 삭제 시 "부록" 섹션 슬라이드 수가 2개에서 1개로 정확히 줄어야 한다', async () => {
    const rawBuf = fs.readFileSync(path.join(samplesDir, 'PPT_A_사업보고_원본.pptx'));
    const zip = await JSZip.loadAsync(rawBuf);

    // 삭제 전 섹션 검사
    const presXmlBefore = await zip.file('ppt/presentation.xml')!.async('text');
    const presDocBefore = parseXmlDoc(presXmlBefore);
    const sectionsBefore = presDocBefore.getElementsByTagName('p14:section');
    expect(sectionsBefore.length).toBe(3); // 3개 구역

    const appendixBefore = Array.from(sectionsBefore).find(
      (s) => s.getAttribute('name') === '부록'
    );
    expect(appendixBefore).toBeDefined();
    const appendixSldIdsBefore = appendixBefore!.getElementsByTagName('p14:sldId');
    expect(appendixSldIdsBefore.length).toBe(2); // 부록에 슬라이드 7, 8 총 2개

    // 슬라이드 7 삭제 (0-based 인덱스 6)
    await deleteSlideFromPptx(zip, 6);

    // 삭제 후 presentation.xml 검사
    const presXmlAfter = await zip.file('ppt/presentation.xml')!.async('text');
    const presDocAfter = parseXmlDoc(presXmlAfter);
    const sectionsAfter = presDocAfter.getElementsByTagName('p14:section');
    const appendixAfter = Array.from(sectionsAfter).find(
      (s) => s.getAttribute('name') === '부록'
    );
    expect(appendixAfter).toBeDefined();
    const appendixSldIdsAfter = appendixAfter!.getElementsByTagName('p14:sldId');
    expect(appendixSldIdsAfter.length).toBe(1); // 2개에서 1개로 감소

    // sldIdLst 슬라이드 수 감소 (8 -> 7)
    const sldIdsAfter = presDocAfter.getElementsByTagName('p:sldId');
    expect(sldIdsAfter.length).toBe(7);

    // presentation.xml.rels 에서도 관계 제거 확인
    const presRelsAfter = await zip.file('ppt/_rels/presentation.xml.rels')!.async('text');
    const relsDoc = parseXmlDoc(presRelsAfter);
    const sldRels = Array.from(relsDoc.getElementsByTagName('Relationship')).filter((r) =>
      (r.getAttribute('Type') || '').endsWith('/slide')
    );
    expect(sldRels.length).toBe(7);

    // [Content_Types].xml 에서도 슬라이드 override 제거 확인
    const ctAfter = await zip.file('[Content_Types].xml')!.async('text');
    const ctDoc = parseXmlDoc(ctAfter);
    const overrides = Array.from(ctDoc.getElementsByTagName('Override')).filter((o) =>
      (o.getAttribute('PartName') || '').includes('/ppt/slides/slide')
    );
    expect(overrides.length).toBe(7);
  });

  it('3. PPT_B 슬라이드 순서 재배열 후에도 "처음으로 이동" 점프 링크가 원래 1번 슬라이드를 정확히 가리켜야 한다', async () => {
    const rawBuf = fs.readFileSync(path.join(samplesDir, 'PPT_B_교육자료_원본.pptx'));
    const zip = await JSZip.loadAsync(rawBuf);

    // 순서 변경: [3, 2, 1, 0] (슬라이드 4개를 역순 배치)
    await reorderSlidesInPptx(zip, [3, 2, 1, 0]);

    // 슬라이드 3 (원래 꺾은선 차트와 "처음으로 이동" 링크가 있는 파일)
    const s3Xml = await zip.file('ppt/slides/slide3.xml')!.async('text');
    const s3Doc = parseXmlDoc(s3Xml);

    const hlinks = s3Doc.getElementsByTagName('a:hlinkClick');
    expect(hlinks.length).toBeGreaterThan(0);
    const jumpLink = Array.from(hlinks).find((h) =>
      (h.getAttribute('action') || '').includes('hlinksldjump')
    );
    expect(jumpLink).toBeDefined();

    // slide3.xml.rels 확인: 점프 링크의 r:id가 여전히 slide1.xml을 가리키는지 확인
    const s3Rels = await zip.file('ppt/slides/_rels/slide3.xml.rels')!.async('text');
    const s3RelsDoc = parseXmlDoc(s3Rels);
    const rId = jumpLink!.getAttribute('r:id');
    const targetRel = Array.from(s3RelsDoc.getElementsByTagName('Relationship')).find(
      (r) => r.getAttribute('Id') === rId
    );
    expect(targetRel).toBeDefined();
    expect(targetRel!.getAttribute('Target')).toContain('slide1.xml');
  });

  it('4. pptxLossDetector가 매크로, OLE, SmartArt, 미디어, 잉크, 3D 모델, 포함 글꼴을 올바르게 감지해야 한다', async () => {
    // 가상 ZIP 버퍼 생성하여 각 위험 요소 폴더 주입 테스트
    const dummyZip = new JSZip();
    dummyZip.file('ppt/presentation.xml', '<p:presentation/>');
    dummyZip.file('ppt/diagrams/data1.xml', '<dgm/>'); // SmartArt
    dummyZip.file('ppt/embeddings/oleObject1.bin', 'ole'); // OLE
    dummyZip.file('ppt/media/movie1.mp4', 'video'); // Video
    dummyZip.file('ppt/vbaProject.bin', 'macro'); // VBA Macro
    dummyZip.file('ppt/ink/ink1.xml', 'ink'); // Ink
    dummyZip.file('ppt/models/model1.glb', '3d'); // 3D model
    dummyZip.file('ppt/fonts/font1.fntdata', 'font'); // Embedded font

    const b64 = await dummyZip.generateAsync({ type: 'base64' });
    const dataUrl = `data:application/vnd.ms-powerpoint;base64,${b64}`;

    const report = await detectPptxUnsupportedFeatures(dataUrl);
    expect(report.hasSmartArt).toBe(true);
    expect(report.hasOleEmbeddings).toBe(true);
    expect(report.hasMediaAudioVideo).toBe(true);
    expect(report.hasVbaMacro).toBe(true);
    expect(report.hasInk).toBe(true);
    expect(report.has3DModel).toBe(true);
    expect(report.hasEmbeddedFonts).toBe(true);
    expect(report.unsupportedList.length).toBe(7);
  });

  it('5. 4개 수정저장결과 파일이 원본 대비 XML 구조적 무결성(마스터, 레이아웃, 차트, 표, 이미지 수)을 100% 유지해야 한다', async () => {
    const filePairs = [
      {
        orig: 'PPT_A_사업보고_원본.pptx',
        edited: 'PPT_A_사업보고_원본_수정저장결과.pptx',
      },
      {
        orig: 'PPT_A_사업보고_한쇼.pptx',
        edited: 'PPT_A_사업보고_한쇼_수정저장결과.pptx',
      },
      {
        orig: 'PPT_B_교육자료_원본.pptx',
        edited: 'PPT_B_교육자료_원본_수정저장결과.pptx',
      },
      {
        orig: 'PPT_B_교육자료_한쇼.pptx',
        edited: 'PPT_B_교육자료_한쇼_수정저장결과.pptx',
      },
    ];

    for (const pair of filePairs) {
      const origBuf = fs.readFileSync(path.join(samplesDir, pair.orig));
      const editBuf = fs.readFileSync(path.join(outDir, pair.edited));

      const origZip = await JSZip.loadAsync(origBuf);
      const editZip = await JSZip.loadAsync(editBuf);

      // 슬라이드 수 비교
      const origSlides = Object.keys(origZip.files).filter((f) =>
        f.match(/^ppt\/slides\/slide\d+\.xml$/)
      );
      const editSlides = Object.keys(editZip.files).filter((f) =>
        f.match(/^ppt\/slides\/slide\d+\.xml$/)
      );
      expect(editSlides.length).toBe(origSlides.length);

      // 마스터 및 레이아웃 수 비교
      const origMasters = Object.keys(origZip.files).filter((f) =>
        f.match(/^ppt\/slideMasters\/slideMaster\d+\.xml$/)
      );
      const editMasters = Object.keys(editZip.files).filter((f) =>
        f.match(/^ppt\/slideMasters\/slideMaster\d+\.xml$/)
      );
      expect(editMasters.length).toBe(origMasters.length);

      const origLayouts = Object.keys(origZip.files).filter((f) =>
        f.match(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/)
      );
      const editLayouts = Object.keys(editZip.files).filter((f) =>
        f.match(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/)
      );
      expect(editLayouts.length).toBe(origLayouts.length);

      // 차트 수 비교
      const origCharts = Object.keys(origZip.files).filter((f) =>
        f.match(/^ppt\/charts\/chart\d+\.xml$/)
      );
      const editCharts = Object.keys(editZip.files).filter((f) =>
        f.match(/^ppt\/charts\/chart\d+\.xml$/)
      );
      expect(editCharts.length).toBe(origCharts.length);

      // DOM 단위 요소 수 일치 검사
      let origShapes = 0;
      let editShapes = 0;
      let origTables = 0;
      let editTables = 0;
      let origPics = 0;
      let editPics = 0;

      for (const sFile of origSlides) {
        const oXml = await origZip.file(sFile)!.async('text');
        const eXml = await editZip.file(sFile)!.async('text');
        const oDoc = parseXmlDoc(oXml);
        const eDoc = parseXmlDoc(eXml);

        origShapes += oDoc.getElementsByTagName('p:sp').length;
        editShapes += eDoc.getElementsByTagName('p:sp').length;

        origTables += oDoc.getElementsByTagName('a:tbl').length;
        editTables += eDoc.getElementsByTagName('a:tbl').length;

        origPics += oDoc.getElementsByTagName('p:pic').length;
        editPics += eDoc.getElementsByTagName('p:pic').length;
      }

      expect(editShapes).toBe(origShapes);
      expect(editTables).toBe(origTables);
      expect(editPics).toBe(origPics);
    }
  });
});
