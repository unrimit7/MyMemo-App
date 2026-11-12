import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { JSDOM } from 'jsdom';

if (typeof globalThis.DOMParser === 'undefined') {
  const dom = new JSDOM();
  globalThis.DOMParser = dom.window.DOMParser;
  globalThis.XMLSerializer = dom.window.XMLSerializer;
}
import {
  replaceRunTextInSlideXml,
  deleteSlideFromPptx,
  reorderSlidesInPptx,
  savePptxPresentation,
} from '../src/utils/pptxAdapter';
import { parsePptxPresentation } from '../src/utils/pptxViewer';
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

      // ZIP 디렉터리 항목 검사: JSZip에 의한 임의의 디렉터리 항목(ppt/, ppt/slides/ 등)이 원본 대비 추가되지 않아야 함
      const origDirs = Object.keys(origZip.files).filter((k) => origZip.files[k].dir).sort();
      const editDirs = Object.keys(editZip.files).filter((k) => editZip.files[k].dir).sort();
      expect(editDirs).toEqual(origDirs);
    }
  });

  it('6. 앱 편집기 및 테스트 공통 savePptxPresentation 함수가 diff->run 치환 및 불필요 폴더 제거를 완벽 수행해야 한다', async () => {
    const rawBuf = fs.readFileSync(path.join(samplesDir, 'PPT_A_사업보고_원본.pptx'));
    const b64 = rawBuf.toString('base64');
    const dataUrl = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${b64}`;

    // 슬라이드 2의 본문 단락 수정 (단일 런 치환)
    const editableSlides = [
      {
        slideNumber: 1,
        title: '사업보고',
        bodyText: '',
      },
      {
        slideNumber: 2,
        title: '2. 사업 추진 경과',
        bodyText: '• 3월: [알림] 신규 서비스 출시 (파랑 테스트 완료)',
      },
    ];

    const savedUrl = await savePptxPresentation(dataUrl, editableSlides);
    expect(savedUrl).toContain('data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,');
    const savedB64 = savedUrl.replace(/^data:[^;]+;base64,/, '');
    const savedZip = await JSZip.loadAsync(Buffer.from(savedB64, 'base64'));

    // 치환된 슬라이드 2 XML 확인
    const s2Xml = await savedZip.file('ppt/slides/slide2.xml')!.async('text');
    expect(s2Xml).toContain('파랑');

    // 형제 run (서식, 하이퍼링크 등) 보존 확인
    expect(s2Xml).toContain('<a:highlight');
    expect(s2Xml).toContain('<a:hlinkClick');

    // 디렉터리 항목 추가 여부 확인 (원본과 동일해야 함)
    const origZip = await JSZip.loadAsync(rawBuf);
    const origDirs = Object.keys(origZip.files).filter((k) => origZip.files[k].dir).sort();
    const savedDirs = Object.keys(savedZip.files).filter((k) => savedZip.files[k].dir).sort();
    expect(savedDirs).toEqual(origDirs);
  });

  it('7. 무변경 저장 테스트(zero-change roundtrip): 4개 파일 모두 편집 없이 저장했을 때 모든 slideN.xml 텍스트가 원본과 100% 동일해야 한다', async () => {
    const sampleFiles = [
      'PPT_A_사업보고_원본.pptx',
      'PPT_A_사업보고_한쇼.pptx',
      'PPT_B_교육자료_원본.pptx',
      'PPT_B_교육자료_한쇼.pptx',
    ];

    for (const fileName of sampleFiles) {
      const origBuf = fs.readFileSync(path.join(samplesDir, fileName));
      const b64 = origBuf.toString('base64');
      const dataUrl = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${b64}`;

      const pres = await parsePptxPresentation(dataUrl);
      const editableSlides = pres.slides.map((s, idx) => {
        const bodyLines = s.paragraphs
          .filter((p) => !p.isTitle && p.text !== s.title)
          .map((p) => {
            const indent = '  '.repeat(Math.max(0, p.level));
            return `${indent}• ${p.text}`;
          });

        const slideTables: { rows: string[][] }[] = [];
        for (const el of s.elements) {
          if (el.type === 'table') {
            const tableRows = el.rows.map((row) =>
              row.cells.map((cell) => cell.text)
            );
            slideTables.push({ rows: tableRows });
          }
        }

        return {
          slideNumber: s.slideNumber,
          originalSlideIndex: idx,
          title: s.title,
          bodyText: bodyLines.join('\n'),
          tables: slideTables,
        };
      });

      // 무변경 상태로 저장 실행
      const savedDataUrl = await savePptxPresentation(dataUrl, editableSlides);
      const savedB64 = savedDataUrl.replace(/^data:[^;]+;base64,/, '');
      const savedZip = await JSZip.loadAsync(Buffer.from(savedB64, 'base64'));
      const origZip = await JSZip.loadAsync(origBuf);

      const slideFiles = Object.keys(origZip.files).filter((f) =>
        f.match(/^ppt\/slides\/slide\d+\.xml$/)
      );

      for (const sFile of slideFiles) {
        const origXml = await origZip.file(sFile)!.async('text');
        const savedXml = await savedZip.file(sFile)!.async('text');
        // 변경이 없는 슬라이드는 원본 XML 텍스트가 바이트 단위로 100% 동일해야 함
        expect(savedXml).toBe(origXml);
      }
    }
  });

  it('8. 4개 파일 실제 편집 저장 및 out/ 생성: PPT_A(빨강->파랑, 영업->영업(수정)) 및 PPT_B(세로쓰기예시->(수정), 첫째 항목->(수정))', async () => {
    // 1) PPT_A_사업보고_원본
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'PPT_A_사업보고_원본.pptx'));
      const dataUrl = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${origBuf.toString('base64')}`;
      const pres = await parsePptxPresentation(dataUrl);
      const editableSlides = pres.slides.map((s, idx) => ({
        slideNumber: s.slideNumber,
        originalSlideIndex: idx,
        title: s.title,
        bodyText: s.paragraphs
          .filter((p) => !p.isTitle && p.text !== s.title)
          .map((p) => `${'  '.repeat(Math.max(0, p.level))}• ${p.text}`)
          .join('\n'),
        tables: s.elements
          .filter((el) => el.type === 'table')
          .map((el: any) => ({ rows: el.rows.map((r: any) => r.cells.map((c: any) => c.text)) })),
      }));

      // Slide 2 본문 "빨강" -> "파랑"
      editableSlides[1].bodyText = editableSlides[1].bodyText.replace('빨강', '파랑');
      // Slide 3 표 셀 "영업" -> "영업(수정)"
      if (editableSlides[2].tables && editableSlides[2].tables.length > 0) {
        editableSlides[2].tables[0].rows[1][0] = '영업(수정)';
      }

      const savedUrl = await savePptxPresentation(dataUrl, editableSlides);
      const outBuf = Buffer.from(savedUrl.replace(/^data:[^;]+;base64,/, ''), 'base64');
      fs.writeFileSync(path.join(outDir, 'PPT_A_사업보고_원본_수정저장결과.pptx'), outBuf);
    }

    // 2) PPT_A_사업보고_한쇼
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'PPT_A_사업보고_한쇼.pptx'));
      const dataUrl = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${origBuf.toString('base64')}`;
      const pres = await parsePptxPresentation(dataUrl);
      const editableSlides = pres.slides.map((s, idx) => ({
        slideNumber: s.slideNumber,
        originalSlideIndex: idx,
        title: s.title,
        bodyText: s.paragraphs
          .filter((p) => !p.isTitle && p.text !== s.title)
          .map((p) => `${'  '.repeat(Math.max(0, p.level))}• ${p.text}`)
          .join('\n'),
        tables: s.elements
          .filter((el) => el.type === 'table')
          .map((el: any) => ({ rows: el.rows.map((r: any) => r.cells.map((c: any) => c.text)) })),
      }));

      editableSlides[1].bodyText = editableSlides[1].bodyText.replace('빨강', '파랑');
      if (editableSlides[2].tables && editableSlides[2].tables.length > 0) {
        editableSlides[2].tables[0].rows[1][0] = '영업(수정)';
      }

      const savedUrl = await savePptxPresentation(dataUrl, editableSlides);
      const outBuf = Buffer.from(savedUrl.replace(/^data:[^;]+;base64,/, ''), 'base64');
      fs.writeFileSync(path.join(outDir, 'PPT_A_사업보고_한쇼_수정저장결과.pptx'), outBuf);
    }

    // 3) PPT_B_교육자료_원본
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'PPT_B_교육자료_원본.pptx'));
      const dataUrl = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${origBuf.toString('base64')}`;
      const pres = await parsePptxPresentation(dataUrl);
      const editableSlides = pres.slides.map((s, idx) => ({
        slideNumber: s.slideNumber,
        originalSlideIndex: idx,
        title: s.title,
        bodyText: s.paragraphs
          .filter((p) => !p.isTitle && p.text !== s.title)
          .map((p) => `${'  '.repeat(Math.max(0, p.level))}• ${p.text}`)
          .join('\n'),
        tables: s.elements
          .filter((el) => el.type === 'table')
          .map((el: any) => ({ rows: el.rows.map((r: any) => r.cells.map((c: any) => c.text)) })),
      }));

      editableSlides[1].bodyText = editableSlides[1].bodyText.replace('세로쓰기예시', '세로쓰기예시(수정)');
      editableSlides[3].bodyText = editableSlides[3].bodyText.replace('첫째 항목', '첫째 항목(수정)');

      const savedUrl = await savePptxPresentation(dataUrl, editableSlides);
      const outBuf = Buffer.from(savedUrl.replace(/^data:[^;]+;base64,/, ''), 'base64');
      fs.writeFileSync(path.join(outDir, 'PPT_B_교육자료_원본_수정저장결과.pptx'), outBuf);
    }

    // 4) PPT_B_교육자료_한쇼
    {
      const origBuf = fs.readFileSync(path.join(samplesDir, 'PPT_B_교육자료_한쇼.pptx'));
      const dataUrl = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${origBuf.toString('base64')}`;
      const pres = await parsePptxPresentation(dataUrl);
      const editableSlides = pres.slides.map((s, idx) => ({
        slideNumber: s.slideNumber,
        originalSlideIndex: idx,
        title: s.title,
        bodyText: s.paragraphs
          .filter((p) => !p.isTitle && p.text !== s.title)
          .map((p) => `${'  '.repeat(Math.max(0, p.level))}• ${p.text}`)
          .join('\n'),
        tables: s.elements
          .filter((el) => el.type === 'table')
          .map((el: any) => ({ rows: el.rows.map((r: any) => r.cells.map((c: any) => c.text)) })),
      }));

      editableSlides[1].bodyText = editableSlides[1].bodyText.replace('세로쓰기예시', '세로쓰기예시(수정)');
      editableSlides[3].bodyText = editableSlides[3].bodyText.replace('첫째 항목', '첫째 항목(수정)');

      const savedUrl = await savePptxPresentation(dataUrl, editableSlides);
      const outBuf = Buffer.from(savedUrl.replace(/^data:[^;]+;base64,/, ''), 'base64');
      fs.writeFileSync(path.join(outDir, 'PPT_B_교육자료_한쇼_수정저장결과.pptx'), outBuf);
    }
  });

  it('9. 슬라이드별 전체 텍스트 비교: 수정한 문단 외의 모든 슬라이드/문단/표 텍스트가 원본과 100% 일치해야 한다', async () => {
    const filePairs = [
      {
        orig: 'PPT_A_사업보고_원본.pptx',
        edited: 'PPT_A_사업보고_원본_수정저장결과.pptx',
        isA: true,
      },
      {
        orig: 'PPT_A_사업보고_한쇼.pptx',
        edited: 'PPT_A_사업보고_한쇼_수정저장결과.pptx',
        isA: true,
      },
      {
        orig: 'PPT_B_교육자료_원본.pptx',
        edited: 'PPT_B_교육자료_원본_수정저장결과.pptx',
        isA: false,
      },
      {
        orig: 'PPT_B_교육자료_한쇼.pptx',
        edited: 'PPT_B_교육자료_한쇼_수정저장결과.pptx',
        isA: false,
      },
    ];

    for (const pair of filePairs) {
      const origBuf = fs.readFileSync(path.join(samplesDir, pair.orig));
      const editBuf = fs.readFileSync(path.join(outDir, pair.edited));
      const origPres = await parsePptxPresentation(`data:application/vnd.ms-powerpoint;base64,${origBuf.toString('base64')}`);
      const editPres = await parsePptxPresentation(`data:application/vnd.ms-powerpoint;base64,${editBuf.toString('base64')}`);

      expect(editPres.slides.length).toBe(origPres.slides.length);

      for (let sIdx = 0; sIdx < origPres.slides.length; sIdx++) {
        const oSlide = origPres.slides[sIdx];
        const eSlide = editPres.slides[sIdx];

        // 1. 모든 슬라이드의 제목은 100% 유지되어야 함 (‹#›나 다른 텍스트로 덮어써지지 않음)
        expect(eSlide.title).toBe(oSlide.title);
        expect(eSlide.title).not.toBe('‹#›');
        expect(eSlide.title).not.toBe('<#>');

        // 2. 표(Table) 검사
        const oTables = oSlide.elements.filter((el) => el.type === 'table') as any[];
        const eTables = eSlide.elements.filter((el) => el.type === 'table') as any[];
        expect(eTables.length).toBe(oTables.length);

        for (let tIdx = 0; tIdx < oTables.length; tIdx++) {
          const ot = oTables[tIdx];
          const et = eTables[tIdx];
          expect(et.rows.length).toBe(ot.rows.length);

          for (let r = 0; r < ot.rows.length; r++) {
            expect(et.rows[r].cells.length).toBe(ot.rows[r].cells.length);
            for (let c = 0; c < ot.rows[r].cells.length; c++) {
              const otText = ot.rows[r].cells[c].text;
              const etText = et.rows[r].cells[c].text;

              if (pair.isA && sIdx === 2 && r === 1 && c === 0) {
                // 수정된 셀: "영업" -> "영업(수정)"
                expect(etText).toBe('영업(수정)');
                expect(otText).toBe('영업');
              } else {
                // 그 외 모든 셀은 100% 원본과 동일해야 함
                expect(etText).toBe(otText);
              }
            }
          }
        }

        // 3. 문단(Paragraphs) 검사
        const oParas = oSlide.paragraphs.filter((p) => !p.isTitle && p.text !== oSlide.title);
        const eParas = eSlide.paragraphs.filter((p) => !p.isTitle && p.text !== eSlide.title);
        expect(eParas.length).toBe(oParas.length);

        for (let pIdx = 0; pIdx < oParas.length; pIdx++) {
          const opText = oParas[pIdx].text;
          const epText = eParas[pIdx].text;

          if (pair.isA && sIdx === 1 && opText.includes('빨강')) {
            // PPT_A 슬라이드 2 수정 문단
            expect(epText).toBe(opText.replace('빨강', '파랑'));
          } else if (!pair.isA && sIdx === 1 && opText.includes('세로쓰기예시')) {
            // PPT_B 슬라이드 2 수정 문단
            expect(epText).toBe(opText.replace('세로쓰기예시', '세로쓰기예시(수정)'));
          } else if (!pair.isA && sIdx === 3 && opText.includes('첫째 항목')) {
            // PPT_B 슬라이드 4 수정 문단
            expect(epText).toBe(opText.replace('첫째 항목', '첫째 항목(수정)'));
          } else {
            // 수정하지 않은 모든 문단은 100% 동일해야 함
            expect(epText).toBe(opText);
          }
        }
      }
    }
  });
});
