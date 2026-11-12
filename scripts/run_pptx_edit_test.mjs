import fs from 'fs';
import path from 'path';
import { JSDOM } from 'jsdom';

if (typeof globalThis.DOMParser === 'undefined') {
  const dom = new JSDOM();
  globalThis.DOMParser = dom.window.DOMParser;
  globalThis.XMLSerializer = dom.window.XMLSerializer;
}

import {
  savePptxPresentation,
  bufferToDataUrl,
  dataUrlToBuffer,
} from '../src/utils/pptxAdapter.js';
import { parsePptxPresentation } from '../src/utils/pptxViewer.js';

const OUT_DIR = path.resolve('docs/office-upgrade/out');
if (!fs.existsSync(OUT_DIR)) {
  fs.mkdirSync(OUT_DIR, { recursive: true });
}

async function editPptAWithAppAdapter(inputPath, outputPath) {
  const buf = fs.readFileSync(inputPath);
  const dataUrl = bufferToDataUrl(buf);
  const pres = await parsePptxPresentation(dataUrl);

  const editableSlides = pres.slides.map((s, idx) => {
    const bodyLines = s.paragraphs
      .filter((p) => !p.isTitle && p.text !== s.title)
      .map((p) => {
        const indent = '  '.repeat(Math.max(0, p.level));
        return `${indent}• ${p.text}`;
      });

    const slideTables = [];
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

  // 1. Slide 2 (인덱스 1): 본문 내 "빨강" -> "파랑" 수정
  editableSlides[1].bodyText = editableSlides[1].bodyText.replace('빨강', '파랑');

  // 2. Slide 3 (인덱스 2): 표 셀 "영업" -> "영업(수정)"
  if (editableSlides[2].tables && editableSlides[2].tables.length > 0) {
    const rows = editableSlides[2].tables[0].rows;
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        if (rows[r][c] === '영업') {
          rows[r][c] = '영업(수정)';
          break;
        }
      }
    }
  }

  // 앱 편집기(PptxEditorModal handleSave)와 100% 동일한 savePptxPresentation 호출
  const updatedDataUrl = await savePptxPresentation(dataUrl, editableSlides);
  const outBuf = dataUrlToBuffer(updatedDataUrl);
  fs.writeFileSync(outputPath, outBuf);
  console.log(`[성공] ${path.basename(outputPath)} 저장 완료 (${outBuf.length.toLocaleString()} 바이트)`);
}

async function editPptBWithAppAdapter(inputPath, outputPath) {
  const buf = fs.readFileSync(inputPath);
  const dataUrl = bufferToDataUrl(buf);
  const pres = await parsePptxPresentation(dataUrl);

  const editableSlides = pres.slides.map((s, idx) => {
    const bodyLines = s.paragraphs
      .filter((p) => !p.isTitle && p.text !== s.title)
      .map((p) => {
        const indent = '  '.repeat(Math.max(0, p.level));
        return `${indent}• ${p.text}`;
      });

    const slideTables = [];
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

  // 1. Slide 2 (인덱스 1): "세로쓰기예시" -> "세로쓰기예시(수정)"
  editableSlides[1].bodyText = editableSlides[1].bodyText.replace('세로쓰기예시', '세로쓰기예시(수정)');

  // 2. Slide 4 (인덱스 3): "첫째 항목" -> "첫째 항목(수정)"
  editableSlides[3].bodyText = editableSlides[3].bodyText.replace('첫째 항목', '첫째 항목(수정)');

  const updatedDataUrl = await savePptxPresentation(dataUrl, editableSlides);
  const outBuf = dataUrlToBuffer(updatedDataUrl);
  fs.writeFileSync(outputPath, outBuf);
  console.log(`[성공] ${path.basename(outputPath)} 저장 완료 (${outBuf.length.toLocaleString()} 바이트)`);
}

async function main() {
  console.log('앱 실제 저장 함수(savePptxPresentation)를 이용한 In-Place 편집 테스트 시작...');
  await editPptAWithAppAdapter(
    'docs/office-upgrade/samples/real/PPT_A_사업보고_원본.pptx',
    path.join(OUT_DIR, 'PPT_A_사업보고_원본_수정저장결과.pptx')
  );
  await editPptAWithAppAdapter(
    'docs/office-upgrade/samples/real/PPT_A_사업보고_한쇼.pptx',
    path.join(OUT_DIR, 'PPT_A_사업보고_한쇼_수정저장결과.pptx')
  );
  await editPptBWithAppAdapter(
    'docs/office-upgrade/samples/real/PPT_B_교육자료_원본.pptx',
    path.join(OUT_DIR, 'PPT_B_교육자료_원본_수정저장결과.pptx')
  );
  await editPptBWithAppAdapter(
    'docs/office-upgrade/samples/real/PPT_B_교육자료_한쇼.pptx',
    path.join(OUT_DIR, 'PPT_B_교육자료_한쇼_수정저장결과.pptx')
  );
  console.log('4개 파일 앱 어댑터 수정 저장 완료.');
}

main().catch((err) => {
  console.error('오류 발생:', err);
  process.exit(1);
});
