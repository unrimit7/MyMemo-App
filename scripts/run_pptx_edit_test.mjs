import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { JSDOM } from 'jsdom';

const OUT_DIR = path.resolve('docs/office-upgrade/out');
if (!fs.existsSync(OUT_DIR)) {
  fs.mkdirSync(OUT_DIR, { recursive: true });
}

function replaceRunTextInDoc(doc, findText, replaceText) {
  const runs = doc.getElementsByTagName('a:r');
  let replaced = false;
  for (let i = 0; i < runs.length; i++) {
    const r = runs[i];
    const tEls = r.getElementsByTagName('a:t');
    for (let j = 0; j < tEls.length; j++) {
      if (tEls[j].textContent && tEls[j].textContent.includes(findText)) {
        tEls[j].textContent = tEls[j].textContent.replace(findText, replaceText);
        replaced = true;
        break;
      }
    }
    if (replaced) break;
  }
  return replaced;
}

async function editPptA(inputPath, outputPath) {
  const buf = fs.readFileSync(inputPath);
  const zip = await JSZip.loadAsync(buf);

  // 1. Slide 2: 본문 텍스트 run-level 수정 ("빨강" -> "파랑")
  const s2Xml = await zip.file('ppt/slides/slide2.xml').async('text');
  const dom2 = new JSDOM(s2Xml, { contentType: 'text/xml' });
  const replacedS2 = replaceRunTextInDoc(dom2.window.document, '빨강', '파랑');
  if (!replacedS2) {
    throw new Error(`slide2.xml에서 "빨강" 텍스트를 찾지 못했습니다 (${inputPath})`);
  }
  const serializer = new dom2.window.XMLSerializer();
  zip.file('ppt/slides/slide2.xml', serializer.serializeToString(dom2.window.document));

  // 2. Slide 3: 표 셀 수정 ("영업" -> "영업(수정)")
  const s3Xml = await zip.file('ppt/slides/slide3.xml').async('text');
  const dom3 = new JSDOM(s3Xml, { contentType: 'text/xml' });
  const replacedS3 = replaceRunTextInDoc(dom3.window.document, '영업', '영업(수정)');
  if (!replacedS3) {
    throw new Error(`slide3.xml에서 "영업" 텍스트를 찾지 못했습니다 (${inputPath})`);
  }
  zip.file('ppt/slides/slide3.xml', serializer.serializeToString(dom3.window.document));

  const outBuf = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  fs.writeFileSync(outputPath, outBuf);
  console.log(`[성공] ${path.basename(outputPath)} 저장 완료 (${outBuf.length.toLocaleString()} 바이트)`);
}

async function editPptB(inputPath, outputPath) {
  const buf = fs.readFileSync(inputPath);
  const zip = await JSZip.loadAsync(buf);

  // 1. Slide 2: 본문 텍스트 run-level 수정 ("세로쓰기예시" -> "세로쓰기예시(수정)")
  const s2Xml = await zip.file('ppt/slides/slide2.xml').async('text');
  const dom2 = new JSDOM(s2Xml, { contentType: 'text/xml' });
  const replacedS2 = replaceRunTextInDoc(dom2.window.document, '세로쓰기예시', '세로쓰기예시(수정)');
  if (!replacedS2) {
    throw new Error(`slide2.xml에서 "세로쓰기예시" 텍스트를 찾지 못했습니다 (${inputPath})`);
  }
  const serializer = new dom2.window.XMLSerializer();
  zip.file('ppt/slides/slide2.xml', serializer.serializeToString(dom2.window.document));

  // 2. Slide 4: 사용자 지정 글머리 본문 수정 ("첫째 항목" -> "첫째 항목(수정)")
  const s4Xml = await zip.file('ppt/slides/slide4.xml').async('text');
  const dom4 = new JSDOM(s4Xml, { contentType: 'text/xml' });
  const replacedS4 = replaceRunTextInDoc(dom4.window.document, '첫째 항목', '첫째 항목(수정)');
  if (!replacedS4) {
    throw new Error(`slide4.xml에서 "첫째 항목" 텍스트를 찾지 못했습니다 (${inputPath})`);
  }
  zip.file('ppt/slides/slide4.xml', serializer.serializeToString(dom4.window.document));

  const outBuf = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  fs.writeFileSync(outputPath, outBuf);
  console.log(`[성공] ${path.basename(outputPath)} 저장 완료 (${outBuf.length.toLocaleString()} 바이트)`);
}

async function main() {
  console.log('PPTX In-Place run-level 편집 테스트 시작...');
  await editPptA(
    'docs/office-upgrade/samples/real/PPT_A_사업보고_원본.pptx',
    path.join(OUT_DIR, 'PPT_A_사업보고_원본_수정저장결과.pptx')
  );
  await editPptA(
    'docs/office-upgrade/samples/real/PPT_A_사업보고_한쇼.pptx',
    path.join(OUT_DIR, 'PPT_A_사업보고_한쇼_수정저장결과.pptx')
  );
  await editPptB(
    'docs/office-upgrade/samples/real/PPT_B_교육자료_원본.pptx',
    path.join(OUT_DIR, 'PPT_B_교육자료_원본_수정저장결과.pptx')
  );
  await editPptB(
    'docs/office-upgrade/samples/real/PPT_B_교육자료_한쇼.pptx',
    path.join(OUT_DIR, 'PPT_B_교육자료_한쇼_수정저장결과.pptx')
  );
  console.log('4개 파일 In-Place 수정 저장 완료.');
}

main().catch(err => {
  console.error('오류 발생:', err);
  process.exit(1);
});
