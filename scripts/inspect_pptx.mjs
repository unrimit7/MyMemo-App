import fs from 'fs';
import JSZip from 'jszip';
import { JSDOM } from 'jsdom';

async function inspectSlides(filePath, name) {
  const buf = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buf);
  const presFile = zip.file('ppt/presentation.xml');
  const presRelsFile = zip.file('ppt/_rels/presentation.xml.rels');
  const pDom = new JSDOM(await presFile.async('text'), { contentType: 'text/xml' });
  const prDom = new JSDOM(await presRelsFile.async('text'), { contentType: 'text/xml' });

  const relMap = new Map();
  for (const r of prDom.window.document.getElementsByTagName('Relationship')) {
    relMap.set(r.getAttribute('Id'), r.getAttribute('Target'));
  }

  const sldIds = pDom.window.document.getElementsByTagName('p:sldId');
  console.log(`\n================== ${name} ==================`);
  for (let i = 0; i < sldIds.length; i++) {
    const sldId = sldIds[i];
    const rId = sldId.getAttribute('r:id');
    const target = relMap.get(rId);
    const fullPath = target.startsWith('ppt/') ? target : `ppt/${target}`;
    const sXml = await zip.file(fullPath).async('text');
    const sDom = new JSDOM(sXml, { contentType: 'text/xml' });
    const sDoc = sDom.window.document;

    const isHidden = sDoc.documentElement.getAttribute('show') === '0';
    const shapes = sDoc.getElementsByTagName('p:sp').length;
    const tables = sDoc.getElementsByTagName('a:tbl').length;
    const charts = sDoc.getElementsByTagName('c:chart').length;
    const pics = sDoc.getElementsByTagName('p:pic').length;
    const cxns = sDoc.getElementsByTagName('p:cxnSp').length;
    const grps = sDoc.getElementsByTagName('p:grpSp').length;

    // First heading text
    const texts = Array.from(sDoc.getElementsByTagName('a:t')).map(t => t.textContent).filter(t => t.trim().length > 0);
    const firstText = texts[0] || '(텍스트 없음)';

    // Check special elements
    const special = [];
    if (isHidden) special.push('숨김(show=0)');
    if (charts > 0) special.push(`차트(${charts})`);
    if (tables > 0) special.push(`표(${tables})`);
    if (pics > 0) special.push(`이미지(${pics})`);
    if (cxns > 0) special.push(`연결선(${cxns})`);
    if (grps > 0) special.push(`그룹(${grps})`);
    if (sDoc.getElementsByTagName('a:highlight').length > 0) special.push('형광펜(highlight)');
    if (sDoc.getElementsByTagName('a:hlinkClick').length > 0) special.push('하이퍼링크(hlinkClick)');
    if (sXml.includes('baseline=')) special.push('위/아래첨자(baseline)');
    if (sXml.includes('strike=')) special.push('취소선(strike)');
    if (sXml.includes('eaVert')) special.push('세로쓰기(eaVert)');
    if (sXml.includes('numCol=')) special.push('다단(numCol)');
    if (sXml.includes('buChar')) special.push('글머리기호(buChar)');
    if (sDoc.getElementsByTagName('a:fld').length > 0) special.push('슬라이드번호필드(fld)');

    console.log(`[Slide ${i + 1}] (${fullPath}) "${firstText.slice(0, 20)}" | sp:${shapes} | ${special.join(', ')}`);
  }
}

async function main() {
  await inspectSlides('docs/office-upgrade/samples/real/PPT_A_사업보고_원본.pptx', 'PPT_A_원본');
  await inspectSlides('docs/office-upgrade/samples/real/PPT_A_사업보고_한쇼.pptx', 'PPT_A_한쇼');
  await inspectSlides('docs/office-upgrade/samples/real/PPT_B_교육자료_원본.pptx', 'PPT_B_원본');
  await inspectSlides('docs/office-upgrade/samples/real/PPT_B_교육자료_한쇼.pptx', 'PPT_B_한쇼');
}
main().catch(console.error);
