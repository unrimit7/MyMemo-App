import JSZip from 'jszip';

export function bufferToDataUrl(buf: Buffer | Uint8Array, mime: string = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'): string {
  const base64 = Buffer.isBuffer(buf) ? buf.toString('base64') : Buffer.from(buf).toString('base64');
  return `data:${mime};base64,${base64}`;
}

export function dataUrlToBuffer(dataUrl: string): Buffer {
  const commaIdx = dataUrl.indexOf(',');
  const b64 = commaIdx !== -1 ? dataUrl.slice(commaIdx + 1) : dataUrl;
  return Buffer.from(b64, 'base64');
}

function parseXml(xmlStr: string): Document {
  if (typeof DOMParser !== 'undefined') {
    return new DOMParser().parseFromString(xmlStr, 'text/xml');
  }
  if (typeof globalThis !== 'undefined' && (globalThis as any).DOMParser) {
    return new (globalThis as any).DOMParser().parseFromString(xmlStr, 'text/xml');
  }
  try {
    const { JSDOM } = require('jsdom');
    const dom = new JSDOM(xmlStr, { contentType: 'text/xml' });
    return dom.window.document;
  } catch {
    throw new Error('DOMParser를 찾을 수 없습니다.');
  }
}

function serializeXml(doc: Document): string {
  if (typeof XMLSerializer !== 'undefined') {
    return new XMLSerializer().serializeToString(doc);
  }
  if (typeof globalThis !== 'undefined' && (globalThis as any).XMLSerializer) {
    return new (globalThis as any).XMLSerializer().serializeToString(doc);
  }
  try {
    const { JSDOM } = require('jsdom');
    const dom = new JSDOM();
    return new dom.window.XMLSerializer().serializeToString(doc);
  } catch {
    throw new Error('XMLSerializer를 찾을 수 없습니다.');
  }
}

/**
 * 슬라이드 XML 내 특정 텍스트를 run(a:r) 단위로 안전하게 치환합니다.
 * 문단(a:p) 전체를 덮어쓰지 않고, 일치하는 run의 a:t 내용만 변경하므로
 * 다른 run의 서식(굵게, 기울임, 밑줄, 빨강 색상, a:highlight 형광펜,
 * a:hlinkClick 하이퍼링크, baseline 위/아래 첨자, a:fld 슬라이드 번호 등)이 온전히 보존됩니다.
 */
export function replaceRunTextInSlideXml(
  slideXmlText: string,
  findText: string,
  replacement: string
): { updatedXml: string; replaced: boolean } {
  const doc = parseXml(slideXmlText);

  let replaced = false;

  // 1. 모든 run (a:r) 탐색
  const runs = doc.getElementsByTagName('a:r');
  for (let i = 0; i < runs.length; i++) {
    const r = runs[i];
    const tEls = r.getElementsByTagName('a:t');
    if (tEls.length > 0 && tEls[0].textContent && tEls[0].textContent.includes(findText)) {
      tEls[0].textContent = tEls[0].textContent.replace(findText, replacement);
      replaced = true;
      break;
    }
  }

  // 2. 만약 일반 run에서 찾지 못했다면 단락(a:p) 내 다중 run에 걸쳐있는지 확인
  if (!replaced) {
    const ps = doc.getElementsByTagName('a:p');
    for (let pIdx = 0; pIdx < ps.length; pIdx++) {
      const p = ps[pIdx];
      const pRuns = p.getElementsByTagName('a:r');
      const fullText = Array.from(pRuns)
        .map((r) => r.getElementsByTagName('a:t')[0]?.textContent || '')
        .join('');

      if (fullText.includes(findText)) {
        // 단일 run에서 부분 일치 우선 시도
        for (let rIdx = 0; rIdx < pRuns.length; rIdx++) {
          const t = pRuns[rIdx].getElementsByTagName('a:t')[0];
          if (t && t.textContent && t.textContent.includes(findText)) {
            t.textContent = t.textContent.replace(findText, replacement);
            replaced = true;
            break;
          }
        }
        if (replaced) break;
      }
    }
  }

  return {
    updatedXml: serializeXml(doc),
    replaced,
  };
}

/**
 * 슬라이드 삭제 및 관련 메타데이터(sldIdLst, sectionLst, rels, Content_Types) 완벽 동기화
 * @param zip PPTX JSZip 인스턴스
 * @param slideIndex 삭제할 슬라이드의 0-based 인덱스
 */
export async function deleteSlideFromPptx(zip: JSZip, slideIndex: number): Promise<void> {
  // 1. presentation.xml
  const presFile = zip.file('ppt/presentation.xml');
  if (!presFile) throw new Error('ppt/presentation.xml을 찾을 수 없습니다.');
  const presDoc = parseXml(await presFile.async('text'));

  // 2. presentation.xml.rels
  const presRelsFile = zip.file('ppt/_rels/presentation.xml.rels');
  if (!presRelsFile) throw new Error('ppt/_rels/presentation.xml.rels를 찾을 수 없습니다.');
  const presRelsDoc = parseXml(await presRelsFile.async('text'));

  const sldIdLst = presDoc.getElementsByTagName('p:sldIdLst')[0];
  if (!sldIdLst) throw new Error('p:sldIdLst를 찾을 수 없습니다.');
  const sldIds = sldIdLst.getElementsByTagName('p:sldId');

  if (slideIndex < 0 || slideIndex >= sldIds.length) {
    throw new Error(`유효하지 않은 슬라이드 인덱스: ${slideIndex} (전체 ${sldIds.length}개)`);
  }

  const targetSld = sldIds[slideIndex];
  const targetId = targetSld.getAttribute('id');
  const targetRId = targetSld.getAttribute('r:id');

  // 관계(Relationship)에서 슬라이드 대상 파일 찾기
  let targetSlidePath = '';
  const rels = presRelsDoc.getElementsByTagName('Relationship');
  for (let i = 0; i < rels.length; i++) {
    if (rels[i].getAttribute('Id') === targetRId) {
      targetSlidePath = rels[i].getAttribute('Target') || '';
      rels[i].parentNode?.removeChild(rels[i]);
      break;
    }
  }

  // sldIdLst에서 삭제
  targetSld.parentNode?.removeChild(targetSld);

  // 3. 구역(p14:sectionLst)에서 슬라이드 ID 제거
  if (targetId) {
    const sections = presDoc.getElementsByTagName('p14:section');
    for (let i = 0; i < sections.length; i++) {
      const sec = sections[i];
      const sldEntries = sec.getElementsByTagName('p14:sldId');
      for (let j = 0; j < sldEntries.length; j++) {
        if (sldEntries[j].getAttribute('id') === targetId) {
          sldEntries[j].parentNode?.removeChild(sldEntries[j]);
          break;
        }
      }
    }
  }

  // 4. [Content_Types].xml에서 슬라이드 Override 제거
  const ctFile = zip.file('[Content_Types].xml');
  if (ctFile && targetSlidePath) {
    const ctDoc = parseXml(await ctFile.async('text'));
    const normalizedTarget = targetSlidePath.startsWith('/')
      ? targetSlidePath
      : targetSlidePath.startsWith('ppt/')
      ? `/${targetSlidePath}`
      : `/ppt/${targetSlidePath}`;

    const overrides = ctDoc.getElementsByTagName('Override');
    for (let i = 0; i < overrides.length; i++) {
      if (overrides[i].getAttribute('PartName') === normalizedTarget) {
        overrides[i].parentNode?.removeChild(overrides[i]);
        break;
      }
    }
    zip.file('[Content_Types].xml', serializeXml(ctDoc));
  }

  // 5. ZIP 내 실제 슬라이드 파일 및 슬라이드 rels 제거
  if (targetSlidePath) {
    const fullSlidePath = targetSlidePath.startsWith('ppt/')
      ? targetSlidePath
      : targetSlidePath.startsWith('/')
      ? targetSlidePath.substring(1)
      : `ppt/${targetSlidePath}`;

    zip.remove(fullSlidePath);

    const parts = fullSlidePath.split('/');
    const fname = parts.pop();
    const relsFname = `${parts.join('/')}/_rels/${fname}.rels`;
    zip.remove(relsFname);
  }

  // 갱신된 presentation.xml 및 presentation.xml.rels 저장
  zip.file('ppt/presentation.xml', serializeXml(presDoc));
  zip.file('ppt/_rels/presentation.xml.rels', serializeXml(presRelsDoc));
}

/**
 * 슬라이드 순서 재배열 (Reorder)
 * p:sldIdLst 내 p:sldId 순서를 재정렬합니다.
 * 슬라이드 고유 파일 및 r:id 관계가 보존되므로,
 * 슬라이드 이동 링크(ppaction://hlinksldjump)가 항상 원래 대상 슬라이드를 가리킵니다.
 */
export async function reorderSlidesInPptx(zip: JSZip, newOrderIndices: number[]): Promise<void> {
  const presFile = zip.file('ppt/presentation.xml');
  if (!presFile) throw new Error('ppt/presentation.xml을 찾을 수 없습니다.');
  const presDoc = parseXml(await presFile.async('text'));

  const sldIdLst = presDoc.getElementsByTagName('p:sldIdLst')[0];
  if (!sldIdLst) throw new Error('p:sldIdLst를 찾을 수 없습니다.');
  const sldIds = Array.from(sldIdLst.getElementsByTagName('p:sldId'));

  if (newOrderIndices.length !== sldIds.length) {
    throw new Error('새 슬라이드 순서 배열 길이가 현재 슬라이드 수와 일치하지 않습니다.');
  }

  // 슬라이드 요소 재배치
  const reordered = newOrderIndices.map((idx) => sldIds[idx]);
  while (sldIdLst.firstChild) {
    sldIdLst.removeChild(sldIdLst.firstChild);
  }
  for (const sld of reordered) {
    sldIdLst.appendChild(sld);
  }

  zip.file('ppt/presentation.xml', serializeXml(presDoc));
}
