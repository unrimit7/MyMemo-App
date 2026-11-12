import JSZip from 'jszip';
import { parsePptxPresentation } from './pptxViewer';

export function bufferToDataUrl(
  buf: Uint8Array | ArrayBuffer | any,
  mime: string = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
): string {
  if (typeof Buffer !== 'undefined' && Buffer.isBuffer && Buffer.isBuffer(buf)) {
    return `data:${mime};base64,${buf.toString('base64')}`;
  }
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  let binary = '';
  const len = u8.byteLength;
  const chunk = 8192;
  for (let i = 0; i < len; i += chunk) {
    binary += String.fromCharCode.apply(null, u8.subarray(i, Math.min(i + chunk, len)) as any);
  }
  return `data:${mime};base64,${btoa(binary)}`;
}

export function dataUrlToBuffer(dataUrl: string): any {
  const commaIdx = dataUrl.indexOf(',');
  const b64 = commaIdx !== -1 ? dataUrl.slice(commaIdx + 1) : dataUrl;
  if (typeof Buffer !== 'undefined' && typeof Buffer.from === 'function') {
    return Buffer.from(b64, 'base64');
  }
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  const commaIdx = dataUrl.indexOf(',');
  const b64 = commaIdx !== -1 ? dataUrl.slice(commaIdx + 1) : dataUrl;
  const binary = atob(b64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function parseXml(xmlStr: string): Document {
  if (typeof DOMParser !== 'undefined') {
    return new DOMParser().parseFromString(xmlStr, 'text/xml');
  }
  if (typeof globalThis !== 'undefined' && (globalThis as any).DOMParser) {
    return new (globalThis as any).DOMParser().parseFromString(xmlStr, 'text/xml');
  }
  throw new Error('DOMParser를 찾을 수 없습니다.');
}

function serializeXml(doc: Document): string {
  if (typeof XMLSerializer !== 'undefined') {
    return new XMLSerializer().serializeToString(doc);
  }
  if (typeof globalThis !== 'undefined' && (globalThis as any).XMLSerializer) {
    return new (globalThis as any).XMLSerializer().serializeToString(doc);
  }
  throw new Error('XMLSerializer를 찾을 수 없습니다.');
}

/**
 * 단락(a:p) 내의 텍스트 runs를 diff 기반으로 정확하게 매핑하여 수정합니다.
 * - 문단 전체를 덮어쓰지 않고 변경된 글자 위치(diff)만 찾아 해당 run의 a:t 내용만 변경합니다.
 * - 인접한 다른 run의 서식(굵게, 기울임, 밑줄, 글자색, a:highlight 형광펜,
 *   a:hlinkClick 하이퍼링크, baseline 위/아래 첨자, a:strike 취소선)과 a:fld 필드가 100% 보존됩니다.
 */
export function updateParagraphElementWithDiff(p: Element, newText: string): boolean {
  // a:r 직속 run들 및 각 run의 a:t 찾기
  const runs: { el: Element; tEl: Element; text: string }[] = [];
  for (let i = 0; i < p.children.length; i++) {
    const child = p.children[i];
    if (child.localName === 'r') {
      const tEls = child.getElementsByTagName('a:t');
      if (tEls.length > 0) {
        runs.push({
          el: child,
          tEl: tEls[0],
          text: tEls[0].textContent || '',
        });
      }
    }
  }

  const origText = runs.map((r) => r.text).join('');
  if (origText === newText) {
    // 내용이 동일하면 DOM 노드를 전혀 건드리지 않음
    return false;
  }

  if (runs.length === 0) {
    // 기존 run이 없는 빈 단락인 경우에만 새 run 생성
    const doc = p.ownerDocument;
    const rEl = doc.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:r');
    const tEl = doc.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:t');
    tEl.textContent = newText;
    rEl.appendChild(tEl);
    p.appendChild(rEl);
    return true;
  }

  // 공통 접두사(prefix) 및 공통 접미사(suffix) 계산
  let prefixLen = 0;
  while (
    prefixLen < origText.length &&
    prefixLen < newText.length &&
    origText[prefixLen] === newText[prefixLen]
  ) {
    prefixLen++;
  }

  let suffixLen = 0;
  while (
    suffixLen < origText.length - prefixLen &&
    suffixLen < newText.length - prefixLen &&
    origText[origText.length - 1 - suffixLen] === newText[newText.length - 1 - suffixLen]
  ) {
    suffixLen++;
  }

  const origChangeStart = prefixLen;
  const origChangeEnd = origText.length - suffixLen;
  const replacement = newText.slice(prefixLen, newText.length - suffixLen);

  // 변경 대상 범위에 속하는 run 식별
  let pos = 0;
  let firstAffected = -1;
  let lastAffected = -1;

  for (let i = 0; i < runs.length; i++) {
    const rStart = pos;
    const rEnd = pos + runs[i].text.length;
    pos = rEnd;

    if (rEnd <= origChangeStart || rStart >= origChangeEnd) {
      continue;
    }
    if (firstAffected === -1) firstAffected = i;
    lastAffected = i;
  }

  if (firstAffected === -1) {
    // 경계 끝에 추가되는 경우 등
    runs[runs.length - 1].tEl.textContent =
      (runs[runs.length - 1].tEl.textContent || '') + replacement;
    return true;
  }

  if (firstAffected === lastAffected) {
    // 단일 run 내에서만 변경이 발생한 경우 (가장 흔한 케이스: 한 단어/서식 수정)
    let rStart = 0;
    for (let i = 0; i < firstAffected; i++) rStart += runs[i].text.length;
    const curRunText = runs[firstAffected].text;
    const before = curRunText.slice(0, origChangeStart - rStart);
    const after = curRunText.slice(origChangeEnd - rStart);
    runs[firstAffected].tEl.textContent = before + replacement + after;
  } else {
    // 여러 run에 걸쳐 변경/삭제된 경우
    let rStartFirst = 0;
    for (let i = 0; i < firstAffected; i++) rStartFirst += runs[i].text.length;
    let rStartLast = 0;
    for (let i = 0; i < lastAffected; i++) rStartLast += runs[i].text.length;

    runs[firstAffected].tEl.textContent =
      runs[firstAffected].text.slice(0, origChangeStart - rStartFirst) + replacement;

    for (let i = firstAffected + 1; i < lastAffected; i++) {
      runs[i].tEl.textContent = '';
    }

    runs[lastAffected].tEl.textContent =
      runs[lastAffected].text.slice(origChangeEnd - rStartLast);
  }

  return true;
}

/**
 * 슬라이드 XML 내 특정 텍스트를 run(a:r) 단위로 치환합니다.
 */
export function replaceRunTextInSlideXml(
  slideXmlText: string,
  findText: string,
  replacement: string
): { updatedXml: string; replaced: boolean } {
  const doc = parseXml(slideXmlText);
  let replaced = false;

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

  if (!replaced) {
    const ps = doc.getElementsByTagName('a:p');
    for (let pIdx = 0; pIdx < ps.length; pIdx++) {
      const p = ps[pIdx];
      const pRuns = p.getElementsByTagName('a:r');
      const fullText = Array.from(pRuns)
        .map((r) => r.getElementsByTagName('a:t')[0]?.textContent || '')
        .join('');

      if (fullText.includes(findText)) {
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
 */
export async function deleteSlideFromPptx(zip: JSZip, slideIndex: number): Promise<void> {
  const presFile = zip.file('ppt/presentation.xml');
  if (!presFile) throw new Error('ppt/presentation.xml을 찾을 수 없습니다.');
  const presDoc = parseXml(await presFile.async('text'));

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

  let targetSlidePath = '';
  const rels = presRelsDoc.getElementsByTagName('Relationship');
  for (let i = 0; i < rels.length; i++) {
    if (rels[i].getAttribute('Id') === targetRId) {
      targetSlidePath = rels[i].getAttribute('Target') || '';
      rels[i].parentNode?.removeChild(rels[i]);
      break;
    }
  }

  targetSld.parentNode?.removeChild(targetSld);

  // 구역(p14:sectionLst)에서 슬라이드 ID 제거
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

  // [Content_Types].xml에서 슬라이드 Override 제거
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
    zip.file('[Content_Types].xml', serializeXml(ctDoc), { createFolders: false });
  }

  // ZIP 내 실제 슬라이드 파일 및 슬라이드 rels 제거
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

  zip.file('ppt/presentation.xml', serializeXml(presDoc), { createFolders: false });
  zip.file('ppt/_rels/presentation.xml.rels', serializeXml(presRelsDoc), { createFolders: false });
}

/**
 * 슬라이드 순서 재배열 (Reorder)
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

  const reordered = newOrderIndices.map((idx) => sldIds[idx]);
  while (sldIdLst.firstChild) {
    sldIdLst.removeChild(sldIdLst.firstChild);
  }
  for (const sld of reordered) {
    sldIdLst.appendChild(sld);
  }

  zip.file('ppt/presentation.xml', serializeXml(presDoc), { createFolders: false });
}

export interface EditableSlideData {
  slideNumber: number;
  originalSlideIndex?: number;
  title: string;
  bodyText: string;
  tables?: { rows: string[][] }[];
}

/**
 * 앱 편집기(PptxEditorModal) 및 테스트 스위트가 공유하는 공식 In-Place 파워포인트 저장 함수.
 * - 원래 단락 텍스트와 수정된 단락 텍스트를 비교하여 바뀐 글자가 속한 run(a:r)만 diff 기반으로 치환합니다.
 * - 제목 및 표(Table) 셀 또한 셀 단위로 run 서식을 유지하며 저장합니다.
 * - 슬라이드 삭제 및 순서 재배열이 발생한 경우 메타데이터(sldIdLst, rels, sectionLst 등)를 완전 동기화합니다.
 * - JSZip 저장 시 원래 없던 디렉터리 엔트리(ppt/, ppt/slides/ 등)가 생성되지 않도록 엄격하게 정규화합니다.
 */
export async function savePptxPresentation(
  originalDataUrl: string,
  slides: EditableSlideData[],
  deletedSlideIndices?: number[]
): Promise<string> {
  const bytes = dataUrlToUint8Array(originalDataUrl);
  const zip = await JSZip.loadAsync(bytes);

  // 원본 ZIP의 디렉터리 엔트리 목록 기록 (불필요한 디렉터리 엔트리 생성 차단용)
  const origFolderEntries = new Set(
    Object.keys(zip.files).filter((k) => zip.files[k].dir)
  );

  // 1. 슬라이드 삭제 처리 (역순으로 삭제하여 인덱스 밀림 방지)
  if (deletedSlideIndices && deletedSlideIndices.length > 0) {
    const sortedDesc = [...deletedSlideIndices].sort((a, b) => b - a);
    for (const delIdx of sortedDesc) {
      await deleteSlideFromPptx(zip, delIdx);
    }
  }

  // 2. 슬라이드 순서 재배열 처리
  const presFile = zip.file('ppt/presentation.xml');
  if (!presFile) throw new Error('ppt/presentation.xml을 찾을 수 없습니다.');
  const presDoc = parseXml(await presFile.async('text'));

  const presRelsFile = zip.file('ppt/_rels/presentation.xml.rels');
  if (!presRelsFile) throw new Error('ppt/_rels/presentation.xml.rels를 찾을 수 없습니다.');
  const presRelsDoc = parseXml(await presRelsFile.async('text'));

  const sldIdLst = presDoc.getElementsByTagName('p:sldIdLst')[0];
  const sldIds = sldIdLst ? Array.from(sldIdLst.getElementsByTagName('p:sldId')) : [];

  const relMap = new Map<string, string>();
  const rels = presRelsDoc.getElementsByTagName('Relationship');
  for (let i = 0; i < rels.length; i++) {
    const id = rels[i].getAttribute('Id');
    const target = rels[i].getAttribute('Target');
    if (id && target) relMap.set(id, target);
  }

  // 3. 원본 프레젠테이션 구조 파싱 (변경 여부 정밀 비교 및 shapeId/placeholder 매핑용)
  const origPres = await parsePptxPresentation(originalDataUrl);

  for (let sIdx = 0; sIdx < slides.length; sIdx++) {
    if (sIdx >= sldIds.length) break;

    const updatedSlide = slides[sIdx];
    const origSlideIdx =
      updatedSlide.originalSlideIndex !== undefined ? updatedSlide.originalSlideIndex : sIdx;
    const origSlide =
      origSlideIdx >= 0 && origSlideIdx < origPres.slides.length
        ? origPres.slides[origSlideIdx]
        : undefined;

    // 변경 여부 정밀 검사
    const titleChanged = origSlide
      ? updatedSlide.title !== origSlide.title
      : updatedSlide.title !== undefined;

    const origBodyParas = origSlide
      ? origSlide.paragraphs.filter(
          (p) => !p.isTitle && p.text !== origSlide.title && p.shapeId !== origSlide.titleShapeId
        )
      : [];
    const origBodyLineTexts = origBodyParas.map((p) => p.text.trim());
    const updatedBodyLines =
      updatedSlide.bodyText !== undefined
        ? updatedSlide.bodyText
            .split('\n')
            .map((l) => l.trim().replace(/^[-*•▪\u2022\u25AA\s]+/, '').trim())
            .filter(Boolean)
        : [];

    let bodyChanged = false;
    if (origSlide) {
      if (updatedBodyLines.length !== origBodyLineTexts.length) {
        bodyChanged = true;
      } else {
        for (let i = 0; i < updatedBodyLines.length; i++) {
          if (updatedBodyLines[i] !== origBodyLineTexts[i]) {
            bodyChanged = true;
            break;
          }
        }
      }
    } else if (updatedSlide.bodyText !== undefined) {
      bodyChanged = true;
    }

    let tablesChanged = false;
    const origTables = origSlide ? origSlide.elements.filter((el) => el.type === 'table') : [];
    if (updatedSlide.tables && updatedSlide.tables.length > 0) {
      if (origSlide) {
        if (updatedSlide.tables.length !== origTables.length) {
          tablesChanged = true;
        } else {
          for (let tIdx = 0; tIdx < updatedSlide.tables.length; tIdx++) {
            const updRows = updatedSlide.tables[tIdx].rows;
            const origT = origTables[tIdx];
            if (origT && origT.type === 'table') {
              if (updRows.length !== origT.rows.length) {
                tablesChanged = true;
                break;
              }
              for (let r = 0; r < updRows.length; r++) {
                if (updRows[r].length !== origT.rows[r].cells.length) {
                  tablesChanged = true;
                  break;
                }
                for (let c = 0; c < updRows[r].length; c++) {
                  if (updRows[r][c] !== origT.rows[r].cells[c].text) {
                    tablesChanged = true;
                    break;
                  }
                }
                if (tablesChanged) break;
              }
            } else {
              tablesChanged = true;
              break;
            }
            if (tablesChanged) break;
          }
        }
      } else {
        tablesChanged = true;
      }
    }

    // 변경된 항목이 전혀 없는 슬라이드는 원본 XML 파일을 절대 건드리지 않음 (100% 원본 유지)
    if (!titleChanged && !bodyChanged && !tablesChanged) {
      continue;
    }

    const sldEl = sldIds[sIdx];
    const rId = sldEl.getAttribute('r:id');
    const target = rId ? relMap.get(rId) : undefined;
    if (!target) continue;

    const fullPath = target.startsWith('ppt/')
      ? target
      : target.startsWith('/')
      ? target.substring(1)
      : `ppt/${target}`;

    const slideFile = zip.file(fullPath);
    if (!slideFile) continue;

    const sXmlText = await slideFile.async('text');
    const slideDoc = parseXml(sXmlText);
    let slideModified = false;

    // (A) 제목(Title) 수정: shapeId 매핑 우선
    if (titleChanged && updatedSlide.title !== undefined) {
      const spList = Array.from(slideDoc.getElementsByTagName('p:sp'));
      let titleSp: Element | null = null;
      if (origSlide?.titleShapeId) {
        titleSp =
          spList.find((sp) => {
            const cNvPr = sp.getElementsByTagName('p:cNvPr')[0];
            return cNvPr?.getAttribute('id') === origSlide.titleShapeId;
          }) || null;
      }
      if (!titleSp) {
        titleSp =
          spList.find((sp) => {
            const ph = sp.getElementsByTagName('p:ph')[0];
            const pType = ph?.getAttribute('type');
            return pType === 'title' || pType === 'ctrTitle';
          }) || null;
      }
      if (titleSp) {
        const titlePs = titleSp.getElementsByTagName('a:p');
        if (titlePs.length > 0) {
          const mod = updateParagraphElementWithDiff(titlePs[0], updatedSlide.title);
          if (mod) slideModified = true;
        }
      }
    }

    // (B) 본문(Body) 문단 수정: 각 문단별 shapeId 및 pIndex 정확 매핑
    if (bodyChanged && updatedSlide.bodyText !== undefined) {
      const spList = Array.from(slideDoc.getElementsByTagName('p:sp'));
      for (let lIdx = 0; lIdx < updatedBodyLines.length; lIdx++) {
        const newLine = updatedBodyLines[lIdx];
        if (lIdx < origBodyParas.length) {
          const origPara = origBodyParas[lIdx];
          if (newLine === origPara.text.trim()) continue; // 변경 없는 문단은 건너뜀

          let targetSp: Element | null = null;
          if (origPara.shapeId) {
            targetSp =
              spList.find((sp) => {
                const cNvPr = sp.getElementsByTagName('p:cNvPr')[0];
                return cNvPr?.getAttribute('id') === origPara.shapeId;
              }) || null;
          }
          if (!targetSp) {
            targetSp =
              spList.find((sp) => {
                const ph = sp.getElementsByTagName('p:ph')[0];
                const pType = ph?.getAttribute('type');
                return pType === 'body' || pType === 'subTitle';
              }) || null;
          }

          if (targetSp) {
            const txBody =
              targetSp.getElementsByTagName('p:txBody')[0] ||
              targetSp.getElementsByTagName('a:txBody')[0];
            const pElements = txBody
              ? txBody.getElementsByTagName('a:p')
              : targetSp.getElementsByTagName('a:p');
            const targetPIdx =
              origPara.pIndex !== undefined && origPara.pIndex < pElements.length
                ? origPara.pIndex
                : Math.min(lIdx, pElements.length - 1);
            if (targetPIdx >= 0 && targetPIdx < pElements.length) {
              const mod = updateParagraphElementWithDiff(pElements[targetPIdx], newLine);
              if (mod) slideModified = true;
            }
          }
        }
      }
    }

    // (C) 표(Table) 셀 데이터 수정: hMerge/vMerge 연속 셀 제외 후 정확한 컬럼 매핑
    if (tablesChanged && updatedSlide.tables && updatedSlide.tables.length > 0) {
      const tblList = Array.from(slideDoc.getElementsByTagName('a:tbl'));
      for (let tIdx = 0; tIdx < updatedSlide.tables.length; tIdx++) {
        if (tIdx >= tblList.length) break;
        const tbl = tblList[tIdx];
        const trList = tbl.getElementsByTagName('a:tr');
        const updatedRows = updatedSlide.tables[tIdx].rows;
        const origT =
          tIdx < origTables.length && origTables[tIdx].type === 'table'
            ? (origTables[tIdx] as any)
            : null;

        for (let rIdx = 0; rIdx < updatedRows.length; rIdx++) {
          if (rIdx >= trList.length) break;
          const tr = trList[rIdx];
          // 중요: hMerge="1" 또는 vMerge="1"인 병합 연속 셀은 뷰어에서 제외되었으므로 매핑 시에도 제외
          const allTcs = Array.from(tr.getElementsByTagName('a:tc'));
          const visibleTcs = allTcs.filter((tc) => {
            const hm = tc.getAttribute('hMerge');
            const vm = tc.getAttribute('vMerge');
            return hm !== '1' && hm !== 'true' && vm !== '1' && vm !== 'true';
          });
          const rowData = updatedRows[rIdx];

          for (let cIdx = 0; cIdx < rowData.length; cIdx++) {
            if (cIdx >= visibleTcs.length) break;
            const newCellText = rowData[cIdx];
            const origCellText = origT?.rows?.[rIdx]?.cells?.[cIdx]?.text;
            if (origCellText !== undefined && newCellText === origCellText) {
              continue; // 수정되지 않은 셀은 건너뜀
            }

            const tc = visibleTcs[cIdx];
            const pElements = tc.getElementsByTagName('a:p');
            if (pElements.length > 0) {
              const mod = updateParagraphElementWithDiff(pElements[0], newCellText);
              if (mod) slideModified = true;
            }
          }
        }
      }
    }

    if (slideModified) {
      const newXml = serializeXml(slideDoc);
      zip.file(fullPath, newXml, { createFolders: false });
    }
  }

  // 4. 원본에 없던 신규 폴더 엔트리(ppt/, ppt/slides/ 등) 제거
  for (const path of Object.keys(zip.files)) {
    if (zip.files[path].dir && !origFolderEntries.has(path)) {
      zip.remove(path);
    }
  }

  // 5. 무손실 압축 바이너리 생성
  const outBuf = await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  return bufferToDataUrl(outBuf);
}
