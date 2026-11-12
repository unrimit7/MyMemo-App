import JSZip from 'jszip';

function getDomParser(): any {
  if (typeof DOMParser !== 'undefined') {
    return new DOMParser();
  }
  // Node.js fallback
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { DOMParser: XmldomParser } = require('@xmldom/xmldom');
  return new XmldomParser();
}

function getXmlSerializer(): any {
  if (typeof XMLSerializer !== 'undefined') {
    return new XMLSerializer();
  }
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const { XMLSerializer: XmldomSerializer } = require('@xmldom/xmldom');
  return new XmldomSerializer();
}

function findElements(parent: Document | Element, tagName: string): Element[] {
  const list = parent.getElementsByTagName(tagName);
  if (list && list.length > 0) {
    return Array.from(list);
  }
  const all = parent.getElementsByTagName('*');
  const res: Element[] = [];
  const cleanName = tagName.includes(':') ? tagName.split(':')[1] : tagName;
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    if (el.localName === cleanName || el.nodeName === tagName) {
      res.push(el);
    }
  }
  return res;
}

/**
 * 문단의 직계 런(hp:run > hp:t) 텍스트만 추출
 * (내포된 표 hp:tbl, 구역 정의 등의 텍스트는 제외하여 순수 문단 텍스트만 비교)
 */
function getDirectParaText(pNode: Element): string {
  let text = '';
  for (let c = pNode.firstChild; c; c = c.nextSibling) {
    if (c.nodeName === 'hp:run' || (c as any).localName === 'run') {
      for (let r = c.firstChild; r; r = r.nextSibling) {
        if (r.nodeName === 'hp:t' || (r as any).localName === 't') {
          text += r.textContent || '';
        }
      }
    }
  }
  return text.replace(/\s+/g, ' ').trim();
}

/**
 * 문단의 직계 hp:linesegarray 자식 요소만 제거 (내포된 표 셀 등의 linesegarray 보존)
 */
function removeDirectLineSegArray(pNode: Element): boolean {
  let removed = false;
  const toRemove: Node[] = [];
  for (let c = pNode.firstChild; c; c = c.nextSibling) {
    if (c.nodeName === 'hp:linesegarray' || (c as any).localName === 'linesegarray') {
      toRemove.push(c);
    }
  }
  toRemove.forEach((n) => {
    pNode.removeChild(n);
    removed = true;
  });
  return removed;
}

/**
 * HWP/HWPX 문서에 변경 추적(Track Changes)이 포함되어 있는지 감지
 */
export async function hasHwpTrackChanges(
  bytes: Uint8Array,
  isHwpx: boolean
): Promise<{ hasTrackChanges: boolean; details: string[] }> {
  const details: string[] = [];
  if (isHwpx) {
    try {
      const zip = await JSZip.loadAsync(bytes);
      for (const name of Object.keys(zip.files)) {
        if (name.endsWith('.xml')) {
          const text = await zip.file(name)!.async('string');
          if (text.includes('trackChange') || text.includes('insertBegin') || text.includes('deleteBegin')) {
            details.push(`${name} 내 변경 추적 태그 포함`);
          }
        }
      }
    } catch {
      // ignore
    }
  } else {
    // HWP 바이너리: OLE2 레코드 및 FileHeader/BodyText 검사
    try {
      const latin = Buffer.from(bytes).toString('latin1');
      if (latin.includes('TrackChange') || latin.includes('TRACKCHANGE')) {
        details.push('HWP 바이너리 내 변경 추적 레코드 포함');
      }
    } catch {
      // ignore
    }
  }
  return {
    hasTrackChanges: details.length > 0,
    details,
  };
}

/**
 * HWPX 문서 저장 후처리 어댑터
 *
 * 1. 원본 문서의 version.xml(버전 정보 1.4/1.5 등) 보존
 * 2. 원본 header.xml의 hh:trackChanges 보존
 * 3. 텍스트가 실제로 수정된 문단(표 셀 문단 포함)만 직계 linesegarray(줄 배치 캐시)를 선별 제거
 * 4. 수정되지 않은 문단 중 변경 추적(deleteBegin/insertBegin)이 포함된 문단은 원본 마크업을 복원하여 보존
 * 5. 수정되지 않은 문단은 원본 linesegarray 및 구조 100% 바이트 보존
 */
export async function postProcessHwpx(
  exportedBytes: Uint8Array,
  originalBytes?: Uint8Array
): Promise<Uint8Array> {
  if (!originalBytes) return exportedBytes;

  try {
    const expZip = await JSZip.loadAsync(exportedBytes);
    const origZip = await JSZip.loadAsync(originalBytes);

    let modified = false;
    const parser = getDomParser();
    const serializer = getXmlSerializer();

    // 1. version.xml 보존 (한컴 한글 호환성 유지)
    const origVer = origZip.file('version.xml');
    if (origVer) {
      const origVerXml = await origVer.async('string');
      expZip.file('version.xml', origVerXml);
      modified = true;
    }

    // 2. Contents/header.xml의 hh:trackChanges 보존
    const origHdrFile = origZip.file('Contents/header.xml');
    const expHdrFile = expZip.file('Contents/header.xml');
    if (origHdrFile && expHdrFile) {
      const origHdrXml = await origHdrFile.async('string');
      if (origHdrXml.includes('trackChanges')) {
        const expHdrXml = await expHdrFile.async('string');
        if (!expHdrXml.includes('trackChanges')) {
          const docOrigHdr = parser.parseFromString(origHdrXml, 'application/xml');
          const docExpHdr = parser.parseFromString(expHdrXml, 'application/xml');
          const tcList = findElements(docOrigHdr, 'hh:trackChanges');
          if (tcList.length > 0) {
            docExpHdr.documentElement.appendChild(tcList[0]);
            expZip.file('Contents/header.xml', serializer.serializeToString(docExpHdr));
            modified = true;
          }
        }
      }
    }

    // 3. Contents/section*.xml 문단 단위 linesegarray 선별 제거 및 미수정 변경추적 복원
    for (const [fname, expFile] of Object.entries(expZip.files)) {
      if (fname.startsWith('Contents/section') && fname.endsWith('.xml')) {
        const origFile = origZip.file(fname);
        if (!origFile) continue;

        const expXml = await expFile.async('string');
        const origXml = await origFile.async('string');

        if (expXml === origXml) continue;

        const expDoc = parser.parseFromString(expXml, 'application/xml');
        const origDoc = parser.parseFromString(origXml, 'application/xml');

        // Live NodeList 변조 방지를 위해 정적 배열로 변환
        const expParas = Array.from(findElements(expDoc, 'hp:p'));
        const origParas = Array.from(findElements(origDoc, 'hp:p'));

        let sectionChanged = false;

        const count = Math.min(expParas.length, origParas.length);
        for (let idx = 0; idx < count; idx++) {
          const expP = expParas[idx];
          const origP = origParas[idx];

          const expText = getDirectParaText(expP);
          const origText = getDirectParaText(origP);

          if (expText !== origText) {
            // 직계 linesegarray만 선별 제거
            if (removeDirectLineSegArray(expP)) {
              sectionChanged = true;
            }
          } else {
            // 직계 텍스트가 수정되지 않은 문단:
            // 원본 문단에 변경 추적 마크업(deleteBegin/insertBegin)이 있는 경우 원본 문단 노드로 복원
            const origPXml = serializer.serializeToString(origP);
            if (origPXml.includes('deleteBegin') || origPXml.includes('insertBegin')) {
              if (expP.parentNode) {
                expP.parentNode.replaceChild(origP, expP);
                sectionChanged = true;
              }
            }
          }
        }

        if (sectionChanged) {
          const newXml = serializer.serializeToString(expDoc);
          expZip.file(fname, newXml);
          modified = true;
        }
      }
    }

    if (modified) {
      return await expZip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    }
    return exportedBytes;
  } catch (err) {
    console.warn('HWPX 후처리 중 오류 (원본 export 유지):', err);
    return exportedBytes;
  }
}
